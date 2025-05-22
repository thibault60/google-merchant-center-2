import streamlit as st
import requests
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO
import re
from decimal import Decimal, InvalidOperation

# ---------------------------------------------------------------------------
# 1)   CSS Streamlit : look & feel
# ---------------------------------------------------------------------------
def add_custom_css():
    st.markdown(
        """
        <style>
        body {background-color:#f8f9fa;font-family:'Arial',sans-serif;}
        .main-title {color:#343a40;text-align:center;font-size:2.5rem;margin-bottom:1rem;}
        </style>
        """,
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
# 2)   Téléchargement du flux (URL ou upload)
# ---------------------------------------------------------------------------
def fetch_xml(url: str) -> bytes | None:
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e:
        st.error(f"Erreur lors du téléchargement du flux : {e}")
        return None

# ---------------------------------------------------------------------------
# 3)   Parsing XML → ElementTree root
# ---------------------------------------------------------------------------
def parse_xml(content: bytes) -> ET.Element | None:
    try:
        return ET.fromstring(content)
    except ET.ParseError as e:
        st.error(f"Erreur lors du parsing XML : {e}")
        return None

# ---------------------------------------------------------------------------
# 4)   Helpers : normalisation prix / GTIN
# ---------------------------------------------------------------------------
_price_re = re.compile(r"^(\d+[.,]?\d*)(?:\s*([A-Z]{3}))?$")  # ex : 44.99 EUR, 44,99EUR, 44 EUR

def normalize_price(raw: str) -> str:
    """
    Nettoie un prix en entrée :
    - remplace la virgule par un point pour garder la convention US
    - isole la devise si présente
    - retourne 'MISSING' si vide
    """
    if not raw:
        return "MISSING"
    
    match = _price_re.match(raw.strip())
    if not match:
        return raw.strip()  # laisser l’utilisateur voir la valeur non conforme
    
    amount, currency = match.groups()
    amount = amount.replace(",", ".")
    try:
        # arrondi à 2 décimales maxi
        amount = f"{Decimal(amount):.2f}".rstrip("0").rstrip(".")
    except InvalidOperation:
        return raw.strip()
    
    return f"{amount} {currency or 'EUR'}".strip()

def normalize_gtin(raw: str) -> str:
    """
    Convertit la notation scientifique éventuelle '8.80609E+12' → '8806090000000'.
    Préserve la valeur d'origine si la conversion échoue.
    """
    if not raw or raw == "MISSING":
        return "MISSING"
    try:
        as_int = int(Decimal(raw))
        return f"{as_int:013d}"  # GTIN-13
    except (InvalidOperation, ValueError):
        return raw

# ---------------------------------------------------------------------------
# 5)   Analyse des produits : double stratégie (Merchant / Français)
# ---------------------------------------------------------------------------
def analyze_products(root: ET.Element) -> list[dict]:
    namespace = {"g": "http://base.google.com/ns/1.0"}
    # --- Cas 1 : flux Google Merchant standard -----------------------------
    merchant_items = root.findall(".//item", namespace)
    if merchant_items:
        return [_parse_google_item(it, namespace) for it in merchant_items]

    # --- Cas 2 : flux « français » ------------------------------------------
    sheet_items = root.findall(".//Sheet1")
    return [_parse_french_item(it) for it in sheet_items]

# ---------- Sous-fonctions de parsing --------------------------------------
def _parse_google_item(item: ET.Element, ns) -> dict:
    def g(tag):  # racourci
        elem = item.find(f"g:{tag}", ns)
        return elem.text.strip() if elem is not None and elem.text else "MISSING"

    def g_or_plain(tag):
        # g:title ou title
        elem = item.find(f"g:{tag}", ns) or item.find(tag)
        return elem.text.strip() if elem is not None and elem.text else "MISSING"

    def get_shipping_block():
        shipping = item.find("g:shipping", ns)
        return "".join(shipping.itertext()).strip() if shipping is not None else "MISSING"

    product = {
        "id"                : g("id"),
        "title"             : g_or_plain("title"),
        "description"       : g_or_plain("description"),
        "link"              : g_or_plain("link"),
        "image_link"        : g("image_link"),
        "price"             : normalize_price(g("price")),
        "availability"      : g("availability"),
        "color"             : g("color"),
        "gender"            : g("gender"),
        "size"              : g("size"),
        "age_group"         : g("age_group"),
        "condition"         : g("condition"),
        "brand"             : g("brand"),
        "gtin"              : normalize_gtin(g("gtin")),
        "mpn"               : g("mpn"),
        "item_group_id"     : g("item_group_id"),
        "shipping"          : get_shipping_block(),
        "shipping_weight"   : g("shipping_weight"),
        "pattern"           : g("pattern"),
        "material"          : g("material"),
        "additional_image_link"   : g("additional_image_link"),
        "size_type"         : g("size_type"),
        "size_system"       : g("size_system"),
        "canonical_link"    : g("canonical_link"),
        "expiration_date"   : g("expiration_date"),
        "sale_price"        : normalize_price(g("sale_price")),
        "sale_price_effective_date": g("sale_price_effective_date"),
        "product_highlight" : g("product_highlight"),
        "ships_from_country": g("ships_from_country"),
        "max_handling_time" : g("max_handling_time"),
        "availability_date" : g("availability_date"),
    }
    return product

def _parse_french_item(item: ET.Element) -> dict:
    txt = lambda tag: (item.findtext(tag) or "").strip() or "MISSING"

    product = {
        "id"                : txt("identifiant"),
        "title"             : txt("titre"),
        "description"       : txt("description"),
        "link"              : txt("lien"),
        "image_link"        : txt("lienimage"),
        "price"             : normalize_price(txt("prix")),
        "availability"      : txt("disponibilite"),      # sans accent dans l’XML fourni
        "color"             : txt("couleur"),
        "gender"            : txt("genre"),
        "size"              : txt("taille"),
        "age_group"         : txt("tranche_age"),
        "condition"         : txt("etat"),
        "brand"             : txt("marque"),
        "gtin"              : normalize_gtin(txt("gtin")),
        "mpn"               : txt("mpn"),
        "item_group_id"     : txt("id_groupe_article"),
        "shipping"          : txt("expedition"),
        "shipping_weight"   : txt("poids_expedition"),
        "pattern"           : txt("motif"),
        "material"          : txt("materiau"),
        "additional_image_link"   : txt("lienimage_suppl"),
        "size_type"         : txt("type_taille"),
        "size_system"       : txt("systeme_taille"),
        "canonical_link"    : txt("lien_canonique"),
        "expiration_date"   : txt("date_expiration"),
        "sale_price"        : normalize_price(txt("prix_promo")),
        "sale_price_effective_date": txt("periode_promo"),
        "product_highlight" : txt("mise_en_avant"),
        "ships_from_country": txt("pays_expedition"),
        "max_handling_time" : txt("delai_max_traitement"),
        "availability_date" : txt("date_disponibilite"),
    }
    return product

# ---------------------------------------------------------------------------
# 6)   Validation des produits
# ---------------------------------------------------------------------------
def validate_products(products: list[dict]) -> list[dict]:
    price_regex = re.compile(r"^\d+(?:\.\d{1,2})?\s?[A-Z]{3}$")
    seen_ids: set[str] = set()
    validated = []

    for p in products:
        pid = p.get("id", "MISSING")
        price = p.get("price", "MISSING")
        desc  = p.get("description", "")

        errors = {
            "duplicate_id"              : "Erreur" if pid in seen_ids else "OK",
            "invalid_or_missing_price"   : "Erreur" if price=="MISSING" or not price_regex.match(price) else "OK",
            "null_price"                : "Erreur" if price.startswith("0") else "OK",
            "missing_title"             : "Erreur" if p.get("title")=="MISSING" else "OK",
            "description_missing_or_short": "Erreur" if len(desc)<20 else "OK",
            "invalid_availability"      : "Erreur" if p.get("availability")=="MISSING" else "OK",
            "missing_or_empty_color"    : "Erreur" if p.get("color")=="MISSING" else "OK",
            "missing_or_empty_gender"   : "Erreur" if p.get("gender")=="MISSING" else "OK",
            "missing_or_empty_size"     : "Erreur" if p.get("size")=="MISSING" else "OK",
            "missing_or_empty_age_group": "Erreur" if p.get("age_group")=="MISSING" else "OK",
            "missing_or_empty_image_link": "Erreur" if p.get("image_link")=="MISSING" else "OK",
        }
        validated.append({**p, **errors})
        seen_ids.add(pid)
    return validated

# ---------------------------------------------------------------------------
# 7)   Export Excel
# ---------------------------------------------------------------------------
_HEADERS = [
    # attributs produit
    "id","title","link","image_link","price","description","availability",
    "condition","brand","gtin","mpn","color","size","age_group","gender",
    "item_group_id","shipping","shipping_weight","pattern","material",
    "additional_image_link","size_type","size_system","canonical_link",
    "expiration_date","sale_price","sale_price_effective_date",
    "product_highlight","ships_from_country","max_handling_time",
    "availability_date",
    # validations
    "duplicate_id","invalid_or_missing_price","null_price","missing_title",
    "description_missing_or_short","invalid_availability","missing_or_empty_color",
    "missing_or_empty_gender","missing_or_empty_size","missing_or_empty_age_group",
    "missing_or_empty_image_link",
]

def generate_excel(data: list[dict]) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Validation"

    ws.append(_HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for prod in data:
        ws.append([prod.get(col,"") for col in _HEADERS])

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# ---------------------------------------------------------------------------
# 8)   Interface Streamlit
# ---------------------------------------------------------------------------
def main():
    add_custom_css()
    st.markdown("<h1 class='main-title'>Audit Flux Google Merchant</h1>", unsafe_allow_html=True)

    url = st.text_input("Entrez l'URL du flux XML :")
    uploaded_file = st.file_uploader("… ou téléchargez un fichier XML :", type=["xml"])

    if st.button("Auditer le flux"):
        content = fetch_xml(url) if url else uploaded_file.read() if uploaded_file else None

        if not content:
            st.warning("Aucun contenu à analyser.")
            st.stop()

        root = parse_xml(content)
        if not root:
            st.stop()

        products = analyze_products(root)
        validated_products = validate_products(products)
        excel_data = generate_excel(validated_products)

        st.success(f"Audit terminé ! {len(products)} produit(s) analysé(s).")
        st.download_button(
            label="Télécharger le rapport Excel",
            data=excel_data,
            file_name="audit_flux_google_merchant.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
