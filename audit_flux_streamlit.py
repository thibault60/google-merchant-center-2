import streamlit as st
import requests
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO
import re
from decimal import Decimal, InvalidOperation

"""
Audit d'un flux Google Merchant (ou équivalent interne en français)
==================================================================
1. Télécharge ou lit un fichier XML.
2. Détecte automatiquement si c'est un flux Merchant (<item>) ou interne FR.
3. Extrait les produits, normalise prix & GTIN.
4. Valide les attributs clés (prix, disponibilité…) + 10 attributs « certification/dimensions ».
5. Génère un Excel : feuille « Validation » + feuille « Recap_Attributs ».
"""

# ---------------------------------------------------------------------------
# 1)  CSS minimal
# ---------------------------------------------------------------------------

def add_custom_css():
    st.markdown(
        """
        <style>
        body {background-color:#f8f9fa;font-family:Arial, sans-serif;}
        .main-title {color:#343a40;text-align:center;font-size:2.4rem;margin-bottom:1rem;}
        </style>
        """,
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
# 2)  Téléchargement du XML
# ---------------------------------------------------------------------------

def fetch_xml(url: str) -> bytes | None:
    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        return r.content
    except requests.exceptions.RequestException as exc:
        st.error(f"Erreur téléchargement : {exc}")
        return None

# ---------------------------------------------------------------------------
# 3)  Parsing XML
# ---------------------------------------------------------------------------

def parse_xml(content: bytes) -> ET.Element | None:
    try:
        return ET.fromstring(content)
    except ET.ParseError as exc:
        st.error(f"Erreur de parsing XML : {exc}")
        return None

# ---------------------------------------------------------------------------
# 4)  Normalisations
# ---------------------------------------------------------------------------

_PRICE_RE = re.compile(r"^(\d+[.,]?\d*)(?:\s*([A-Z]{3}))?$")
_DIMENSION_RE = re.compile(r"^\d+(?:[.,]\d+)?\s?(?:mm|cm|in|kg|g)?$", re.I)

def normalize_price(raw: str) -> str:
    if not raw or raw == "MISSING":
        return "MISSING"
    m = _PRICE_RE.match(raw.strip())
    if not m:
        return raw.strip()
    amount, currency = m.groups()
    amount = amount.replace(",", ".")
    try:
        amount = f"{Decimal(amount):.2f}".rstrip("0").rstrip(".")
    except InvalidOperation:
        return raw.strip()
    return f"{amount} {currency or 'EUR'}".strip()

def normalize_gtin(raw: str) -> str:
    if not raw or raw == "MISSING":
        return "MISSING"
    try:
        val = int(Decimal(raw))
        return f"{val:013d}"
    except (InvalidOperation, ValueError):
        return raw

# ---------------------------------------------------------------------------
# 5)  Extraction produits
# ---------------------------------------------------------------------------

_NAMESPACE = {"g": "http://base.google.com/ns/1.0"}

def analyze_products(root: ET.Element) -> list[dict]:
    merchant_items = root.findall(".//item", _NAMESPACE)
    if merchant_items:
        return [_parse_google_item(it) for it in merchant_items]
    sheet_items = root.findall(".//Sheet1")
    return [_parse_french_item(it) for it in sheet_items]

# 5a) Google Merchant
def _parse_google_item(item: ET.Element) -> dict:
    g = lambda tag: (item.findtext(f"g:{tag}", namespaces=_NAMESPACE) or "").strip() or "MISSING"
    g_or_plain = lambda tag: ((item.findtext(f"g:{tag}", _NAMESPACE) or item.findtext(tag) or "").strip() or "MISSING")

    shipping_elem = item.find("g:shipping", _NAMESPACE)
    shipping_block = "".join(shipping_elem.itertext()).strip() if shipping_elem is not None else "MISSING"

    return {
        "id": g("id"),
        "title": g_or_plain("title"),
        "description": g_or_plain("description"),
        "link": g_or_plain("link"),
        "image_link": g("image_link"),
        "price": normalize_price(g("price")),
        "sale_price": normalize_price(g("sale_price")),
        "availability": g("availability"),
        "color": g("color"),
        "gender": g("gender"),
        "size": g("size"),
        "age_group": g("age_group"),
        "condition": g("condition"),
        "brand": g("brand"),
        "gtin": normalize_gtin(g("gtin")),
        "mpn": g("mpn"),
        "item_group_id": g("item_group_id"),
        "shipping": shipping_block,
        "shipping_weight": g("shipping_weight"),
        # --- Nouveaux attributs ------------------------------------------------
        "certification_authority": g("certification_authority"),
        "certification_name": g("certification_name"),
        "certification_code": g("certification_code"),
        "product_length": g("product_length"),
        "product_width": g("product_width"),
        "product_height": g("product_height"),
        "product_weight": g("product_weight"),
        "shipping_length": g("shipping_length"),
        "shipping_width": g("shipping_width"),
        "shipping_height": g("shipping_height"),
        # ----------------------------------------------------------------------
        "pattern": g("pattern"),
        "material": g("material"),
        "additional_image_link": g("additional_image_link"),
        "size_type": g("size_type"),
        "size_system": g("size_system"),
        "canonical_link": g("canonical_link"),
        "expiration_date": g("expiration_date"),
        "sale_price_effective_date": g("sale_price_effective_date"),
        "product_highlight": g("product_highlight"),
        "ships_from_country": g("ships_from_country"),
        "max_handling_time": g("max_handling_time"),
        "availability_date": g("availability_date"),
    }

# 5b) Flux interne français
def _parse_french_item(item: ET.Element) -> dict:
    def txt(tag: str, *alts) -> str:
        for t in (tag, *alts):
            v = (item.findtext(t) or "").strip()
            if v:
                return v
        return "MISSING"

    return {
        "id": txt("identifiant"),
        "title": txt("titre"),
        "description": txt("description"),
        "link": txt("lien"),
        "image_link": txt("lienimage"),
        "price": normalize_price(txt("prix")),
        "sale_price": normalize_price(txt("prixsoldé", "prixsolde")),
        "availability": txt("disponibilité", "disponibilite"),
        "condition": txt("état", "etat"),
        "brand": txt("marque"),
        "gtin": normalize_gtin(txt("gtin")),
        "mpn": txt("mpn"),
        "color": txt("couleur"),
        "size": txt("taille"),
        "age_group": txt("tranche_age"),
        "gender": txt("genre"),
        "item_group_id": txt("id_groupe_article"),
        "shipping": txt("expedition"),
        "shipping_weight": txt("poids_expedition"),
        # --- Nouveaux attributs ------------------------------------------------
        "certification_authority": txt("autorité_certification"),
        "certification_name": txt("nom_certification"),
        "certification_code": txt("code_certification"),
        "product_length": txt("longueur_produit"),
        "product_width": txt("largeur_produit"),
        "product_height": txt("hauteur_produit"),
        "product_weight": txt("poids_produit"),
        "shipping_length": txt("longueur_colis"),
        "shipping_width": txt("largeur_colis"),
        "shipping_height": txt("hauteur_colis"),
        # ----------------------------------------------------------------------
        "pattern": txt("motif"),
        "material": txt("materiau", "matériau"),
        "additional_image_link": txt("lienimage_suppl"),
        "size_type": txt("type_taille"),
        "size_system": txt("systeme_taille", "système_taille"),
        "canonical_link": txt("lien_canonique"),
        "expiration_date": txt("date_expiration"),
        "sale_price_effective_date": txt("periode_promo"),
        "product_highlight": txt("mise_en_avant"),
        "ships_from_country": txt("pays_expedition"),
        "max_handling_time": txt("delai_max_traitement"),
        "availability_date": txt("date_disponibilite"),
    }

# ---------------------------------------------------------------------------
# 6)  Validation
# ---------------------------------------------------------------------------

_PRICE_VALID_RE = re.compile(r"^\d+(?:\.\d{1,2})?\s?[A-Z]{3}$")

def validate_products(products: list[dict]) -> list[dict]:
    seen_ids: set[str] = set()
    validated: list[dict] = []

    for prod in products:
        pid = prod.get("id", "MISSING")
        price = prod.get("price", "MISSING")
        desc  = prod.get("description", "")

        def missing(attr: str) -> bool:
            return prod.get(attr) in ("", "MISSING")

        # validation des nouveaux attributs dimensions/poids/certif
        dims_invalid = any(
            not _DIMENSION_RE.match(prod.get(attr, "")) for attr in (
                "product_length", "product_width", "product_height",
                "shipping_length", "shipping_width", "shipping_height",
                "product_weight",
            ) if not missing(attr)
        )

        errors = {
            "duplicate_id":             "Erreur" if pid in seen_ids else "OK",
            "invalid_or_missing_price": "Erreur" if price == "MISSING" or not _PRICE_VALID_RE.match(price) else "OK",
            "null_price":               "Erreur" if price.startswith("0") else "OK",
            "missing_title":            "Erreur" if prod.get("title") == "MISSING" else "OK",
            "description_missing_or_short": "Erreur" if len(desc) < 20 else "OK",
            "invalid_availability":     "Erreur" if prod.get("availability") == "MISSING" else "OK",
            "missing_or_empty_color":   "Erreur" if missing("color") else "OK",
            "missing_or_empty_gender":  "Erreur" if missing("gender") else "OK",
            "missing_or_empty_size":    "Erreur" if missing("size") else "OK",
            "missing_or_empty_age_group": "Erreur" if missing("age_group") else "OK",
            "missing_or_empty_image_link": "Erreur" if missing("image_link") else "OK",
            # --- nouvelles règles --------------------------------------------
            "missing_certification": "Erreur" if any(missing(a) for a in (
                "certification_authority", "certification_name", "certification_code"
            )) else "OK",
            "missing_dimensions_weight": "Erreur" if any(missing(a) for a in (
                "product_length", "product_width", "product_height", "product_weight",
                "shipping_length", "shipping_width", "shipping_height"
            )) else "OK",
            "invalid_dimension_format": "Erreur" if dims_invalid else "OK",
            # -----------------------------------------------------------------
        }
        validated.append({**prod, **errors})
        seen_ids.add(pid)
    return validated

# ---------------------------------------------------------------------------
# 7)  Export Excel
# ---------------------------------------------------------------------------

_PRODUCT_ATTRS = [
    # attributs pré-existants
    "id", "title", "link", "image_link", "price", "description", "availability",
    "condition", "brand", "gtin", "mpn", "color", "size", "age_group", "gender",
    "item_group_id", "shipping", "shipping_weight", "pattern", "material",
    "additional_image_link", "size_type", "size_system", "canonical_link",
    "expiration_date", "sale_price", "sale_price_effective_date", "product_highlight",
    "ships_from_country", "max_handling_time", "availability_date",
    # --- nouveaux attributs --------------------------------------------------
    "certification_authority", "certification_name", "certification_code",
    "product_length", "product_width", "product_height", "product_weight",
    "shipping_length", "shipping_width", "shipping_height",
]

_VALIDATION_ATTRS = [
    "duplicate_id", "invalid_or_missing_price", "null_price", "missing_title",
    "description_missing_or_short", "invalid_availability", "missing_or_empty_color",
    "missing_or_empty_gender", "missing_or_empty_size", "missing_or_empty_age_group",
    "missing_or_empty_image_link", "missing_certification",
    "missing_dimensions_weight", "invalid_dimension_format",
]

_HEADERS = _PRODUCT_ATTRS + _VALIDATION_ATTRS

def generate_excel(data: list[dict]) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Validation"

    ws.append(_HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True)

    for prod in data:
        ws.append([prod.get(col, "") for col in _HEADERS])

    recap = wb.create_sheet("Recap_Attributs")
    recap.append(["Attribut", "Présents", "Manquants", "% manquant"])
    for c in recap[1]:
        c.font = Font(bold=True)

    total = len(data) or 1
    for attr in _PRODUCT_ATTRS:
        missing = sum(1 for p in data if p.get(attr) in ("", "MISSING"))
        present = total - missing
        recap.append([attr, present, missing, f"{missing/total*100:.1f}"])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------------------------------------------------------------------------
# 8)  Interface Streamlit
# ---------------------------------------------------------------------------

def main():
    add_custom_css()
    st.markdown("<h1 class='main-title'>Audit Flux Google Merchant</h1>", unsafe_allow_html=True)

    url = st.text_input("Entrez l'URL du flux XML :")
    uploaded_file = st.file_uploader("… ou téléchargez un fichier XML :", type=["xml"])

    if st.button("Auditer le flux"):
        content = None
        if url:
            content = fetch_xml(url)
        elif uploaded_file is not None:
            content = uploaded_file.read()

        if not content:
            st.warning("Veuillez fournir une URL ou un fichier XML.")
            st.stop()

        root = parse_xml(content)
        if not root:
            st.stop()

        products = analyze_products(root)
        validated = validate_products(products)
        xlsx = generate_excel(validated)

        st.success(f"Audit terminé ! {len(products)} produit(s) analysé(s).")
        st.download_button(
            label="Télécharger le rapport Excel",
            data=xlsx,
            file_name="audit_flux_google_merchant.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if __name__ == "__main__":
    main()
