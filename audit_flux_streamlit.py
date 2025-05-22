import streamlit as st
import requests
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO
import re
from decimal import Decimal, InvalidOperation

"""Audit d'un flux Google Merchant (ou équivalent interne en français)
====================================================================
Ce script Streamlit :
1. Télécharge ou lit un fichier XML.
2. Détecte automatiquement si c'est un vrai flux Google Merchant (<item>)
   ou un flux interne français (<Sheet1>).
3. Extrait les produits, normalise les prix et GTIN.
4. Applique plusieurs règles de validation (prix, disponibilité, etc.).
5. Génère un fichier Excel avec deux onglets :
   • « Validation » : toutes les données + colonnes OK/Erreur.
   • « Recap_Attributs » : pour chaque attribut, nombre/%% de valeurs manquantes.
"""

# ---------------------------------------------------------------------------
# 1) Helper : CSS pour Streamlit
# ---------------------------------------------------------------------------

def add_custom_css():
    """Applique un thème léger dans l'app Streamlit."""
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
# 2) Téléchargement ou lecture du fichier XML
# ---------------------------------------------------------------------------

def fetch_xml(url: str) -> bytes | None:  # pragma: no cover
    """Retourne le contenu brut du XML à partir d'une URL."""
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as exc:
        st.error(f"Erreur lors du téléchargement du flux : {exc}")
        return None

# ---------------------------------------------------------------------------
# 3) Parsing XML
# ---------------------------------------------------------------------------

def parse_xml(content: bytes) -> ET.Element | None:
    """Transforme le contenu en ElementTree root."""
    try:
        return ET.fromstring(content)
    except ET.ParseError as exc:
        st.error(f"Erreur lors du parsing XML : {exc}")
        return None

# ---------------------------------------------------------------------------
# 4) Normalisation prix & GTIN
# ---------------------------------------------------------------------------

_PRICE_RE = re.compile(r"^(\d+[.,]?\d*)(?:\s*([A-Z]{3}))?$")


def normalize_price(raw: str) -> str:
    """Convertit « 44,99EUR » → « 44.99 EUR ». Retourne MISSING si vide."""
    if not raw or raw == "MISSING":
        return "MISSING"

    m = _PRICE_RE.match(raw.strip())
    if not m:
        return raw.strip()  # valeur non conforme, on la laisse telle quelle

    amount, currency = m.groups()
    amount = amount.replace(",", ".")
    try:
        amount = f"{Decimal(amount):.2f}".rstrip("0").rstrip(".")
    except InvalidOperation:
        return raw.strip()

    return f"{amount} {currency or 'EUR'}".strip()


def normalize_gtin(raw: str) -> str:
    """Convertit 8.8061E+12 en 8806100000000 (GTIN‑13)."""
    if not raw or raw == "MISSING":
        return "MISSING"
    try:
        value = int(Decimal(raw))
        return f"{value:013d}"
    except (InvalidOperation, ValueError):
        return raw

# ---------------------------------------------------------------------------
# 5) Analyse des produits : détection automatique Merchant vs Français
# ---------------------------------------------------------------------------

_NAMESPACE = {"g": "http://base.google.com/ns/1.0"}


def analyze_products(root: ET.Element) -> list[dict]:
    """Retourne une liste de dictionnaires produits."""
    merchant_items = root.findall(".//item", _NAMESPACE)
    if merchant_items:
        return [_parse_google_item(it) for it in merchant_items]

    sheet_items = root.findall(".//Sheet1")
    return [_parse_french_item(it) for it in sheet_items]


# ---------------------------------------------------------------------------
# 5a) Parsing Google Merchant standard
# ---------------------------------------------------------------------------

def _parse_google_item(item: ET.Element) -> dict:
    g = lambda tag: (item.findtext(f"g:{tag}", namespaces=_NAMESPACE) or "").strip() or "MISSING"

    def g_or_plain(tag: str) -> str:
        txt = item.findtext(f"g:{tag}", namespaces=_NAMESPACE) or item.findtext(tag)
        return (txt or "").strip() or "MISSING"

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

# ---------------------------------------------------------------------------
# 5b) Parsing flux interne français
# ---------------------------------------------------------------------------

def _parse_french_item(item: ET.Element) -> dict:
    def txt(tag: str, *alternatives: str) -> str:
        for t in (tag, *alternatives):
            value = (item.findtext(t) or "").strip()
            if value:
                return value
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
# 6) Validation des produits
# ---------------------------------------------------------------------------

_PRICE_VALID_RE = re.compile(r"^\d+(?:\.\d{1,2})?\s?[A-Z]{3}$")


def validate_products(products: list[dict]) -> list[dict]:
    """Ajoute les colonnes de validation à chaque produit."""
    seen_ids: set[str] = set()
    validated: list[dict] = []

    for prod in products:
        pid = prod.get("id", "MISSING")
        price = prod.get("price", "MISSING")
        desc = prod.get("description", "")

        errors = {
            "duplicate_id": "Erreur" if pid in seen_ids else "OK",
            "invalid_or_missing_price": (
                "Erreur" if price == "MISSING" or not _PRICE_VALID_RE.match(price) else "OK"
            ),
            "null_price": "Erreur" if price.startswith("0") else "OK",
            "missing_title": "Erreur" if prod.get("title") == "MISSING" else "OK",
            "description_missing_or_short": "Erreur" if len(desc) < 20 else "OK",
            "invalid_availability": "Erreur" if prod.get("availability") == "MISSING" else "OK",
            "missing_or_empty_color": "Erreur" if prod.get("color") == "MISSING" else "OK",
            "missing_or_empty_gender": "Erreur" if prod.get("gender") == "MISSING" else "OK",
            "missing_or_empty_size": "Erreur" if prod.get("size") == "MISSING" else "OK",
            "missing_or_empty_age_group": "Erreur" if prod.get("age_group") == "MISSING" else "OK",
            "missing_or_empty_image_link": "Erreur" if prod.get("image_link") == "MISSING" else "OK",
        }
        validated.append({**prod, **errors})
        seen_ids.add(pid)

    return validated

# ---------------------------------------------------------------------------
# 7) Export Excel avec récapitulatif
# ---------------------------------------------------------------------------

_PRODUCT_ATTRS = [
    "id", "title", "link", "image_link", "price", "description", "availability",
    "condition", "brand", "gtin", "mpn", "color", "size", "age_group", "gender",
    "item_group_id", "shipping", "shipping_weight", "pattern", "material",
    "additional_image_link", "size_type", "size_system", "canonical_link",
    "expiration_date", "sale_price", "sale_price_effective_date", "product_highlight",
    "ships_from_country", "max_handling_time", "availability_date",
]

_VALIDATION_ATTRS = [
    "duplicate_id", "invalid_or_missing_price", "null_price", "missing_title",
    "description_missing_or_short", "invalid_availability", "missing_or_empty_color",
    "missing_or_empty_gender", "missing_or_empty_size", "missing_or_empty_age_group",
    "missing_or_empty_image_link",
]

_HEADERS = _PRODUCT_ATTRS + _VALIDATION_ATTRS


def generate_excel(data: list[dict]) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Validation"

    # --- Feuille Validation -------------------------------------------------
    ws.append(_HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for prod in data:
        ws.append([prod.get(col, "") for col in _HEADERS])

    # --- Feuille Recap_Attributs -------------------------------------------
    recap = wb.create_sheet("Recap_Attributs")
    recap.append(["Attribut", "Présents", "Manquants", "% manquant"])
    for cell in recap[1]:
        cell.font = Font(bold=True)

    total = len(data) or 1  # éviter division par zéro
    for attr in _PRODUCT_ATTRS:
        missing = sum(1 for p in data if p.get(attr) in ("", "MISSING"))
        present = total - missing
        pct_missing = missing / total * 100
        recap.append([attr, present, missing, f"{pct_missing:.1f}"])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------------------------------------------------------------------------
# 8) Interface Streamlit
# ---------------------------------------------------------------------------

def main():  # pragma: no cover
    add_custom_css()
    st.markdown("<h1 class='main-title'>Audit Flux Google Merchant</h1>", unsafe_allow_html=True)

    url = st.text_input("Entrez l'URL du flux XML :")
    uploaded_file = st.file_uploader("… ou téléchargez un fichier XML :", type=["xml"])

    if st.button("Auditer le flux"):
        if url:
            content = fetch_xml(url)
        elif uploaded_file is not None:
            content = uploaded_file.read()
        else:
            st.warning("Veuillez fournir une URL ou un fichier XML.")
            st.stop()

        if not content:
            st.stop()

        root = parse_xml(content)
        if not root:
            st.stop()

        products = analyze_products(root)
        validated = validate_products(products)
        xlsx = generate_excel(validated)

        st.success(f"Audit terminé ! {len(products)} produit(s) analysé(s).")
        st.download_button(
            label="Télécharger le rapport Excel",
            data=xlsx,
            file_name="audit_flux_google_merchant.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
