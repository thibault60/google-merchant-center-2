import streamlit as st
import requests
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO
import re
from decimal import Decimal, InvalidOperation

"""Audit d'un flux Google Merchant ou interne (FR)
==================================================
Fonctions clés
-------------
* Détection automatique de deux structures XML : Google Merchant standard (`<item>` + balises `g:`) ou flux interne (`<Sheet1>` + balises FR).
* Normalisation : prix (`price`, `sale_price`) et GTIN (notation scientifique ➜ chaîne 13 chiffres).
* Validation : règles de complétude + format sur des attributs obligatoires/recommandés.
* Export Excel : deux onglets – `Validation` (données & erreurs) et `Recap_Attributs` (statistiques de valeurs manquantes).
* Couverture élargie : prise en charge des attributs « Electronics & Household Equipment ».
"""

# ---------------------------------------------------------------------------
# 1)   CSS Streamlit
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
# 2)   Téléchargement / lecture XML
# ---------------------------------------------------------------------------

def fetch_xml(url: str) -> bytes | None:
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e:
        st.error(f"Erreur lors du téléchargement : {e}")
        return None

# ---------------------------------------------------------------------------
# 3)   Parsing XML (ElementTree)
# ---------------------------------------------------------------------------

def parse_xml(content: bytes) -> ET.Element | None:
    try:
        return ET.fromstring(content)
    except ET.ParseError as e:
        st.error(f"Erreur lors du parsing XML : {e}")
        return None

# ---------------------------------------------------------------------------
# 4)   Helpers : normalisation prix / GTIN
# ---------------------------------------------------------------------------

_price_re = re.compile(r"^(\d+[.,]?\d*)(?:\s*([A-Z]{3}))?$")  # 44.99 EUR ou 44,99EUR


def normalize_price(raw: str) -> str:
    """Retourne <montant> <devise> ou MISSING."""
    if not raw:
        return "MISSING"
    m = _price_re.match(raw.strip())
    if not m:
        return raw.strip()
    amount, cur = m.groups()
    amount = amount.replace(",", ".")
    try:
        amount = f"{Decimal(amount):.2f}".rstrip("0").rstrip(".")
    except InvalidOperation:
        return raw.strip()
    return f"{amount} {cur or 'EUR'}".strip()


def normalize_gtin(raw: str) -> str:
    if not raw or raw == "MISSING":
        return "MISSING"
    try:
        return f"{int(Decimal(raw)):013d}"
    except (InvalidOperation, ValueError):
        return raw

# ---------------------------------------------------------------------------
# 5)   Analyse des produits (auto‑détection)
# ---------------------------------------------------------------------------

def analyze_products(root: ET.Element) -> list[dict]:
    ns = {"g": "http://base.google.com/ns/1.0"}
    merchant_items = root.findall(".//item", ns)
    if merchant_items:
        return [_parse_google_item(it, ns) for it in merchant_items]
    # sinon, flux français interne
    sheet_items = root.findall(".//Sheet1")
    return [_parse_french_item(it) for it in sheet_items]

# ---------- Parsing Google Merchant ---------------------------------------

def _parse_google_item(item: ET.Element, ns) -> dict:
    def g(tag):
        el = item.find(f"g:{tag}", ns)
        return el.text.strip() if el is not None and el.text else "MISSING"

    def g_or_plain(tag):
        el = item.find(f"g:{tag}", ns) or item.find(tag)
        return el.text.strip() if el is not None and el.text else "MISSING"

    def shipping_block():
        sp = item.find("g:shipping", ns)
        return "".join(sp.itertext()).strip() if sp is not None else "MISSING"

    return {
        # Core
        "id": g("id"),
        "title": g_or_plain("title"),
        "description": g_or_plain("description"),
        "link": g_or_plain("link"),
        "image_link": g("image_link"),
        "price": normalize_price(g("price")),
        "sale_price": normalize_price(g("sale_price")),
        "availability": g("availability"),
        "condition": g("condition"),
        "brand": g("brand"),
        "gtin": normalize_gtin(g("gtin")),
        "mpn": g("mpn"),
        "item_group_id": g("item_group_id"),
        "shipping": shipping_block(),
        "shipping_weight": g("shipping_weight"),
        # Apparel / autres
        "color": g("color"),
        "gender": g("gender"),
        "size": g("size"),
        "age_group": g("age_group"),
        "pattern": g("pattern"),
        "material": g("material"),
        "additional_image_link": g("additional_image_link"),
        "size_type": g("size_type"),
        "size_system": g("size_system"),
        # Divers
        "canonical_link": g("canonical_link"),
        "expiration_date": g("expiration_date"),
        "sale_price_effective_date": g("sale_price_effective_date"),
        "product_highlight": g("product_highlight"),
        "ships_from_country": g("ships_from_country"),
        "max_handling_time": g("max_handling_time"),
        "availability_date": g("availability_date"),
        # Electronics & Household Equipment
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
    }

# ---------- Parsing flux interne (FR) -------------------------------------

def _parse_french_item(item: ET.Element) -> dict:
    def txt(tag, *alts):
        for t in (tag, *alts):
            val = (item.findtext(t) or "").strip()
            if val:
                return val
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
        "item_group_id": txt("id_groupe_article"),
        "shipping": txt("expedition"),
        "shipping_weight": txt("poids_expedition"),
        "color": txt("couleur"),
        "gender": txt("genre"),
        "size": txt("taille"),
        "age_group": txt("tranche_age"),
        "pattern": txt("motif"),
        "material": txt("materiau"),
        "additional_image_link": txt("lienimage_suppl"),
        "size_type": txt("type_taille"),
        "size_system": txt("systeme_taille"),
        "canonical_link": txt("lien_canonique"),
        "expiration_date": txt("date_expiration"),
        "sale_price_effective_date": txt("periode_promo"),
        "product_highlight": txt("mise_en_avant"),
        "ships_from_country": txt("pays_expedition"),
        "max_handling_time": txt("delai_max_traitement"),
        "availability_date": txt("date_disponibilite"),
        # Electronics & Household Equipment
        "certification_authority": txt("certification_authority", "autorite_certification"),
        "certification_name": txt("certification_name", "nom_certification"),
        "certification_code": txt("certification_code", "code_certification"),
        "product_length": txt("product_length", "longueur_produit"),
        "product_width": txt("product_width", "largeur_produit"),
        "product_height": txt("product_height", "hauteur_produit"),
        "product_weight": txt("product_weight", "poids_produit"),
        "shipping_length": txt("shipping_length", "longueur_expedition"),
        "shipping_width": txt("shipping_width", "largeur_expedition"),
        "shipping_height": txt("shipping_height", "hauteur_expedition"),
    }

# ---------------------------------------------------------------------------
# 6)   Validation des produits
# ---------------------------------------------------------------------------

def validate_products(products: list[dict]) -> list[dict]:
    price_re = re.compile(r"^\d+(?:\.\d{1,2})?\s?[A-Z]{3}$")
    seen = set()
    validated = []
    for p in products:
        pid = p.get("id", "MISSING")
        price = p.get("price", "MISSING")
        desc = p.get("description", "")
        errs = {
            "duplicate
