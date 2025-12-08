import streamlit as st
import requests
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO
import re
from decimal import Decimal, InvalidOperation
from collections import defaultdict

"""
Audit d'un flux Google Merchant – version TXT pipe '|' ➜ EN
===========================================================
• Télécharge ou lit un TXT (lignes pipe-séparées).
• Extrait les produits et alimente les attributs EN attendus.
• Normalise prix, GTIN, dimensions.
• Valide attributs clés + certification/dimensions.
• Exporte un rapport Excel, avec récap par statut (Mandatory/Recommended…).
"""

# ---------------------------------------------------------------------------
# 1)  Helpers visuels
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
# 2)  Téléchargement / lecture TXT
# ---------------------------------------------------------------------------

def fetch_txt(url: str) -> bytes | None:
    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        return r.content
    except requests.exceptions.RequestException as exc:
        st.error(f"Erreur téléchargement : {exc}")
        return None

# ---------------------------------------------------------------------------
# 3)  Normalisations
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
# 4)  Extraction produits depuis un TXT pipe '|'
# ---------------------------------------------------------------------------

def parse_txt(content: bytes) -> list[dict]:
    text = content.decode("utf-8", errors="replace")
    lines = [ln.strip() for ln in text.splitlines() if ln.strip() and not ln.strip().startswith("#")]
    products: list[dict] = []

    for ln in lines:
        fields = ln.split("|")

        # Initialise toutes les colonnes exportées à MISSING
        prod = {key: "MISSING" for key in _PRODUCT_ATTRS}

        # Mapping minimal basé sur l'exemple fourni
        def get(i: int) -> str:
            return (fields[i].strip() if i < len(fields) else "")

        # Positions d'après l'exemple:
        # 0:id | 1:title | 2:link | 3:price | 4:description | 5:condition
        # 6:gtin | 7:color | 8:mpn | 9:image_link
        # 11:availability | 12:shipping | 13:shipping_weight | 14:google_product_category
        # 16:item_group_id | 17:gender | 18:age_group | 19:pattern | 20:size
        prod["id"]          = get(0) or "MISSING"
        prod["title"]       = get(1) or "MISSING"
        prod["link"]        = get(2) or "MISSING"
        prod["price"]       = normalize_price(get(3) or "MISSING")
        prod["description"] = get(4) or "MISSING"
        prod["condition"]   = get(5) or "MISSING"
        prod["gtin"]        = normalize_gtin(get(6) or "MISSING")
        prod["color"]       = get(7) or "MISSING"
        prod["mpn"]         = get(8) or "MISSING"
        prod["image_link"]  = get(9) or "MISSING"
        prod["availability"]= get(11) or "MISSING"
        prod["shipping"]    = get(12) or "MISSING"
        prod["shipping_weight"] = get(13) or "MISSING"
        prod["google_product_category"] = get(14) or "MISSING"
        prod["item_group_id"] = get(16) or "MISSING"
        prod["gender"]      = get(17) or "MISSING"
        prod["age_group"]   = get(18) or "MISSING"
        prod["pattern"]     = get(19) or "MISSING"
        prod["size"]        = get(20) or "MISSING"

        # Brand : si un champ dédié n'existe pas dans la ligne d'exemple, on tente un fallback léger
        # (ex.: valeur non vide après size) sinon laisse MISSING pour que la validation le signale.
        candidate_brand = get(21)
        if candidate_brand and candidate_brand.lower() not in ("0", "na", "null"):
            prod["brand"] = candidate_brand

        # additional_image_link : on cherche un champ contenant des URLs séparées par des virgules
        if prod["additional_image_link"] == "MISSING":
            for f in fields[22:]:
                s = f.strip()
                if "http" in s and "," in s:
                    prod["additional_image_link"] = s
                    break

        # ships_from_country : on prend le dernier token non vide de 2–3 lettres majuscules
        for f in reversed(fields):
            s = f.strip()
            if len(s) in (2, 3) and s.isupper():
                prod["ships_from_country"] = s
                break

        # Normalisations finales (sécurité)
        prod["price"] = normalize_price(prod.get("price", "MISSING"))
        prod["gtin"]  = normalize_gtin(prod.get("gtin", "MISSING"))

        products.append(prod)

    return products

# ---------------------------------------------------------------------------
# 5)  Validation
# ---------------------------------------------------------------------------

_PRICE_VALID_RE = re.compile(r"^\d+(?:\.\d{1,2})?\s?[A-Z]{3}$")

def validate_products(products: list[dict]) -> list[dict]:
    seen_ids: set[str] = set()
    validated: list[dict] = []

    for prod in products:
        pid = prod["id"]
        price = prod["price"]
        desc  = prod["description"]

        def missing(attr: str) -> bool:
            return prod.get(attr) in ("", "MISSING")

        dims_invalid = any(
            not _DIMENSION_RE.match(prod[attr]) for attr in (
                "product_length", "product_width", "product_height",
                "shipping_length", "shipping_width", "shipping_height",
                "product_weight",
            ) if not missing(attr)
        )

        errors = {
            "duplicate_id":                 "Erreur" if pid in seen_ids else "OK",
            "invalid_or_missing_price":     "Erreur" if price == "MISSING" or not _PRICE_VALID_RE.match(price) else "OK",
            "null_price":                   "Erreur" if price.startswith("0") else "OK",
            "missing_title":                "Erreur" if prod["title"] == "MISSING" else "OK",
            "description_missing_or_short": "Erreur" if len(desc) < 20 else "OK",
            "invalid_availability":         "Erreur" if prod["availability"] == "MISSING" else "OK",
            "missing_or_empty_color":       "Erreur" if missing("color") else "OK",
            "missing_or_empty_gender":      "Erreur" if missing("gender") else "OK",
            "missing_or_empty_size":        "Erreur" if missing("size") else "OK",
            "missing_or_empty_age_group":   "Erreur" if missing("age_group") else "OK",
            "missing_or_empty_image_link":  "Erreur" if missing("image_link") else "OK",
            # certification / dimensions
            "missing_certification":        "Erreur" if any(missing(a) for a in (
                                                "certification_authority", "certification_name", "certification_code"
                                              )) else "OK",
            "missing_dimensions_weight":    "Erreur" if any(missing(a) for a in (
                                                "product_length", "product_width", "product_height", "product_weight",
                                                "shipping_length", "shipping_width", "shipping_height"
                                              )) else "OK",
            "invalid_dimension_format":     "Erreur" if dims_invalid else "OK",
            # ✅ Nouveaux contrôles demandés
            "missing_google_product_category": "Erreur" if missing("google_product_category") else "OK",
            "missing_minimum_handling_time":   "Erreur" if missing("minimum_handling_time") else "OK",
        }

        validated.append({**prod, **errors})
        seen_ids.add(pid)

    return validated

# ---------------------------------------------------------------------------
# 6)  Export Excel – Statuts & Récap
# ---------------------------------------------------------------------------

# Liste des attributs exportés (ordre colonnes)
_PRODUCT_ATTRS = [
    "id", "title", "link", "image_link", "price", "sale_price",
    "description", "availability", "condition", "brand",
    "gtin", "mpn", "color", "size", "age_group", "gender",
    "item_group_id", "google_product_category",
    # certification & dimensions
    "certification_authority", "certification_name", "certification_code",
    "product_length", "product_width", "product_height", "product_weight",
    "shipping_length", "shipping_width", "shipping_height",
    # autres
    "shipping", "shipping_weight", "pattern", "material",
    "additional_image_link", "size_type", "size_system", "canonical_link",
    "expiration_date", "sale_price_effective_date", "product_highlight",
    "ships_from_country", "minimum_handling_time", "max_handling_time",
    "availability_date", "product_detail",
]

# Colonnes de validation
_VALIDATION_ATTRS = [
    "duplicate_id", "invalid_or_missing_price", "null_price", "missing_title",
    "description_missing_or_short", "invalid_availability", "missing_or_empty_color",
    "missing_or_empty_gender", "missing_or_empty_size", "missing_or_empty_age_group",
    "missing_or_empty_image_link", "missing_certification",
    "missing_dimensions_weight", "invalid_dimension_format",
    # ✅ nouveaux checks
    "missing_google_product_category", "missing_minimum_handling_time",
]

_HEADERS = _PRODUCT_ATTRS + _VALIDATION_ATTRS

# Table des statuts (inchangée)
FIELD_STATUS = {
    "id": "Mandatory",
    "title": "Mandatory",
    "link": "Mandatory",
    "image_link": "Mandatory",
    "price": "Mandatory",
    "description": "Mandatory",
    "availability": "Mandatory",
    "condition": "Mandatory IF SECOND-HAND or RECONDITIONED",
    "brand": "Mandatory",
    "gtin": "Mandatory if GTIN existing",
    "mpn": "Mandatory if GTIN already exists",
    "color": "Mandatory if CLOTHING AND ACCESSORIES",
    "size": "Mandatory if CLOTHING AND ACCESSORIES",
    "age_group": "Mandatory if CLOTHING AND ACCESSORIES",
    "gender": "Mandatory if CLOTHING AND ACCESSORIES",
    "item_group_id": "Mandatory FOR VARIANTS",
    "shipping": "Mandatory if not configured in GMC",
    "shipping_weight": "Recommended",
    "pattern": "Recommended if CLOTHING AND ACCESSORIES",
    "material": "Recommended if CLOTHING AND ACCESSORIES",
    "additional_image_link": "Recommended",
    "size_type": "Recommended if CLOTHING AND ACCESSORIES",
    "size_system": "Recommended if CLOTHING AND ACCESSORIES",
    "canonical_link": "Recommended",
    "expiration_date": "Recommended to stop displaying a product",
    "sale_price": "Recommended for promotions",
    "sale_price_effective_date": "Recommended for promotions",
    "product_highlight": "Recommended",
    "ships_from_country": "Recommended",
    "minimum_handling_time": "Recommended",
    "max_handling_time": "Recommended",
    "availability_date": "Mandatory if product is pre-ordered",
    "certification_authority": "Recommended for ELECTRONICS & HOUSEHOLD EQUIPEMENT",
    "certification_name": "Recommended for ELECTRONICS & HOUSEHOLD EQUIPEMENT",
    "certification_code": "Recommended for ELECTRONICS & HOUSEHOLD EQUIPEMENT",
    "product_detail": "Recommended",
    "product_length": "Recommended for ELECTRONICS & HOUSEHOLD EQUIPEMENT",
    "product_width": "Recommended for ELECTRONICS & HOUSEHOLD EQUIPEMENT",
    "product_height": "Recommended for ELECTRONICS & HOUSEHOLD EQUIPEMENT",
    "product_weight": "Recommended for ELECTRONICS & HOUSEHOLD EQUIPEMENT",
    "shipping_length": "Recommended for ELECTRONICS & HOUSEHOLD EQUIPEMENT",
    "shipping_width": "Recommended for ELECTRONICS & HOUSEHOLD EQUIPEMENT",
    "shipping_height": "Recommended for ELECTRONICS & HOUSEHOLD EQUIPEMENT",
    "google_product_category": "Mandatory",
}

def generate_excel(data: list[dict]) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Validation"

    # Feuille 1 : données + flags
    ws.append(_HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True)
    for prod in data:
        ws.append([prod.get(col, "") for col in _HEADERS])

    # Feuille 2 : récap par attribut
    recap = wb.create_sheet("Recap_Attributs")
    recap.append(["Attribut", "Statut", "Présents", "Manquants", "% manquant"])
    for c in recap[1]:
        c.font = Font(bold=True)

    total = len(data) or 1
    by_status_counts = defaultdict(list)

    for attr in _PRODUCT_ATTRS:
        missing = sum(1 for p in data if p.get(attr) in ("", "MISSING"))
        status = FIELD_STATUS.get(attr, "")
        missing_pct = (missing / total) * 100
        recap.append([attr, status, total - missing, missing, f"{missing_pct:.1f}"])
        if status:
            by_status_counts[status].append(100 - missing_pct)

    # Feuille 3 : synthèse par statut
    synth = wb.create_sheet("Synthese_par_statut")
    synth.append(["Statut", "Nb attributs", "Taux de complétion moyen (%)", "Attributs"])
    for c in synth[1]:
        c.font = Font(bold=True)

    for status, completion_list in by_status_counts.items():
        attrs = [a for a, s in FIELD_STATUS.items() if s == status and a in _PRODUCT_ATTRS]
        avg_completion = sum(completion_list) / len(completion_list) if completion_list else 0.0
        synth.append([status, len(attrs), f"{avg_completion:.1f}", ", ".join(attrs)])

    # Feuille 4 : règles (tableau brut des statuts)
    rules = wb.create_sheet("Regles_Attributs")
    rules.append(["Field Name", "Status"])
    for c in rules[1]:
        c.font = Font(bold=True)
    for field, status in FIELD_STATUS.items():
        rules.append([field, status])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------------------------------------------------------------------------
# 7)  Interface Streamlit
# ---------------------------------------------------------------------------

def main():
    add_custom_css()
    st.markdown("<h1 class='main-title'>Audit Flux Google Merchant (TXT)</h1>", unsafe_allow_html=True)

    url = st.text_input("Entrez l'URL du fichier TXT :")
    uploaded_file = st.file_uploader("… ou téléchargez un fichier TXT :", type=["txt"])

    if st.button("Auditer le flux"):
        content = None
        if url:
            content = fetch_txt(url)
        elif uploaded_file is not None:
            content = uploaded_file.read()

        if not content:
            st.warning("Veuillez fournir une URL ou un fichier TXT.")
            st.stop()

        products = parse_txt(content)
        validated = validate_products(products)
        xlsx = generate_excel(validated)

        st.success(f"Audit terminé : {len(products)} produit(s) analysé(s).")
        st.download_button(
            "Télécharger le rapport Excel",
            data=xlsx,
            file_name="audit_flux_google_merchant.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if __name__ == "__main__":
    main()
