import streamlit as st
import requests, re
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO

# ---------- Apparence ----------
def add_custom_css():
    st.markdown("""
        <style>
        body {background:#f8f9fa;font-family:Arial,sans-serif;}
        .main-title {color:#343a40;text-align:center;font-size:2.5rem;margin-bottom:1rem;}
        </style>
        """, unsafe_allow_html=True)

# ---------- 1. Téléchargement ----------
def fetch_xml(url: str) -> bytes | None:
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        return r.content
    except requests.exceptions.RequestException as e:
        st.error(f"Erreur téléchargement : {e}")
        return None

# ---------- 2. Parsing ----------
def parse_xml(data: bytes) -> ET.Element | None:
    try:
        return ET.fromstring(data)
    except ET.ParseError as e:
        st.error(f"Erreur parsing XML : {e}")
        return None

# ---------- 3. Extraction ----------
def analyze_products(root: ET.Element) -> list[dict]:
    ns = {"g": "http://base.google.com/ns/1.0"}

    def gtext(it, tag):
        el = it.find(f"g:{tag}", ns)
        return el.text.strip() if el is not None and el.text else "MISSING"

    def get_link(it):
        for tag in ("g:link", "link"):
            el = it.find(tag, ns) if ":" in tag else it.find(tag)
            if el is not None and el.text:
                return el.text.strip()
        return "MISSING"

    def gettext(it, gtag, tag):
        for t in (f"g:{gtag}", tag):
            el = it.find(t, ns) if ":" in t else it.find(t)
            if el is not None and el.text:
                return el.text.strip()
        return "MISSING"

    def get_shipping(it):
        s = it.find("g:shipping", ns)
        return "".join(s.itertext()).strip() if s is not None else "MISSING"

    products = []
    for it in root.findall(".//item", ns):
        products.append({
            "id": gtext(it, "id"),
            "title": gettext(it, "title", "title"),
            "description": gettext(it, "description", "description"),
            "link": get_link(it),                     # ← seule URL utilisée
            "image_link": gtext(it, "image_link"),
            "price": gtext(it, "price"),
            "availability": gtext(it, "availability"),
            "color": gtext(it, "color"),
            "gender": gtext(it, "gender"),
            "size": gtext(it, "size"),
            "age_group": gtext(it, "age_group"),
            # Champs additionnels
            "condition": gtext(it, "condition"),
            "brand": gtext(it, "brand"),
            "gtin": gtext(it, "gtin"),
            "mpn": gtext(it, "mpn"),
            "item_group_id": gtext(it, "item_group_id"),
            "shipping": get_shipping(it),
            "shipping_weight": gtext(it, "shipping_weight"),
            "pattern": gtext(it, "pattern"),
            "material": gtext(it, "material"),
            "additional_image_link": gtext(it, "additional_image_link"),
            "size_type": gtext(it, "size_type"),
            "size_system": gtext(it, "size_system"),
            "canonical_link": gtext(it, "canonical_link"),
            "expiration_date": gtext(it, "expiration_date"),
            "sale_price": gtext(it, "sale_price"),
            "sale_price_effective_date": gtext(it, "sale_price_effective_date"),
            "product_highlight": gtext(it, "product_highlight"),
            "ships_from_country": gtext(it, "ships_from_country"),
            "max_handling_time": gtext(it, "max_handling_time"),
            "availability_date": gtext(it, "availability_date"),
        })
    return products

# ---------- 4. Validation ----------
def validate_products(products: list[dict]) -> tuple[list[dict], list[dict]]:
    price_re = re.compile(r"^\d+(\.\d{1,2})? [A-Z]{3}$")  # devise obligatoire
    seen, validated = set(), []
    for p in products:
        errs = {
            "duplicate_id": "Erreur" if p["id"] in seen else "OK",
            "invalid_or_missing_price": "Erreur" if p["price"] == "MISSING" or not price_re.match(p["price"]) else "OK",
            "null_price": "Erreur" if str(p["price"]).startswith("0") else "OK",
            "missing_title": "Erreur" if p["title"] == "MISSING" else "OK",
            "description_missing_or_short": "Erreur" if len(p["description"]) < 20 else "OK",
            "invalid_availability": "Erreur" if p["availability"] == "MISSING" else "OK",
            "missing_or_empty_color": "Erreur" if p["color"] == "MISSING" else "OK",
            "missing_or_empty_gender": "Erreur" if p["gender"] == "MISSING" else "OK",
            "missing_or_empty_size": "Erreur" if p["size"] == "MISSING" else "OK",
            "missing_or_empty_age_group": "Erreur" if p["age_group"] == "MISSING" else "OK",
            "missing_or_empty_image_link": "Erreur" if p["image_link"] == "MISSING" else "OK",
            "missing_or_empty_link": "Erreur" if p["link"] == "MISSING" else "OK",
        }
        validated.append({**p, **errs})
        seen.add(p["id"])
    return validated, validated  # compatibilité avec votre appel

# ---------- 5. Excel ----------
def generate_excel(rows: list[dict]) -> BytesIO:
    wb, ws = Workbook(), Workbook().active
    ws.title = "Validation Results"
    header = list(rows[0].keys())
    ws.append(header)
    for row in rows:
        ws.append([row.get(h, "") for h in header])
    for c in ws[1]:
        c.font = Font(bold=True)
    buf = BytesIO(); wb.save(buf); buf.seek(0); return buf

# ---------- 6. Interface ----------
def main():
    add_custom_css()
    st.markdown("<h1 class='main-title'>Audit Flux Google Merchant</h1>", unsafe_allow_html=True)
    url = st.text_input("URL du flux XML")
    upload = st.file_uploader("…ou importez un fichier XML :", type=["xml"])
    if st.button("Auditer le flux"):
        xml = fetch_xml(url) if url else (upload.read() if upload else None)
        if xml:
            root = parse_xml(xml)
            if root:
                products = analyze_products(root)
                errors, validated = validate_products(products)
                excel = generate_excel(validated)
                st.success("Audit terminé !")
                st.download_button(
                    "Télécharger le fichier Excel",
                    data=excel,
                    file_name="audit_flux_google_merchant.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

if __name__ == "__main__":
    main()
