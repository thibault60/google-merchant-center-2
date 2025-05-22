import streamlit as st
import requests
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO
import re

# --------------------------------------------------
# Mise en forme Streamlit
# --------------------------------------------------
def add_custom_css() -> None:
    """Ajoute un peu de CSS pour un rendu plus agréable."""
    st.markdown(
        """
        <style>
        body {background-color: #f8f9fa; font-family: Arial, sans-serif;}
        .main-title {color:#343a40; text-align:center; font-size:2.5rem; margin-bottom:1rem;}
        </style>
        """,
        unsafe_allow_html=True,
    )

# --------------------------------------------------
# 1. Téléchargement du flux
# --------------------------------------------------
def fetch_xml(url: str) -> bytes | None:
    """Télécharge le XML à partir d’une URL."""
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        return resp.content
    except requests.exceptions.RequestException as e:
        st.error(f"Erreur lors du téléchargement du flux : {e}")
        return None

# --------------------------------------------------
# 2. Parsing XML
# --------------------------------------------------
def parse_xml(content: bytes) -> ET.Element | None:
    """Parse le contenu XML et renvoie la racine."""
    try:
        return ET.fromstring(content)
    except ET.ParseError as e:
        st.error(f"Erreur de parsing XML : {e}")
        return None

# --------------------------------------------------
# 3. Extraction des produits
# --------------------------------------------------
def analyze_products(root: ET.Element) -> list[dict]:
    """Crée une liste de dictionnaires représentant chaque produit."""
    ns = {"g": "http://base.google.com/ns/1.0"}

    def gtext(item, tag: str) -> str:
        elt = item.find(f"g:{tag}", ns)
        return elt.text.strip() if elt is not None and elt.text else "MISSING"

    def get_link(item) -> str:
        g_link = item.find("g:link", ns)
        if g_link is not None and g_link.text:
            return g_link.text.strip()
        alt_link = item.find("link")
        if alt_link is not None and alt_link.text:
            return alt_link.text.strip()
        return "MISSING"

    def get_title(item) -> str:
        g_title = item.find("g:title", ns)
        if g_title is not None and g_title.text:
            return g_title.text.strip()
        title = item.find("title")
        return title.text.strip() if title is not None and title.text else "MISSING"

    def get_description(item) -> str:
        g_desc = item.find("g:description", ns)
        if g_desc is not None and g_desc.text:
            return g_desc.text.strip()
        desc = item.find("description")
        return desc.text.strip() if desc is not None and desc.text else "MISSING"

    def get_shipping_block(item) -> str:
        ship = item.find("g:shipping", ns)
        return "".join(ship.itertext()).strip() if ship is not None else "MISSING"

    products = []
    for it in root.findall(".//item", ns):
        products.append({
            # Champs obligatoires ou fréquents
            "id": gtext(it, "id"),
            "title": get_title(it),
            "description": get_description(it),
            "link": get_link(it),
            "image_link": gtext(it, "image_link"),
            "price": gtext(it, "price"),
            "availability": gtext(it, "availability"),
            "color": gtext(it, "color"),
            "gender": gtext(it, "gender"),
            "size": gtext(it, "size"),
            "age_group": gtext(it, "age_group"),

            # Attributs additionnels
            "condition": gtext(it, "condition"),
            "brand": gtext(it, "brand"),
            "gtin": gtext(it, "gtin"),
            "mpn": gtext(it, "mpn"),
            "item_group_id": gtext(it, "item_group_id"),
            "shipping": get_shipping_block(it),
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

# --------------------------------------------------
# 4. Validation des produits
# --------------------------------------------------
def validate_products(products: list[dict]) -> list[dict]:
    """Ajoute les colonnes de contrôle (OK / Erreur) à chaque produit."""
    results = []
    price_re = re.compile(r"^\d+(\.\d{1,2})? [A-Z]{3}$")  # devise obligatoire
    seen_ids: set[str] = set()

    for p in products:
        pid = p.get("id", "MISSING")
        price = p.get("price", "MISSING")
        desc = p.get("description", "")
        link = p.get("link", "MISSING")

        checks = {
            "duplicate_id": "Erreur" if pid in seen_ids else "OK",
            "invalid_or_missing_price": (
                "Erreur" if price == "MISSING" or not price_re.match(price) else "OK"
            ),
            "null_price": "Erreur" if str(price).startswith("0") else "OK",
            "missing_title": "Erreur" if p.get("title") == "MISSING" else "OK",
            "description_missing_or_short": "Erreur" if len(desc) < 20 else "OK",
            "invalid_availability": "Erreur" if p.get("availability") == "MISSING" else "OK",
            "missing_or_empty_color": "Erreur" if p.get("color") == "MISSING" else "OK",
            "missing_or_empty_gender": "Erreur" if p.get("gender") == "MISSING" else "OK",
            "missing_or_empty_size": "Erreur" if p.get("size") == "MISSING" else "OK",
            "missing_or_empty_age_group": "Erreur" if p.get("age_group") == "MISSING" else "OK",
            "missing_or_empty_image_link": "Erreur" if p.get("image_link") == "MISSING" else "OK",
            "missing_or_empty_link": "Erreur" if link == "MISSING" else "OK",
        }

        results.append({**p, **checks})
        seen_ids.add(pid)
    return results

# --------------------------------------------------
# 5. Génération du fichier Excel
# --------------------------------------------------
def generate_excel(data: list[dict]) -> BytesIO:
    """Crée un Excel en mémoire à partir de la liste de produits."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Validation Results"

    header = list(data[0].keys())
    ws.append(header)

    for prod in data:
        ws.append([prod.get(col, "") for col in header])

    # Entêtes en gras
    for cell in ws[1]:
        cell.font = Font(bold=True)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# --------------------------------------------------
# 6. Interface Streamlit
# --------------------------------------------------
def main() -> None:
    add_custom_css()
    st.markdown("<h1 class='main-title'>Audit Flux Google Merchant</h1>", unsafe_allow_html=True)

    url = st.text_input("Entrez l’URL du flux XML :")
    uploaded = st.file_uploader("…ou téléchargez un fichier XML :", type=["xml"])

    if st.button("Auditer le flux"):
        xml_content: bytes | None = None
        if url:
            xml_content = fetch_xml(url)
        elif uploaded is not None:
            xml_content = uploaded.read()

        if xml_content:
            root = parse_xml(xml_content)
            if root is not None:
                products = analyze_products(root)
                validated = validate_products(products)
                xl_file = generate_excel(validated)

                st.success("Audit terminé ! Téléchargez le rapport :")
                st.download_button(
                    label="Télécharger le fichier Excel",
                    data=xl_file,
                    file_name="audit_flux_google_merchant.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

if __name__ == "__main__":
    main()
