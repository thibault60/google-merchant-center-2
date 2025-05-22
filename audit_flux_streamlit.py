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
def add_custom_css():
    """Injecte un peu de CSS pour un rendu plus propre."""
    st.markdown(
        """
        <style>
        body {
            background-color: #f8f9fa;
            font-family: 'Arial', sans-serif;
        }
        .main-title {
            color: #343a40;
            text-align: center;
            font-size: 2.5rem;
            margin-bottom: 1rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

# --------------------------------------------------
# 1. Téléchargement du flux
# --------------------------------------------------
def fetch_xml(url: str) -> bytes | None:
    """Télécharge le fichier XML et renvoie son contenu brut."""
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e:
        st.error(f"Erreur lors du téléchargement du flux : {e}")
        return None

# --------------------------------------------------
# 2. Parsing XML
# --------------------------------------------------
def parse_xml(content: bytes) -> ET.Element | None:
    """Parse le XML et renvoie la racine de l’arbre."""
    try:
        return ET.fromstring(content)
    except ET.ParseError as e:
        st.error(f"Erreur lors du parsing XML : {e}")
        return None

# --------------------------------------------------
# 3. Extraction des produits
# --------------------------------------------------
def analyze_products(root: ET.Element) -> list[dict]:
    """
    Extrait tous les attributs utiles.    
    Si un attribut est manquant, on retourne « MISSING ».
    """
    ns = {"g": "http://base.google.com/ns/1.0"}

    def gtext(item, tag: str) -> str:
        """Retourne le texte de g:tag ou 'MISSING'."""
        elem = item.find(f"g:{tag}", ns)
        return elem.text.strip() if elem is not None and elem.text else "MISSING"

    def get_link(item) -> str:
        """Essaye g:link puis link."""
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
        alt_title = item.find("title")
        if alt_title is not None and alt_title.text:
            return alt_title.text.strip()
        return "MISSING"

    def get_description(item) -> str:
        g_desc = item.find("g:description", ns)
        if g_desc is not None and g_desc.text:
            return g_desc.text.strip()
        alt_desc = item.find("description")
        if alt_desc is not None and alt_desc.text:
            return alt_desc.text.strip()
        return "MISSING"

    def get_shipping_block(item) -> str:
        ship = item.find("g:shipping", ns)
        return "".join(ship.itertext()).strip() if ship is not None else "MISSING"

    products = []
    for item in root.findall(".//item", ns):
        product = {
            # Champs principaux
            "id": gtext(item, "id"),
            "title": get_title(item),
            "description": get_description(item),
            "link": get_link(item),
            "image_link": gtext(item, "image_link"),
            "price": gtext(item, "price"),
            "availability": gtext(item, "availability"),
            "color": gtext(item, "color"),
            "gender": gtext(item, "gender"),
            "size": gtext(item, "size"),
            "age_group": gtext(item, "age_group"),

            # Champs additionnels
            "condition": gtext(item, "condition"),
            "brand": gtext(item, "brand"),
            "gtin": gtext(item, "gtin"),
            "mpn": gtext(item, "mpn"),
            "item_group_id": gtext(item, "item_group_id"),
            "shipping": get_shipping_block(item),
            "shipping_weight": gtext(item, "shipping_weight"),
            "pattern": gtext(item, "pattern"),
            "material": gtext(item, "material"),
            "additional_image_link": gtext(item, "additional_image_link"),
            "size_type": gtext(item, "size_type"),
            "size_system": gtext(item, "size_system"),
            "canonical_link": gtext(item, "canonical_link"),
            "expiration_date": gtext(item, "expiration_date"),
            "sale_price": gtext(item, "sale_price"),
            "sale_price_effective_date": gtext(item, "sale_price_effective_date"),
            "product_highlight": gtext(item, "product_highlight"),
            "ships_from_country": gtext(item, "ships_from_country"),
            "max_handling_time": gtext(item, "max_handling_time"),
            "availability_date": gtext(item, "availability_date"),
        }
        products.append(product)

    return products

# --------------------------------------------------
# 4. Validation des produits
# --------------------------------------------------
def validate_products(products: list[dict]) -> list[dict]:
    """
    Ajoute des colonnes de validation (OK / Erreur) à chaque produit.
    """
    errors = []
    price_pattern = re.compile(r"^\d+(\.\d{1,2})? [A-Z]{3}$")  # devise obligatoire
    seen_ids: set[str] = set()

    for product in products:
        current_id = product.get("id", "MISSING")
        current_price = product.get("price", "MISSING")
        current_desc = product.get("description", "")
        current_link = product.get("link", "MISSING")

        product_errors = {
            "duplicate_id": "Erreur" if current_id in seen_ids else "OK",
            "invalid_or_missing_price": (
                "Erreur"
                if current_price == "MISSING" or not price_pattern.match(current_price)
                else "OK"
            ),
            "null_price": "Erreur" if str(current_price).startswith("0") else "OK",
            "missing_title": "Erreur" if product.get("title") == "MISSING" else "OK",
            "description_missing_or_short": "Erreur" if len(current_desc) < 20 else "OK",
            "invalid_availability": "Erreur" if product.get("availability") == "MISSING" else "OK",
            "missing_or_empty_color": "Erreur" if product.get("color") == "MISSING" else "OK",
            "missing_or_empty_gender": "Erreur" if product.get("gender") == "MISSING" else "OK",
            "missing_or_empty_size": "Erreur" if product.get("size") == "MISSING" else "OK",
            "missing_or_empty_age_group": "Erreur" if product.get("age_group") == "MISSING" else "OK",
            "missing_or_empty_image_link": "Erreur" if product.get("image_link") == "MISSING" else "OK",
            "missing_or_empty_link": "Erreur" if current_link == "MISSING" else "OK",
        }
        errors.append({**product, **product_errors})
        seen_ids.add(current_id)

    return errors

# --------------------------------------------------
# 5. Génération Excel
# --------------------------------------------------
def generate_excel(data: list[dict]) -> BytesIO:
    """Transforme la liste de dicts en fichier Excel en mémoire."""
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Validation Results"

    headers = list(data[0].keys())  # attributs + validations
    sheet.append(headers)

    for product in data:
        sheet.append([product.get(h, "") for h in headers])

    # Header en gras
    for cell in sheet[1]:
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
    uploaded_file = st.file_uploader("…ou téléchargez un fichier XML :", type=["xml"])

    if st.button("Auditer le flux"):
        content = None
        if url:
            content = fetch_xml(url)
        elif uploaded_file is not None:
            content = uploaded_file.read()

        if content:
            root = parse_xml(content)
            if root is not None:
                products = analyze_products(root)
                validated_products = validate_products(products)
                excel_file = generate_excel(validated_products)

                st.success("Audit terminé ! Téléchargez le rapport Excel :")
                st.download_button(
                    label="Télécharger le fichier Excel",
                    data=excel_file,
                    file_name="audit_flux_google_merchant.xlsx",
                    mime=("application/"
                          "vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                )

if __name__ == "__main__":
    main()
