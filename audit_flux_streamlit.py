import streamlit as st
import requests
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO
import re

# Injecter du CSS personnalisé
def add_custom_css():
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

# Étape 1 : Télécharger le flux XML
def fetch_xml(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e:
        st.error(f"Erreur lors du téléchargement du flux : {e}")
        return None

# Étape 2 : Parser le contenu XML
def parse_xml(content):
    try:
        root = ET.fromstring(content)
        return root
    except ET.ParseError as e:
        st.error(f"Erreur lors du parsing XML : {e}")
        return None

# Étape 3 : Analyser les produits
def analyze_products(root):
    namespace = {"g": "http://base.google.com/ns/1.0"}
    products = []
    for item in root.findall(".//item", namespace):
        product = {
            'id': item.find("g:id", namespace).text if item.find("g:id", namespace) is not None else "MISSING",
            'title': (
                item.find("g:title", namespace).text if item.find("g:title", namespace) is not None else
                item.find("title").text if item.find("title") is not None else "MISSING"
            ),
            'description': (
                item.find("g:description", namespace).text if item.find("g:description", namespace) is not None else
                item.find("description").text if item.find("description") is not None else "MISSING"
            ),
            'product_url': (
                item.find("g:link", namespace).text if item.find("g:link", namespace) is not None else
                item.find("link").text if item.find("link") is not None else "MISSING"
            ),
            'color': item.find("g:color", namespace).text if item.find("g:color", namespace) is not None else "MISSING",
            'gender': item.find("g:gender", namespace).text if item.find("g:gender", namespace) is not None else "MISSING",
            'size': item.find("g:size", namespace).text if item.find("g:size", namespace) is not None else "MISSING",
            'age_group': item.find("g:age_group", namespace).text if item.find("g:age_group", namespace) is not None else "MISSING",
            'image_link': item.find("g:image_link", namespace).text if item.find("g:image_link", namespace) is not None else "MISSING",
            'price': item.find("g:price", namespace).text if item.find("g:price", namespace) is not None else "MISSING",
            'availability': item.find("g:availability", namespace).text if item.find("g:availability", namespace) is not None else "MISSING",
        }
        products.append(product)
    return products

# Étape 4 : Validation des produits
def validate_products(products):
    errors = []
    price_pattern = re.compile(r"^\d+(\.\d{1,2})?( [A-Z]{3})?$")  # Accepte les prix avec devise (ex: "44.99 EUR")
    seen_ids = set()

    for product in products:
        product_errors = {
            'duplicate_id': "Erreur" if product['id'] in seen_ids else "OK",
            'invalid_or_missing_price': "Erreur" if product.get('price', 'MISSING') == "MISSING" or not price_pattern.match(product.get('price', '')) else "OK",
            'null_price': "Erreur" if product.get('price', '').startswith("0") else "OK",
            'missing_title': "Erreur" if product.get('title', 'MISSING') == "MISSING" else "OK",
            'description_missing_or_short': "Erreur" if len(product.get('description', '')) < 20 else "OK",
            'invalid_availability': "Erreur" if product.get('availability', 'MISSING') == "MISSING" else "OK",
            'missing_or_empty_color': "Erreur" if product.get('color', 'MISSING') == "MISSING" else "OK",
            'missing_or_empty_gender': "Erreur" if product.get('gender', 'MISSING') == "MISSING" else "OK",
            'missing_or_empty_size': "Erreur" if product.get('size', 'MISSING') == "MISSING" else "OK",
            'missing_or_empty_age_group': "Erreur" if product.get('age_group', 'MISSING') == "MISSING" else "OK",
            'missing_or_empty_image_link': "Erreur" if product.get('image_link', 'MISSING') == "MISSING" else "OK",
        }
        errors.append({**product, **product_errors})
        seen_ids.add(product['id'])

    return errors

# Étape 5 : Générer le fichier Excel
def generate_excel(data):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Validation Results"

    headers = [
        "Product ID", "Title", "Description", "Product URL", "Color", "Gender", "Size", "Age Group", "Image Link",
        "Duplicate ID", "Invalid or Missing Price", "Prix nul", "Missing Title",
        "Description Missing or Too Short", "Invalid Availability", "Missing or Empty Color",
        "Missing or Empty Gender", "Missing or Empty Size", "Missing or Empty Age Group",
        "Missing or Empty Image Link"
    ]
    sheet.append(headers)

    for product in data:
        sheet.append([
            product['id'], product['title'], product['description'], product['product_url'], product['color'],
            product['gender'], product['size'], product['age_group'], product['image_link'],
            product['duplicate_id'], product['invalid_or_missing_price'], product['null_price'],
            product['missing_title'], product['description_missing_or_short'], product['invalid_availability'],
            product['missing_or_empty_color'], product['missing_or_empty_gender'], product['missing_or_empty_size'],
            product['missing_or_empty_age_group'], product['missing_or_empty_image_link']
        ])

    for col in sheet.iter_cols(min_row=1, max_row=1, min_col=1, max_col=len(headers)):
        for cell in col:
            cell.font = Font(bold=True)

    excel_data = BytesIO()
    workbook.save(excel_data)
    excel_data.seek(0)
    return excel_data

# Fonction principale
def main():
    add_custom_css()
    st.markdown("<h1 class='main-title'>Audit Flux Google Merchant</h1>", unsafe_allow_html=True)

    url = st.text_input("Entrez l'URL du flux XML :")
    uploaded_file = st.file_uploader("Téléchargez un fichier XML :", type=["xml"])

    if st.button("Auditer le flux"):
        content = None

        if url:
            content = fetch_xml(url)
        elif uploaded_file is not None:
            content = uploaded_file.read()

        if content:
            root = parse_xml(content)
            if root:
                products = analyze_products(root)
                validated_products = validate_products(products)
                excel_file = generate_excel(validated_products)

                st.success("Audit terminé. Téléchargez le fichier Excel ci-dessous :")
                st.download_button(
                    label="Télécharger le fichier Excel",
                    data=excel_file,
                    file_name="audit_flux_google_merchant.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
