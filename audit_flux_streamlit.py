import streamlit as st
import requests
import xml.etree.ElementTree as ET
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO
import re  # Pour valider le format du prix avec devise

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

# Étape 1 : Télécharger le flux XML depuis l'URL
def fetch_xml(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e:
        st.error(f"Erreur lors du téléchargement du flux : {e}")
        return None

# Étape 2 : Parser le contenu XML avec gestion des espaces de noms
def parse_xml(content):
    try:
        root = ET.fromstring(content)
        return root
    except ET.ParseError as e:
        st.error(f"Erreur lors du parsing XML : {e}")
        return None

# Étape 3 : Analyser les données des produits
def analyze_products(root):
    namespace = {"g": "http://base.google.com/ns/1.0"}
    products = []
    for item in root.findall(".//item", namespace):
        product = {
            'id': item.find("g:id", namespace).text if item.find("g:id", namespace) is not None else "MISSING",
            'title': item.find("g:title", namespace).text if item.find("g:title", namespace) is not None else "MISSING",
            'description': item.find("g:description", namespace).text if item.find("g:description", namespace) is not None else "MISSING",
            'price': item.find("g:price", namespace).text if item.find("g:price", namespace) is not None else "MISSING",
            'availability': item.find("g:availability", namespace).text if item.find("g:availability", namespace) is not None else "MISSING",
            'condition': item.find("g:condition", namespace).text if item.find("g:condition", namespace) is not None else "MISSING",
            'brand': item.find("g:brand", namespace).text if item.find("g:brand", namespace) is not None else "MISSING",
            'gtin': item.find("g:gtin", namespace).text if item.find("g:gtin", namespace) is not None else "MISSING",
            'mpn': item.find("g:mpn", namespace).text if item.find("g:mpn", namespace) is not None else "MISSING",
            'color': item.find("g:color", namespace).text if item.find("g:color", namespace) is not None else "MISSING",
            'size': item.find("g:size", namespace).text if item.find("g:size", namespace) is not None else "MISSING",
            'age_group': item.find("g:age_group", namespace).text if item.find("g:age_group", namespace) is not None else "MISSING",
            'gender': item.find("g:gender", namespace).text if item.find("g:gender", namespace) is not None else "MISSING",
            'item_group_id': item.find("g:item_group_id", namespace).text if item.find("g:item_group_id", namespace) is not None else "MISSING",
            'shipping': item.find("g:shipping", namespace).text if item.find("g:shipping", namespace) is not None else "MISSING",
            'shipping_weight': item.find("g:shipping_weight", namespace).text if item.find("g:shipping_weight", namespace) is not None else "MISSING",
            'pattern': item.find("g:pattern", namespace).text if item.find("g:pattern", namespace) is not None else "MISSING",
            'material': item.find("g:material", namespace).text if item.find("g:material", namespace) is not None else "MISSING",
            'additional_image_link': item.find("g:additional_image_link", namespace).text if item.find("g:additional_image_link", namespace) is not None else "MISSING",
            'size_type': item.find("g:size_type", namespace).text if item.find("g:size_type", namespace) is not None else "MISSING",
            'size_system': item.find("g:size_system", namespace).text if item.find("g:size_system", namespace) is not None else "MISSING",
            'canonical_link': item.find("g:canonical_link", namespace).text if item.find("g:canonical_link", namespace) is not None else "MISSING",
            'expiration_date': item.find("g:expiration_date", namespace).text if item.find("g:expiration_date", namespace) is not None else "MISSING",
            'sale_price': item.find("g:sale_price", namespace).text if item.find("g:sale_price", namespace) is not None else "MISSING",
            'sale_price_effective_date': item.find("g:sale_price_effective_date", namespace).text if item.find("g:sale_price_effective_date", namespace) is not None else "MISSING",
            'product_highlight': item.find("g:product_highlight", namespace).text if item.find("g:product_highlight", namespace) is not None else "MISSING",
            'ships_from_country': item.find("g:ships_from_country", namespace).text if item.find("g:ships_from_country", namespace) is not None else "MISSING",
            'max_handling_time': item.find("g:max_handling_time", namespace).text if item.find("g:max_handling_time", namespace) is not None else "MISSING",
            'availability_date': item.find("g:availability_date", namespace).text if item.find("g:availability_date", namespace) is not None else "MISSING"
        }
        products.append(product)
    return products

# Étape 4 : Validation des données
def validate_products(products):
    errors = []
    seen_ids = set()
    price_pattern = re.compile(r"^\d+(\.\d{1,2})?( [A-Z]{3})?$")  # Accepte les prix avec devise (ex: "44.99 EUR")

    for product in products:
        # Vérification des champs obligatoires
        if product['id'] == "MISSING":
            errors.append(("Missing ID", product['id']))
        if product['title'] == "MISSING":
            errors.append(("Missing Title", product['id']))
        if product['product_url'] == "MISSING":
            errors.append(("Missing Product URL", product['id']))
        if product['image_link'] == "MISSING":
            errors.append(("Missing Image Link", product['id']))
        if product['price'] == "MISSING" or not price_pattern.match(product['price']):
            errors.append(("Invalid or Missing Price", product['id']))
        
        # Conditions spécifiques
        if product['condition'] in ["used", "refurbished"] and product['condition'] == "MISSING":
            errors.append(("Missing Condition", product['id']))
        if product['brand'] == "MISSING":
            errors.append(("Missing Brand", product['id']))

        # Vêtements et accessoires
        if product['color'] == "MISSING" or product['size'] == "MISSING" or product['gender'] == "MISSING":
            errors.append(("Missing Clothing Attributes", product['id']))

    return errors, products

# Main Function
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
                errors, validated_products = validate_products(products)

                if errors:
                    st.write(f"Nombre total de produits analysés : {len(products)}")
                    st.write(f"Nombre total d'erreurs : {len(errors)}")
                    for error_type, count in summarize_errors(errors).items():
                        st.write(f"- {error_type}: {count}")

if __name__ == "__main__":
    main()
