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
    """
    Récupère tous les attributs demandés.
    Si un attribut n'est pas trouvé, on renvoie "MISSING" par défaut.
    """
    namespace = {"g": "http://base.google.com/ns/1.0"}
    
    def get_text(item, tag):
        """Retourne le texte de l'élément XML (namespace Google) ou 'MISSING' si vide."""
        element = item.find(f"g:{tag}", namespace)
        return element.text if element is not None else "MISSING"
    
    # Pour 'link', on essaie d'abord la balise g:link puis link "classique"
    def get_link(item):
        g_link = item.find("g:link", namespace)
        if g_link is not None and g_link.text:
            return g_link.text
        alt_link = item.find("link")
        if alt_link is not None and alt_link.text:
            return alt_link.text
        return "MISSING"

    # Idem pour 'title' et 'description', on check g:tag puis tag normal
    def get_title(item):
        t = item.find("g:title", namespace)
        if t is not None and t.text:
            return t.text
        alt_t = item.find("title")
        if alt_t is not None and alt_t.text:
            return alt_t.text
        return "MISSING"

    def get_description(item):
        d = item.find("g:description", namespace)
        if d is not None and d.text:
            return d.text
        alt_d = item.find("description")
        if alt_d is not None and alt_d.text:
            return alt_d.text
        return "MISSING"
    
    # Pour l'élément shipping, on renvoie tout le bloc XML sous forme de texte (simplifié),
    # ou MISSING s'il n'existe pas
    def get_shipping_block(item):
        shipping_elem = item.find("g:shipping", namespace)
        if shipping_elem is not None:
            return "".join(shipping_elem.itertext()).strip()
        return "MISSING"

    products = []
    for item in root.findall(".//item", namespace):
        product = {
            # Champs déjà existants ou équivalents
            'id': get_text(item, "id"),
            'title': get_title(item),
            'description': get_description(item),
            'link': get_link(item),
            'color': get_text(item, "color"),
            'gender': get_text(item, "gender"),
            'size': get_text(item, "size"),
            'age_group': get_text(item, "age_group"),
            'image_link': get_text(item, "image_link"),
            'price': get_text(item, "price"),
            'availability': get_text(item, "availability"),

            # Nouveaux attributs demandés
            'condition': get_text(item, "condition"),
            'brand': get_text(item, "brand"),
            'gtin': get_text(item, "gtin"),
            'mpn': get_text(item, "mpn"),
            'item_group_id': get_text(item, "item_group_id"),
            'shipping': get_shipping_block(item),  # Le bloc <shipping> en brut
            'shipping_weight': get_text(item, "shipping_weight"),
            'pattern': get_text(item, "pattern"),
            'material': get_text(item, "material"),
            'additional_image_link': get_text(item, "additional_image_link"),
            'size_type': get_text(item, "size_type"),
            'size_system': get_text(item, "size_system"),
            'canonical_link': get_text(item, "canonical_link"),
            'expiration_date': get_text(item, "expiration_date"),
            'sale_price': get_text(item, "sale_price"),
            'sale_price_effective_date': get_text(item, "sale_price_effective_date"),
            'product_highlight': get_text(item, "product_highlight"),
            'ships_from_country': get_text(item, "ships_from_country"),
            'max_handling_time': get_text(item, "max_handling_time"),
            'availability_date': get_text(item, "availability_date"),
        }
        products.append(product)
    return products

# Étape 4 : Validation des produits
def validate_products(products):
    """
    Effectue les contrôles déjà existants et ajoute les résultats
    (OK / Erreur) dans le dictionnaire de chaque produit.
    """
    errors = []
    price_pattern = re.compile(r"^\d+(\.\d{1,2})?( [A-Z]{3})?$")  # ex: "44.99 EUR"
    seen_ids = set()

    for product in products:
        current_id = product.get('id', 'MISSING')
        current_price = product.get('price', 'MISSING')
        current_desc = product.get('description', '')
        
        product_errors = {
            'duplicate_id': "Erreur" if current_id in seen_ids else "OK",
            'invalid_or_missing_price': (
                "Erreur"
                if current_price == "MISSING" or not price_pattern.match(current_price)
                else "OK"
            ),
            'null_price': "Erreur" if current_price.startswith("0") else "OK",
            'missing_title': "Erreur" if product.get('title', 'MISSING') == "MISSING" else "OK",
            'description_missing_or_short': "Erreur" if len(current_desc) < 20 else "OK",
            'invalid_availability': "Erreur" if product.get('availability', 'MISSING') == "MISSING" else "OK",
            'missing_or_empty_color': "Erreur" if product.get('color', 'MISSING') == "MISSING" else "OK",
            'missing_or_empty_gender': "Erreur" if product.get('gender', 'MISSING') == "MISSING" else "OK",
            'missing_or_empty_size': "Erreur" if product.get('size', 'MISSING') == "MISSING" else "OK",
            'missing_or_empty_age_group': "Erreur" if product.get('age_group', 'MISSING') == "MISSING" else "OK",
            'missing_or_empty_image_link': "Erreur" if product.get('image_link', 'MISSING') == "MISSING" else "OK",
        }
        errors.append({**product, **product_errors})
        seen_ids.add(current_id)

    return errors

# Étape 5 : Générer le fichier Excel
def generate_excel(data):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Validation Results"

    # Les en-têtes (colonnes) : d'abord tous les attributs, puis les colonnes de validation
    headers = [
        "id", 
        "title", 
        "link", 
        "image_link", 
        "price", 
        "description", 
        "availability",
        "condition",
        "brand",
        "gtin",
        "mpn",
        "color",
        "size",
        "age_group",
        "gender",
        "item_group_id",
        "shipping",
        "shipping_weight",
        "pattern",
        "material",
        "additional_image_link",
        "size_type",
        "size_system",
        "canonical_link",
        "expiration_date",
        "sale_price",
        "sale_price_effective_date",
        "product_highlight",
        "ships_from_country",
        "max_handling_time",
        "availability_date",
        # Validations existantes
        "duplicate_id",
        "invalid_or_missing_price",
        "null_price",
        "missing_title",
        "description_missing_or_short",
        "invalid_availability",
        "missing_or_empty_color",
        "missing_or_empty_gender",
        "missing_or_empty_size",
        "missing_or_empty_age_group",
        "missing_or_empty_image_link",
    ]
    sheet.append(headers)

    for product in data:
        row = [
            product.get("id", ""),
            product.get("title", ""),
            product.get("link", ""),
            product.get("image_link", ""),
            product.get("price", ""),
            product.get("description", ""),
            product.get("availability", ""),
            product.get("condition", ""),
            product.get("brand", ""),
            product.get("gtin", ""),
            product.get("mpn", ""),
            product.get("color", ""),
            product.get("size", ""),
            product.get("age_group", ""),
            product.get("gender", ""),
            product.get("item_group_id", ""),
            product.get("shipping", ""),
            product.get("shipping_weight", ""),
            product.get("pattern", ""),
            product.get("material", ""),
            product.get("additional_image_link", ""),
            product.get("size_type", ""),
            product.get("size_system", ""),
            product.get("canonical_link", ""),
            product.get("expiration_date", ""),
            product.get("sale_price", ""),
            product.get("sale_price_effective_date", ""),
            product.get("product_highlight", ""),
            product.get("ships_from_country", ""),
            product.get("max_handling_time", ""),
            product.get("availability_date", ""),
            # Résultats validation
            product.get("duplicate_id", ""),
            product.get("invalid_or_missing_price", ""),
            product.get("null_price", ""),
            product.get("missing_title", ""),
            product.get("description_missing_or_short", ""),
            product.get("invalid_availability", ""),
            product.get("missing_or_empty_color", ""),
            product.get("missing_or_empty_gender", ""),
            product.get("missing_or_empty_size", ""),
            product.get("missing_or_empty_age_group", ""),
            product.get("missing_or_empty_image_link", ""),
        ]
        sheet.append(row)

    # Mettre les titres en gras
    for col in sheet.iter_cols(min_row=1, max_row=1):
        for cell in col:
            cell.font = Font(bold=True)

    excel_data = BytesIO()
    workbook.save(excel_data)
    excel_data.seek(0)
    return excel_data

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
