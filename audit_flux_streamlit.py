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
            'invalid_or_missing_price': ("Erreur"
                if product.get('price', 'MISSING') == "MISSING" or not price_pattern.match(product.get('price', ''))
                else "OK"
            ),
            'null_price': "Erreur" if product.get('price', '').startswith("0") else "OK",
            'missing_title': "Erreur" if product.get('title', 'MISSING') == "MISSING" else "OK",
            'description_missing_or_short': "Erreur" if len(product.get('description', '') or '') < 20 else "OK",
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

# Étape 5 : Générer le fichier Excel pour le détail des produits
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

# Tableau récapitulatif (statique) que l’on souhaite générer en Excel
SUMMARY_DATA = [
    {
        "Nom du champ": "id",
        "Existant/Manquants": "Existant",
        "Status": "Obligatoire",
        "Nb. de produits Optimisables": "0",
        "Commentaire": ""
    },
    {
        "Nom du champ": "title",
        "Existant/Manquants": "Manquants",
        "Status": "Obligatoire",
        "Nb. de produits Optimisables": "1",
        "Commentaire": "34183556563084"
    },
    {
        "Nom du champ": "link",
        "Existant/Manquants": "Existant",
        "Status": "Obligatoire",
        "Nb. de produits Optimisables": "0",
        "Commentaire": ""
    },
    {
        "Nom du champ": "image_link",
        "Existant/Manquants": "Existant",
        "Status": "Obligatoire",
        "Nb. de produits Optimisables": "0",
        "Commentaire": ""
    },
    {
        "Nom du champ": "price",
        "Existant/Manquants": "Incohérents / Nuls",
        "Status": "Obligatoire",
        "Nb. de produits Optimisables": "6 / 5",
        "Commentaire": (
            "Prix incohérents / Nuls\n"
            "40468974043276\n"
            "41209498108044\n"
            "48654943486289\n"
            "48654915404113\n"
            "48654943912273\n"
            "41082134855820\n"
            "48004016275793\n"
            "55093296693622\n"
            "53565093806454\n"
            "48004005822801\n"
            "48004019224913\n"
            "53565089546614"
        )
    },
    {
        "Nom du champ": "description",
        "Existant/Manquants": "Manquants",
        "Status": "Obligatoire",
        "Nb. de produits Optimisables": "10",
        "Commentaire": (
            "48004005822801\n"
            "48004016275793\n"
            "48004019224913\n"
            "48654719779153\n"
            "48654719844689\n"
            "48654915633489\n"
            "53565089546614\n"
            "53565093806454\n"
            "55093292597622"
        )
    },
    {
        "Nom du champ": "availability",
        "Existant/Manquants": "Existant",
        "Status": "Obligatoire",
        "Nb. de produits Optimisables": "0",
        "Commentaire": ""
    },
    {
        "Nom du champ": "condition",
        "Existant/Manquants": "Existant",
        "Status": "Obligatoire SI OCCASION ou RECONDITIONNEE",
        "Nb. de produits Optimisables": "0",
        "Commentaire": ""
    },
    {
        "Nom du champ": "brand",
        "Existant/Manquants": "Existant",
        "Status": "Obligatoire",
        "Nb. de produits Optimisables": "0",
        "Commentaire": ""
    },
    {
        "Nom du champ": "gtin",
        "Existant/Manquants": "Manquants / Doublons",
        "Status": "Obligatoire si GTIN existant",
        "Nb. de produits Optimisables": "7 / 192",
        "Commentaire": (
            "53558283436406\n"
            "53558282060150\n"
            "47925273395537\n"
            "47925270643025\n"
            "47925268447569\n"
            "47925266743633\n"
            "47925262352721"
        )
    },
    {
        "Nom du champ": "mpn",
        "Existant/Manquants": "",
        "Status": "Obligatoire si pas de GTIN",
        "Nb. de produits Optimisables": "0",
        "Commentaire": ""
    },
    {
        "Nom du champ": "color",
        "Existant/Manquants": "Manquants",
        "Status": "Obligatoire si VETEMENT ET ACCESSOIRES",
        "Nb. de produits Optimisables": "17",
        "Commentaire": (
            "53558283436406\n"
            "41178288390284\n"
            "53565093806454\n"
            "53558282060150\n"
            "47925273395537\n"
            "47925268447569\n"
            "47984860856657\n"
            "55093292597622\n"
            "34183556563084\n"
            "53565089546614\n"
            "48004019224913\n"
            "48004016275793\n"
            "47925262352721\n"
            "47925270643025\n"
            "55093296693622\n"
            "47925266743633\n"
            "48004005822801"
        )
    },
    {
        "Nom du champ": "size",
        "Existant/Manquants": "Inexistant",
        "Status": "Obligatoire si VETEMENT ET ACCESSOIRES",
        "Nb. de produits Optimisables": "1995",
        "Commentaire": "A lister par type de produit"
    },
    {
        "Nom du champ": "age_group",
        "Existant/Manquants": "Existant",
        "Status": "Obligatoire si VETEMENT ET ACCESSOIRES",
        "Nb. de produits Optimisables": "0",
        "Commentaire": ""
    },
    {
        "Nom du champ": "gender",
        "Existant/Manquants": "Existant",
        "Status": "Obligatoire si VETEMENT ET ACCESSOIRES",
        "Nb. de produits Optimisables": "0",
        "Commentaire": "Bien distinguer Frère & Soeur"
    },
    {
        "Nom du champ": "item_group_id",
        "Existant/Manquants": "Existant",
        "Status": "Obligatoire POUR VARIANTES",
        "Nb. de produits Optimisables": "0",
        "Commentaire": ""
    },
    {
        "Nom du champ": "shipping",
        "Existant/Manquants": "",
        "Status": "Obligatoire pour remplacer la data GMC",
        "Nb. de produits Optimisables": "0",
        "Commentaire": ""
    },
    {
        "Nom du champ": "shipping_weight",
        "Existant/Manquants": "Inexistant",
        "Status": "Recommandé",
        "Nb. de produits Optimisables": "1995",
        "Commentaire": "A normaliser / lister"
    },
    {
        "Nom du champ": "pattern",
        "Existant/Manquants": "Inexistant",
        "Status": "Recommandé si VETEMENT ET ACCESSOIRES",
        "Nb. de produits Optimisables": "1995",
        "Commentaire": "A normaliser / lister"
    },
    {
        "Nom du champ": "material",
        "Existant/Manquants": "Inexistant",
        "Status": "Recommandé si VETEMENT ET ACCESSOIRES",
        "Nb. de produits Optimisables": "1995",
        "Commentaire": "A lister par type de produit"
    },
    {
        "Nom du champ": "additional_image_link",
        "Existant/Manquants": "Inexistant",
        "Status": "Recommandé",
        "Nb. de produits Optimisables": "1995",
        "Commentaire": "A lister par type de produit"
    },
    {
        "Nom du champ": "size_type",
        "Existant/Manquants": "Inexistant",
        "Status": "Recommandé si VETEMENT ET ACCESSOIRES",
        "Nb. de produits Optimisables": "1995",
        "Commentaire": "A lister par type de produit"
    },
    {
        "Nom du champ": "size_system",
        "Existant/Manquants": "Inexistant",
        "Status": "Recommandé si VETEMENT ET ACCESSOIRES",
        "Nb. de produits Optimisables": "1995",
        "Commentaire": "A lister par type de produit"
    },
    {
        "Nom du champ": "canonical_link",
        "Existant/Manquants": "Inexistant",
        "Status": "Recommandé pour tous les produits",
        "Nb. de produits Optimisables": "1995",
        "Commentaire": "A voir comment remplir le champ en reprenant le g:link"
    },
    {
        "Nom du champ": "expiration_date",
        "Existant/Manquants": "Inexistant",
        "Status": "Recommandé pour arrêter de display un produit",
        "Nb. de produits Optimisables": "-",
        "Commentaire": ""
    },
    {
        "Nom du champ": "sale_price",
        "Existant/Manquants": "Existant",
        "Status": "Recommandé lors de promotions",
        "Nb. de produits Optimisables": "-",
        "Commentaire": ""
    },
    {
        "Nom du champ": "sale_price_effective_date",
        "Existant/Manquants": "Non Existant",
        "Status": "Recommandé lors de promotions",
        "Nb. de produits Optimisables": "-",
        "Commentaire": ""
    },
    {
        "Nom du champ": "product_highlight",
        "Existant/Manquants": "Inexistant",
        "Status": "Recommandé",
        "Nb. de produits Optimisables": "1995",
        "Commentaire": "A lister par produit ou type de produit pour remplir le champ"
    },
    {
        "Nom du champ": "ships_from_country",
        "Existant/Manquants": "Inexistant",
        "Status": "Recommandé",
        "Nb. de produits Optimisables": "1995",
        "Commentaire": "A ajouter manuellement"
    },
    {
        "Nom du champ": "max_handling_time",
        "Existant/Manquants": "Inexistant",
        "Status": "Recommandé",
        "Nb. de produits Optimisables": "1995",
        "Commentaire": "Dans GMC, a voir"
    },
    {
        "Nom du champ": "availability_date",
        "Existant/Manquants": "Inexistant",
        "Status": "Obligatoire si produit en pré-order",
        "Nb. de produits Optimisables": "-",
        "Commentaire": ""
    },
]

# Fonction pour générer le fichier Excel du récapitulatif
def generate_summary_excel(summary_data):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Tableau récap"

    headers = [
        "Nom du champ",
        "Existant/Manquants",
        "Status",
        "Nb. de produits Optimisables",
        "Commentaire"
    ]
    sheet.append(headers)

    for row in summary_data:
        sheet.append([
            row["Nom du champ"],
            row["Existant/Manquants"],
            row["Status"],
            row["Nb. de produits Optimisables"],
            row["Commentaire"]
        ])

    # Mettre les en-têtes en gras
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
                    label="Télécharger le fichier Excel (Détails produits)",
                    data=excel_file,
                    file_name="audit_flux_google_merchant.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # 2ème CTA pour télécharger le récap
                st.info("Téléchargez le récapitulatif ci-dessous :")
                summary_file = generate_summary_excel(SUMMARY_DATA)
                st.download_button(
                    label="Télécharger le récap",
                    data=summary_file,
                    file_name="recap_flux_google_merchant.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
