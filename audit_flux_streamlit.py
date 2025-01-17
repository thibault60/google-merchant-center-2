# Chemin principal
import streamlit as st
import requests
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from io import BytesIO

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

# Télécharger le flux XML depuis l'URL
def fetch_xml(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e:
        st.error(f"Erreur lors du téléchargement du flux : {e}")
        return None

# Parser le contenu XML
def parse_xml(content):
    try:
        root = ET.fromstring(content)
        return root
    except ET.ParseError as e:
        st.error(f"Erreur lors du parsing XML : {e}")
        return None

# Extraire les titres et descriptions des produits
def extract_titles_and_descriptions(root):
    namespace = {"g": "http://base.google.com/ns/1.0"}
    data = []
    for item in root.findall(".//item", namespace):
        title = item.find("g:title", namespace)
        description = item.find("g:description", namespace)
        data.append({
            'title': title.text if title is not None else "MISSING",
            'description': description.text if description is not None else "MISSING",
        })
    return data

# Exporter les données dans un fichier Excel
def export_to_excel(data):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Produits"
    sheet.append(["Titre", "Description"])

    for product in data:
        sheet.append([product['title'], product['description']])

    # Ajuster la largeur des colonnes
    for col in sheet.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        sheet.column_dimensions[col_letter].width = max_length + 2

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()

# Fonction principale
def main():
    add_custom_css()
    st.markdown("<h1 class='main-title'>Extraction des Titres et Descriptions</h1>", unsafe_allow_html=True)

    url = st.text_input("Entrez l'URL du flux XML :")
    uploaded_file = st.file_uploader("Téléchargez un fichier XML :", type=["xml"])

    if st.button("Extraire les données"):
        content = None

        if url:
            content = fetch_xml(url)
        elif uploaded_file is not None:
            content = uploaded_file.read()

        if content:
            root = parse_xml(content)
            if root:
                data = extract_titles_and_descriptions(root)
                st.success(f"Extraction réussie : {len(data)} produits trouvés.")
                excel_file = export_to_excel(data)
                st.download_button(
                    label="Télécharger le fichier Excel",
                    data=excel_file,
                    file_name="titres_descriptions.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

if __name__ == "__main__":
    main()
