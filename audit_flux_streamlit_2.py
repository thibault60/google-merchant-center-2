import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="TSV → XLSX Converter", page_icon="📊", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&display=swap');
    .block-container { max-width: 900px; padding-top: 2rem; }
    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
    .main-title { font-size: 2rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.2rem; }
    .subtitle { font-size: 1rem; color: #6b7280; margin-bottom: 2rem; }
    .metric-card {
        background: linear-gradient(135deg, #eff6ff, #dbeafe);
        border: 1px solid #bfdbfe;
        border-radius: 12px;
        padding: 1.2rem;
        text-align: center;
    }
    .metric-card h3 { font-size: 2rem; color: #1e40af; margin: 0; }
    .metric-card p { font-size: 0.85rem; color: #4b5563; margin: 0; }
    .stDownloadButton > button {
        background-color: #1e40af !important;
        color: white !important;
        border-radius: 8px !important;
        padding: 0.6rem 2rem !important;
        font-weight: 600 !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 TSV → XLSX Converter</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Convertit un fichier TSV (ou CSV) en fichier Excel formaté</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Configuration")
    separator = st.selectbox(
        "Séparateur du fichier source",
        options=["Tab (\\t)", "Virgule (,)", "Point-virgule (;)"],
        index=0,
    )
    sep_map = {"Tab (\\t)": "\t", "Virgule (,)": ",", "Point-virgule (;)": ";"}
    sep_char = sep_map[separator]

    sheet_name = st.text_input("Nom de la feuille Excel", value="Sheet1")
    freeze_header = st.checkbox("Figer la ligne d'en-tête", value=True)
    header_color = st.color_picker("Couleur d'en-tête", value="#1e40af")

uploaded_file = st.file_uploader(
    "📂 Importe ton fichier TSV / CSV / TXT",
    type=["tsv", "csv", "txt"],
)

if uploaded_file:
    try:
        df = pd.read_csv(uploaded_file, sep=sep_char, dtype=str)
        df.columns = df.columns.str.strip().str.strip('"')
        for col in df.columns:
            df[col] = df[col].str.strip().str.strip('"')
    except Exception as e:
        st.error(f"Erreur de lecture : {e}")
        st.stop()

    total_rows, total_cols = df.shape

    st.markdown("---")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f'<div class="metric-card"><h3>{total_rows}</h3><p>Lignes</p></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><h3>{total_cols}</h3><p>Colonnes</p></div>', unsafe_allow_html=True)
    with c3:
        fname = uploaded_file.name.rsplit(".", 1)[0] + ".xlsx"
        st.markdown(f'<div class="metric-card"><h3>XLSX</h3><p>{fname}</p></div>', unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("📋 Aperçu")
    st.dataframe(df, use_container_width=True, height=350)

    # ── Build XLSX ──
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]  # Excel max sheet name length

    # Header style
    hex_color = header_color.lstrip("#")
    header_fill = PatternFill("solid", fgColor=hex_color)
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ci, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=ci, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border
    ws.row_dimensions[1].height = 28

    # Data rows
    even_fill = PatternFill("solid", fgColor="EFF6FF")
    odd_fill = PatternFill("solid", fgColor="FFFFFF")
    data_font = Font(name="Arial", size=10)

    for ri, row in enumerate(df.itertuples(index=False), 2):
        fill = even_fill if ri % 2 == 0 else odd_fill
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value="" if pd.isna(val) else str(val))
            cell.font = data_font
            cell.fill = fill
            cell.border = border
            cell.alignment = Alignment(vertical="center")

    # Auto column width (capped at 80)
    for ci, col_name in enumerate(df.columns, 1):
        col_vals = df[col_name].dropna().astype(str)
        max_len = max([len(col_name)] + col_vals.map(len).tolist()) if len(col_vals) else len(col_name)
        ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 4, 80)

    if freeze_header:
        ws.freeze_panes = "A2"

    wb.save(output)
    output.seek(0)

    st.markdown("---")
    st.download_button(
        label="⬇️ Télécharger le fichier XLSX",
        data=output,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.markdown("---")
    st.info("👆 Importe un fichier TSV, CSV ou TXT pour commencer.")
    with st.expander("📖 Comment ça marche ?"):
        st.markdown("""
1. **Sélectionne** le séparateur dans la barre latérale (tab par défaut)
2. **Importe** ton fichier
3. **Vérifie** l'aperçu des données
4. **Télécharge** le fichier XLSX formaté
        """)
