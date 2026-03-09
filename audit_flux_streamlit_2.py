import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── Page config ───
st.set_page_config(
    page_title="Canonical & SKU Extractor",
    page_icon="🔗",
    layout="wide",
)

# ─── Custom CSS ───
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&display=swap');

    .block-container { max-width: 1100px; padding-top: 2rem; }
    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

    .main-title {
        font-size: 2rem;
        font-weight: 700;
        color: #1a1a2e;
        margin-bottom: 0.2rem;
    }
    .subtitle {
        font-size: 1rem;
        color: #6b7280;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #f0fdf4, #dcfce7);
        border: 1px solid #bbf7d0;
        border-radius: 12px;
        padding: 1.2rem;
        text-align: center;
    }
    .metric-card h3 {
        font-size: 2rem;
        color: #166534;
        margin: 0;
    }
    .metric-card p {
        font-size: 0.85rem;
        color: #4b5563;
        margin: 0;
    }
    .stDownloadButton > button {
        background-color: #166534 !important;
        color: white !important;
        border-radius: 8px !important;
        padding: 0.6rem 2rem !important;
        font-weight: 600 !important;
    }
</style>
""", unsafe_allow_html=True)

# ─── Header ───
st.markdown('<div class="main-title">🔗 Canonical & SKU Extractor</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Feedonomics — Extrait les SKUs, titres et génère l\'attribut <code>canonical_link</code> en supprimant <code>?variant=</code></div>', unsafe_allow_html=True)

# ─── Sidebar config ───
with st.sidebar:
    st.header("⚙️ Configuration")
    variant_pattern = st.text_input(
        "Regex pattern à supprimer",
        value=r"\?variant=\d+",
        help="Expression régulière pour identifier et supprimer le paramètre variant des URLs"
    )
    separator = st.selectbox(
        "Séparateur du fichier source",
        options=["Tab (\\t)", "Virgule (,)", "Point-virgule (;)"],
        index=0,
    )
    sep_map = {"Tab (\\t)": "\t", "Virgule (,)": ",", "Point-virgule (;)": ";"}
    sep_char = sep_map[separator]

    st.divider()
    st.markdown("**Colonnes attendues :**")
    st.code("link | title | sku", language=None)
    st.markdown("Le champ `canonical_link` sera généré automatiquement.")

# ─── File upload ───
uploaded_file = st.file_uploader(
    "📂 Importe ton fichier feed (TXT, CSV, TSV)",
    type=["txt", "csv", "tsv"],
    help="Le fichier doit contenir au minimum les colonnes 'link', 'title' et 'sku'"
)

if uploaded_file is not None:
    # ─── Read file ───
    try:
        df = pd.read_csv(uploaded_file, sep=sep_char, dtype=str)
        df.columns = df.columns.str.strip().str.strip('"').str.lower()
    except Exception as e:
        st.error(f"Erreur de lecture du fichier : {e}")
        st.stop()

    # ─── Validate columns ───
    if "link" not in df.columns:
        st.error("❌ Colonne `link` introuvable. Colonnes détectées : " + ", ".join(df.columns))
        st.stop()

    # ─── Clean values ───
    for col in df.columns:
        df[col] = df[col].str.strip().str.strip('"')

    # ─── Generate canonical_link ───
    df["canonical_link"] = df["link"].apply(
        lambda url: re.sub(variant_pattern, "", str(url)) if pd.notna(url) else ""
    )

    # ─── Metrics ───
    total = len(df)
    transformed = (df["canonical_link"] != df["link"]).sum()
    has_sku = "sku" in df.columns
    has_title = "title" in df.columns
    unique_skus = df["sku"].nunique() if has_sku else 0

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f'<div class="metric-card"><h3>{total}</h3><p>Produits dans le fichier</p></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="metric-card"><h3>{unique_skus}</h3><p>SKUs uniques</p></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="metric-card"><h3>{transformed}</h3><p>URLs transformées</p></div>', unsafe_allow_html=True)

    st.markdown("---")

    # ─── Build display columns ───
    display_cols = []
    col_config = {}
    if has_sku:
        display_cols.append("sku")
        col_config["sku"] = st.column_config.TextColumn("SKU", width="small")
    if has_title:
        display_cols.append("title")
        col_config["title"] = st.column_config.TextColumn("Titre du produit", width="large")
    display_cols += ["link", "canonical_link"]
    col_config["link"] = st.column_config.TextColumn("Link (original)", width="large")
    col_config["canonical_link"] = st.column_config.TextColumn("Canonical Link (généré)", width="large")

    # ─── Preview table ───
    st.subheader("📋 Aperçu des résultats")
    st.dataframe(
        df[display_cols],
        use_container_width=True,
        height=400,
        column_config=col_config,
    )

    # ─── Export ───
    st.markdown("---")
    st.subheader("📥 Export")
    col_a, col_b = st.columns(2)

    # ── XLSX export with formatting ──
    with col_a:
        output_xlsx = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "SKU & Canonical Links"

        # Header style
        header_fill = PatternFill("solid", fgColor="166534")
        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin = Side(style="thin", color="BBBBBB")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        headers = [c.upper() for c in display_cols]
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=ci, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = border

        ws.row_dimensions[1].height = 30

        # Data rows
        row_fill_even = PatternFill("solid", fgColor="F0FDF4")
        data_font = Font(name="Arial", size=10)
        for ri, row in enumerate(df[display_cols].itertuples(index=False), 2):
            fill = row_fill_even if ri % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
            for ci, val in enumerate(row, 1):
                cell = ws.cell(row=ri, column=ci, value=str(val) if pd.notna(val) else "")
                cell.font = data_font
                cell.fill = fill
                cell.border = border
                cell.alignment = Alignment(vertical="center", wrap_text=False)

        # Column widths
        col_widths = {"sku": 16, "title": 55, "link": 65, "canonical_link": 65}
        for ci, col_name in enumerate(display_cols, 1):
            ws.column_dimensions[get_column_letter(ci)].width = col_widths.get(col_name, 20)

        # Freeze header row
        ws.freeze_panes = "A2"

        wb.save(output_xlsx)
        output_xlsx.seek(0)

        st.download_button(
            label="⬇️ Télécharger en XLSX",
            data=output_xlsx,
            file_name="MaxWarehouse_SKU_Canonical.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ── TSV export ──
    with col_b:
        csv_data = df[display_cols].to_csv(index=False, sep="\t")
        st.download_button(
            label="⬇️ Télécharger en TSV",
            data=csv_data,
            file_name="MaxWarehouse_SKU_Canonical.tsv",
            mime="text/tab-separated-values",
        )

else:
    st.markdown("---")
    st.info("👆 Importe ton fichier feed Feedonomics pour commencer. Le fichier doit contenir les colonnes `link`, `title` et `sku`.")

    with st.expander("📖 Comment ça marche ?"):
        st.markdown("""
1. **Importe** ton fichier TXT/CSV/TSV exporté depuis Feedonomics
2. **Vérifie** l'aperçu : SKU, titre, lien original et canonical link généré
3. **Exporte** le résultat en XLSX formaté ou TSV
4. **Réimporte** dans Feedonomics pour mapper `canonical_link` vers Google Merchant Center
        """)
