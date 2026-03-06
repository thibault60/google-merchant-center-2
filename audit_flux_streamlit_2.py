import streamlit as st
import pandas as pd
import re
from io import BytesIO

# ─── Page config ───
st.set_page_config(
    page_title="Canonical Link Generator",
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
st.markdown('<div class="main-title">🔗 Canonical Link Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Feedonomics — Supprime les paramètres <code>?variant=</code> pour générer l\'attribut <code>canonical_link</code></div>', unsafe_allow_html=True)

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
    st.code("id | link", language=None)
    st.markdown("Le champ `canonical_link` sera généré automatiquement.")

# ─── File upload ───
uploaded_file = st.file_uploader(
    "📂 Importe ton fichier feed (TXT, CSV, TSV)",
    type=["txt", "csv", "tsv"],
    help="Le fichier doit contenir au minimum les colonnes 'id' et 'link'"
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

    if "id" not in df.columns:
        st.warning("⚠️ Colonne `id` introuvable — le fichier sera traité sans IDs.")

    # ─── Clean link values (remove surrounding quotes) ───
    df["link"] = df["link"].str.strip().str.strip('"')
    if "id" in df.columns:
        df["id"] = df["id"].str.strip().str.strip('"')

    # ─── Generate canonical_link ───
    df["canonical_link"] = df["link"].apply(lambda url: re.sub(variant_pattern, "", str(url)) if pd.notna(url) else "")

    # ─── Metrics ───
    total = len(df)
    transformed = (df["canonical_link"] != df["link"]).sum()
    unchanged = total - transformed

    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f'<div class="metric-card"><h3>{total}</h3><p>Produits dans le fichier</p></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="metric-card"><h3>{transformed}</h3><p>URLs transformées</p></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="metric-card"><h3>{unchanged}</h3><p>URLs inchangées</p></div>', unsafe_allow_html=True)

    st.markdown("---")

    # ─── Preview table ───
    st.subheader("📋 Aperçu des résultats")

    display_cols = ["id", "link", "canonical_link"] if "id" in df.columns else ["link", "canonical_link"]
    st.dataframe(
        df[display_cols],
        use_container_width=True,
        height=400,
        column_config={
            "id": st.column_config.TextColumn("ID", width="small"),
            "link": st.column_config.TextColumn("Link (original)", width="large"),
            "canonical_link": st.column_config.TextColumn("Canonical Link (généré)", width="large"),
        }
    )

    # ─── Export buttons ───
    st.markdown("---")
    st.subheader("📥 Export")

    col_a, col_b = st.columns(2)

    # XLSX export
    with col_a:
        output_xlsx = BytesIO()
        with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
            df[display_cols].to_excel(writer, index=False, sheet_name="Canonical Links")
            ws = writer.sheets["Canonical Links"]
            for col_cells in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col_cells)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 80)
        output_xlsx.seek(0)

        st.download_button(
            label="⬇️ Télécharger en XLSX",
            data=output_xlsx,
            file_name="MaxWarehouse_Canonical_Links.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # CSV export
    with col_b:
        csv_data = df[display_cols].to_csv(index=False, sep="\t")
        st.download_button(
            label="⬇️ Télécharger en TSV",
            data=csv_data,
            file_name="MaxWarehouse_Canonical_Links.tsv",
            mime="text/tab-separated-values",
        )

else:
    # ─── Empty state ───
    st.markdown("---")
    st.info("👆 Importe ton fichier feed Feedonomics pour commencer. Le fichier doit contenir les colonnes `id` et `link`.")

    with st.expander("📖 Comment ça marche ?"):
        st.markdown("""
1. **Importe** ton fichier TXT/CSV/TSV exporté depuis Feedonomics
2. **Vérifie** l'aperçu : chaque URL `link` est nettoyée du paramètre `?variant=XXXXX`
3. **Exporte** le résultat en XLSX ou TSV
4. **Réimporte** le fichier dans Feedonomics pour mapper le champ `canonical_link` vers Google Merchant Center
        """)
