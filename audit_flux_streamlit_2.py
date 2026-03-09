"""
merchant_feed_to_xlsx.py — Extraction du flux Merchant Center vers XLSX
Glisse ton fichier CSV sur GLISSER_FEED_ICI.bat
"""
import sys
import time
import json
import re
from pathlib import Path


def extract_metafields(json_str):
    """Extrait les metafields product_meta en colonnes clé→valeur."""
    try:
        items = json.loads(json_str)
        return {item["namespace"] + "." + item["key"]: item["value"] for item in items if isinstance(item, dict)}
    except Exception:
        return {}


def extract_publications(json_str):
    """Retourne les noms des publications séparés par | ."""
    try:
        items = json.loads(json_str)
        return " | ".join(item["name"] for item in items if isinstance(item, dict) and "name" in item)
    except Exception:
        return json_str


def extract_variant_names(json_str):
    """Retourne les noms de variantes ex: Color: Red | Size: L"""
    try:
        d = json.loads(json_str)
        return " | ".join(f"{k}: {v}" for k, v in d.items())
    except Exception:
        return json_str


def clean_description(text):
    """Supprime le balisage markdown basique des descriptions."""
    if not isinstance(text, str):
        return text
    text = re.sub(r"\* ", "", text)
    text = re.sub(r"\n+", " / ", text)
    return text.strip(" /")


def main():
    if len(sys.argv) < 2:
        input("❌ Utilisation : glisse ton fichier CSV sur GLISSER_FEED_ICI.bat\n\nAppuie sur Entrée pour fermer...")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        input(f"❌ Fichier introuvable : {input_path}\n\nAppuie sur Entrée pour fermer...")
        sys.exit(1)

    output_path = input_path.with_stem(input_path.stem + "_extracted").with_suffix(".xlsx")

    print(f"\n📂 Source  : {input_path.name}  ({input_path.stat().st_size / 1e6:.1f} Mo)")
    print(f"📊 Cible   : {output_path.name}")
    print(f"\n⏳ Lecture du fichier...\n")

    try:
        import pandas as pd
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError as e:
        input(f"\n❌ Module manquant : {e}\nLance : pip install pandas openpyxl\n\nAppuie sur Entrée pour fermer...")
        sys.exit(1)

    t0 = time.time()

    # ── Lecture ──
    df = pd.read_csv(
        input_path,
        dtype=str,
        encoding="utf-8-sig",
        on_bad_lines="skip",
        low_memory=False,
    )
    df.columns = df.columns.str.strip().str.strip('"').str.lower()
    df = df.fillna("")
    print(f"  ✅ {len(df):,} lignes × {len(df.columns)} colonnes lues")

    # ── Colonnes principales à conserver ──
    CORE_COLS = [
        "id", "item_group_id", "sku", "gtin",
        "parent_title", "child_title",
        "brand", "product_type",
        "price", "sale_price", "availability",
        "inventory_quantity", "inventory_management", "inventory_policy",
        "link", "image_link", "additional_image_link",
        "color", "size", "material",
        "weight", "weight_unit", "shipping_weight",
        "taxable", "requires_shipping", "fulfillment_service",
        "published_status",
        "custom_collections_title",
        "smart_collections_title",
        "tags",
    ]

    present_core = [c for c in CORE_COLS if c in df.columns]

    # ── Nettoyage colonnes texte lourdes ──
    for col in ["description", "es_body_html"]:
        if col in df.columns:
            df[col] = df[col].apply(clean_description)

    # ── Publications → noms lisibles ──
    if "publications" in df.columns:
        df["publications_names"] = df["publications"].apply(extract_publications)

    # ── variant_names → lisible ──
    if "variant_names" in df.columns:
        df["variant_names_clean"] = df["variant_names"].apply(extract_variant_names)

    # ── Extraction metafields product_meta ──
    meta_cols_data = []
    if "product_meta" in df.columns:
        print("  🔍 Extraction des metafields product_meta...")
        meta_extracted = df["product_meta"].apply(extract_metafields)
        meta_df = pd.DataFrame(meta_extracted.tolist()).fillna("")
        # Préfixer les colonnes
        meta_df.columns = ["meta." + c for c in meta_df.columns]
        meta_cols_data = list(meta_df.columns)
        df = pd.concat([df.reset_index(drop=True), meta_df.reset_index(drop=True)], axis=1)
        print(f"  ✅ {len(meta_cols_data)} metafields extraits : {', '.join(meta_cols_data[:6])}{'...' if len(meta_cols_data) > 6 else ''}")

    # ── Colonnes ES (traductions espagnol) ──
    es_cols = [c for c in df.columns if c.startswith("es_")]

    # ── Ordre final des colonnes ──
    extra_cols = []
    if "publications_names" in df.columns:
        extra_cols.append("publications_names")
    if "variant_names_clean" in df.columns:
        extra_cols.append("variant_names_clean")
    if "description" in df.columns:
        extra_cols.append("description")

    final_cols = present_core + extra_cols + meta_cols_data + es_cols

    # Dédupliquer en conservant l'ordre
    seen = set()
    final_cols = [c for c in final_cols if not (c in seen or seen.add(c))]

    df_out = df[final_cols]
    print(f"\n  📋 {len(final_cols)} colonnes dans le fichier final")

    # ── Écriture XLSX ──
    print(f"\n⏳ Écriture du fichier Excel ({len(df_out):,} lignes)...")

    CHUNK     = 50_000
    MAX_ROWS  = 1_048_575
    HDR_CLR   = "1e3a8a"
    EVEN_CLR  = "EFF6FF"
    META_CLR  = "fef9c3"   # jaune pâle pour metafields
    ES_CLR    = "f0fdf4"   # vert pâle pour colonnes ES

    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    thin   = Side(style="thin", color="D1D5DB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def get_header_fill(col_name):
        if col_name.startswith("meta."):
            return PatternFill("solid", fgColor="92400e")  # brun
        if col_name.startswith("es_"):
            return PatternFill("solid", fgColor="065f46")  # vert foncé
        return PatternFill("solid", fgColor=HDR_CLR)

    def get_row_fill(col_name, ri):
        if col_name.startswith("meta."):
            return PatternFill("solid", fgColor=META_CLR if ri % 2 == 0 else "fffbeb")
        if col_name.startswith("es_"):
            return PatternFill("solid", fgColor=ES_CLR if ri % 2 == 0 else "f7fef9")
        return PatternFill("solid", fgColor=EVEN_CLR if ri % 2 == 0 else "FFFFFF")

    sheet_index = 1
    total_written = 0
    batch = []

    def flush(rows, name):
        ws = wb.create_sheet(title=name)
        # En-tête
        for ci, col in enumerate(final_cols, 1):
            cell = ws.cell(row=1, column=ci, value=col)
            cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
            cell.fill      = get_header_fill(col)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = border
        ws.row_dimensions[1].height = 32
        ws.freeze_panes = "A2"

        data_font = Font(name="Arial", size=10)
        for ri, row in enumerate(rows, 2):
            for ci, (col, val) in enumerate(zip(final_cols, row), 1):
                cell           = ws.cell(row=ri, column=ci, value=val)
                cell.font      = data_font
                cell.fill      = get_row_fill(col, ri)
                cell.border    = border
                cell.alignment = Alignment(vertical="center", wrap_text=False)

        # Largeurs par type de colonne
        width_map = {
            "id": 14, "item_group_id": 14, "sku": 14, "gtin": 16,
            "parent_title": 45, "child_title": 45,
            "brand": 16, "product_type": 20,
            "price": 10, "sale_price": 10, "availability": 14,
            "link": 60, "image_link": 60, "additional_image_link": 40,
            "color": 14, "size": 12, "material": 16,
            "description": 70, "publications_names": 50,
            "variant_names_clean": 25,
        }
        for ci, col in enumerate(final_cols, 1):
            if col in width_map:
                w = width_map[col]
            elif col.startswith("meta."):
                w = 22
            elif col.startswith("es_"):
                w = 35
            else:
                w = 18
            ws.column_dimensions[get_column_letter(ci)].width = min(w, 80)

    for _, row_data in df_out.iterrows():
        batch.append(list(row_data))
        total_written += 1
        if len(batch) >= MAX_ROWS:
            name = f"Data_{sheet_index}"
            flush(batch, name)
            print(f"  ✅ Feuille '{name}' ({len(batch):,} lignes)")
            sheet_index += 1
            batch = []
        if total_written % 10_000 == 0:
            print(f"  ↳ {total_written:>10,} lignes traitées...", end="\r")

    if batch:
        name = "Data" if sheet_index == 1 else f"Data_{sheet_index}"
        flush(batch, name)

    # ── Feuille légende ──
    ws_legend = wb.create_sheet(title="Légende")
    ws_legend["A1"] = "Couleur"
    ws_legend["B1"] = "Type de colonne"
    for cell in ws_legend["1:1"]:
        cell.font = Font(name="Arial", bold=True, size=11)

    legend = [
        ("1e3a8a", "Attributs Merchant Center principaux"),
        ("92400e", "Metafields produit (product_meta.*)"),
        ("065f46", "Attributs espagnol (es_*)"),
    ]
    for ri, (color, label) in enumerate(legend, 2):
        c = ws_legend.cell(row=ri, column=1, value="")
        c.fill = PatternFill("solid", fgColor=color)
        ws_legend.cell(row=ri, column=2, value=label).font = Font(name="Arial", size=10)
    ws_legend.column_dimensions["A"].width = 6
    ws_legend.column_dimensions["B"].width = 45

    wb.save(output_path)

    elapsed = time.time() - t0
    size_mb = output_path.stat().st_size / 1e6
    print(f"\n\n{'─'*50}")
    print(f"✅ Terminé en {elapsed:.0f}s")
    print(f"   {total_written:,} produits  •  {len(final_cols)} colonnes  •  {size_mb:.1f} Mo")
    print(f"   Fichier : {output_path.name}")
    print(f"{'─'*50}\n")
    input("Appuie sur Entrée pour fermer...")


if __name__ == "__main__":
    main()
