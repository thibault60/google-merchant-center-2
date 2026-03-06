"""
MaxWarehouse — Canonical Link Extractor
Reads a TSV/TXT feed export, strips ?variant= params, and outputs an XLSX.

Usage:
    python extract_canonical_links.py input_feed.txt output_canonical.xlsx
"""

import re
import csv
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


def extract_canonical(url):
    """Remove ?variant=XXXXX from a product URL."""
    return re.sub(r'\?variant=\d+', '', url)


def main():
    if len(sys.argv) < 3:
        print("Usage: python extract_canonical_links.py <input.txt> <output.xlsx>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    # Read TSV input
    rows = []
    with open(input_file, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter='\t', quotechar='"')
        for row in reader:
            pid = row.get('id', '').strip('"')
            link = row.get('link', '').strip('"')
            if pid and link:
                rows.append((pid, link, extract_canonical(link)))

    # Create XLSX
    wb = Workbook()
    ws = wb.active
    ws.title = "Canonical Links"

    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    header_fill = PatternFill("solid", fgColor="2D5B2D")
    header_align = Alignment(horizontal="center", vertical="center")
    cell_font = Font(name="Arial", size=10)

    for col, h in enumerate(["id", "link", "canonical_link"], 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    for row_idx, (pid, link, canonical) in enumerate(rows, 2):
        ws.cell(row=row_idx, column=1, value=pid).font = cell_font
        ws.cell(row=row_idx, column=2, value=link).font = cell_font
        ws.cell(row=row_idx, column=3, value=canonical).font = cell_font

    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 95
    ws.column_dimensions['C'].width = 85
    ws.auto_filter.ref = f"A1:C{len(rows) + 1}"

    wb.save(output_file)
    print(f"Done — {len(rows)} products exported to {output_file}")


if __name__ == "__main__":
    main()
