#!/usr/bin/env python3
"""
generate_lookup_tool.py

Regenerates index.html for the SME Compensation Matrix Rate Lookup Tool
from the current version of the SME Excel grid.

Run this any time the underlying SME Excel file changes, then re-upload
the freshly generated index.html to your host (Netlify/Vercel/GitHub
Pages/your own server) to replace the old one.

USAGE
-----
    python3 generate_lookup_tool.py "SME_Upload_Grid_UPDATED.xlsx"

Optional: specify a different output filename
    python3 generate_lookup_tool.py "SME_Upload_Grid_UPDATED.xlsx" -o index.html

REQUIREMENTS
------------
    pip install openpyxl

WHAT IT DOES
------------
1. Reads the Excel file's "SME Compensation Matrix" sheet.
2. Reads the header row to find the columns it needs, so it works with
   either of these two layouts:
       - Grid ID, Company, Category, Product, Occupancy Code, Occupancy Name,
         Recivable Net, Payable Net, ...   (older layout, two occupancy columns)
       - Grid ID, Company, Category, Product, Occupancy, Recivable Net,
         Payable Net, ...                   (current layout, single Occupancy
         column — this is what the live production export uses. A value like
         "Sum Insured - Below 150 Crs" in this column is treated as an
         Occupancy Name and surfaced through the same 3rd dropdown as any
         other occupancy value, so CAR/EAR Sum Insured tiers need no change
         to the page's search logic.)
3. Builds a nested lookup structure: Company -> Product -> Occupancy Name -> [rate rows]
4. Injects that structure as JSON into index_template.html (which must be
   in the same folder as this script) to produce a new, fully self-contained
   index.html.

NOTE: index_template.html must stay in the same folder as this script —
it is the page shell (design, search logic) with a placeholder where the
data gets inserted. Do not edit the placeholder text __DATA_JSON__ inside it.
"""

import argparse
import json
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit(
        "Missing dependency 'openpyxl'.\n"
        "Install it first with:  pip install openpyxl"
    )

SHEET_NAME = "SME Compensation Matrix"
TEMPLATE_FILENAME = "index_template.html"
PLACEHOLDER = "__DATA_JSON__"


def _find_columns(header_row):
    """Map header labels (case-insensitive) to their 0-based column index.
    Supports both the older 'Occupancy Code' + 'Occupancy Name' layout and
    the current single 'Occupancy' column layout."""
    idx = {}
    for i, h in enumerate(header_row):
        if h is None:
            continue
        key = str(h).strip().lower()
        idx[key] = i

    required = ["grid id", "company", "category", "product", "recivable net", "payable net"]
    missing = [r for r in required if r not in idx]
    if missing:
        sys.exit(f"Could not find required column(s) in the header row: {missing}")

    has_split_occupancy = "occupancy code" in idx and "occupancy name" in idx
    has_single_occupancy = "occupancy" in idx

    if not has_split_occupancy and not has_single_occupancy:
        sys.exit(
            "Could not find an occupancy column. Expected either "
            "'Occupancy Code' + 'Occupancy Name', or a single 'Occupancy' column."
        )

    return {
        "grid_id": idx["grid id"],
        "company": idx["company"],
        "category": idx["category"],
        "product": idx["product"],
        "occ_code": idx.get("occupancy code"),
        "occ_name": idx.get("occupancy name"),
        "occupancy": idx.get("occupancy"),
        "recv": idx["recivable net"],
        "pay": idx["payable net"],
        "split_occupancy": has_split_occupancy,
    }


def build_lookup_json(xlsx_path: Path) -> str:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        sys.exit(
            f"Sheet '{SHEET_NAME}' not found in {xlsx_path.name}. "
            f"Sheets present: {wb.sheetnames}"
        )
    ws = wb[SHEET_NAME]

    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    cols = _find_columns(header_row)

    data = {}
    companies_set = set()
    row_count = 0

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        company = row[cols["company"]]
        product = row[cols["product"]]
        if not company or not product:
            continue

        category = row[cols["category"]]
        recv = row[cols["recv"]]
        pay = row[cols["pay"]]

        if cols["split_occupancy"]:
            occ_code = row[cols["occ_code"]]
            occ_name = row[cols["occ_name"]]
        else:
            occ_code = ""
            occ_name = row[cols["occupancy"]]

        companies_set.add(company)
        occ_key = occ_name if occ_name else "(No Occupancy Classification / Standard Rate)"

        data.setdefault(company, {}).setdefault(product, {}).setdefault(occ_key, []).append(
            {
                "code": str(occ_code) if occ_code else "",
                "cat": str(category) if category else "",
                "recv": str(recv) if recv is not None else "",
                "pay": str(pay) if pay is not None else "",
            }
        )
        row_count += 1

    payload = {"companies": sorted(companies_set), "data": data}
    print(f"Read {row_count} data rows across {len(companies_set)} companies.")
    return json.dumps(payload, separators=(",", ":"))


def main():
    parser = argparse.ArgumentParser(description="Regenerate the SME Rate Lookup Tool's index.html")
    parser.add_argument("excel_file", help="Path to the current SME Compensation Matrix .xlsx file")
    parser.add_argument(
        "-o", "--output", default="index.html", help="Output HTML filename (default: index.html)"
    )
    args = parser.parse_args()

    xlsx_path = Path(args.excel_file)
    if not xlsx_path.exists():
        sys.exit(f"File not found: {xlsx_path}")

    script_dir = Path(__file__).parent
    template_path = script_dir / TEMPLATE_FILENAME
    if not template_path.exists():
        sys.exit(
            f"Could not find {TEMPLATE_FILENAME} next to this script.\n"
            f"Expected it at: {template_path}"
        )

    print(f"Reading data from: {xlsx_path}")
    data_json = build_lookup_json(xlsx_path)

    if "</script>" in data_json:
        sys.exit(
            "Safety check failed: embedded data unexpectedly contains a "
            "'</script>' sequence. Please report this before proceeding."
        )

    template_html = template_path.read_text(encoding="utf-8")
    if PLACEHOLDER not in template_html:
        sys.exit(f"Placeholder {PLACEHOLDER} not found in {TEMPLATE_FILENAME}. Has it been edited?")

    final_html = template_html.replace(PLACEHOLDER, data_json)

    out_path = Path(args.output)
    out_path.write_text(final_html, encoding="utf-8")

    size_mb = out_path.stat().st_size / (1024 * 1024)
    print(f"Done. Wrote {out_path} ({size_mb:.2f} MB).")
    print("Upload this file to your host to replace the old version.")


if __name__ == "__main__":
    main()
