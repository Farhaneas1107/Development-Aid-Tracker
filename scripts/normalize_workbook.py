#!/usr/bin/env python3
"""
Normalises Naue_Project_Opportunity_Tracker.xlsx in place:

* Project Opportunity Tracker  - uniform Naue-branded headers, clean alternating rows,
  tier colour coding, accurate wrapped row heights, freeze pane A5, full auto-filter,
  landscape print setup, and cleaned whitespace.
* Opportunity Dashboard & KPIs - one row per programme aligned row-for-row with the
  tracker sheet, tier colour badges, and live Excel formulas for actionable counts and
  portfolio totals.

Run:  python scripts/normalize_workbook.py [--backup]
Then: python scripts/build_data.py
"""

import argparse
import math
import re
import shutil
import sys
from pathlib import Path

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_data as bd  # noqa: E402

XLSX = bd.EXCEL_PATH
BRAND_DARK = "FF283C32"
BRAND_GREEN = "FF2BAC70"
BRAND_SOFT = "FFD6EBE2"
SUBTITLE_BG = "FFEDF7F2"
TEXT_DARK = "FF1F2937"
BORDER_COLOR = "FFCBD5E0"

THIN_SIDE = Side(style="thin", color=BORDER_COLOR)
MED_DARK = Side(style="medium", color=BRAND_DARK)
MED_GREEN = Side(style="medium", color=BRAND_GREEN)
DOUBLE_DARK = Side(style="double", color=BRAND_DARK)

BOX = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
SUBTITLE_BORDER = Border(bottom=MED_GREEN)
TOTAL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=MED_DARK, bottom=DOUBLE_DARK)

BAND = ("FFFFFFFF", "FFF6F8F7")

# (fill_argb, badge_text_argb)
TIER_STYLE = {
    "Very High": ("FFD6EBE2", "FF14532D"),
    "High": ("FFDBEAFE", "FF1E3A8A"),
    "Medium-to-High": ("FFCFFAFE", "FF155E75"),
    "Moderate": ("FFFEF3C7", "FF92400E"),
    "Low-to-Medium": ("FFE0E7FF", "FF3730A3"),
    "Low": ("FFFEE2E2", "FF991B1B"),
    "None": ("FFF1F5F9", "FF475569"),
}
DASH_RATING = {"Moderate": "Medium"}  # dashboard vocabulary
FIRST = 5

# Short dashboard references where automatic extraction would be noisy (keyed by sheet row)
REF_OVERRIDE = {
    12: "P121648 / P171328 (TF 16835)",
    13: "TF 16835 (GZ-MDLF-16163-CW-RFB)",
    14: "P155732 (IBRD 85200 / 87930)",
    15: "P155732 (IBRD 85200 / 87930)",
    16: "P155732 (IBRD 87930)",
    17: "P171426 (Parent P147689)",
    18: "PSD Technical Cooperation (TC)",
    20: "SOL-669-15-000014 (USAID Pipeline)",
    22: "P155732 (IBRD 85200 / 87930)",
    23: "USAID Pipeline (6 Countries)",
    25: "USAID Pipeline (21 Countries)",
    27: "ENI / 2016 / 039-601",
    28: "IsDB General Procurement Notice",
    31: "JICA ODA Schemes",
    33: "FAO Corporate Output Indicators",
    36: "ADB 40540-016 to -020 (MFF 0103)",
    37: "ADB 53382-BAN (MFF 0112)",
    38: "ADB 53260-001 (Loan 4178-BAN)",
    39: "ADB 57296-002 (Loans 4665 / 4666-BAN)",
}


def clean(text):
    if not isinstance(text, str):
        return text
    t = text.replace("\r\n", "\n").replace("\r", "\n").replace("\u00a0", " ")
    lines = [re.sub(r"[ \t]{2,}", " ", ln).strip() for ln in t.split("\n")]
    return "\n".join(lines).strip()


def detect_last_row(ws, first_row=FIRST, ncols=10):
    last = first_row - 1
    for r in range(first_row, ws.max_row + 1):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, ncols + 1)):
            last = r
    return last


def est_height(ws, row, cols, min_h=38.0, max_h=409.5, line_h=14.2, pad=12.0):
    lines_max = 1
    for col in cols:
        v = ws[f"{col}{row}"].value
        if not isinstance(v, str) or v.startswith("="):
            continue
        width = ws.column_dimensions[col].width or 8.43
        cpl = max(4, int(width * 0.88))
        n = sum(max(1, math.ceil(len(seg) / cpl)) for seg in v.split("\n"))
        lines_max = max(lines_max, n)
    return round(max(min_h, min(max_h, lines_max * line_h + pad)), 1)


def ensure_merge(ws, ref):
    existing = {str(rng) for rng in ws.merged_cells.ranges}
    if ref not in existing:
        ws.merge_cells(ref)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.fill = PatternFill("solid", fgColor=BRAND_DARK)
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX


def style_titles(ws, ncols):
    end_col = openpyxl.utils.get_column_letter(ncols)
    ensure_merge(ws, f"A1:{end_col}1")
    ensure_merge(ws, f"A2:{end_col}2")

    ws["A1"].font = Font(name="Calibri", size=15, bold=True, color=BRAND_DARK)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30

    for c in range(1, ncols + 1):
        cell = ws.cell(2, c)
        cell.fill = PatternFill("solid", fgColor=SUBTITLE_BG)
        cell.border = SUBTITLE_BORDER
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color=BRAND_DARK)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 8


def page_setup(ws, title_row=4):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{title_row}:{title_row}"
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.sheet_view.zoomScale = 85
    ws.sheet_view.showGridLines = True
    ws.sheet_properties.tabColor = BRAND_GREEN[2:]


def product_lines(scope, rel, inferred):
    found, _ = bd.extract_products(scope or "", rel or "", "", inferred)
    cat = {p["id"]: p for p in bd.NAUE_PRODUCTS_CATALOG}
    out = []
    for pid in found:
        p = cat.get(pid)
        if p:
            out.append(f'{p["name"]} – {p["category"]}')
    return ", ".join(out) if out else "No direct geosynthetic fit identified"


def ref_short(text, row=None):
    if row in REF_OVERRIDE:
        return REF_OVERRIDE[row]
    codes = bd.extract_tender_codes(text)
    if codes:
        return " / ".join(codes[:2])[:48]
    first = re.split(r"[(;\n]", text)[0].strip()
    return first[:44].rstrip(" ,:-")


def main():
    parser = argparse.ArgumentParser(description="Normalize Naue Excel Opportunity Tracker.")
    parser.add_argument("--backup", action="store_true", help="Write a .backup.xlsx copy before editing.")
    args = parser.parse_args()

    if args.backup:
        backup = XLSX.with_suffix(".backup.xlsx")
        shutil.copy2(XLSX, backup)
        print(f"Backup created: {backup.name}")

    wb = openpyxl.load_workbook(XLSX)
    s1, s2 = wb["Project Opportunity Tracker"], wb["Opportunity Dashboard & KPIs"]
    last = detect_last_row(s1, FIRST, 10)
    n_programs = last - FIRST + 1
    changed = 0

    # Preserve original rating & bespoke text for the initial 7 rows if already present on Sheet 2
    old_rating = {r: s2.cell(r, 5).value for r in range(5, min(12, last + 1))}
    keep = {r: [s2.cell(r, c).value for c in range(1, 12)] for r in range(5, min(12, last + 1))}

    # ---------------- Sheet 1: Project Opportunity Tracker ----------------
    for r in range(FIRST, last + 1):
        for c in range(1, 11):
            cell = s1.cell(r, c)
            new = clean(cell.value)
            if new != cell.value:
                changed += 1
                cell.value = new

    # Remove stray trailing empty columns/rows beyond the active table
    if s1.max_column > 10:
        s1.delete_cols(11, s1.max_column - 10)
    if s1.max_row > last:
        s1.delete_rows(last + 1, s1.max_row - last)

    s1["A2"] = (
        f"Comprehensive Project-Level Review of All {n_programs} Programs Identified Across the Shared "
        "Documents | Technical Scope, Opportunity Counts & Geosynthetic Relevance (Low to Very High)"
    )
    s1.conditional_formatting = ConditionalFormattingList()
    s1.freeze_panes = "A5"
    s1.auto_filter.ref = f"A4:J{last}"

    for col, w in zip("ABCDEFGHIJ", (20, 22, 36, 36, 28, 38, 34, 64, 36, 64)):
        s1.column_dimensions[col].width = w

    style_titles(s1, 10)
    style_header(s1, 4, 10)
    s1.row_dimensions[4].height = 34

    tiers = {}
    for r in range(FIRST, last + 1):
        rel = s1.cell(r, 10).value or ""
        tier, _, _ = bd.normalize_relevance(rel, old_rating.get(r))
        tiers[r] = tier
        tier_fill, _ = TIER_STYLE[tier]
        fill_band = BAND[(r - FIRST) % 2]
        for c in range(1, 11):
            cell = s1.cell(r, c)
            bold = c in (2, 6)
            center = c in (1, 2)
            cell.font = Font(name="Calibri", size=10, bold=bold, color=TEXT_DARK)
            cell.alignment = Alignment(
                horizontal="center" if center else "left",
                vertical="top",
                wrap_text=True,
            )
            cell.border = BOX
            cell.fill = PatternFill("solid", fgColor=tier_fill if c == 10 else fill_band)
        s1.row_dimensions[r].height = est_height(s1, r, "ABCDEFGHIJ", min_h=38.0, line_h=14.2, pad=12.0)

    page_setup(s1)

    # ---------------- Sheet 2: Opportunity Dashboard & KPIs ----------------
    if s2.max_column > 11:
        s2.delete_cols(12, s2.max_column - 11)
    if s2.max_row > last + 1:
        s2.delete_rows(last + 2, s2.max_row - (last + 1))

    for col, w in zip("ABCDEFGHIJK", (28, 24, 42, 28, 18, 16, 18, 16, 16, 18, 56)):
        s2.column_dimensions[col].width = w

    for r in range(FIRST, last + 2):
        for c in range(1, 12):
            s2.cell(r, c).value = None

    for r in range(FIRST, last + 1):
        meta = bd.CURATED_META.get(r, {})
        stats = meta.get("stats", {"total": 1, "high": 0, "medium": 1, "low": 0})
        if r in keep and keep[r][0]:
            vals = list(keep[r])
        else:
            row1 = {c: s1.cell(r, c).value or "" for c in range(1, 11)}
            countries = meta.get("countries") or [row1[2]]
            vals = [
                ref_short(row1[5], r),
                ", ".join(countries)[:60],
                meta.get("shortName") or row1[6][:60],
                " / ".join(meta.get("donors", [])) or row1[7][:40],
                DASH_RATING.get(tiers[r], tiers[r]),
                stats["total"],
                stats["high"],
                stats["medium"],
                stats["low"],
                None,
                product_lines(row1[8], row1[10], meta.get("inferredProducts")),
            ]

        vals[4] = DASH_RATING.get(tiers[r], tiers[r])
        vals[9] = f"=G{r}+H{r}"

        tier_fill, tier_text = TIER_STYLE[tiers[r]]
        fill_band = BAND[(r - FIRST) % 2]

        for c in range(1, 12):
            cell = s2.cell(r, c)
            cell.value = vals[c - 1]
            cell.border = BOX
            if c == 5:
                cell.fill = PatternFill("solid", fgColor=tier_fill)
                cell.font = Font(name="Calibri", size=10, bold=True, color=tier_text)
            else:
                cell.fill = PatternFill("solid", fgColor=fill_band)
                cell.font = Font(
                    name="Calibri",
                    size=10,
                    bold=c in (1, 3, 10),
                    color=BRAND_DARK if c == 10 else TEXT_DARK,
                )
            cell.alignment = Alignment(
                horizontal="center" if 5 <= c <= 10 else "left",
                vertical="center",
                wrap_text=True,
            )
            if 6 <= c <= 10:
                cell.number_format = "#,##0"
        s2.row_dimensions[r].height = est_height(s2, r, "ABCDEK", min_h=32.0, max_h=160.0, line_h=14.0, pad=10.0)

    # Total Portfolio Row
    tr = last + 1
    countries = {
        c
        for r in range(FIRST, last + 1)
        for c in bd.CURATED_META.get(r, {}).get("countries", [])
        if "multi" not in c.lower()
    }
    total_vals = [
        "TOTAL PORTFOLIO",
        f"{len(countries)} Countries + Multi-Country Programs",
        f'=COUNTA(A{FIRST}:A{last})&" Distinct Programs Identified in Shared Documents"',
        "World Bank / ADB / USAID / KfW / Others",
        "Portfolio Mix",
        f"=SUM(F{FIRST}:F{last})",
        f"=SUM(G{FIRST}:G{last})",
        f"=SUM(H{FIRST}:H{last})",
        f"=SUM(I{FIRST}:I{last})",
        f"=SUM(J{FIRST}:J{last})",
        f'=J{tr}&" Actionable Geosynthetic Opportunities out of "&F{tr}&" Total Sub-Projects/Packages ("&TEXT(J{tr}/F{tr},"0.0%")&" Hit Rate)"',
    ]
    for c in range(1, 12):
        cell = s2.cell(tr, c)
        cell.value = total_vals[c - 1]
        cell.border = TOTAL_BORDER
        cell.fill = PatternFill("solid", fgColor=BRAND_SOFT)
        cell.font = Font(name="Calibri", size=10, bold=True, color=BRAND_DARK)
        cell.alignment = Alignment(
            horizontal="center" if 5 <= c <= 10 else "left",
            vertical="center",
            wrap_text=True,
        )
        if 6 <= c <= 10:
            cell.number_format = "#,##0"
    s2.row_dimensions[tr].height = 36

    s2["A2"] = (
        f"Summary of All {n_programs} Identified Project Programs (rows match the tracker sheet) | "
        "Sub-Project Counts, Opportunity Tiers & Primary Naue Product Fit"
    )
    style_titles(s2, 11)
    style_header(s2, 4, 11)
    s2.row_dimensions[4].height = 36
    s2.freeze_panes = "A5"
    s2.auto_filter.ref = f"A4:K{last}"
    page_setup(s2)

    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(XLSX)
    print(f"Normalized {XLSX.name} ({n_programs} programs, rows {FIRST}..{last}); {changed} text cells cleaned.")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
