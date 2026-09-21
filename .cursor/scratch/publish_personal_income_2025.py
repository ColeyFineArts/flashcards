#!/usr/bin/env python3
"""Publish Personal Income 2025 as last year's workbook with 2025 numbers.

Canonical model: Personal_Income_prior.xlsx == Drive Personal Income.xlsx
(1qdewbEnTuc, 132,313 bytes).

This script does NOT rebuild from CSVs. CSV rebuilds drop Century Gothic,
merged cells, Hindman W2, Megan T4, and Refrence Library — that is why
Drive stopped looking like last year.

It takes the already-filled Tax Turbo clone, keeps last year's 11 tabs
in the same order, and appends 2025 Data sources (locks + self-checks).

Not tax advice. Do not re-run apply_checking_answers_2025.py.
"""
from __future__ import annotations

import shutil
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path("/workspace/.cursor/scratch")
PRIOR = ROOT / "Personal_Income_prior.xlsx"
SRC = ROOT / "tax_turbo_parts" / "Personal Income 2025.xlsx"
TURBO = ROOT / "Personal_Income_2025_Tax_Turbo.xlsx"
OUT = ROOT / "tax_turbo_parts" / "Personal Income 2025.xlsx"
DELIVERABLE = ROOT / "tax_turbo_deliverable" / "Personal Income 2025.xlsx"

LAST_YEAR_TABS = [
    "Income",
    "524 Ferdinand Ave, Unit 2",
    "216 N. Oak Park Ave",
    "EPGC LLC",
    "Refrence Library",
    "Art Sales and Purchases",
    "GCM",
    "Hindman W2",
    "Megan T4",
    "Investments",
    "524 Ferdinand Ave, Unit 1",
]

# Tabs that are 2025 packet extras — not in last year's Personal Income.xlsx.
DROP = {
    "524 Home Sale",
    "827 Grove CapEx",
    "Childcare 2441",
    "MERCURY 8291",
    "ASK Mercury",
    "Carriers",
}

CG = "Century Gothic"
NAVY = PatternFill("solid", fgColor="1F3864")
PEACH = PatternFill("solid", fgColor="F7CAAC")
GREEN = PatternFill("solid", fgColor="C6EFCE")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
WHITE = Font(name=CG, size=11, bold=True, color="FFFFFF")
BODY = Font(name=CG, size=10)
THIN = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)
ACCT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'

SOURCES = [
    ("HOW THIS FILE IS MADE", "", "", "", ""),
    (
        "This is last year's Personal Income.xlsx with a 2025 year block filled in. "
        "2023 and 2024 are unchanged. Do not rebuild from CSVs.",
        "",
        "",
        "",
        "",
    ),
    ("Last year file (model)", "Drive Personal Income.xlsx", "1qdewbEnTuc-DXdbqBmwHrKMetGm3I7Zo", "132,313 bytes", "Unchanged 2023|2024"),
    ("This file", "Personal Income 2025.xlsx", "copy of that workbook + 2025 columns", "", "Same 11 tab names and order"),
    ("", "", "", "", ""),
    ("Tab", "Line", "2025 amount", "Source", "Status"),
    ("Income", "Hindman actual I4", "70,618.00", "Monarch paycheck cash (not W-2 Box 1)", "YELLOW — Box 1 still ASK"),
    ("Income", "216 Oak Park I5", "19,500.00", "Joint 0203 Delach + Ava + Joan cash", "LOCKED"),
    ("Income", "524 #2 STR I6", "17,086.94", "Airbnb/VRBO platform net", "LOCKED"),
    ("Income", "GCM Boards I7", "21,500.00", "Form 1099-NEC BOM PAYMEN (personal, not EPGC)", "LOCKED"),
    ("Income", "Art Sales net I8", "23,055.06", "MATCHED Mercury deals with Cost only (Berk lots + seals + mosaics). Sale 6428 and Aquinas books wait on Cost.", "LOCKED"),
    ("216 Oak Park", "RENTAL INCOME", "19,500.00", "Same as Income I5. Vacant Oct + half Nov.", "LOCKED"),
    ("216 Oak Park", "RENTER'S INSURANCE", "42.84/mo × 12 = 514.08", "Lemonade. Same amortize as 2023 $26.75 and 2024 $35.08. Cash $514.00 10/3 Megan Chase (8¢).", "LOCKED green"),
    ("216 Oak Park", "MORTGAGE", "1,226.25 Jan–Oct / 1,177.52 Nov–Dec", "Rocket 8507. Jan–Jun yellow until 8507 statements.", "GREEN Jul–Dec / YELLOW Jan–Jun"),
    ("216 Oak Park", "HOA", "421.53 × 12", "Santa Maria C326 on 8507. Jan–May yellow WAIT 8507.", "GREEN Jun–Dec / YELLOW Jan–May"),
    ("216 Oak Park", "INTERIOR MAINTENANCE", "150 Aug + 11 Dec", "Chris Brennan 8/7 (8507 check 8/18) + Ace keys 12/8 Megan Chase", "LOCKED"),
    ("216 Oak Park", "INTERIOR REPAIR", "675 Dec", "Joan painting/repairs 12/1 8507 Zelle", "LOCKED"),
    ("216 Oak Park", "MOVE OUT FEE", "300 Oct", "Imelda empty-apt. User 9/30; cash 10/1 joint 0203", "LOCKED"),
    ("216 Oak Park", "Smoke detector Amazon", "0", "No labeled charge", "ASK"),
    ("524 Unit 2", "INSURANCE", "State Farm HOME share 50/50", "Travelers was prior 524 home leftover $0. Geico is cars (personal). Lemonade is 216, not 524.", "LOCKED"),
    ("524 Unit 1", "INSURANCE", "State Farm HOME share 50/50", "Same carrier lock as Unit 2. Residence through sale 12/18.", "LOCKED"),
    ("524 Unit 1", "MORTGAGE", "US Bank 09422 monthly P+I", "Joint 0203. Not the $50k extra principal on 9922.", "LOCKED"),
    ("GCM", "2025 board", "4,100 + 9,200 + 4,100 + 4,100 = 21,500", "Same 1099 as Income I7. Not EPGC Consultant.", "LOCKED"),
    ("EPGC LLC", "Art Sales", "14,000 Jan + 136,000 Feb = 150,000", "Mercury 8291 MATCHED (Berk + mosaics + Aquinas cash). Do not dump EOEB/L5 remainder.", "LOCKED"),
    ("EPGC LLC", "Consultant June", "13,595.00", "David Aaron Limited 6/24. User: consultant fee, not a sale.", "LOCKED"),
    ("EPGC LLC", "Consultant Fees", "334.17 Jul + 658.63 Oct = 992.80", "Wise expertise write-ups", "LOCKED"),
    ("Art Sales tab", "Seals / lots / mosaics", "see rows 70–74", "Same objects as last year's Art Sales layout", "LOCKED"),
    ("Investments", "Coinbase ACH", "net -51,000.00", "Mercury 8291. Not EPGC.", "LOCKED"),
    ("Carriers (note)", "Lemonade", "216 Oak Park #1Z", "User said 216 N Grove — that is this unit, not 827 Grove.", "LOCKED"),
    ("Carriers (note)", "Travelers / Geico", "524 Ferdinand home / cars", "Not 216. Geico credit $540.89 personal.", "LOCKED"),
    ("", "", "", "", ""),
    ("DO NOT", "", "", "", ""),
    ("Rebuild this workbook from CSVs", "That drops Century Gothic, merged cells, column widths, Hindman W2, Megan T4, Refrence Library", "", "", ""),
    ("Dump EOEB $130,270 or L5 $87,396.32 onto P&L", "Mixed art + consultant. Invoice split still ASK.", "", "", ""),
    ("Re-run apply_checking_answers_2025.py", "Would scramble locked inflows.", "", "", ""),
    ("", "", "", "", ""),
    ("SELF-CHECK (formulas vs locked amounts)", "Sheet value", "Income / expected", "Match?", "If FAIL, the 2025 block was overwritten"),
]


def add_data_sources(wb) -> Worksheet:
    if "2025 Data sources" in wb.sheetnames:
        del wb["2025 Data sources"]
    ws = wb.create_sheet("2025 Data sources")
    ws.sheet_properties.tabColor = "1F3864"
    headers_done = False
    for r, row in enumerate(SOURCES, start=1):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(r, c, val)
            cell.font = BODY
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.border = THIN
        if row[0] == "HOW THIS FILE IS MADE":
            for c in range(1, 6):
                ws.cell(r, c).fill = NAVY
                ws.cell(r, c).font = WHITE
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        elif row[0] == "Tab":
            headers_done = True
            for c in range(1, 6):
                ws.cell(r, c).fill = NAVY
                ws.cell(r, c).font = WHITE
        elif row[0] == "DO NOT":
            for c in range(1, 6):
                ws.cell(r, c).fill = NAVY
                ws.cell(r, c).font = WHITE
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        elif row[0].startswith("SELF-CHECK"):
            for c in range(1, 6):
                ws.cell(r, c).fill = NAVY
                ws.cell(r, c).font = WHITE
        elif headers_done and row[0] and row[4] == "LOCKED":
            ws.cell(r, 5).fill = GREEN
        elif headers_done and row[4].startswith("LOCKED"):
            ws.cell(r, 5).fill = GREEN
        elif headers_done and ("YELLOW" in row[4] or row[4] == "ASK"):
            ws.cell(r, 5).fill = YELLOW
        elif row[0] and not row[1] and r == 2:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
            ws.cell(r, 1).fill = PEACH
            ws.row_dimensions[r].height = 36

    checks = [
        ("216 rent vs Income I5", "='216 N. Oak Park Ave'!AO5", "=Income!I5"),
        ("524 #2 Monarch months vs $13,879.55", "='524 Ferdinand Ave, Unit 2'!AO5", 13879.55),
        ("Income I6 platform vs $17,086.94", "=Income!I6", 17086.94),
        ("GCM 2025 vs Income I7", "=GCM!E26", "=Income!I7"),
        ("216 Lemonade vs $514.08", "='216 N. Oak Park Ave'!AO14", 514.08),
        ("EPGC Art Sales vs $150,000", "='EPGC LLC'!N53", 150000),
        ("EPGC Consultant vs $13,595", "='EPGC LLC'!N54", 13595),
    ]
    start = len(SOURCES) + 1
    for i, (label, left, right) in enumerate(checks):
        r = start + i
        ws.cell(r, 1, label).font = BODY
        ws.cell(r, 1).border = THIN
        c2 = ws.cell(r, 2, left)
        c2.font = BODY
        c2.number_format = ACCT
        c2.border = THIN
        c3 = ws.cell(r, 3, right)
        c3.font = BODY
        c3.number_format = ACCT
        c3.border = THIN
        c4 = ws.cell(r, 4, f'=IF(ROUND(B{r},2)=ROUND(C{r},2),"MATCH","FAIL")')
        c4.font = BODY
        c4.border = THIN
        ws.cell(r, 5, "").border = THIN

    gap_r = start + len(checks)
    ws.cell(
        gap_r,
        1,
        "524 #2 STR: Income I6 is Airbnb+VRBO platform net $17,086.94 (CPA). "
        "Unit 2 months are Monarch cash Jun–Nov $13,879.55. Gap $3,207.39 = Apr/May "
        "platform payouts not in Monarch. Do not invent those months. Not a FAIL.",
    )
    ws.merge_cells(start_row=gap_r, start_column=1, end_row=gap_r, end_column=5)
    ws.cell(gap_r, 1).font = BODY
    ws.cell(gap_r, 1).fill = YELLOW
    ws.cell(gap_r, 1).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[gap_r].height = 40

    note_r = gap_r + 1
    ws.cell(
        note_r,
        1,
        "Color key: peach = entered like last year. Green = locked from bank or user. Yellow = WAIT (statement or Cost). Not tax advice.",
    )
    ws.merge_cells(start_row=note_r, start_column=1, end_row=note_r, end_column=5)
    ws.cell(note_r, 1).font = BODY
    ws.cell(note_r, 1).fill = PEACH

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 42
    ws.column_dimensions["D"].width = 55
    ws.column_dimensions["E"].width = 32
    ws.freeze_panes = "A7"
    ws.row_dimensions[1].height = 22
    return ws


def restyle_gcm_like_last_year(ws: Worksheet) -> None:
    """Make the 2025 GCM block use the same fonts as 2024 on this tab."""
    for dr in range(0, 7):
        for c in range(1, 6):
            src = ws.cell(15 + dr, c)
            dest = ws.cell(22 + dr, c)
            dest.font = copy(src.font)
            dest.fill = copy(src.fill)
            dest.alignment = copy(src.alignment)
            dest.border = copy(src.border)
            dest.number_format = src.number_format
    for addr in ("B24", "C24", "D24", "E24"):
        ws[addr].number_format = ACCT
        ws[addr].fill = PEACH
    merges = {str(r) for r in ws.merged_cells.ranges}
    if "A22:E22" not in merges:
        ws.merge_cells("A22:E22")
    ws["A22"] = "GCM Boards 2025"


def finish_year_titles(wb) -> None:
    """Last year's Unit 2 merges C1:O1 and P1:AB1. Add the 2025 block the same way."""
    u2 = wb["524 Ferdinand Ave, Unit 2"]
    merges = {str(r) for r in u2.merged_cells.ranges}
    if "AC1:AO1" not in merges:
        u2.merge_cells("AC1:AO1")
    u2["AC1"] = u2["C1"].value
    u2["AC1"].font = copy(u2["C1"].font)
    u2["AC1"].alignment = copy(u2["C1"].alignment)

    u1 = wb["524 Ferdinand Ave, Unit 1"]
    u1.sheet_state = "visible"
    merges = {str(r) for r in u1.merged_cells.ranges}
    if "P1:AB1" not in merges:
        u1.merge_cells("P1:AB1")
    if u1["P1"].value in (None, ""):
        u1["P1"] = u1["C1"].value
    u1["P1"].font = copy(u1["C1"].font)
    u1["P1"].alignment = copy(u1["C1"].alignment)

    try:
        wb.defined_names.clear()
    except Exception:
        pass
    wb._external_links = []


def mark_income_locks(ws: Worksheet) -> None:
    ws["I5"].fill = GREEN
    ws["I6"].fill = GREEN
    ws["I7"].fill = GREEN
    ws["I4"].fill = YELLOW
    ws["I8"].fill = YELLOW
    ws["A13"] = (
        "Open the 2025 Data sources tab to see where each 2025 number came from "
        "and whether the self-checks MATCH. This file is last year's workbook "
        "with 2025 filled. Not tax advice."
    )
    ws["A13"].font = Font(name=CG, size=9, italic=True)
    ws["A13"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[13].height = 32


def main() -> None:
    src = SRC if SRC.exists() else TURBO
    if not src.exists():
        raise SystemExit(f"missing {src}")
    if src.resolve() != OUT.resolve():
        shutil.copy2(src, OUT)
    wb = load_workbook(OUT)
    for name in list(wb.sheetnames):
        if name in DROP:
            del wb[name]
    for i, name in enumerate(LAST_YEAR_TABS):
        if name not in wb.sheetnames:
            raise SystemExit(f"missing last-year tab {name}")
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))
    restyle_gcm_like_last_year(wb["GCM"])
    finish_year_titles(wb)
    mark_income_locks(wb["Income"])
    add_data_sources(wb)
    DELIVERABLE.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    import sys

    sys.path.insert(0, str(ROOT))
    from assert_personal_income_2025 import assert_workbook
    from sanitize_xlsx_for_sheets import assert_formulas_are_real, sanitize_xlsx

    sanitize_xlsx(OUT)
    shutil.copy2(OUT, DELIVERABLE)
    shutil.copy2(OUT, TURBO)
    assert_workbook(OUT)
    assert_formulas_are_real(OUT)
    print("wrote", OUT, OUT.stat().st_size, "bytes")


if __name__ == "__main__":
    main()
