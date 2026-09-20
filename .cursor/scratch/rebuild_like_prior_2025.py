#!/usr/bin/env python3
"""Rebuild 2025 Tax Turbo on last year's Personal Income.xlsx look.

Copies Personal_Income_prior.xlsx (2023 | 2024 columns, Century Gothic,
MORTGAGE / HOA / monthly rent) and adds a 2025 year block.

216 Oak Park 2025:
  RENTAL INCOME from joint 0203 cash (Delach + Ava + Joan)
  MORTGAGE = Rocket from Checking 8507 (Jul–Dec known; Jan–Jun WAIT 8507)
  HOA = Santa Maria C326 from 8507 (Jun–Dec known; Jan–May WAIT 8507)
  INTERIOR Dec $675 Joan repair (confirmed)
  RENTER'S INSURANCE Oct Lemonade $514 LOCKED 216 LTR (cash 10/3 Megan Chase)

Not tax advice. Do not re-run apply_checking_answers_2025.py.
"""
from __future__ import annotations

import shutil
from copy import copy
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path("/workspace/.cursor/scratch")
PRIOR = ROOT / "Personal_Income_prior.xlsx"
TURBO = ROOT / "Personal_Income_2025_Tax_Turbo.xlsx"
CLEAN_EXTRA = Path("/tmp/Personal_Income_2025_Tax_Turbo_clean.xlsx")
OUT = ROOT / "Personal_Income_2025_Tax_Turbo.xlsx"
PARTS = ROOT / "tax_turbo_parts"
DELIVERABLE = ROOT / "tax_turbo_deliverable" / "Personal_Income_2025_Tax_Turbo.xlsx"

MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEPT", "OCT", "NOV", "DEC"]
PEACH = PatternFill("solid", fgColor="F7CAAC")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GREEN = PatternFill("solid", fgColor="C6EFCE")
NAVY = PatternFill("solid", fgColor="1F3864")
ORANGE = PatternFill("solid", fgColor="833C0B")
ACCT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
CG = "Century Gothic"


def D(x) -> Decimal:
    return Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def money(x) -> float:
    return float(D(x))


def copy_style(src, dest) -> None:
    if src.has_style:
        dest.font = copy(src.font)
        dest.fill = copy(src.fill)
        dest.border = copy(src.border)
        dest.alignment = copy(src.alignment)
        dest.number_format = src.number_format
        dest.protection = copy(src.protection)


def last_used_row(ws: Worksheet, scan_cols: int = 28) -> int:
    for r in range(ws.max_row, 0, -1):
        for c in range(1, scan_cols + 1):
            v = ws.cell(r, c).value
            if v not in (None, ""):
                return r
    return 1


def trim_sheet(ws: Worksheet, pad: int = 2) -> None:
    used = last_used_row(ws)
    cap = used + pad
    if ws.max_row > cap:
        ws.delete_rows(cap + 1, ws.max_row - cap)


def add_year_block(ws: Worksheet, src_start: int, dest_start: int, year: int, last_row: int) -> None:
    """Copy a 13-col year block (12 months + total) and retarget formulas."""
    n = 13
    for r in range(1, last_row + 1):
        for i in range(n):
            src = ws.cell(r, src_start + i)
            dest = ws.cell(r, dest_start + i)
            copy_style(src, dest)
            val = src.value
            if isinstance(val, str) and val.startswith("="):
                dest.value = Translator(val, origin=src.coordinate).translate_formula(dest.coordinate)
            elif r in (2, 8) and i < 12:
                dest.value = MONTHS[i]
            elif r in (2, 8) and i == 12:
                dest.value = f"{year} YR TOTAL" if r == 2 else "YR TOTAL"
            elif isinstance(val, (int, float, Decimal)):
                dest.value = 0
            else:
                dest.value = val
    for i, m in enumerate(MONTHS):
        for r in (2, 8):
            cell = ws.cell(r, dest_start + i)
            if ws.cell(r, src_start + i).value:
                cell.value = m
    ws.cell(2, dest_start + 12).value = f"{year} YR TOTAL"
    if ws.cell(8, src_start + 12).value:
        ws.cell(8, dest_start + 12).value = "YR TOTAL"


def write_months(ws: Worksheet, row: int, start_col: int, values, fill=PEACH, zero_fill=None, wait_idx=None):
    wait_idx = set(wait_idx or [])
    for i, v in enumerate(values):
        cell = ws.cell(row, start_col + i)
        cell.value = money(v)
        cell.number_format = ACCT
        if i in wait_idx:
            cell.fill = YELLOW
        elif D(v) != 0:
            cell.fill = fill
        elif zero_fill is not None:
            cell.fill = zero_fill


def col_widths_from(src_ws: Worksheet, dest_ws: Worksheet, src_start: int, dest_start: int, n: int = 13) -> None:
    for i in range(n):
        letter_src = get_column_letter(src_start + i)
        letter_dest = get_column_letter(dest_start + i)
        dim = src_ws.column_dimensions[letter_src]
        dest_ws.column_dimensions[letter_dest].width = dim.width or 10.6


# --- 2025 216 Oak Park (cash-basis) ---
OAK_RENT = [1950, 1950, 1950, 1950, 1950, 1950, 1450, 2450, 0, 0, 0, 3900]
# Rocket 8507: missing Jan–Jun in Monarch (Jan–May statements not in export; June 8507 has no Rocket)
OAK_MORTGAGE = [1226.25, 1226.25, 1226.25, 1226.25, 1226.25, 1226.25, 1226.25, 1226.25, 1226.25, 1226.25, 1177.52, 1177.52]
OAK_MORTGAGE_WAIT = [0, 1, 2, 3, 4, 5]
# HOA Santa Maria: Jun–Dec on 8507
OAK_HOA = [421.53, 421.53, 421.53, 421.53, 421.53, 421.53, 421.53, 421.53, 421.53, 421.53, 421.53, 421.53]
OAK_HOA_WAIT = [0, 1, 2, 3, 4]
OAK_INTERIOR = [0, 0, 0, 0, 0, 0, 0, 150.0, 0, 0, 0, 11.0]  # Brennan Aug + Ace keys Dec
OAK_REPAIR = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 675.0]  # Joan painting + repairs 12/1
OAK_MOVEOUT = [0, 0, 0, 0, 0, 0, 0, 0, 0, 300.0, 0, 0]  # Imelda empty-apt cash 10/1
OAK_INS = [42.84] * 12  # user $42.84/mo; Lemonade cash $514 10/3 amortized like 2023/2024

# 524 Unit 2 2025 (from turbo; yard split back onto YARD SERVICE like 2024)
U2_RENT = [0, 0, 0, 0, 0, 3479.60, 3180.81, 1757.50, 2690.16, 1261.64, 1509.84, 0]
U2_INS = [0, 78.77, 78.76, 78.76, 78.76, 78.76, 78.76, 80.19, 0, 0, 0, 0]
U2_INTERIOR = [7.46, 0, 52.78, 131.40, 0, 0, 0, 0, 0, 45.60, 12.63, 0]
U2_EXTERIOR = [0, 66.57, 700.77, 39.58, 0, 0, 150.06, 16.54, 34.87, 76.24, 32.33, 0]
U2_WATER = [0, 100.56, 0, 90.68, 0, 112.84, 0, 137.64, 0, 132.77, 0, 0]
U2_GAS = [39.70, 31.37, 32.09, 26.31, 25.41, 24.17, 23.53, 22.70, 22.71, 22.72, 0, 53.71]
U2_ELEC = [59.04, 57.47, 49.15, 54.52, 49.73, 57.54, 109, 123.54, 106.93, 62.88, 46.47, 52.88]
U2_INT = [45, 45, 45, 45, 50, 50, 50, 50, 50, 50, 50, 0]
U2_EQUIP = [0, 0, 73.24, 0, 0, 0, 0, 0, 82.27, 0, 0, 0]
U2_YARD = [0, 0, 60.95, 60.95, 445.20, 60.95, 60.95, 60.95, 249.00, 60.95, 0, 0]
U2_SUPPLIES = [110.80, 0, 92.96, 105.25, 0, 68.73, 128.37, 176.17, 108.93, 0, 0, 0]

U1_INS = [0, 78.76, 78.75, 78.75, 78.75, 78.75, 78.75, 80.19, 0, 0, 0, 0]
U1_EXTERIOR = [0, 66.56, 688.60, 39.57, 0, 0, 50.99, 0, 34.86, 68.48, 32.33, 0]
U1_WATER = [0, 100.55, 0, 90.67, 0, 112.83, 0, 137.63, 0, 132.76, 0, 0]
U1_GAS = [98.57, 133.33, 115.82, 77.80, 79.95, 48.08, 0, 79.33, 36.25, 38.22, 0, 139.35]
U1_ELEC = [132.94, 216.44, 184.09, 60.95, 83.76, 69.69, 164.88, 203.43, 187.63, 139.78, 71.29, 50.38]
# 524 USB loan 09422 regular monthly P+I from joint 0203 (not $50k extra principal on 9922)
U1_MORTGAGE = [4287.51, 4287.51, 4287.51, 4287.51, 4287.51, 4287.51, 4287.51, 5089.77, 5089.77, 4438.69, 4438.69, 4377.54]

# MATCHED Mercury 8291 art-sale cash only (Berk $14k Jan + Berk $30k/mosaics $105k Feb).
# Do not dump EOEB/L5/Coinbase/Newstar here — see apply_mercury_2025.py.
EPGC_BIZ = [14000, 135000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
EPGC_AD = [0, 0, 755.98, 0, 199, 0, 0, 0, 0, 0, 0, 0]
EPGC_TRAVEL = [0, 20, 0, 0, 20, 0, 20, 20, 0, 20, 53, 0]
EPGC_CELL = [51.12, 93.26, 98.26, 98.26, 98.26, 98.26, 98.23, 98.34, 98.34, 98.39, 98.39, 97.91]
EPGC_SOFT = [22.17, 22.17, 22.17, 22.17, 22.17, 22.17, 22.17, 2.17, 57.17, 37.17, 17.17, 57.17]
EPGC_TECH = [0, 0, 0, 20, 192, 0, 0, 0, 0, 0, 0, 0]
EPGC_LIB = [153.44, 78.47, 0, 63.69, 0, 0, 0, 0, 0, 0, 0, 0]


def fill_216(ws: Worksheet) -> None:
    ws["B4"] = "OTHER INCOME"
    ws["B5"] = "RENTAL INCOME"
    ws["B4"].font = Font(name=CG, size=10)
    ws["B5"].font = Font(name=CG, size=10)
    add_year_block(ws, src_start=16, dest_start=29, year=2025, last_row=44)
    col_widths_from(ws, ws, 16, 29)
    ac = 29
    write_months(ws, 5, ac, OAK_RENT)
    write_months(ws, 14, ac, OAK_INS, fill=GREEN)  # $42.84/mo amortized Lemonade
    write_months(ws, 15, ac, OAK_INTERIOR)
    write_months(ws, 16, ac, OAK_REPAIR)
    write_months(ws, 35, ac, OAK_MORTGAGE, wait_idx=OAK_MORTGAGE_WAIT)
    write_months(ws, 36, ac, OAK_HOA, wait_idx=OAK_HOA_WAIT)
    write_months(ws, 38, ac, OAK_MOVEOUT)
    # Ask-Megan checklist → 2025 status
    ws["B60"] = "2025 (this column)"
    ws["B61"] = "Rent — filled from joint 0203 (Delach $1,450 + Ava $500; Joan $1,950 Dec + prepaid Jan). Vacant Oct + half Nov. Half-Nov $975 WAIT 8507."
    ws["B62"] = "Mortgage — user $1,226.25/mo. Cash 8507 Jul–Oct $1,226.25 / Nov–Dec $1,177.52 (green). Jan–Jun $1,226.25 yellow WAIT 8507."
    ws["B63"] = "HOA — user $421.53/mo Santa Maria C326. Jun–Dec cash 8507 green. Jan–May yellow WAIT 8507."
    ws["B64"] = "Insurance — $42.84/mo × 12 = $514.08 amortized like 2023/2024. Lemonade cash $514.00 10/3 Megan Chase. Not Travelers/Geico/SF."
    ws["B65"] = "Interior — Brennan $150 Aug (8507 check 8/18); Ace keys $11 Dec. Joan $675 painting → INTERIOR REPAIR Dec."
    ws["B66"] = "Move-out — Imelda $300 empty-apt (user 9/30; cash 10/1) MOVE OUT FEE October. Other Imelda $150s not on 216."
    ws["B67"] = "Smoke detector Amazon ASK. Sch E: deduct Rocket 1098 interest, not full P+I."
    ws["B68"] = "Not tax advice."
    for r in range(60, 69):
        ws.cell(r, 2).alignment = Alignment(wrap_text=True)
    ws.row_dimensions[61].height = 32
    ws.row_dimensions[62].height = 32


def fill_unit2(ws: Worksheet) -> None:
    add_year_block(ws, src_start=16, dest_start=29, year=2025, last_row=37)
    col_widths_from(ws, ws, 16, 29)
    ac = 29
    write_months(ws, 5, ac, U2_RENT)
    write_months(ws, 14, ac, U2_INS)
    write_months(ws, 15, ac, U2_INTERIOR)
    write_months(ws, 16, ac, U2_EXTERIOR)
    write_months(ws, 17, ac, U2_WATER)
    write_months(ws, 18, ac, U2_GAS)
    write_months(ws, 19, ac, U2_ELEC)
    write_months(ws, 20, ac, U2_INT)
    write_months(ws, 24, ac, U2_EQUIP)
    write_months(ws, 33, ac, U2_YARD)
    write_months(ws, 34, ac, U2_SUPPLIES)
    ws["B34"] = "SUPPLIES (Grove Collaborative)"
    ws["B34"].font = Font(name=CG, size=10)
    ws["A39"] = (
        "2025: platform STR net on RENTAL INCOME (Jun–Nov). TruGreen/Alsip on YARD SERVICE (same as 2024). "
        "Grove Collaborative on SUPPLIES (≠ 827 N Grove the house). "
        "INSURANCE LOCK 2026-09-20: Travelers was prior 524 home; Geico was the cars. "
        "Only State Farm HOME share on this INSURANCE line (50/50 U1/U2). Auto + Geico $540.89 credit personal, not EPGC. "
        "Travelers leftover Jan 2025 is 524 home ($0 on cards/Monarch). ELECTRIC yellow — confirm ComEd with Megan. Not tax advice."
    )


def fill_unit1(ws: Worksheet) -> None:
    # 2023 only in prior file — add 2025 as the second year block (cols P–AB)
    add_year_block(ws, src_start=3, dest_start=16, year=2025, last_row=43)
    col_widths_from(ws, ws, 3, 16)
    p = 16
    write_months(ws, 14, p, U1_INS)
    write_months(ws, 16, p, U1_EXTERIOR)
    write_months(ws, 21, p, U1_WATER)
    write_months(ws, 22, p, U1_GAS)
    write_months(ws, 23, p, U1_ELEC)
    write_months(ws, 34, p, U1_MORTGAGE)
    ws["P2"] = "JAN"
    ws["AB2"] = "2025 YR TOTAL"
    ws["A45"] = (
        "2025 columns P–AB. Residence through sale 2025-12-18. MORTGAGE = US Bank loan 09422 monthly P+I from joint 0203 "
        "(not $50k extra principal on 9922 Jul/Sep — those are not on this row). HOA n/a. "
        "INSURANCE LOCK 2026-09-20: Travelers was prior 524 home; Geico was the cars. "
        "Only State Farm HOME share on this INSURANCE line (50/50 with Unit 2). Travelers leftover Jan is 524 home "
        "($0 on cards/Monarch). Auto + Geico credit personal — not EPGC. Not tax advice."
    )


def copy_row_block(ws: Worksheet, src_r0: int, dest_r0: int, n_rows: int, n_cols: int = 14) -> None:
    for dr in range(n_rows):
        for c in range(1, n_cols + 1):
            src = ws.cell(src_r0 + dr, c)
            dest = ws.cell(dest_r0 + dr, c)
            copy_style(src, dest)
            val = src.value
            if isinstance(val, str) and val.startswith("="):
                dest.value = Translator(val, origin=src.coordinate).translate_formula(dest.coordinate)
            else:
                dest.value = val


def fill_epgc(ws: Worksheet) -> None:
    # 2024 block is rows 26–48. Copy to 51–73 as 2025.
    copy_row_block(ws, 26, 51, 23, n_cols=14)
    ws["A51"] = 2025
    ws["A51"].font = Font(name=CG, size=14, bold=True, color="FFFFFF")
    write_months(ws, 53, 2, EPGC_BIZ)  # Art Sales / business income monthly
    write_months(ws, 54, 2, [0] * 12)  # Consultant $0
    write_months(ws, 63, 2, EPGC_AD)
    write_months(ws, 64, 2, EPGC_TRAVEL)
    write_months(ws, 66, 2, EPGC_CELL)
    write_months(ws, 67, 2, EPGC_SOFT)
    write_months(ws, 68, 2, EPGC_TECH)
    write_months(ws, 70, 2, EPGC_LIB)
    ws["A75"] = (
        "2025: Art Sales = MATCHED Mercury cash only (Berk $14k + $30k + mosaics $105k). "
        "Consultant $0 until EOEB/L5 answered (GCM 1099 is personal — GCM tab). "
        "Coinbase ACH is Investments; Newstar is Art Sales COGS. "
        "Cellphone includes AT&T. Canva + Park Chicago + Field Museum in Software / Travel. Not tax advice."
    )


def fill_income(ws: Worksheet) -> None:
    # 2022 B/C, 2023 D/E, 2024 F/G → add 2025 H/I
    for c in (8, 9):
        copy_style(ws.cell(1, 6), ws.cell(1, c))
        copy_style(ws.cell(2, 6 if c == 8 else 7), ws.cell(2, c))
        ws.cell(1, c).value = 2025
        ws.cell(2, c).value = "Projected" if c == 8 else "Actual"
        ws.cell(10, c).value = f"=SUM({get_column_letter(c)}4:{get_column_letter(c)}9)"
        copy_style(ws.cell(10, 6), ws.cell(10, c))
    # projected
    ws["H4"] = 110000
    ws["H5"] = 19500
    ws["H6"] = 17000
    ws["H7"] = 21500
    ws["H8"] = 20000
    ws["H9"] = 0
    # actual
    ws["I4"] = 70618  # Monarch paycheck cash
    ws["I5"] = 19500
    ws["I6"] = 17086.94
    ws["I7"] = 21500
    ws["I8"] = 0  # net until Sale 6428 Cost
    ws["I9"] = 0
    for addr in ("H4", "H5", "H6", "H7", "H8", "H9", "I4", "I5", "I6", "I7", "I8", "I9"):
        ws[addr].number_format = ACCT
        ws[addr].font = Font(name=CG, size=10)
    ws["I4"].fill = YELLOW
    ws["I8"].fill = YELLOW
    ws["A12"] = (
        "2025 Actual: I4 Monarch paycheck cash ≠ W-2 Box 1. I5 216 LTR cash $19,500. "
        "I6 524 #2 STR platform net LOCKED. I7 GCM 1099-NEC $21,500. I8 Art net $0 until Sale 6428 Cost. Not tax advice."
    )
    ws["A12"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[12].height = 36
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 14


def fill_gcm(ws: Worksheet) -> None:
    r0 = 22
    ws.cell(r0, 1).value = "GCM Boards 2025"
    ws.cell(r0, 1).font = Font(name=CG, size=12, bold=True)
    ws.cell(r0 + 1, 1).value = "Fees"
    for i, q in enumerate(["Q1", "Q2", "Q3", "Q4"], start=2):
        ws.cell(r0 + 1, i).value = q
    ws.cell(r0 + 2, 1).value = "GCM Board"
    for i, amt in enumerate([4100, 9200, 4100, 4100], start=2):
        cell = ws.cell(r0 + 2, i)
        cell.value = amt
        cell.number_format = ACCT
        cell.fill = PEACH
    ws.cell(r0 + 3, 1).value = "Total"
    for i, col in enumerate(["B", "C", "D", "E"], start=2):
        ws.cell(r0 + 3, i).value = f"=SUM({col}{r0 + 3})"  # B24 etc — wait row is r0+2 for board
    # fix totals to board row
    board_row = r0 + 2
    total_row = r0 + 3
    for i, col in enumerate(["B", "C", "D", "E"], start=2):
        ws.cell(total_row, i).value = f"=SUM({col}{board_row})"
    ws.cell(r0 + 4, 1).value = "Grand Total"
    ws.cell(r0 + 4, 5).value = f"=SUM(B{total_row}:E{total_row})"
    ws.cell(r0 + 6, 1).value = (
        "BOM PAYMEN $21,500 = Form 1099-NEC (personal, Jacob). NOT EPGC Consultant. "
        "JAKE C / SEFOF reimbursements excluded. Not tax advice."
    )


def fill_art_sale_6428(ws: Worksheet) -> None:
    # Insert after existing 2025 lots, before Total row 77.
    # Use row 76 (currently blank) for Sale 6428.
    r = 76
    ws.cell(r, 1).value = "Sale 6428 Contract 303468 (object TBD)"
    ws.cell(r, 2).value = "Cost TBD"
    ws.cell(r, 5).value = "Freeman's LLC — check to Megan (joint 0203)"
    ws.cell(r, 6).value = 11000
    ws.cell(r, 6).number_format = ACCT
    ws.cell(r, 6).fill = PEACH
    ws.cell(r, 7).value = "2025-12-24"
    ws.cell(r, 8).value = f"=IF(C{r}=\"\",\"\",F{r}-C{r})"
    ws.cell(r, 6).fill = PEACH
    ws.cell(r, 3).fill = YELLOW
    # inventory purchases after existing inventory section
    start = last_used_row(ws, scan_cols=9) + 2
    ws.cell(start, 1).value = "2025 CC inventory (not in Cost total until object is named)"
    ws.cell(start, 1).font = Font(name=CG, size=11, bold=True)
    extra = [
        ("Los Angeles Modern Auctions (inventory purchase)", "Sapphire 5423", 6821.06, "2025-07-24"),
        ("Hindman LLC (inventory purchase)", "Sapphire 5423", 282.24, "2025-04-03"),
        ("Hindman LLC (inventory purchase)", "Sapphire 5423", 403.51, "2025-07-29"),
        ("The Great Frame Up (inventory framing)", "Sapphire 5423", 536.41, "2025-02-16"),
    ]
    for i, (obj, src, cost, dt) in enumerate(extra, start=start + 1):
        ws.cell(i, 1).value = obj
        ws.cell(i, 2).value = src
        ws.cell(i, 3).value = cost
        ws.cell(i, 3).number_format = ACCT
        ws.cell(i, 3).fill = YELLOW
        ws.cell(i, 4).value = dt


def copy_extra_tabs(dest_wb, turbo_wb) -> None:
    for name in ("524 Home Sale", "827 Grove CapEx", "Childcare 2441"):
        if name not in turbo_wb.sheetnames:
            continue
        src = turbo_wb[name]
        if name in dest_wb.sheetnames:
            del dest_wb[name]
        ws = dest_wb.create_sheet(name)
        for r in src.iter_rows(min_row=1, max_row=min(src.max_row, 80), max_col=min(src.max_column or 1, 15)):
            for cell in r:
                d = ws.cell(cell.row, cell.column, cell.value)
                copy_style(cell, d)
        for col in src.column_dimensions:
            ws.column_dimensions[col].width = src.column_dimensions[col].width


def export_sheet(wb, name: str, path: Path) -> None:
    from openpyxl import Workbook

    out = Workbook()
    out.remove(out.active)
    src = wb[name]
    ws = out.create_sheet(name)
    for r in src.iter_rows(min_row=1, max_row=src.max_row, max_col=src.max_column):
        for cell in r:
            d = ws.cell(cell.row, cell.column, cell.value)
            copy_style(cell, d)
    for col, dim in src.column_dimensions.items():
        ws.column_dimensions[col].width = dim.width
    for rng in src.merged_cells.ranges:
        ws.merge_cells(str(rng))
    if src.sheet_properties.tabColor:
        ws.sheet_properties.tabColor = src.sheet_properties.tabColor.rgb
    out.save(path)


def main() -> None:
    shutil.copy2(PRIOR, OUT)
    wb = load_workbook(OUT)
    extra_src = CLEAN_EXTRA if CLEAN_EXTRA.exists() else None
    turbo = load_workbook(extra_src) if extra_src else None

    for name in wb.sheetnames:
        trim_sheet(wb[name])

    fill_216(wb["216 N. Oak Park Ave"])
    fill_unit2(wb["524 Ferdinand Ave, Unit 2"])
    fill_unit1(wb["524 Ferdinand Ave, Unit 1"])
    fill_epgc(wb["EPGC LLC"])
    fill_income(wb["Income"])
    fill_gcm(wb["GCM"])
    fill_art_sale_6428(wb["Art Sales and Purchases"])
    if turbo is not None:
        copy_extra_tabs(wb, turbo)

    # put Income first like last year (already is)
    wb.save(OUT)
    DELIVERABLE.parent.mkdir(exist_ok=True)
    shutil.copy2(OUT, DELIVERABLE)

    PARTS.mkdir(exist_ok=True)
    export_sheet(wb, "216 N. Oak Park Ave", PARTS / "216 N. Oak Park Ave.xlsx")
    # small 3-tab properties pack
    from openpyxl import Workbook

    props = Workbook()
    props.remove(props.active)
    for name in ("Income", "216 N. Oak Park Ave", "524 Ferdinand Ave, Unit 2", "524 Ferdinand Ave, Unit 1"):
        src = wb[name]
        ws = props.create_sheet(name)
        for r in src.iter_rows(min_row=1, max_row=src.max_row, max_col=src.max_column):
            for cell in r:
                d = ws.cell(cell.row, cell.column, cell.value)
                copy_style(cell, d)
        for col, dim in src.column_dimensions.items():
            ws.column_dimensions[col].width = dim.width
        for rng in src.merged_cells.ranges:
            try:
                ws.merge_cells(str(rng))
            except Exception:
                pass
        if src.sheet_properties.tabColor:
            ws.sheet_properties.tabColor = src.sheet_properties.tabColor.rgb
    props.save(PARTS / "Personal Income 2025 — like last year (Income+Properties).xlsx")

    print("saved", OUT, OUT.stat().st_size)
    print("216 part", (PARTS / "216 N. Oak Park Ave.xlsx").stat().st_size)
    print("props part", (PARTS / "Personal Income 2025 — like last year (Income+Properties).xlsx").stat().st_size)
    oak = load_workbook(OUT, data_only=False)["216 N. Oak Park Ave"]
    print("216 B5", oak["B5"].value, "AC5 rent Jan", oak.cell(5, 29).value, "AN5 Dec", oak.cell(5, 40).value)
    print("216 mortgage Jul", oak.cell(35, 35).value, "HOA Jun", oak.cell(36, 34).value)
    print("216 sheets", load_workbook(OUT).sheetnames)


if __name__ == "__main__":
    main()
