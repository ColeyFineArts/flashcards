#!/usr/bin/env python3
"""DO NOT USE for the user-facing Personal Income 2025 file.

CSV rebuilds drop Century Gothic, merged C1:O1, Hindman W2, Megan T4,
and Refrence Library. That is why Drive stopped looking like last year.

Canonical path: copy last year's Personal_Income_prior.xlsx, fill 2025,
run publish_personal_income_2025.py + assert_personal_income_2025.py.
"""
from __future__ import annotations

import sys

sys.exit(
    "Refused: do not rebuild Personal Income 2025 from CSVs. "
    "Use publish_personal_income_2025.py (copy of last year's xlsx)."
)

# Original CSV builder retained below for history only.
_ = """Build a modest Drive-upload workbook: one file, last-year tab names.

Omits styles (keeps zip small). 2025 cash months on property tabs.
Not tax advice.
"""

import csv
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.dom import minidom

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

ROOT = Path("/workspace/.cursor/scratch")
PARTS = ROOT / "tax_turbo_parts"
CSV_DIR = ROOT / "cc_fill_csv"
OUT_XLSX = PARTS / "Personal Income 2025 — like last year (all tabs).xlsx"
OUT_XML = PARTS / "Personal_Income_2025_all_tabs.xml"
GREEN = PatternFill("solid", fgColor="C6EFCE")

MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEPT", "OCT", "NOV", "DEC"]


def csv_line(name: str, filename: str) -> list[str]:
    path = CSV_DIR / filename
    for row in csv.reader(path.open(encoding="utf-8")):
        if len(row) > 1 and row[1] == name:
            return row
    raise SystemExit(f"missing {name} in {filename}")


def months_from(row: list[str], start: int = 2) -> list[float]:
    return [float(row[start + i] or 0) for i in range(12)]


def write_prop(ws, address: str, income: list[float], lines: list[tuple[str, list[float]]], note: str) -> None:
    ws["B1"] = "PROPERTY ADDRESS"
    ws["C1"] = address
    ws["B2"] = "INCOME"
    for i, m in enumerate(MONTHS):
        ws.cell(2, 3 + i, m)
    ws.cell(2, 15, "2025 YR TOTAL")
    ws["B3"] = "STR / RENTAL INCOME"
    for i, v in enumerate(income):
        ws.cell(3, 3 + i, v)
    ws.cell(3, 15, "=SUM(C3:N3)")
    ws["B4"] = "OTHER INCOME"
    for i in range(12):
        ws.cell(4, 3 + i, 0)
    ws.cell(4, 15, "=SUM(C4:N4)")
    ws["B5"] = "GROSS RENTAL INCOME"
    for i in range(12):
        col = get_column_letter(3 + i)
        ws.cell(5, 3 + i, f"=SUM({col}3:{col}4)")
    ws.cell(5, 15, "=SUM(C5:N5)")
    ws["B7"] = "EXPENSES"
    for i, m in enumerate(MONTHS):
        ws.cell(7, 3 + i, m)
    ws.cell(7, 15, "2025 YR TOTAL")
    for r, (label, vals) in enumerate(lines, start=8):
        ws.cell(r, 2, label)
        for i, v in enumerate(vals):
            cell = ws.cell(r, 3 + i, v)
            if label.upper().find("INSUR") >= 0 and v:
                cell.fill = GREEN
        col_end = get_column_letter(14)
        ws.cell(r, 15, f"=SUM(C{r}:{col_end}{r})")
    last = 7 + len(lines)
    tot = last + 1
    ws.cell(tot, 2, "TOTAL EXPENSES")
    for i in range(12):
        col = get_column_letter(3 + i)
        ws.cell(tot, 3 + i, f"=SUM({col}8:{col}{last})")
    ws.cell(tot, 15, f"=SUM(C{tot}:N{tot})")
    ytd = tot + 2
    ws.cell(ytd, 2, "YEAR-TO-DATE INCOME LESS EXPENSES")
    for i in range(12):
        col = get_column_letter(3 + i)
        ws.cell(ytd, 3 + i, f"={col}5-{col}{tot}")
    ws.cell(ytd, 15, f"=SUM(C{ytd}:N{ytd})")
    ws.cell(ytd + 2, 2, note)
    ws.cell(ytd + 2, 2).alignment = Alignment(wrap_text=True)
    ws.column_dimensions["B"].width = 36
    ws.row_dimensions[ytd + 2].height = 72


def sheet_from_csv_rows(ws, path: Path, max_rows: int = 80, max_cols: int = 14) -> None:
    with path.open(encoding="utf-8") as f:
        for r, row in enumerate(csv.reader(f), start=1):
            if r > max_rows:
                break
            for c, val in enumerate(row[:max_cols], start=1):
                if val == "":
                    continue
                if val.startswith("="):
                    ws.cell(r, c, val)
                else:
                    try:
                        if val.replace(".", "", 1).replace("-", "", 1).isdigit():
                            ws.cell(r, c, float(val) if "." in val else int(val))
                        else:
                            ws.cell(r, c, val)
                    except Exception:
                        ws.cell(r, c, val)
            if r == 1:
                ws.cell(r, 1).font = Font(bold=True)


def build_xlsx() -> Path:
    oak = csv_line("STR / RENTAL INCOME", "216_Oak_Park.csv")
    oak_ins = csv_line("INSURANCE", "216_Oak_Park.csv")
    oak_int = csv_line("INTERIOR MAINTENANCE", "216_Oak_Park.csv")
    oak_mort = csv_line("MORTGAGE (cash P+I — memo only)", "216_Oak_Park.csv")
    oak_hoa = csv_line("HOA", "216_Oak_Park.csv")
    u2 = csv_line("STR / RENTAL INCOME", "524_Unit_2.csv")
    u2_ins = csv_line("INSURANCE", "524_Unit_2.csv")
    u2_int = csv_line("INTERIOR MAINTENANCE", "524_Unit_2.csv")
    u2_ext = csv_line("EXTERIOR REPAIRS", "524_Unit_2.csv")
    u2_w = csv_line("VILLAGE WATER AND REFUSE", "524_Unit_2.csv")
    u2_gas = csv_line("GAS", "524_Unit_2.csv")
    u2_el = csv_line("ELECTRIC", "524_Unit_2.csv")
    u2_net = csv_line("INTERNET", "524_Unit_2.csv")
    u2_sup = csv_line("SUPPLIES", "524_Unit_2.csv")
    u1_ins = csv_line("INSURANCE", "524_Unit_1.csv")
    u1_ext = csv_line("EXTERIOR REPAIRS", "524_Unit_1.csv")
    u1_w = csv_line("VILLAGE WATER AND REFUSE", "524_Unit_1.csv")
    u1_gas = csv_line("GAS", "524_Unit_1.csv")
    u1_el = csv_line("ELECTRIC", "524_Unit_1.csv")

    wb = Workbook()
    wb.calculation = CalcProperties(fullCalcOnLoad=True)

    inc = wb.active
    inc.title = "Income"
    sheet_from_csv_rows(inc, CSV_DIR / "Income.csv", max_rows=12, max_cols=9)
    inc.column_dimensions["A"].width = 24
    inc.column_dimensions["I"].width = 14

    oak_ws = wb.create_sheet("216 N. Oak Park Ave")
    write_prop(
        oak_ws,
        "216 N. Oak Park Avenue #1Z (LTR)",
        months_from(oak),
        [
            ("MARKETING / ADVERTISING", [0] * 12),
            ("SERVICE FEES", [0] * 12),
            ("REMITTED TAX", [0] * 12),
            ("RENTER'S INSURANCE", months_from(oak_ins)),
            ("INTERIOR MAINTENANCE", months_from(oak_int)),
            ("EXTERIOR REPAIRS", [0] * 12),
            ("VILLAGE WATER AND REFUSE", [0] * 12),
            ("GAS", [0] * 12),
            ("ELECTRIC", [0] * 12),
            ("INTERNET", [0] * 12),
            ("CLEANING / TURNOVER", [0] * 12),
            ("SUPPLIES", [0] * 12),
            ("EQUIPMENT/APPLIANCE PURCHASES", [0] * 12),
            ("LAND TAX", [0] * 12),
            ("LOANS INTEREST (from 1098)", [0] * 12),
            ("MORTGAGE (cash P+I — memo only)", months_from(oak_mort)),
            ("HOA", months_from(oak_hoa)),
            ("OTHER", [0] * 12),
        ],
        "LOCKED 2026-09-20: Lemonade $514 cash 2025-10-03 Megan Chase 1702 is 216 LTR "
        "RENTER'S INSURANCE (October cash, not amortized). 2023 $321 / 2024 $421 same Oct 3 "
        "Lemonade. Travelers=524 home, Geico=cars, State Farm home share=524 only. Not tax advice.",
    )

    u2_ws = wb.create_sheet("524 Ferdinand Ave, Unit 2")
    write_prop(
        u2_ws,
        "524 Ferdinand Avenue, Unit 2 (STR)",
        months_from(u2),
        [
            ("MARKETING / ADVERTISING", [0] * 12),
            ("SERVICE FEES", [0] * 12),
            ("REMITTED TAX", [0] * 12),
            ("INSURANCE", months_from(u2_ins)),
            ("INTERIOR MAINTENANCE", months_from(u2_int)),
            ("EXTERIOR REPAIRS", months_from(u2_ext)),
            ("VILLAGE WATER AND REFUSE", months_from(u2_w)),
            ("GAS", months_from(u2_gas)),
            ("ELECTRIC", months_from(u2_el)),
            ("INTERNET", months_from(u2_net)),
            ("CLEANING / TURNOVER", [0] * 12),
            ("SUPPLIES", months_from(u2_sup)),
            ("EQUIPMENT/APPLIANCE PURCHASES", [0] * 12),
            ("LAND TAX", [0] * 12),
            ("LOANS INTEREST (from 1098)", [0] * 12),
            ("MORTGAGE (cash P+I — memo only)", [0] * 12),
            ("HOA", [0] * 12),
            ("OTHER", [0] * 12),
        ],
        "INSURANCE = State Farm HOME share only (Travelers was prior 524 home / Geico = cars). "
        "Not 216. Not tax advice.",
    )

    u1_ws = wb.create_sheet("524 Ferdinand Ave, Unit 1")
    write_prop(
        u1_ws,
        "524 Ferdinand Avenue, Unit 1 (residence through sale 2025-12-18)",
        [0] * 12,
        [
            ("MARKETING / ADVERTISING", [0] * 12),
            ("SERVICE FEES", [0] * 12),
            ("REMITTED TAX", [0] * 12),
            ("INSURANCE", months_from(u1_ins)),
            ("INTERIOR MAINTENANCE", [0] * 12),
            ("EXTERIOR REPAIRS", months_from(u1_ext)),
            ("VILLAGE WATER AND REFUSE", months_from(u1_w)),
            ("GAS", months_from(u1_gas)),
            ("ELECTRIC", months_from(u1_el)),
            ("INTERNET", [0] * 12),
            ("CLEANING / TURNOVER", [0] * 12),
            ("SUPPLIES", [0] * 12),
            ("EQUIPMENT/APPLIANCE PURCHASES", [0] * 12),
            ("LAND TAX", [0] * 12),
            ("LOANS INTEREST (from 1098)", [0] * 12),
            ("MORTGAGE (cash P+I — memo only)", [0] * 12),
            ("HOA", [0] * 12),
            ("OTHER", [0] * 12),
        ],
        "INSURANCE = State Farm HOME share only (Travelers was prior 524 home / Geico = cars). "
        "Not 216. Not tax advice.",
    )

    epgc = wb.create_sheet("EPGC LLC")
    epgc["A1"] = "2025"
    epgc["A2"] = "INCOME"
    for i, m in enumerate(MONTHS):
        epgc.cell(2, 2 + i, m)
    epgc.cell(2, 14, "YEARLY")
    epgc["A3"] = "Art Sales"
    for i, v in enumerate([14000, 136000] + [0] * 10):
        epgc.cell(3, 2 + i, v)
    epgc["N3"] = "=SUM(B3:M3)"
    epgc["A4"] = "Consultant"
    for i, v in enumerate([0, 0, 0, 0, 0, 13595, 0, 0, 0, 0, 0, 0]):
        epgc.cell(4, 2 + i, v)
        if v:
            epgc.cell(4, 2 + i).fill = GREEN
    epgc["N4"] = "=SUM(B4:M4)"
    epgc["A6"] = "EXPENSES"
    epgc["A7"] = "Consultant Fees"
    for i, v in enumerate([0, 0, 0, 0, 0, 0, 334.17, 0, 0, 658.63, 0, 0]):
        epgc.cell(7, 2 + i, v)
        if v:
            epgc.cell(7, 2 + i).fill = GREEN
    epgc["N7"] = "=SUM(B7:M7)"
    epgc["A9"] = (
        "LOCKED: Art Sales cash $150,000 (Jan $14,000 / Feb $136,000). "
        "David Aaron Consultant June $13,595. Wise Consultant Fees $992.80. "
        "Koziol/Ariadne $150k pass-through. Aquinas books $1,000 in Feb cash / Cost TBD. "
        "EOEB $130,270 and L5 $87,396.32 mixed unallocated (not dumped on P&L). "
        "Coinbase is Investments. Newstar $23,581 is jewelry COGS. Not tax advice."
    )
    epgc["A9"].alignment = Alignment(wrap_text=True)
    epgc.row_dimensions[9].height = 60
    epgc.column_dimensions["A"].width = 22

    art = wb.create_sheet("Art Sales and Purchases")
    art["A1"] = "2025 Art Sales (matched cash $150,000) + Cost-TBD rows. Full object register is in git xlsx."
    headers = ["Object", "From", "Cost", "Purchase Date", "To", "Sale Price", "Sale Date", "Note"]
    for i, h in enumerate(headers, start=1):
        art.cell(2, i, h)
    rows = [
        ["A Collection of Seals", "Erdal Dere", 13000, "2025-01-08", "Harlan J. Berk", 14000, "2025-01-10", "MATCHED"],
        ["Berk lots (Feb 7)", "Hindman et al.", 22944.94, "2024/2025", "Harlan J. Berk", 30000, "2025-02-07", "MATCHED Cost of sold lots"],
        ["Three Ancient Mosaics", "Jamal Rifai / Antiquarium", 90000, "2025-02-21", "Jonathan Yantis", 105000, "2025-02-21", "Plutus paid $50k of Cost; $40k ASK"],
        ["Books (titles TBD)", "Cost TBD", None, None, "Aquinas Hobor", 1000, "2025-02-04", "LOCKED book sale; Cost TBD — not in I8"],
        ["Sale 6428 Contract 303468", "Cost TBD", None, None, "Freeman's LLC (Megan 0203)", 11000, "2025-12-24", "Cost TBD — not in I8"],
        ["Canosan Terracotta Horse", "Hindman", 900, None, "Erdal Dere", 1000, None, "Sale date TBD"],
    ]
    for r, row in enumerate(rows, start=3):
        for c, v in enumerate(row, start=1):
            if v is not None:
                art.cell(r, c, v)
    art["A10"] = "Matched cash $150,000 = Berk $14,000 + Berk $30,000 + mosaics $105,000 + Aquinas $1,000. I8 $23,055.06 uses Cost-known deals only."
    art["A12"] = "Newstar Jewelers $23,581 jewelry COGS (10/14 $8,000 + 11/12 $8,000 + 12/24 $7,581) — parked, not dumped on P&L."
    art["A13"] = "Fortuna Venus $20,000 (5/9) + Roman Gold Belt $45,000 (7/21) LOCKED inventory Cost. Aug $20,000 unnamed ASK. Not sold 2025."
    art["A14"] = "David Aaron $13,595 is EPGC Consultant June — NOT an art sale."
    art.column_dimensions["A"].width = 36
    art.column_dimensions["H"].width = 42

    gcm = wb.create_sheet("GCM")
    gcm["A1"] = "GCM Boards 2025 — personal 1099-NEC $21,500 (NOT EPGC Consultant)"
    gcm["A2"] = "Feb"; gcm["B2"] = 4100
    gcm["A3"] = "Apr"; gcm["B3"] = 9200
    gcm["A4"] = "Sep"; gcm["B4"] = 4100
    gcm["A5"] = "Nov"; gcm["B5"] = 4100
    gcm["A6"] = "Total"; gcm["B6"] = 21500
    gcm["A8"] = "BOM PAYMEN on Jake 9922. Reimbursements excluded. Not tax advice."
    gcm.column_dimensions["A"].width = 64

    inv = wb.create_sheet("Investments")
    inv["A1"] = "Coinbase ACH from Mercury 8291 — investment, not EPGC"
    for i, m in enumerate(["SEP", "OCT", "NOV", "DEC"]):
        inv.cell(2, 2 + i, m)
    inv["A3"] = "To Coinbase"
    for i, v in enumerate([-9000, -10000, -35000, -5000]):
        inv.cell(3, 2 + i, v)
    inv["A4"] = "From Coinbase"
    inv["E4"] = 8000
    inv["A6"] = "2025 funded $59,000 / back $8,000 / net $-51,000. Not tax advice."
    inv.column_dimensions["A"].width = 48

    ask = wb.create_sheet("ASK Mercury")
    ask["A1"] = "Q"; ask["B1"] = "Status"; ask["C1"] = "What"
    ask_rows = [
        ("6", "LOCKED", "David Aaron $13,595 Consultant June — not Art Sales, not I8"),
        ("7", "LOCKED", "Wise $992.80 Consultant Fees (Jul $334.17 / Oct $658.63)"),
        ("8", "LOCKED", "Koziol/Ariadne $150k pass-through — not P&L"),
        ("9", "LOCKED", "Aquinas books $1,000 Feb; Cost TBD — not in I8"),
        ("1", "UNALLOCATED", "EOEB remainder $130,270 mixed — do not dump on P&L"),
        ("2", "UNALLOCATED", "L5 $87,396.32 mixed — do not dump on P&L"),
        ("Fortuna", "LOCKED", "Venus $20k (5/9) + Gold Belt $45k (7/21) inventory Cost"),
        ("3", "ASK", "Fortuna unnamed Aug $20k (Monarch 8/15 = native 8/18)"),
    ]
    for r, row in enumerate(ask_rows, start=2):
        for c, v in enumerate(row, start=1):
            ask.cell(r, c, v)
    ask.column_dimensions["A"].width = 12
    ask.column_dimensions["C"].width = 72

    PARTS.mkdir(exist_ok=True)
    wb.save(OUT_XLSX)
    return OUT_XLSX


def xml_cell(parent, col: int, val, typ: str | None = None) -> None:
    cell = ET.SubElement(parent, "{urn:schemas-microsoft-com:office:spreadsheet}Cell")
    cell.set("{urn:schemas-microsoft-com:office:spreadsheet}Index", str(col))
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        data = ET.SubElement(cell, "{urn:schemas-microsoft-com:office:spreadsheet}Data")
        data.set("{urn:schemas-microsoft-com:office:spreadsheet}Type", "Number")
        data.text = str(val)
    else:
        s = "" if val is None else str(val)
        if s.startswith("="):
            cell.set("{urn:schemas-microsoft-com:office:spreadsheet}Formula", s)
            data = ET.SubElement(cell, "{urn:schemas-microsoft-com:office:spreadsheet}Data")
            data.set("{urn:schemas-microsoft-com:office:spreadsheet}Type", "Number")
            data.text = "0"
        else:
            data = ET.SubElement(cell, "{urn:schemas-microsoft-com:office:spreadsheet}Data")
            data.set("{urn:schemas-microsoft-com:office:spreadsheet}Type", "String")
            data.text = s


def xml_row(table, cells: list, start_col: int = 1) -> None:
    ns = "urn:schemas-microsoft-com:office:spreadsheet"
    row = ET.SubElement(table, f"{{{ns}}}Row")
    for i, val in enumerate(cells):
        if val is None or val == "":
            continue
        xml_cell(row, start_col + i, val)


def build_xml() -> Path:
    """SpreadsheetML 2003 — text upload fallback if xlsx base64 fails."""
    ns = "urn:schemas-microsoft-com:office:spreadsheet"
    ET.register_namespace("", ns)
    ET.register_namespace("ss", ns)
    wb = ET.Element(f"{{{ns}}}Workbook")
    wb.set("xmlns", ns)

    def add_sheet(name: str, rows: list[list]):
        ws = ET.SubElement(wb, f"{{{ns}}}Worksheet")
        ws.set(f"{{{ns}}}Name", name)
        table = ET.SubElement(ws, f"{{{ns}}}Table")
        for row in rows:
            xml_row(table, row)

    add_sheet(
        "Income",
        [
            ["", "2022", "2022", "2023", "2023", "2024", "2024", "2025", "2025"],
            ["", "Projected", "Actual", "Projected", "Actual", "Projected", "Actual", "Projected", "Actual"],
            ["Income"],
            ["Hindman", 92000, 97876.85, 110000, 115000, 110000, 110000, 110000, 70618],
            ["216 N. Oak Park Ave", 0, 125, 1500, 1500, 1500, 1500, 19500, 19500],
            ["524 Ferdinand Ave #2", "", "", "", "", 20000, 17000, 17000, 17086.94],
            ["GCM Boards", 14000, 25875, 14700, 14700, 16000, 21500, 21500, 21500],
            ["Art Sales", 10000, 15000, 15000, 54777.25, 20000, 20000, 20000, 23055.06],
            ["Megan's Income", 67000, 67000, 74000, 74000, 0, 0, 0, 0],
            ["I8 $23,055.06 MATCHED Mercury Cost-known deals. Aquinas + Sale 6428 wait Cost. Not tax advice."],
        ],
    )
    oak_ins = [0, 0, 0, 0, 0, 0, 0, 0, 0, 514, 0, 0]
    oak_rent = [1950, 1950, 1950, 1950, 1950, 1950, 1450, 2450, 0, 0, 0, 3900]
    add_sheet(
        "216 N. Oak Park Ave",
        [
            ["PROPERTY ADDRESS", "216 N. Oak Park Avenue #1Z (LTR)"],
            ["LINE", *MONTHS, "2025 TOTAL"],
            ["STR / RENTAL INCOME", *oak_rent, 19500],
            ["RENTER'S INSURANCE", *oak_ins, 514],
            ["INTERIOR MAINTENANCE", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 675, 675],
            ["HOA", 0, 0, 0, 0, 0, 421.53, 421.53, 421.53, 421.53, 421.53, 421.53, 421.53, 2950.71],
            [
                "LOCKED: Lemonade $514 cash 2025-10-03 Megan Chase → October. "
                "Travelers/Geico/State Farm are not 216. Not tax advice."
            ],
        ],
    )
    u2_ins = [0, 78.77, 78.76, 78.76, 78.76, 78.76, 78.76, 80.19, 0, 0, 0, 0]
    u2_rent = [0, 0, 0, 0, 0, 3479.6, 3180.81, 1757.5, 2690.16, 1261.64, 1509.84, 0]
    add_sheet(
        "524 Ferdinand Ave, Unit 2",
        [
            ["PROPERTY ADDRESS", "524 Ferdinand Avenue, Unit 2 (STR)"],
            ["LINE", *MONTHS],
            ["STR / RENTAL INCOME", *u2_rent],
            ["INSURANCE (State Farm HOME share; Travelers was prior 524 home)", *u2_ins],
        ],
    )
    u1_ins = [0, 78.76, 78.75, 78.75, 78.75, 78.75, 78.75, 80.19, 0, 0, 0, 0]
    add_sheet(
        "524 Ferdinand Ave, Unit 1",
        [
            ["PROPERTY ADDRESS", "524 Ferdinand Avenue, Unit 1"],
            ["LINE", *MONTHS],
            ["INSURANCE (State Farm HOME share; not 216)", *u1_ins],
        ],
    )
    add_sheet(
        "EPGC LLC",
        [
            ["2025"],
            ["INCOME", *MONTHS, "YEARLY"],
            ["Art Sales", 14000, 136000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 150000],
            ["Consultant", 0, 0, 0, 0, 0, 13595, 0, 0, 0, 0, 0, 0, 13595],
            ["EXPENSES"],
            ["Consultant Fees", 0, 0, 0, 0, 0, 0, 334.17, 0, 0, 658.63, 0, 0, 992.80],
            [
                "LOCKED: Art Sales cash $150,000; David Aaron Consultant June $13,595; "
                "Wise $992.80 Consultant Fees; Koziol/Ariadne $150k pass-through; "
                "EOEB $130,270 and L5 $87,396.32 mixed unallocated (not dumped on P&L). Not tax advice."
            ],
        ],
    )
    add_sheet(
        "Art Sales and Purchases",
        [
            ["Object", "To", "Sale Price", "Date", "Cost", "Note"],
            ["Seals", "Harlan J. Berk", 14000, "2025-01-10", 13000, "MATCHED"],
            ["Berk lots", "Harlan J. Berk", 30000, "2025-02-07", 22944.94, "MATCHED"],
            ["Mosaics", "Jonathan Yantis", 105000, "2025-02-21", 90000, "Plutus $50k of Cost; $40k ASK"],
            ["Books (titles TBD)", "Aquinas Hobor", 1000, "2025-02-04", "TBD", "LOCKED; not in I8 until Cost"],
            ["Sale 6428", "Freeman's LLC", 11000, "2025-12-24", "TBD", "not in I8 until Cost"],
            ["Fortuna Venus", "inventory", "", "2025-05-09", 20000, "LOCKED parked Cost; not sold 2025"],
            ["Roman Gold Belt", "inventory", "", "2025-07-21", 45000, "LOCKED parked Cost; not sold 2025"],
        ],
    )
    add_sheet(
        "GCM",
        [
            ["GCM Boards 2025 — personal 1099-NEC $21,500 (NOT EPGC Consultant)"],
            ["Feb", 4100],
            ["Apr", 9200],
            ["Sep", 4100],
            ["Nov", 4100],
            ["Total", 21500],
        ],
    )
    add_sheet(
        "Investments",
        [
            ["Coinbase ACH from Mercury 8291 — investment, not EPGC"],
            ["", *["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]],
            ["To Coinbase", 0, 0, 0, 0, 0, 0, 0, 0, -9000, -10000, -35000, -5000],
            ["From Coinbase", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 8000],
            ["2025 funded $59,000 / back $8,000 / net $-51,000. Not tax advice."],
        ],
    )
    add_sheet(
        "ASK Mercury",
        [
            ["Q", "Status", "What"],
            ["6", "LOCKED", "David Aaron $13,595 Consultant June"],
            ["7", "LOCKED", "Wise $992.80 Consultant Fees"],
            ["8", "LOCKED", "Koziol/Ariadne $150k pass-through"],
            ["9", "LOCKED", "Aquinas books $1,000 Feb; Cost TBD"],
            ["1", "UNALLOCATED", "EOEB remainder $130,270 mixed — do not dump"],
            ["2", "UNALLOCATED", "L5 $87,396.32 mixed — do not dump"],
            ["Fortuna", "LOCKED", "Venus $20k (5/9) + Gold Belt $45k (7/21) inventory"],
        ],
    )

    xml_bytes = ET.tostring(wb, encoding="utf-8", xml_declaration=True)
    # keep compact (no pretty-print) for a smaller upload
    OUT_XML.write_bytes(xml_bytes)
    return OUT_XML


if __name__ == "__main__":
    x = build_xlsx()
    xmlp = build_xml()
    print("xlsx", x, x.stat().st_size)
    print("xml", xmlp, xmlp.stat().st_size)
    from openpyxl import load_workbook

    wb = load_workbook(x, data_only=False)
    print("sheets", wb.sheetnames)
    oak = wb["216 N. Oak Park Ave"]
    print("B11", oak["B11"].value, "L11", oak["L11"].value)
    print("Income I8", wb["Income"]["I8"].value)
    print("EPGC B53", wb["EPGC LLC"]["B53"].value, "C53", wb["EPGC LLC"]["C53"].value)
    wb.close()
