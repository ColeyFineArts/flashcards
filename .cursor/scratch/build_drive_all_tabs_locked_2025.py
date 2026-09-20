#!/usr/bin/env python3
"""Build the Drive all-tabs workbook from locked last-year CSVs + extra turbo tabs.

Looks like last year’s Personal Income.xlsx (2023|2024|2025 on property tabs;
GCM / EPGC / Art Sales in the same file).

216 insurance is amortized Lemonade $42.84/mo (not October lump).
Travelers = 524 home; Geico = 524 cars.

Not tax advice.
"""
from __future__ import annotations

import csv
import xml.etree.ElementTree as ET
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path("/workspace/.cursor/scratch")
CSV_DIR = ROOT / "cc_fill_csv"
TURBO = ROOT / "Personal_Income_2025_Tax_Turbo.xlsx"
PARTS = ROOT / "tax_turbo_parts"
OUT_XLSX = PARTS / "Personal Income 2025 — like last year (all tabs).xlsx"
OUT_XML = PARTS / "Personal_Income_2025_all_tabs.xml"

CSV_SHEETS = [
    ("Income", "Income.csv"),
    ("524 Ferdinand Ave, Unit 2", "524_Unit_2_like_last_year.csv"),
    ("216 N. Oak Park Ave", "216_Oak_Park_like_last_year.csv"),
    ("EPGC LLC", "EPGC_LLC.csv"),
    ("Art Sales and Purchases", "Art_Sales.csv"),
    ("GCM", "GCM.csv"),
    ("Investments", "Investments.csv"),
    ("524 Ferdinand Ave, Unit 1", "524_Unit_1_like_last_year.csv"),
    ("ASK Mercury", "ASK_Mercury.csv"),
    ("Carriers", "LOCK_carriers_2025.csv"),
]

EXTRA_FROM_TURBO = [
    "Refrence Library",
    "Hindman W2",
    "Megan T4",
    "524 Home Sale",
    "827 Grove CapEx",
    "Childcare 2441",
    "MERCURY 8291",
]

GREEN = PatternFill("solid", fgColor="C6EFCE")


def parse_cell(val: str):
    if val is None or val == "":
        return None
    if isinstance(val, (int, float)):
        return val
    s = str(val)
    if s.startswith("="):
        return s
    try:
        if s.replace(".", "", 1).replace("-", "", 1).isdigit():
            return float(s) if "." in s else int(s)
    except Exception:
        pass
    return s


def sheet_from_csv(ws: Worksheet, path: Path) -> None:
    with path.open(encoding="utf-8", newline="") as f:
        for r, row in enumerate(csv.reader(f), start=1):
            for c, val in enumerate(row, start=1):
                parsed = parse_cell(val)
                if parsed is None:
                    continue
                cell = ws.cell(r, c, parsed)
                if r == 1:
                    cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 36


def copy_sheet(src: Worksheet, dest: Worksheet) -> None:
    for row in src.iter_rows():
        for cell in row:
            d = dest.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                d.font = copy(cell.font)
                d.fill = copy(cell.fill)
                d.border = copy(cell.border)
                d.alignment = copy(cell.alignment)
                d.number_format = cell.number_format
    for col, dim in src.column_dimensions.items():
        dest.column_dimensions[col].width = dim.width
    dest.sheet_view.showGridLines = src.sheet_view.showGridLines


def paint_216_insurance(ws: Worksheet) -> None:
    # 2025 block starts at AC (29). RENTER'S INSURANCE is row 14.
    for c in range(29, 41):
        cell = ws.cell(14, c)
        if cell.value not in (None, "", 0):
            cell.fill = GREEN


def build_xlsx() -> Path:
    PARTS.mkdir(exist_ok=True)
    wb = Workbook()
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    first = True
    for title, filename in CSV_SHEETS:
        if first:
            ws = wb.active
            ws.title = title
            first = False
        else:
            ws = wb.create_sheet(title)
        sheet_from_csv(ws, CSV_DIR / filename)
        if title == "216 N. Oak Park Ave":
            paint_216_insurance(ws)
            ws.row_dimensions[64].height = 48
            ws.cell(64, 2).alignment = Alignment(wrap_text=True)
        if title.startswith("524"):
            for r in range(1, ws.max_row + 1):
                v = ws.cell(r, 1).value
                if v and "INSURANCE LOCK" in str(v):
                    ws.cell(r, 1).alignment = Alignment(wrap_text=True)
                    ws.row_dimensions[r].height = 64

    # Extra turbo tabs (W2/T4/sale/CapEx/childcare/Mercury) stay in git xlsx.
    # Drive convert of those styled sheets previously corrupted; CSV tabs below
    # are the live Google Sheet (same tab set as 1BFOQzy0, last-year columns).

    order = [
        "Income",
        "524 Ferdinand Ave, Unit 2",
        "216 N. Oak Park Ave",
        "EPGC LLC",
        "Art Sales and Purchases",
        "GCM",
        "Investments",
        "524 Ferdinand Ave, Unit 1",
        "ASK Mercury",
        "Carriers",
    ]
    for i, name in enumerate(order):
        if name in wb.sheetnames:
            wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb.save(OUT_XLSX)
    return OUT_XLSX


def xml_cell(row_el, col: int, val) -> None:
    ns = "urn:schemas-microsoft-com:office:spreadsheet"
    cell = ET.SubElement(row_el, f"{{{ns}}}Cell")
    cell.set(f"{{{ns}}}Index", str(col))
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        data = ET.SubElement(cell, f"{{{ns}}}Data")
        data.set(f"{{{ns}}}Type", "Number")
        data.text = str(val)
        return
    s = "" if val is None else str(val)
    if s.startswith("="):
        cell.set(f"{{{ns}}}Formula", s)
        data = ET.SubElement(cell, f"{{{ns}}}Data")
        data.set(f"{{{ns}}}Type", "Number")
        data.text = "0"
        return
    data = ET.SubElement(cell, f"{{{ns}}}Data")
    data.set(f"{{{ns}}}Type", "String")
    data.text = s


def build_xml() -> Path:
    """Compact SpreadsheetML for Drive convert (full CSV XML is too large)."""
    ns = "urn:schemas-microsoft-com:office:spreadsheet"
    ET.register_namespace("", ns)
    ET.register_namespace("ss", ns)
    root = ET.Element(f"{{{ns}}}Workbook")
    root.set("xmlns", ns)
    months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEPT", "OCT", "NOV", "DEC"]

    def add_sheet(name: str, rows: list[list]):
        ws = ET.SubElement(root, f"{{{ns}}}Worksheet")
        ws.set(f"{{{ns}}}Name", name)
        table = ET.SubElement(ws, f"{{{ns}}}Table")
        for row in rows:
            row_el = ET.SubElement(table, f"{{{ns}}}Row")
            for i, val in enumerate(row, start=1):
                if val in (None, "", 0, 0.0):
                    continue
                xml_cell(row_el, i, parse_cell(val) if isinstance(val, str) else val)

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
            ["Megan's Income", 67000, 67000, 74000, 74000],
            [
                "I8 $23,055.06 MATCHED Cost-known deals. Aquinas + Sale 6428 wait Cost. "
                "David Aaron is EPGC Consultant not I8. Not tax advice."
            ],
        ],
    )
    y23_ins = [26.75] * 12
    y24_ins = [35.08] * 12
    y25_ins = [42.84] * 12
    y23_rent = [0, 0, 0, 0, 0, 0, 0, 0, 0, 1775, 1775, 1775]
    y24_rent = [1775] * 9 + [1950, 1950, 1950]
    y25_rent = [1950, 1950, 1950, 1950, 1950, 1950, 1450, 2450, 0, 0, 0, 3900]
    y23_mort = [1066.27] * 10 + [1215.82, 1215.82]
    y24_mort = [1318.02] * 12
    y25_mort = [1226.25] * 10 + [1177.52, 1177.52]
    y23_hoa = [397.48] * 12
    y24_hoa = [401.46] * 12
    y25_hoa = [421.53] * 12
    y25_int = [0, 0, 0, 0, 0, 0, 0, 150, 0, 0, 0, 11]
    y25_rep = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 675]
    y25_move = [0, 0, 0, 0, 0, 0, 0, 0, 0, 300, 0, 0]
    hdr = ["LINE", *months, "2023 YR", *months, "2024 YR", *months, "2025 YR"]
    add_sheet(
        "216 N. Oak Park Ave",
        [
            ["PROPERTY ADDRESS", "216 N. Oak Park Avenue #1Z (LTR)"],
            hdr,
            ["RENTAL INCOME", *y23_rent, 5325, *y24_rent, 21975, *y25_rent, 19500],
            ["RENTER'S INSURANCE", *y23_ins, 321, *y24_ins, 420.96, *y25_ins, 514.08],
            ["INTERIOR MAINTENANCE", *[0] * 12, 0, 250, *[0] * 11, 250, *y25_int, 161],
            ["INTERIOR REPAIR", *[0] * 12, 0, 0, 0, 0, 0, 80, 220, 0, 0, 0, 0, 250, 0, 550, *y25_rep, 675],
            ["MORTGAGE", *y23_mort, 13094.34, *y24_mort, 15816.24, *y25_mort, 14617.54],
            ["HOA", *y23_hoa, 4769.76, *y24_hoa, 4817.52, *y25_hoa, 5058.36],
            ["MOVE OUT FEE", *[0] * 9, 200, 0, 0, 200, *[0] * 9, 200, 0, 0, 200, *y25_move, 300],
            [
                "LOCKED 2026-09-20: Lemonade $42.84/mo × 12 = $514.08 amortized like 2023 $26.75 / 2024 $35.08 "
                "(cash $514.00 10/3 Megan Chase, 8¢). User said 216 N Grove — this Oak Park #1Z, not 827 Grove. "
                "Mortgage $1,226.25 Jan–Oct / $1,177.52 Nov–Dec (Jan–Jun yellow WAIT 8507). HOA $421.53 "
                "(Jan–May yellow WAIT 8507). Brennan $150 Aug INTERIOR. Ace keys $11 Dec INTERIOR. "
                "Imelda $300 MOVE OUT October. Joan $675 INTERIOR REPAIR December. Smoke detector Amazon ASK. "
                "Travelers+Geico=524. Not tax advice."
            ],
        ],
    )
    u2_ins = [0, 78.77, 78.76, 78.76, 78.76, 78.76, 78.76, 80.19, 0, 0, 0, 0]
    u2_rent = [0, 0, 0, 0, 0, 3479.6, 3180.81, 1757.5, 2690.16, 1261.64, 1509.84, 0]
    add_sheet(
        "524 Ferdinand Ave, Unit 2",
        [
            ["PROPERTY ADDRESS", "524 Ferdinand Avenue, Unit 2 (STR)"],
            ["LINE", *months, "2025 TOTAL"],
            ["STR / RENTAL INCOME", *u2_rent, 13879.55],
            ["INSURANCE (State Farm HOME share 50/50)", *u2_ins, 552.76],
            [
                "INSURANCE LOCK: Travelers was prior 524 home leftover Jan $0. Geico was the cars (personal; "
                "$540.89 credit excluded). Only State Farm HOME share on this line. Lemonade is 216 Oak Park, "
                "not 524, not 827 Grove. Grove Collaborative ≠ 827 N Grove the house. Not tax advice."
            ],
        ],
    )
    u1_ins = [0, 78.76, 78.75, 78.75, 78.75, 78.75, 78.75, 80.19, 0, 0, 0, 0]
    add_sheet(
        "524 Ferdinand Ave, Unit 1",
        [
            ["PROPERTY ADDRESS", "524 Ferdinand Avenue, Unit 1 (residence through sale 2025-12-18 §121)"],
            ["LINE", *months, "2025 TOTAL"],
            ["INSURANCE (State Farm HOME share 50/50)", *u1_ins, 552.70],
            [
                "INSURANCE LOCK: Travelers was prior 524 home leftover Jan $0. Geico was the cars (personal). "
                "Only State Farm HOME share on this line. Lemonade is 216 Oak Park, not 524. Not tax advice."
            ],
        ],
    )
    add_sheet(
        "EPGC LLC",
        [
            ["2025"],
            ["INCOME", *months, "YEARLY"],
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
            ["Jan", 0],
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
            ["216", "ASK", "Smoke detector Amazon $0; Ace $17 12/8; Joan half-Nov $975 WAIT 8507"],
        ],
    )
    add_sheet(
        "Carriers",
        [
            ["Carrier", "Property", "Status", "Where the dollars are", "Not"],
            [
                "Lemonade",
                "216 N. Oak Park Avenue #1Z (LTR)",
                "LOCKED 2026-09-20",
                "216 RENTER'S INSURANCE $42.84/mo × 12 = $514.08 (cash $514.00 10/3 Megan Chase)",
                "Not 524. Not 827 N Grove. Not Grove Collaborative.",
            ],
            [
                "Travelers",
                "524 Ferdinand (home)",
                "LOCKED 2026-09-20",
                "524 INSURANCE leftover Jan 2025 $0.00 on cards/Monarch. Units 1 & 2 keep 50/50 home-share.",
                "Not 216. Not Geico (cars). Not Lemonade.",
            ],
            [
                "Geico",
                "524 Ferdinand household cars",
                "LOCKED 2026-09-20",
                "Personal auto. 2025-02-11 AUTO credit $540.89 EXCLUDED. Not on 524 home INSURANCE line.",
                "Not 216. Not EPGC. Not Lemonade.",
            ],
            [
                "State Farm 2025 bundle",
                "524 Ferdinand car+home (replaced Travelers/Geico)",
                "LOCKED 2026-09-18 / 2026-09-20",
                "HOME share $1,105.46 on 524 INSURANCE 50/50 U1/U2. AUTO $1,019.33 personal.",
                "Not 216.",
            ],
            [
                "827 N Grove the house",
                "Closed 2025-12-18",
                "NOT Lemonade",
                "CapEx $53,660 locked. HD 1/3 $604.12 pre-close materials.",
                "User said 216 N Grove — that is Oak Park #1Z, not this purchase.",
            ],
        ],
    )

    xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    OUT_XML.write_bytes(xml_bytes)
    slim = Path("/workspace/.cursor/scratch/tax_turbo_parts/Personal_Income_2025_all_tabs_slim.xml")
    slim.write_bytes(xml_bytes)
    return OUT_XML


if __name__ == "__main__":
    x = build_xlsx()
    xmlp = build_xml()
    print("xlsx", x, x.stat().st_size)
    print("xml", xmlp, xmlp.stat().st_size)
    wb = load_workbook(x, data_only=False)
    print("sheets", wb.sheetnames)
    oak = wb["216 N. Oak Park Ave"]
    print("216 r14 2023", [oak.cell(14, c).value for c in range(3, 15)])
    print("216 r14 2025", [oak.cell(14, c).value for c in range(29, 41)])
    print("216 note", str(oak.cell(64, 2).value)[:160])
    print("Income I8", wb["Income"]["I8"].value)
    print("GCM B5", wb["GCM"]["B5"].value)
    print("Carriers A2", wb["Carriers"]["A2"].value)
    wb.close()
