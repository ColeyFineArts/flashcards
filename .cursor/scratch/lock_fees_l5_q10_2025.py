#!/usr/bin/env python3
"""Lock confirmed FEE recasts. Do not dump remaining EOEB.

User 2026-09-21 Agent-mode lock-in:
  Q10 EOEB 2/21 $105,000 = Jamal mosaics SALE $40,000 + Antiquarium/Plutus
  SALE $50,000 + $15,000 FEE (not I8 profit). Working dollars $50k / $40k /
  $15k. Jamal BoA Cost still ASK — do not invent a $40k BoA Cost wire.
  L5 leftover six invoices $37,396.32 ALL FEES → EPGC Consultant, not I8 /
  Art Sales. Belt 7/15 $50,000 stays SALE. Koziol/Ariadne $150,000 stays PASS.
  L5 10/21 $7,500 is 5% of that $150k window (FEE). L5 7/15 $2,500 is 5% of
  the belt $50k (FEE, not attached to the belt SALE).
  EOEB 1/16 $1,000 Canosan FEE (cash is EOEB; register horse is Hindman/Erdal
  — do not double-count as a second horse SALE).
  EOEB 5/1 $21,000 FEE on the NBerk deal, not Venus (Venus is Fortuna 5/9
  −$20k / Erdal 7/24 +$26k).
  Rest of EOEB after $1k+$21k = $108,270 stays UNALLOCATED. Do not dump.
  Belt / Venus / August SALE unchanged. Aysel still ASK.

P&L recast
  I8 $24,055.06 (drop mosaics $15k fee). I10 → $152,760.
  Feb Art Sales $121,000 (Berk $30k + Aquinas $1k + mosaics SALE $90k).
  Year Art Sales $236,000.
  Consultant $87,991.32 (Jan 1,000 + Feb 15,000 + May 21,000 + Jun 13,595
  + Jul 11,034.79 + Sep 14,161.27 + Oct 7,500 + Dec 4,700.26).

Not tax advice. Do not re-run apply_checking_answers_2025.py.
"""
from __future__ import annotations

import csv
import json
import shutil
from copy import copy
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
import sys

sys.path.insert(0, str(Path("/workspace/.cursor/scratch")))
from sanitize_xlsx_for_sheets import sanitize_xlsx

ROOT = Path("/workspace/.cursor/scratch")
PARTS = ROOT / "tax_turbo_parts"
CSV_DIR = ROOT / "cc_fill_csv"
XLSX = PARTS / "Personal Income 2025.xlsx"
DELIVERABLE = ROOT / "tax_turbo_deliverable" / "Personal Income 2025.xlsx"
PACKET = ROOT / "cpa_packet_v1.json"
MERCURY_JSON = ROOT / "mercury_2025.json"
DRIVE_MAP = ROOT / "drive_folder_map.json"

GREEN = PatternFill("solid", fgColor="C6EFCE")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
PEACH = PatternFill("solid", fgColor="F7CAAC")
CG = Font(name="Century Gothic", size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
ACCT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'

# Live Drive IDs — patched after upload in the same run.
OPEN_XLSX = "https://docs.google.com/spreadsheets/d/1rSWqsoa_mxvA5fEXw_0yj8oa0IZ-5rs2KwNpKsd2XMs/edit"
OPEN_07 = "https://docs.google.com/spreadsheets/d/1apJDDSGYEx2CIA1eQ62fj9lK44AYBKCy0GlhyX8B_pE/edit"
OPEN_ROOT = "https://docs.google.com/spreadsheets/d/1XqpWHqdoSTZBJnzT8frhVME-IvRqOoVt-d-Utirb6YE/edit"
LABELED_ART = "https://docs.google.com/spreadsheets/d/1xJGlxCmigKBhR0XnYQbjZGGzl_m64aalDfDrP2l8KlQ/edit"
CLASSIFIER = "https://docs.google.com/spreadsheets/d/1NEa7qvWhduqndg9tdGYVvvinZ-NSB8mEwFW1l-a-8bk/edit"
ASK = "https://docs.google.com/spreadsheets/d/12PxZu780mmvpD44pLve5Vy6dvYQo_dXjVOMvVow7S5U/edit"
FORTUNA = "https://docs.google.com/spreadsheets/d/1TK93q6iwXuoWodkikxwFNKgGkfDiGgyk1RpvMEKmPPw/edit"
START_HERE = "https://docs.google.com/document/d/1Kup5ndlhMX_JRhLzs6BddeNUfNH95V6G3lVgIrqrv5I/edit"
LOCK = "https://docs.google.com/spreadsheets/d/1lki8Cy3_InEUK9ZCraTAcgxQBllz_YOahwmoAGDHKa8/edit"

I8 = 24055.06
I10_VALUE = 152760.00
FEB_ART = 121000.0
ART_GROSS = 236000.0
CONSULTANT = 87991.32
EOEB_REMAINING = 108270.0
L5_FEES = 37396.32
MOSAICS_SALE = 90000.0
MOSAICS_FEE = 15000.0
PLUTUS_SALE = 50000.0
JAMAL_SALE = 40000.0
CANOSAN_FEE = 1000.0
NBERK_FEE = 21000.0

CONSULTANT_MONTHLY = [
    1000.0,  # Jan Canosan
    15000.0,  # Feb mosaics fee
    0.0,
    0.0,
    21000.0,  # May NBerk
    13595.0,  # Jun David Aaron
    11034.79,  # Jul L5 8534.79 + 2500
    0.0,
    14161.27,  # Sep L5 5042 + 9119.27
    7500.0,  # Oct L5 5% of Koziol/Ariadne
    0.0,
    4700.26,  # Dec L5
]

L5_INVOICES = [
    ("2025-07-11", 8534.79, "L5 invoice — FEE (not belt SALE)"),
    ("2025-07-15", 2500.00, "L5 invoice — FEE; 5% of belt $50,000 window (not attached to belt SALE)"),
    ("2025-09-02", 5042.00, "L5 invoice — FEE"),
    ("2025-09-23", 9119.27, "L5 invoice — FEE"),
    ("2025-10-21", 7500.00, "L5 invoice — FEE; 5% of Koziol/Ariadne $150,000 window (not an Ariadne object SALE)"),
    ("2025-12-04", 4700.26, "L5 invoice — FEE"),
]

DEAL_B_NOTE = (
    "LOCKED 2026-09-21 SALE Deal B. Plutus & Mnemosyne / DBA Antiquarium. "
    "Cost $50,000 on Mercury 8291 (2/24 $10 + 2/25 $49,990 native COGS). "
    "Sale $50,000 inside EOEB 2/21 $105,000 MAKE A PAYMENT (pooled with Deal A "
    "$40,000 SALE + $15,000 FEE). Net $0. Do NOT treat as Jamal DBA. "
    "2/24 $90k EPGC→9922 is owner transfer, not extra Cost. Classifier: SALE. "
    "Not tax advice."
)
DEAL_A_NOTE = (
    "LOCKED 2026-09-21 SALE Deal A cash / ASK Cost. Three Ancient Mosaics / "
    "Jamal Rifai. Sale $40,000 inside EOEB 2/21 $105,000 (working dollars: "
    "$50k Antiquarium + $40k Jamal + $15k FEE). Cost from personal BoA still "
    "ASK — Monarch 9922 has no labeled $40k/$90k OUT to Jamal in Feb. Do not "
    "invent a $40k BoA Cost wire. $10 Mercury tests 2/21 and 2/24 only. "
    "Net excluded from I8 until Cost. The $15k is FEE (Consultant February), "
    "not art profit. Classifier: SALE. Not tax advice."
)
ART_FOOTER = (
    "Mercury 8291 MATCHED/LOCKED: Berk $14,000 (1/10 seals) + $30,000 (2/7 lots). "
    "EOEB 2/21 $105,000 = Deal B Plutus SALE $50,000 + Deal A Jamal SALE $40,000 "
    "+ $15,000 FEE (Consultant February — not I8). Plutus Cost $50,000 on 8291. "
    "Jamal Cost ASK (personal BoA). Fortuna $13,000 (1/8 seals Cost). "
    "Roman Gold Belt SALE to L5: Cost $45,000 / Sale $50,000 / net $5,000. "
    "Venus joint SALE EPGC share: Cost $20,000 / proceeds $26,000 / net $6,000. "
    "August Fortuna joint SALE EPGC share: Cost $20,000 / proceeds $25,000 / net $5,000 "
    "(Sale Price = remittance, not full hammer; object unnamed). "
    "L5 leftover $37,396.32 LOCKED FEE (Consultant; 7/15 $2,500 = 5% of belt; "
    "10/21 $7,500 = 5% of Koziol/Ariadne). EOEB 1/16 $1,000 Canosan FEE (not horse "
    "SALE). EOEB 5/1 $21,000 NBerk FEE (not Venus). Newstar $23,581 jewelry "
    "fabrication COGS parked below. David Aaron $13,595 LOCKED consultant June. "
    "Aquinas Hobor $1,000 LOCKED book sale (Feb cash; Cost TBD, not in I8). "
    "Do not dump EOEB remainder $108,270 (see Deal Classifier per invoice). "
    "Not tax advice."
)
INCOME_NOTE = (
    "2025 Actual: I4 Monarch paycheck cash ≠ W-2 Box 1. I5 216 LTR cash $19,500. "
    "I6 524 #2 STR platform net LOCKED. I7 GCM 1099-NEC $21,500. "
    "I8 Art net $24,055.06 = Berk lots + seals $1,000 + Belt net $5,000 + Venus "
    "EPGC-share net $6,000 + August Fortuna joint net $5,000. Mosaics $15,000 is "
    "FEE (EPGC Consultant February) — not I8. Deal B Plutus net $0. Deal A Jamal "
    "Cost ASK so not in I8. Excludes Sale 6428 $11,000 until Cost. Excludes "
    "Aquinas books $1,000 until Cost. David Aaron $13,595 is EPGC Consultant June "
    "(not I8). L5 leftover $37,396.32 / EOEB Canosan $1,000 / EOEB NBerk $21,000 "
    "are Consultant FEEs (not I8). EOEB remainder $108,270 stays UNALLOCATED "
    "(per-invoice Deal Classifier — not dumped). Erdal 11/7 $6,000 / Aysel $50,000 "
    "still ASK. Not tax advice."
)
EPGC_NOTE = (
    "2025 Mercury 2026-09-21 FEE lock: Art Sales gross cash by month = Jan $14,000 "
    "+ Feb $121,000 (Berk $30,000 + Aquinas $1,000 + mosaics SALE $90,000) + July "
    "$76,000 (Belt L5 $50,000 + Venus remittance $26,000) + October $25,000 "
    "(August Fortuna joint remittance) = $236,000. EOEB 2/21 $105,000 split: "
    "Deal B $50,000 SALE + Deal A $40,000 SALE + $15,000 FEE. "
    "Consultant $87,991.32 = Jan $1,000 Canosan + Feb $15,000 mosaics fee + May "
    "$21,000 NBerk + June $13,595 David Aaron + July $11,034.79 L5 + Sep "
    "$14,161.27 L5 + Oct $7,500 L5 + Dec $4,700.26 L5. Consultant never to I8. "
    "Consultant Fees July $334.17 + October $658.63 = Wise expertise write-ups "
    "LOCKED ($992.80). EOEB remainder $108,270 is per-invoice UNALLOCATED on the "
    "Deal Classifier (proposed FEE) — not dumped. Erdal 11/7 $6,000 FORTUNA "
    "PAYMENT still ASK. Aysel Monarch +$50k vs native Failed −$50k still ASK "
    "(not the belt). Koziol/Ariadne $150,000 LOCKED pass-through — not P&L. "
    "Coinbase is Investments. Newstar $23,581 is Art Sales COGS. Dec Wise "
    "$565.75 reimbursed. Aquinas books Cost TBD. Joint Sale Price = EPGC "
    "proceeds, not full hammer. GCM 1099 is personal. Not tax advice."
)
LEDGER_BANNER = (
    "Mercury Checking 8291 — EPGC LLC 2025. 64 unique cash txns. "
    "Deal Classifier: SALE vs FEE vs PASS vs UNALLOCATED. "
    "EOEB 2/21 $105,000 = Deal B Plutus SALE $50,000 + Deal A Jamal SALE $40,000 "
    "+ $15,000 FEE. Plutus 2/24–25 $50,000 = Deal B Cost. Jamal Deal A Cost from "
    "personal BoA ASK (do not invent $40k wire). 2/24 $90k EPGC→9922 = owner "
    "transfer, not extra Cost (draws $261,172). Roman Gold Belt = SALE. Venus = "
    "joint SALE. August Fortuna = joint SALE. L5 leftover $37,396.32 = FEE. "
    "EOEB 1/16 $1,000 Canosan = FEE (not horse SALE). EOEB 5/1 $21,000 NBerk = "
    "FEE (not Venus). Remaining EOEB $108,270 per-invoice UNALLOCATED (proposed "
    "FEE — not dumped). David Aaron = FEE. Koziol/Ariadne = PASS. 11/7 $6,000 "
    "FORTUNA PAYMENT still ASK. Not tax advice."
)
DS_NOTE = (
    "Mosaics 2026-09-21 LOCKED split: EOEB 2/21 $105,000 = Deal B Plutus/Antiquarium "
    "SALE $50,000 (Cost $50,000 on 8291, net $0) + Deal A Jamal SALE $40,000 (Cost "
    "ASK, not in I8) + $15,000 FEE (Consultant February, not I8 profit). Do not "
    "invent a $40k BoA Jamal Cost wire. 2/24 $90k EPGC→9922 is transfer, not extra "
    "Cost. Do not dump EOEB remainder $108,270. Not tax advice."
)


def start_here_mercury() -> str:
    return f"""START HERE — Mercury Bank 8291 (FEE lock)

Not tax advice. Same Personal Income.xlsx tabs as last year. Organizational packet only.

Mercury Checking 8291 is the EPGC LLC operating account (Choice Financial).
64 unique 2025 cash transactions from Monarch IDs.

OPEN THIS 11-tab workbook (numbers; Google convert dropped Art object labels):
{OPEN_XLSX}
07_income copy:
{OPEN_07}
root copy:
{OPEN_ROOT}

Art Sales NAMES (use this for object labels — Deal A / Deal B / Belt / Venus / August):
{LABELED_ART}

Deal Classifier (per-invoice EOEB; remainder $108,270 UNALLOCATED — do not dump):
{CLASSIFIER}

ASK (Q10 split LOCKED; Jamal BoA Cost still ASK):
{ASK}

LOCK (Q6–Q10 + Canosan + NBerk + L5 leftover FEEs):
{LOCK}

Fortuna inventory LOCKED SOLD 2025 (Belt, Venus, August — not parked unsold):
{FORTUNA}

LOCKED this pass
- Q10 mosaics. EOEB 2/21 $105,000 MAKE A PAYMENT = Deal B Plutus/Antiquarium SALE $50,000 + Deal A Jamal SALE $40,000 + $15,000 FEE.
  Deal B — Plutus & Mnemosyne / DBA Antiquarium. Cost $50,000 on Mercury 8291 (native COGS 2/24 $10 + 2/25 $49,990). Net $0.
  Deal A — Three Ancient Mosaics / Jamal Rifai. Sale $40,000. Cost from personal BoA ASK. $10 Mercury tests to Jamal 2/21 and 2/24 only. Do not invent a $40k BoA Cost wire.
  $15,000 FEE → EPGC Consultant February. Not I8 profit.
  2/24 $90k EPGC→9922 is owner transfer, not extra Cost. BoA 9922 draws $261,172.
- L5 leftover $37,396.32 ALL FEES → EPGC Consultant (not I8 / not Art Sales).
  7/11 $8,534.79; 7/15 $2,500 (5% of belt $50k window); 9/2 $5,042; 9/23 $9,119.27; 10/21 $7,500 (5% of Koziol/Ariadne $150k window); 12/4 $4,700.26.
  Belt 7/15 $50,000 stays SALE. Koziol/Ariadne $150,000 stays PASS. No Ariadne object SALE on the Art register.
- EOEB 1/16 $1,000 Canosan FEE. Cash is EOEB. Register horse is Hindman $900 / Erdal $1,000 — do not double-count as a second horse SALE.
- EOEB 5/1 $21,000 NBerk FEE. Not Venus (Venus is Fortuna 5/9 −$20,000 / Erdal 7/24 +$26,000).
- Roman Gold Belt SALE: Fortuna Cost $45,000 / L5 Sale $50,000 / net $5,000 → EPGC Art Sales July + I8.
- Venus joint SALE: Fortuna Cost $20,000 / Erdal proceeds $26,000 / net $6,000 → EPGC Art Sales July + I8.
- August Fortuna joint SALE: Fortuna Cost $20,000 / Erdal proceeds $25,000 / net $5,000 → EPGC Art Sales October + I8.
- David Aaron Limited $13,595 (6/24) → EPGC Consultant June. FEE, not a sale.
- Koziol $150,000 (10/20) / Ariadne $150,000 (10/21) → PASS. Not P&L.
- Wise expertise write-ups $992.80 → EPGC Consultant Fees.
- EOEB 12/2 $565.75 reimbursed — not income.

EPGC 2025 Art Sales LOCKED cash $236,000 (Jan $14,000 / Feb $121,000 / July $76,000 / Oct $25,000).
EPGC Consultant $87,991.32 (Canosan + mosaics fee + NBerk + David Aaron + L5 leftover).
Income I8 Art net is $24,055.06 (mosaics $15k is FEE, not I8). I10 EBITDA =SUM(I4:I9) → $152,760.

ASK remaining — do not dump
1. EOEB remainder $108,270.00 — per-invoice UNALLOCATED (proposed FEE, no matching Cost). Fortuna outs are NOT EOEB Cost.
2. Erdal 11/7 $6,000 FORTUNA PAYMENT — not August, not Venus.
3. Aysel Dere — Monarch 7/18 +$50,000 IN vs native Failed 7/17 −$50,000 OUT. Not the belt.
4. Jamal Deal A Cost amount on the 9922 statement (no labeled $40k OUT in Feb).

Sale 6428 $11,000 and Aquinas books $1,000 wait on Cost (not in I8).
August object name still unknown.
"""


def start_here_all() -> str:
    return f"""START HERE — last year's spreadsheet, 2025 numbers filled
Not tax advice. Organizational packet only.

OPEN THIS ONE spreadsheet (same 11 tab names as last year, 2023 | 2024 | 2025 year blocks, 2025 numbers, formulas that calculate; I8 $24,055.06; EPGC July $76,000 + Oct $25,000; mosaics SALE $90,000 + FEE $15,000). Google convert kept Art numbers but dropped object labels:
{OPEN_XLSX}

07_income copy:
{OPEN_07}

2025 Taxes root copy:
{OPEN_ROOT}

Art Sales NAMES (2022–2024 lots + Deal A Jamal + Deal B Plutus + Belt / Venus / August):
{LABELED_ART}

Deal Classifier (per-invoice EOEB — SALE/FEE/PASS/UNALLOCATED; remainder $108,270 NOT dumped):
{CLASSIFIER}

ASK (Q10 split LOCKED; Jamal BoA Cost still ASK):
{ASK}

Fortuna inventory LOCKED SOLD 2025 (Belt, Venus, August — not parked unsold):
{FORTUNA}

Tabs in last year's order (all inside that one file)
1. Income — 2022–2025. 2025 Actual: Hindman $70,618 / 216 $19,500 / 524 #2 $17,086.94 / GCM $21,500 / Art $24,055.06. I10 EBITDA =SUM(I4:I9) → $152,760.
2. 524 Ferdinand Ave, Unit 2
3. 216 N. Oak Park Ave
4. EPGC LLC — 2025 Art Sales $14,000 + $121,000 + July $76,000 + October $25,000 = $236,000; Consultant $87,991.32 (Canosan + mosaics fee + NBerk + David Aaron + L5 leftover FEEs).
5. Refrence Library
6. Art Sales and Purchases — 2022–2024 lots, then Sales 2025 including Deal B Plutus/Antiquarium (Cost $50k / Sale $50k) + Deal A Jamal (Sale $40k, Cost ASK) + Roman Gold Belt + Venus joint + August Fortuna joint. Do not dump EOEB $108,270.
7. GCM
8. Hindman W2
9. Megan T4
10. Investments — Coinbase ACH net -$51,000.
11. 524 Ferdinand Ave, Unit 1

Mosaics (user 2026-09-21 lock): EOEB 2/21 $105,000 = $50k Antiquarium SALE + $40k Jamal SALE + $15k FEE. $15k is Consultant February, not I8 profit. Plutus $50k on Mercury is Deal B Cost. Jamal Cost from personal BoA ASK — do not invent a $40k BoA Cost wire. Do not double-count the 2/24 $90k EPGC→9922 transfer as extra Cost. BoA 9922 draws $261,172.

Deal Classifier: SALE vs FEE vs PASS vs UNALLOCATED.
Belt = SALE. Venus = SALE (joint). August Fortuna = SALE (joint). Mosaics = two SALE deals ($90k) + FEE ($15k). L5 leftover = FEE. Canosan EOEB $1k = FEE. NBerk EOEB $21k = FEE. EOEB remainder = UNALLOCATED per invoice (proposed FEE). David Aaron = FEE. Koziol/Ariadne = PASS.

START HERE Doc:
{START_HERE}

Not tax advice.
"""


def paint(cell, value=None, fill=None, num=None):
    if value is not None:
        cell.value = value
    if fill is not None:
        cell.fill = fill
    if num is not None:
        cell.number_format = num
    cell.font = CG


def copy_style(src, dest):
    if src.has_style:
        dest.font = copy(src.font)
        dest.border = copy(src.border)
        dest.fill = copy(src.fill)
        dest.number_format = src.number_format
        dest.alignment = copy(src.alignment)


def find_row(ws, pred, start=1, end=140):
    for r in range(start, end + 1):
        if pred(str(ws.cell(r, 1).value or "")):
            return r
    return None


def patch_xlsx() -> None:
    wb = load_workbook(XLSX)
    art = wb["Art Sales and Purchases"]
    if art["A74"].value is None or "mosaic" not in str(art["A74"].value).lower():
        raise SystemExit(f"Art A74 is {art['A74'].value!r}")

    already = str(art["A74"].value or "").startswith("Mosaic Deal B")
    if str(art["A78"].value or "") == "Roman Gold Belt":
        art.insert_rows(75)
        for c in range(1, 12):
            copy_style(art.cell(74, c), art.cell(75, c))
        art.row_dimensions[75].height = 96
    elif not already:
        raise SystemExit(f"Art A78 is {art['A78'].value!r} — insert already applied?")
    # insert_rows does not always shift merges; Total landed under A82:I82.
    if "A82:I82" in {str(m) for m in art.merged_cells.ranges}:
        art.unmerge_cells("A82:I82")
    if "A83:I83" not in {str(m) for m in art.merged_cells.ranges}:
        try:
            art.merge_cells("A83:I83")
        except Exception:
            pass

    paint(art["A74"], "Mosaic Deal B — Plutus/Antiquarium")
    paint(art["B74"], "Plutus & Mnemosyne / DBA Antiquarium")
    paint(art["C74"], 50000, fill=GREEN, num=ACCT)
    paint(art["D74"], date(2025, 2, 25))
    paint(art["E74"], "EOEB LLC")
    paint(art["F74"], 50000, fill=GREEN, num=ACCT)
    paint(art["G74"], date(2025, 2, 21))
    paint(art["H74"], "=F74-C74", num=ACCT)
    paint(art["I74"], DEAL_B_NOTE)
    art["I74"].alignment = WRAP
    paint(art["K74"], "SALE")
    art.row_dimensions[74].height = 96

    paint(art["A75"], "Mosaic Deal A — Three Ancient Mosaics / Jamal Rifai")
    paint(art["B75"], "Jamal M. Rifai (Cost ASK — personal BoA)")
    art["C75"].value = None
    paint(art["C75"], fill=YELLOW, num=ACCT)
    art["D75"].value = None
    paint(art["E75"], "EOEB LLC")
    paint(art["F75"], 40000, fill=GREEN, num=ACCT)
    paint(art["G75"], date(2025, 2, 21))
    paint(art["H75"], '=IF(OR(F75="",C75=""),"",F75-C75)', num=ACCT)
    paint(art["I75"], DEAL_A_NOTE, fill=YELLOW)
    art["I75"].alignment = WRAP
    paint(art["K75"], "SALE")

    # Shifted H formulas (insert_rows does not rewrite references).
    formula_by_name = {
        "Canosan Terracotta Horse": '=SUM(F{r}-C{r})',
        "Sale 6428 Contract 303468 (object TBD)": '=IF(C{r}="","",F{r}-C{r})',
        "Books (titles TBD)": '=IF(OR(F{r}="",C{r}=""),"",F{r}-C{r})',
        "Roman Gold Belt": "=F{r}-C{r}",
        "Bronze head of a goddess, likely Venus (joint; EPGC share)": "=F{r}-C{r}",
        "August Fortuna joint — unnamed object (EPGC share)": "=F{r}-C{r}",
        "Total": "=SUM(H51:H76)+H79+H80+H81",
    }
    for r in range(76, 90):
        a = str(art.cell(r, 1).value or "")
        for key, tmpl in formula_by_name.items():
            if a == key or a.startswith(key):
                paint(art.cell(r, 8), tmpl.format(r=r), num=ACCT)
                break

    canosan_r = find_row(art, lambda a: a.startswith("Canosan Terracotta Horse"))
    if canosan_r:
        note = (
            "Register horse is Hindman Cost $900 sold to Erdal $1,000 (OPEN vs Berk 2/7 lump). "
            "EOEB 1/16 $1,000 is a FEE on Mercury 8291 — not this horse SALE. Do not double-count."
        )
        paint(art.cell(canosan_r, 9), note, fill=YELLOW)
        art.cell(canosan_r, 9).alignment = WRAP

    footer_r = find_row(art, lambda a: a.startswith("Mercury 8291 MATCHED"))
    if footer_r:
        paint(art.cell(footer_r, 1), ART_FOOTER, fill=YELLOW)
        art.cell(footer_r, 1).alignment = WRAP
        art.row_dimensions[footer_r].height = 96

    belt_r = find_row(art, lambda a: a == "Roman Gold Belt", start=70)
    venus_r = find_row(art, lambda a: "Venus" in a, start=70)
    aug_r = find_row(art, lambda a: a.startswith("August Fortuna"), start=70)
    total_r = find_row(art, lambda a: a == "Total", start=70)
    if belt_r != 79 or venus_r != 80 or aug_r != 81 or total_r != 82:
        raise SystemExit(f"row shift unexpected: belt={belt_r} venus={venus_r} aug={aug_r} total={total_r}")

    inc = wb["Income"]
    paint(inc["I8"], I8, fill=GREEN, num=ACCT)
    paint(inc["A12"], INCOME_NOTE)
    inc["A12"].alignment = WRAP

    ep = wb["EPGC LLC"]
    paint(ep["C53"], FEB_ART, fill=GREEN, num=ACCT)
    for i, amt in enumerate(CONSULTANT_MONTHLY):
        cell = ep.cell(54, 2 + i)
        fill = GREEN if amt else PEACH
        paint(cell, amt, fill=fill, num=ACCT)
    for r in range(70, 90):
        v = str(ep.cell(r, 1).value or "")
        if "2025 Mercury" in v or "Art Sales gross" in v:
            paint(ep.cell(r, 1), EPGC_NOTE)
            ep.cell(r, 1).alignment = WRAP
            break

    ds = wb["2025 Data sources"]
    paint(ds["C11"], "24,055.06")
    paint(
        ds["D11"],
        "Berk lots + seals $1,000 + Belt net $5,000 + Venus EPGC-share net $6,000 + August Fortuna joint net $5,000. "
        "Mosaics $15,000 is FEE (Consultant February), not I8. Deal A Cost ASK. Sale 6428 and Aquinas wait on Cost.",
    )
    paint(ds["E11"], "LOCKED green")
    paint(
        ds["C24"],
        "14,000 Jan + 121,000 Feb + 76,000 July + 25,000 Oct = 236,000",
    )
    paint(
        ds["D24"],
        "Gross sale cash. Feb = Berk $30,000 + Aquinas $1,000 + mosaics SALE $90,000 (not the $15k FEE). "
        "July = Belt L5 $50,000 + Venus $26,000. October = August Fortuna joint remittance $25,000. "
        "Do not dump EOEB $108,270.",
    )
    paint(ds["B25"], "Consultant year")
    paint(ds["C25"], "87,991.32")
    paint(
        ds["D25"],
        "Jan $1,000 Canosan + Feb $15,000 mosaics fee + May $21,000 NBerk + June $13,595 David Aaron + "
        "July $11,034.79 L5 + Sep $14,161.27 L5 + Oct $7,500 L5 + Dec $4,700.26 L5. Never I8.",
    )
    paint(ds["C43"], 236000)
    paint(ds["C44"], 87991.32)
    paint(
        ds["A34"],
        "Dump EOEB $108,270 onto P&L",
    )
    paint(
        ds["B34"],
        "EOEB remainder is per-invoice UNALLOCATED (proposed FEE, no matching Cost). Not dumped. "
        "Named FEEs are locked (Canosan $1k, NBerk $21k, mosaics $15k, L5 leftover $37,396.32). "
        "Belt $50k SALE; remaining L5 is FEE not attached to the belt.",
    )
    paint(ds["D50"], "Jamal Deal A Cost ASK")
    paint(
        ds["C52"],
        "David Aaron June $13,595; mosaics $15,000 Feb; Canosan $1,000 Jan; NBerk $21,000 May; "
        "L5 leftover $37,396.32; Wise write-ups $992.80 Fees",
    )
    paint(ds["D52"], "EOEB remainder $108,270 proposed FEE but not locked")
    paint(ds["C54"], "EOEB remainder $108,270.00 per-invoice (see DEAL_CLASSIFIER.csv)")
    paint(ds["D54"], "Erdal 11/7 $6,000; Aysel $50,000")
    paint(
        ds["C55"],
        "EOEB remainder $108,270.00; Erdal 11/7 $6,000; Aysel $50,000",
    )
    wrote = False
    for r in range(1, 60):
        v = str(ds.cell(r, 1).value or "")
        if v.startswith("Mosaics 2026-09-21"):
            paint(ds.cell(r, 1), DS_NOTE, fill=YELLOW)
            wrote = True
    if not wrote:
        ds.cell((ds.max_row or 1) + 1, 1, DS_NOTE)

    try:
        wb.defined_names.clear()
    except Exception:
        pass
    wb._external_links = []
    wb.save(XLSX)
    sanitize_xlsx(XLSX)
    DELIVERABLE.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(XLSX, DELIVERABLE)

    def sheet_to_csv(ws, path: Path) -> None:
        max_r, max_c = ws.max_row or 1, ws.max_column or 1
        rows = []
        for r in range(1, max_r + 1):
            rows.append([ws.cell(r, c).value for c in range(1, max_c + 1)])
        with path.open("w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(rows)

    sheet_to_csv(art, CSV_DIR / "Art_Sales.csv")
    sheet_to_csv(inc, CSV_DIR / "Income.csv")
    sheet_to_csv(ep, CSV_DIR / "EPGC_LLC.csv")


def patch_ledger() -> None:
    path = CSV_DIR / "MERCURY_LEDGER.csv"
    rows = list(csv.reader(path.open(encoding="utf-8")))
    rows[0] = [LEDGER_BANNER] + [""] * max(0, len(rows[0]) - 1)
    header = rows[2]
    idx = {name: i for i, name in enumerate(header)}
    date_i, who_i, amt_i = idx["Date"], idx["Counterparty"], idx["Amount"]
    notes_i = idx["Notes"]
    deal_i = idx["Deal / object"]
    bkt_i = idx["Bucket"]
    st_i = idx["Status"]
    clf_i = idx.get("Classifier")
    hits_art = idx.get("Hits EPGC Art Sales?")
    hits_con = idx.get("Hits EPGC Consultant?")

    def amt_of(r):
        try:
            return abs(float(r[amt_i]))
        except Exception:
            return None

    l5_notes = {inv[0]: inv for inv in L5_INVOICES}

    for r in rows[3:]:
        if len(r) <= notes_i or not r[date_i]:
            continue
        date_s, who, amt = r[date_i], r[who_i], r[amt_i]
        if date_s == "2025-02-21" and who == "EOEB LLC" and amt.startswith("105000"):
            r[bkt_i] = "ART_SALE+FEE"
            if clf_i is not None:
                r[clf_i] = "SALE $90k + FEE $15k"
            r[st_i] = "LOCKED"
            r[deal_i] = "Two mosaic deals + fee — $50k Plutus SALE + $40k Jamal SALE + $15k FEE"
            if hits_art is not None:
                r[hits_art] = "YES ($90,000)"
            if hits_con is not None:
                r[hits_con] = "YES ($15,000)"
            r[notes_i] = (
                "EOEB LLC MAKE A PAYMENT $105,000. User lock: working dollars $50k Antiquarium "
                "+ $40k Jamal + $15k FEE. Hits EPGC Art Sales February $90,000 + Consultant "
                "February $15,000. Deal B Cost $50,000 on 8291. Deal A Cost ASK (personal BoA). "
                "Do not invent a $40k BoA Cost wire. $15k is FEE not I8 profit. Not tax advice."
            )
        if date_s == "2025-02-25" and "Plutus" in who:
            r[bkt_i] = "ART_PURCHASE"
            if clf_i is not None:
                r[clf_i] = "SALE"
            r[st_i] = "LOCKED Cost"
            r[deal_i] = "Mosaic Deal B — Plutus/Antiquarium Cost $50,000 (Sale $50,000, net $0)"
            r[notes_i] = (
                "Native Mercury: DBA Antiquarium (Plutus & Mnemosyne, LLC) COGS. "
                "With the 2/24 $10 test this is $50,000 Cost of Deal B. Matching Sale is "
                "$50,000 of EOEB 2/21 $105,000. Net $0. Not Jamal. Not tax advice."
            )
        if date_s == "2025-02-24" and "Plutus" in who:
            r[deal_i] = "Mosaic Deal B — $10 test of Plutus/Antiquarium $50,000 Cost"
            r[st_i] = "LOCKED Cost"
            r[notes_i] = (
                "Native: DBA Antiquarium (Plutus & Mnemosyne, LLC). First $10 of Deal B Cost. "
                "Not Jamal. $49,990 follows 2/25."
            )
        if date_s in ("2025-02-21", "2025-02-24") and "Jamal" in who:
            r[deal_i] = "$10 test wire to Jamal M. Rifai (Deal A counterparty)"
            r[notes_i] = (
                "Not P&L. Mercury could not send the Jamal Cost here — user paid Deal A from "
                "personal BoA (amount ASK). Distinct from Plutus/Antiquarium Deal B. Do not "
                "invent a $40k BoA Cost wire."
            )
        if date_s == "2025-01-16" and who == "EOEB LLC" and amt.startswith("1000"):
            r[bkt_i] = "ADVISORY"
            if clf_i is not None:
                r[clf_i] = "FEE"
            r[st_i] = "LOCKED"
            r[deal_i] = "EOEB Canosan FEE $1,000 (not horse SALE)"
            if hits_art is not None:
                r[hits_art] = "no — not dumped as horse SALE"
            if hits_con is not None:
                r[hits_con] = "YES"
            r[notes_i] = (
                "LOCKED FEE. Cash is EOEB. Art Sales register Canosan horse is Hindman $900 / "
                "Erdal $1,000 (and Berk 2/7 lump) — do not double-count this $1,000 as a second "
                "horse SALE. Hits EPGC Consultant January. Not tax advice."
            )
        if date_s == "2025-05-01" and who == "EOEB LLC" and amt.startswith("21000"):
            r[bkt_i] = "ADVISORY"
            if clf_i is not None:
                r[clf_i] = "FEE"
            r[st_i] = "LOCKED"
            r[deal_i] = "EOEB NBerk FEE $21,000 (not Venus)"
            if hits_art is not None:
                r[hits_art] = "no"
            if hits_con is not None:
                r[hits_con] = "YES"
            r[notes_i] = (
                "LOCKED FEE on the NBerk deal. Not Venus (Venus Cost is Fortuna 5/9 −$20,000 / "
                "Erdal 7/24 +$26,000). Hits EPGC Consultant May. Not tax advice."
            )
        if who == "L5" and amt_of(r) == 50000.0:
            continue
        if who == "L5" and date_s in l5_notes:
            _d, dollars, label = l5_notes[date_s]
            r[bkt_i] = "ADVISORY"
            if clf_i is not None:
                r[clf_i] = "FEE"
            r[st_i] = "LOCKED"
            r[deal_i] = label
            if hits_art is not None:
                r[hits_art] = "no — not I8 / not Art Sales"
            if hits_con is not None:
                r[hits_con] = "YES"
            extra = ""
            if dollars == 2500.0:
                extra = " 5% of belt $50,000 SALE window — FEE, not attached to the belt SALE."
            if dollars == 7500.0:
                extra = (
                    " Same calendar day as Ariadne −$150,000 but different time/direction/"
                    "counterparty. 5% of Koziol/Ariadne $150,000 — FEE, not an Ariadne object SALE."
                )
            r[notes_i] = (
                f"LOCKED FEE ${dollars:,.2f}. User: L5 leftovers are all fees. Hits EPGC "
                f"Consultant (not I8). Belt $50k SALE and Koziol/Ariadne PASS stay.{extra} "
                "Not tax advice."
            )
    with path.open("w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)


def patch_classifier() -> None:
    path = CSV_DIR / "DEAL_CLASSIFIER.csv"
    rows = list(csv.reader(path.open(encoding="utf-8")))
    out = []
    for r in rows:
        if not r:
            out.append(r)
            continue
        line0 = str(r[0])
        # Banner
        if line0.startswith("Deal Classifier"):
            r[0] = (
                "Deal Classifier — Mercury 8291 / EPGC 2025. Columns: date, amount, "
                "object/invoice, matching Cost out?, proposed bucket SALE/FEE/PASS/UNALLOCATED, "
                "evidence, lock status. Last-year model: Art Sales = gross sale cash by month; "
                "Income I8 = art NET for lots with Cost; Consultant never goes to I8. Do NOT "
                "dump EOEB remainder $108,270 as a lump onto Consultant or Art Sales. Not tax advice."
            )
            out.append(r)
            continue
        if line0 == "2025-01-16":
            out.append(
                [
                    "2025-01-16",
                    "1000.00",
                    "EOEB Canosan FEE (cash is EOEB; register horse is Hindman/Erdal)",
                    "none on 8291 — this is not the horse SALE",
                    "FEE",
                    "LOCKED. User: 1k Canosan is a fee. Do not double-count the Art Sales Canosan horse (Hindman $900 / Erdal $1,000, also in Berk 2/7 lump). Hits EPGC Consultant January. Not tax advice.",
                    "LOCKED",
                ]
            )
            continue
        if line0 == "2025-02-21" and "mosaic" in ",".join(r).lower():
            out.append(
                [
                    "2025-02-21",
                    "50000.00",
                    "Mosaic Deal B / Plutus & Mnemosyne DBA Antiquarium — SALE slice of EOEB $105k",
                    "Deal B Cost LOCKED: Plutus/Antiquarium 2/24–25 $50,000 on 8291",
                    "SALE",
                    "User lock: working dollars $50k Antiquarium / $40k Jamal / $15k FEE. Net $0. Not Jamal DBA. Hits EPGC Art Sales February. Not tax advice.",
                    "LOCKED",
                ]
            )
            out.append(
                [
                    "2025-02-21",
                    "40000.00",
                    "Mosaic Deal A / Jamal Rifai — SALE slice of EOEB $105k",
                    "Deal A Cost ASK: personal BoA (do not invent a $40k wire). $10 tests 2/21 and 2/24 only. 2/24 $90k EPGC→9922 is transfer, not extra Cost.",
                    "SALE",
                    "Sale $40,000 LOCKED. Cost unknown so not in I8. Not dumped. Not tax advice.",
                    "LOCKED cash / ASK Cost",
                ]
            )
            out.append(
                [
                    "2025-02-21",
                    "15000.00",
                    "Mosaics FEE slice of EOEB $105k (not I8 profit)",
                    "n/a — fee not a sale",
                    "FEE",
                    "User: $15k is my fee. Hits EPGC Consultant February. Never I8. Not tax advice.",
                    "LOCKED",
                ]
            )
            continue
        if line0 == "2025-02-25" and "Plutus" in (r[2] if len(r) > 2 else ""):
            r[3] = "this row is Deal B Cost ($10 on 2/24 + $49,990 on 2/25)"
            r[4] = "SALE"
            r[5] = (
                "Native Mercury COGS. Matching Sale is the $50,000 Deal B slice of EOEB 2/21. "
                "Net $0. Not Jamal. Not tax advice."
            )
            r[6] = "LOCKED Cost"
            out.append(r)
            continue
        if line0 == "2025-05-01":
            out.append(
                [
                    "2025-05-01",
                    "21000.00",
                    "EOEB NBerk FEE $21,000 (NOT Venus)",
                    "none on 8291. Venus Cost is Fortuna 5/9 −$20,000 — different object, different counterparty.",
                    "FEE",
                    "LOCKED. User: 21k from NBerk deal. Hits EPGC Consultant May. Not a Berk object SALE (Berk cash is 1/10 $14k + 2/7 $30k). Not tax advice.",
                    "LOCKED",
                ]
            )
            continue
        if line0 == "remainder total":
            r[1] = "108270.00"
            r[2] = (
                "EOEB remainder after mosaics $105,000, Canosan FEE $1,000, NBerk FEE $21,000, "
                "and Wise reimburse $565.75"
            )
            r[4] = "UNALLOCATED"
            r[5] = (
                "Sum of remaining proposed-FEE invoices. Shown per invoice — not posted to EPGC "
                "Consultant or Art Sales. Do not dump."
            )
            r[6] = "UNALLOCATED (do not dump)"
            out.append(r)
            continue
        if line0 == "various" or (
            len(r) > 2 and "L5 remainder" in str(r[2])
        ):
            for d, amt, label in L5_INVOICES:
                out.append(
                    [
                        d,
                        f"{amt:.2f}",
                        label,
                        "n/a — fee not a sale",
                        "FEE",
                        "LOCKED. User: the L5 leftovers are all fees. Hits EPGC Consultant, not I8 / not Art Sales. Belt $50k SALE stays. Koziol/Ariadne PASS stays. Not tax advice.",
                        "LOCKED",
                    ]
                )
            continue
        out.append(r)
    with path.open("w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(out)


def patch_ask() -> None:
    path = CSV_DIR / "ASK_Mercury.csv"
    rows = [
        [
            "ASK — remaining Mercury 8291 questions after Q10 / Canosan / NBerk / L5 leftover FEE lock. "
            "EOEB remainder per-invoice UNALLOCATED (proposed FEE) — not dumped. Q6–Q10 LOCKED. Not tax advice.",
            "",
            "",
        ],
        ["", "", ""],
        ["Question", "What I need", "Classifier"],
        [
            "1. EOEB LLC remainder after named locks — per-invoice UNALLOCATED (proposed FEE, not dumped)",
            "$108,270.00 of EOEB IN is not mosaics $105,000, not Canosan FEE $1,000, not NBerk FEE $21,000, and not Wise reimburse $565.75. See DEAL_CLASSIFIER.csv for every invoice. Hypothesis: no matching dealer Cost → looks like FEE. Do NOT dump this remainder onto Consultant or Art Sales. Fortuna outs are NOT EOEB Cost.",
            "UNALLOCATED",
        ],
        [
            "2. L5 leftover $37,396.32 — LOCKED FEE",
            "User: the remainder are all fees. 7/11 $8,534.79; 7/15 $2,500 (5% of belt $50k); 9/2 $5,042; 9/23 $9,119.27; 10/21 $7,500 (5% of Koziol/Ariadne $150k); 12/4 $4,700.26. Hits EPGC Consultant, not I8. Belt $50k SALE stays. No Ariadne object SALE.",
            "FEE (LOCKED)",
        ],
        [
            "3. August Fortuna joint — LOCKED SALE (object still unnamed)",
            "Fortuna 8/15 (native 8/18) −$20,000 Cost / Erdal 10/9 +$25,000 proceeds ($20k capital + $5k earnings). Hits EPGC Art Sales October. Object name still unknown. Not consultant. Not unsold inventory.",
            "SALE (joint, EPGC share)",
        ],
        [
            "4. Erdal Dere 11/7 $6,000 FORTUNA PAYMENT — still ASK",
            "$6,000.00. Not the August joint (10/9 $25,000). Not Venus (7/24 $26,000).",
            "UNALLOCATED",
        ],
        [
            "5. Aysel Dere $50,000 — Monarch IN vs native Failed OUT",
            "Monarch 7/18 +$50,000 IN is in the 64. Native Failed $50,000 OUT 7/17. Not the belt. Do not dump.",
            "UNALLOCATED",
        ],
        [
            "6. David Aaron Limited $13,595 IN (6/24) — LOCKED consultant fee",
            "CONFIRMED: consultant fee, not a sale. Booked EPGC Consultant June $13,595. Not Art Sales. Not I8.",
            "FEE",
        ],
        [
            "7. Wise $334.17 (7/9) and $658.63 (10/30) — LOCKED expertise write-ups",
            "EPGC Consultant Fees $992.80. Distinct from Dec Wise $565.75 reimbursed by EOEB.",
            "FEE",
        ],
        [
            "8. Jack Koziol & Tracy Hoffman $150,000 IN (10/20) / Ariadne Demirjian $150,000 OUT (10/21) — LOCKED pass-through",
            "CONFIRMED: pass-through, not P&L. L5 10/21 $7,500 is a FEE (5% of this window), not a second Ariadne SALE.",
            "PASS",
        ],
        [
            "9. Aquinas Hobor $1,000 IN (2/4) — LOCKED book sale",
            "Hits EPGC Art Sales February. Cost / titles TBD (not in I8 until Cost).",
            "SALE",
        ],
        [
            "10. TWO mosaic deals + fee — LOCKED split (Jamal BoA Cost still ASK)",
            "User: it's Jamal, Antiquarium, and my 15k fee. Working dollars $50k Antiquarium SALE + $40k Jamal SALE + $15k FEE. Plutus Cost $50k on 8291 LOCKED. Jamal Cost from personal BoA ASK — do not invent a $40k BoA wire. $15k FEE → Consultant February, not I8. Aysel $50k Failed (not mosaics Cost).",
            "SALE $90k + FEE $15k (Cost ASK on Deal A)",
        ],
        ["", "", ""],
        ["LOCKED / MATCHED this pass (do not recast without saying so)", "", "Classifier"],
        [
            "Roman Gold Belt",
            "Fortuna 7/21 −$45,000 Cost / L5 7/15 +$50,000 Sale. Net $5,000. Remaining L5 is FEE, not attached.",
            "SALE",
        ],
        [
            "Venus (bronze head, joint)",
            "Fortuna 5/9 −$20,000 / Erdal 7/24 +$26,000. Sale Price = EPGC proceeds.",
            "SALE (joint, EPGC share)",
        ],
        [
            "August Fortuna joint (unnamed)",
            "Fortuna 8/15 −$20,000 / Erdal 10/9 +$25,000. EPGC Art Sales October. Net $5,000.",
            "SALE (joint, EPGC share)",
        ],
        [
            "Mosaic Deal B (Plutus/Antiquarium)",
            "Cost $50,000 on 8291 / Sale $50,000 of EOEB 2/21. Net $0.",
            "SALE",
        ],
        [
            "Mosaic Deal A (Jamal)",
            "Sale $40,000 of EOEB 2/21. Cost from personal BoA ASK (not in I8).",
            "SALE (Cost ASK)",
        ],
        [
            "Mosaics $15,000",
            "FEE slice of EOEB 2/21 → EPGC Consultant February. Not I8 profit.",
            "FEE",
        ],
        [
            "EOEB Canosan $1,000",
            "1/16 FEE → Consultant January. Not a second horse SALE.",
            "FEE",
        ],
        [
            "EOEB NBerk $21,000",
            "5/1 FEE → Consultant May. Not Venus.",
            "FEE",
        ],
        [
            "L5 leftover $37,396.32",
            "Six invoices LOCKED FEE → Consultant. Belt $50k SALE stays.",
            "FEE",
        ],
        [
            "David Aaron Limited",
            "$13,595 (6/24) consultant fee → EPGC Consultant June. Not a sale.",
            "FEE",
        ],
        [
            "Koziol / Ariadne",
            "$150,000 IN 10/20 + $150,000 OUT 10/21 LOCKED pass-through — not P&L",
            "PASS",
        ],
        [
            "Wise/EOEB 12/2",
            "$565.75 reimbursed — not income, not expense",
            "n/a",
        ],
        [
            "EOEB remainder",
            "$108,270.00 per-invoice UNALLOCATED (proposed FEE). Not dumped.",
            "UNALLOCATED",
        ],
    ]
    with path.open("w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)


def patch_lock_q() -> None:
    path = CSV_DIR / "LOCK_Mercury_Q6_Q9.csv"
    rows = [
        ["Q", "Status", "What", "Where it hits"],
        [
            "6",
            "LOCKED",
            "David Aaron Limited $13,595 IN (6/24) consultant fee",
            "EPGC Consultant June $13,595. Not Art Sales. Not I8. Classifier: FEE.",
        ],
        [
            "7",
            "LOCKED",
            "Wise $334.17 (7/9) + $658.63 (10/30) expertise write-ups",
            "EPGC Consultant Fees $992.80. Classifier: FEE.",
        ],
        [
            "8",
            "LOCKED",
            "Koziol $150,000 IN (10/20) / Ariadne $150,000 OUT (10/21)",
            "Pass-through — not P&L. Classifier: PASS. L5 10/21 $7,500 is FEE (5%), not a second SALE.",
        ],
        [
            "9",
            "LOCKED",
            "Aquinas Hobor $1,000 IN (2/4) book SALE",
            "EPGC Art Sales February. Cost TBD. Classifier: SALE.",
        ],
        [
            "10",
            "LOCKED split / ASK Jamal Cost",
            "EOEB 2/21 $105,000 = Deal B $50k SALE + Deal A $40k SALE + $15k FEE",
            "Art Sales February +$90,000. Consultant February +$15,000. Deal B Cost $50k net $0. Deal A Cost ASK (not in I8). I8 drops the $15k.",
        ],
        [
            "Canosan",
            "LOCKED",
            "EOEB 1/16 $1,000 Canosan FEE (not horse SALE)",
            "EPGC Consultant January $1,000. Do not double-count Hindman/Erdal horse register.",
        ],
        [
            "NBerk",
            "LOCKED",
            "EOEB 5/1 $21,000 NBerk FEE (not Venus)",
            "EPGC Consultant May $21,000. Classifier: FEE.",
        ],
        [
            "L5 leftover",
            "LOCKED",
            f"L5 leftover ${L5_FEES:,.2f} ALL FEES",
            "EPGC Consultant Jul $11,034.79 / Sep $14,161.27 / Oct $7,500 / Dec $4,700.26. Not I8. Belt $50k SALE stays.",
        ],
        [
            "Belt",
            "LOCKED",
            "Roman Gold Belt Fortuna $45,000 / L5 $50,000",
            "EPGC Art Sales July +$50,000. I8 net +$5,000. Classifier: SALE. L5 7/15 $2,500 is FEE (5%), not attached.",
        ],
        [
            "Venus",
            "LOCKED",
            "Venus Fortuna $20,000 / Erdal remittance $26,000",
            "EPGC Art Sales July +$26,000. I8 net +$6,000. Classifier: SALE (joint, EPGC share).",
        ],
        [
            "August",
            "LOCKED",
            "August Fortuna joint Fortuna $20,000 / Erdal 10/9 $25,000",
            "EPGC Art Sales October +$25,000. I8 net +$5,000. Classifier: SALE (joint, EPGC share). 11/7 $6,000 still ASK.",
        ],
        [
            "EOEB 12/2",
            "LOCKED not income",
            "EOEB/Wise $565.75 reimburse",
            "Not P&L.",
        ],
        [
            "1",
            "UNALLOCATED per invoice (proposed FEE)",
            f"EOEB remainder ${EOEB_REMAINING:,.2f}",
            "DEAL_CLASSIFIER.csv. Do NOT dump onto Consultant or Art Sales.",
        ],
        [
            "LOCKED totals",
            "LOCKED",
            "Art Sales gross $236,000 (Jan 14k + Feb 121k + July 76k + Oct 25k); "
            "I8 $24,055.06 (mosaics $15k is FEE); Consultant $87,991.32",
            "EOEB remainder $108,270 not on P&L.",
        ],
    ]
    with path.open("w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)


def patch_json() -> None:
    packet = json.loads(PACKET.read_text(encoding="utf-8"))
    packet["updated"] = "2026-09-21"
    note = (
        "2026-09-21 FEE lock: EOEB 2/21 $105,000 = Deal B Plutus SALE $50k + Deal A Jamal "
        "SALE $40k + $15k FEE (Consultant February, not I8). I8 $24,055.06. Feb Art Sales "
        "$121,000. Year Art $236,000. Consultant $87,991.32 (Canosan $1k + mosaics fee $15k "
        "+ NBerk $21k + David Aaron $13,595 + L5 leftover $37,396.32). EOEB remainder "
        "$108,270 UNALLOCATED — not dumped. Jamal BoA Cost ASK (do not invent $40k wire). "
        "Belt/Venus/August SALE unchanged. Not tax advice."
    )
    packet["xlsx_note"] = note
    m = packet.setdefault("mercury_8291", {})
    m["updated"] = "2026-09-21"
    locked = m.setdefault("locked", {})
    locked["mosaics_sale"] = MOSAICS_SALE
    locked["mosaics_sale_wire"] = 105000.0
    locked["mosaics_fee"] = MOSAICS_FEE
    locked["mosaics_plutus_sale"] = PLUTUS_SALE
    locked["mosaics_jamal_sale"] = JAMAL_SALE
    locked["mosaics_plutus_8291"] = 50000.0
    locked["mosaics_plutus_net"] = 0.0
    locked["mosaics_cost"] = 50000.0
    locked.pop("mosaics_profit", None)
    locked["mosaics_profit_in_i8"] = 0.0
    locked["mosaics_two_deals"] = True
    locked["mosaics_cost_note"] = (
        "Deal B Plutus/Antiquarium Cost $50,000 LOCKED on 8291 (Sale $50,000, net $0). "
        "Deal A Jamal Sale $40,000 LOCKED; Cost from personal BoA ASK — do not invent a "
        "$40k BoA Cost wire. Combined $90,000 was a WORKING Cost that treated the $15k "
        "as art profit; recast as FEE."
    )
    locked["boa_9922_draws"] = 261172.0
    locked["l5_leftover_fees"] = L5_FEES
    locked["eoeb_canosan_fee"] = CANOSAN_FEE
    locked["eoeb_nberk_fee"] = NBERK_FEE
    locked["art_net_locked"] = I8
    locked["art_sale_locked_gross"] = ART_GROSS
    locked["epgc_consultant"] = CONSULTANT
    ask = m.setdefault("ask", {})
    ask["eoeb_remainder"] = EOEB_REMAINING
    ask.pop("l5", None)
    ask.pop("l5_after_belt", None)
    ask.pop("l5_note", None)
    ask.pop("eoeb_jan16_canosan_horse", None)
    ask.pop("mosaics_two_deal_split", None)
    ask["mosaics_jamal_boa_cost"] = True
    ask["mosaics_two_deal_note"] = (
        "Split LOCKED ($50k / $40k / $15k). Jamal Cost amount on 9922 statement still ASK. "
        "Do not dump EOEB remainder $108,270. Not tax advice."
    )
    ask["eoeb_remainder_note"] = (
        "Per-invoice UNALLOCATED (proposed FEE, no matching Cost). Named FEEs locked "
        "(Canosan $1k, NBerk $21k, mosaics $15k). Do not dump the rest."
    )
    ask["l5_leftover"] = "LOCKED FEE $37,396.32"
    m["epgc_art_sales_2025_monthly"] = [
        14000.0,
        FEB_ART,
        0.0,
        0.0,
        0.0,
        0.0,
        76000.0,
        0.0,
        0.0,
        25000.0,
        0.0,
        0.0,
    ]
    m["epgc_consultant_2025_monthly"] = CONSULTANT_MONTHLY
    m["epgc_consultant"] = CONSULTANT
    dc = packet.setdefault("deal_classifier", {})
    dc["eoeb_mosaics"] = "SALE $90k (Deal B $50k + Deal A $40k) + FEE $15k"
    dc["plutus_antiquarium"] = "SALE Deal B Cost $50k / Sale $50k LOCKED net $0"
    dc["jamal_mosaics"] = "SALE Deal A $40k; Cost from personal BoA ASK"
    dc["mosaics_fee"] = "FEE $15k Consultant February"
    dc["l5_remainder"] = "FEE $37,396.32 LOCKED"
    dc["eoeb_canosan"] = "FEE $1,000 LOCKED"
    dc["eoeb_nberk"] = "FEE $21,000 LOCKED"
    dc["eoeb_remainder"] = "UNALLOCATED $108,270 (proposed FEE per invoice; not dumped)"
    PACKET.write_text(json.dumps(packet, indent=2) + "\n", encoding="utf-8")

    mercury = json.loads(MERCURY_JSON.read_text(encoding="utf-8"))
    mercury["updated"] = "2026-09-21"
    mercury.setdefault("locked", {}).update(locked)
    mercury["locked"].pop("mosaics_profit", None)
    mercury.setdefault("ask", {}).update(ask)
    for k in ("l5", "l5_after_belt", "l5_note", "eoeb_jan16_canosan_horse", "mosaics_two_deal_split"):
        mercury["ask"].pop(k, None)
    mercury["epgc_art_sales_2025_monthly"] = m["epgc_art_sales_2025_monthly"]
    mercury["epgc_consultant_2025_monthly"] = CONSULTANT_MONTHLY
    mercury["epgc_consultant"] = CONSULTANT
    mercury.setdefault("deal_classifier", {}).update(
        {
            "eoeb_mosaics": dc["eoeb_mosaics"],
            "plutus_antiquarium": dc["plutus_antiquarium"],
            "jamal_mosaics": dc["jamal_mosaics"],
            "mosaics_fee": dc["mosaics_fee"],
            "l5_remainder": dc["l5_remainder"],
            "eoeb_canosan": dc["eoeb_canosan"],
            "eoeb_nberk": dc["eoeb_nberk"],
            "eoeb_remainder": dc["eoeb_remainder"],
        }
    )
    MERCURY_JSON.write_text(json.dumps(mercury, indent=2) + "\n", encoding="utf-8")


def write_start_here() -> None:
    (CSV_DIR / "START_HERE_mercury.txt").write_text(start_here_mercury(), encoding="utf-8")
    (CSV_DIR / "START_HERE_all_tabs.txt").write_text(start_here_all(), encoding="utf-8")


def patch_drive_ids(ids: dict[str, str]) -> None:
    """Rewrite START HERE + packet live pointers after Drive upload."""
    global OPEN_XLSX, OPEN_07, OPEN_ROOT, LABELED_ART, CLASSIFIER, ASK, START_HERE, LOCK
    OPEN_XLSX = ids.get("open_xlsx", OPEN_XLSX)
    OPEN_07 = ids.get("open_07", OPEN_07)
    OPEN_ROOT = ids.get("open_root", OPEN_ROOT)
    LABELED_ART = ids.get("labeled_art", LABELED_ART)
    CLASSIFIER = ids.get("classifier", CLASSIFIER)
    ASK = ids.get("ask", ASK)
    START_HERE = ids.get("start_here", START_HERE)
    LOCK = ids.get("lock", LOCK)
    write_start_here()

    packet = json.loads(PACKET.read_text(encoding="utf-8"))
    sheet = packet.setdefault("sheet_ids", {})
    mapping = {
        "ask_mercury": ids.get("ask_id"),
        "ask_mercury_mosaics_boa": ids.get("ask_id"),
        "deal_classifier": ids.get("classifier_id"),
        "art_sales_labeled": ids.get("labeled_id"),
        "like_last_year_income_properties": ids.get("xlsx_id"),
        "like_last_year_income_properties_packet": ids.get("xlsx_id"),
        "like_last_year_all_tabs": ids.get("xlsx_id"),
        "personal_income_2025": ids.get("xlsx_id"),
        "income_2025_sheet": ids.get("xlsx_id"),
        "oak_park_2025_sheet": ids.get("xlsx_id"),
        "unit2_2025_sheet": ids.get("xlsx_id"),
        "unit1_2025_sheet": ids.get("xlsx_id"),
        "gcm_2025_sheet": ids.get("xlsx_id"),
        "epgc_2025_sheet": ids.get("xlsx_id"),
        "art_2025_sheet": ids.get("xlsx_id"),
        "w2_2025_sheet": ids.get("xlsx_id"),
        "t4_2025_sheet": ids.get("xlsx_id"),
        "like_last_year_all_tabs_07": ids.get("xlsx_07_id"),
        "personal_income_2025_07": ids.get("xlsx_07_id"),
        "like_last_year_all_tabs_root": ids.get("xlsx_root_id"),
        "personal_income_2025_root": ids.get("xlsx_root_id"),
        "start_here": ids.get("start_here_id"),
        "start_here_packet": ids.get("start_here_id"),
        "start_here_all_tabs": ids.get("start_here_id"),
        "start_here_mercury": ids.get("start_here_mercury_id"),
        "start_here_mercury_eoeb": ids.get("start_here_mercury_id"),
        "start_here_all_tabs_07": ids.get("start_here_07_id"),
        "start_here_mercury_eoeb_07": ids.get("start_here_mercury_07_id"),
        "start_here_all_tabs_root": ids.get("start_here_root_id"),
        "art_sales_labeled_07": ids.get("labeled_07_id"),
        "art_sales_labeled_root": ids.get("labeled_root_id"),
        "deal_classifier_07": ids.get("classifier_07_id"),
        "deal_classifier_bank": ids.get("classifier_bank_id"),
        "mercury_lock_two_deals": ids.get("lock_id"),
    }
    for k, v in mapping.items():
        if v:
            sheet[k] = v
    if ids.get("xlsx_id"):
        packet["like_last_year_income_properties"] = ids["xlsx_id"]
        packet["like_last_year_income_properties_packet"] = ids["xlsx_id"]
    if ids.get("start_here_id"):
        packet["start_here"] = ids["start_here_id"]
        packet["start_here_packet"] = ids["start_here_id"]
    PACKET.write_text(json.dumps(packet, indent=2) + "\n", encoding="utf-8")

    mercury = json.loads(MERCURY_JSON.read_text(encoding="utf-8"))
    drive = mercury.setdefault("drive", {})
    nested = packet.setdefault("mercury_8291", {}).setdefault("drive", {})
    if ids.get("ask_id"):
        drive["ask_mercury"] = ids["ask_id"]
        nested["ask_mercury"] = ids["ask_id"]
        packet["start_here_mercury"] = ids.get("start_here_mercury_id") or packet.get("start_here_mercury")
    if ids.get("classifier_id"):
        drive["deal_classifier"] = ids["classifier_id"]
        nested["deal_classifier"] = ids["classifier_id"]
        mercury.setdefault("deal_classifier", {})["sheet_id"] = ids["classifier_id"]
    if ids.get("classifier_bank_id"):
        drive["deal_classifier_bank"] = ids["classifier_bank_id"]
        nested["deal_classifier_bank"] = ids["classifier_bank_id"]
    if ids.get("classifier_07_id"):
        drive["deal_classifier_07"] = ids["classifier_07_id"]
        nested["deal_classifier_07"] = ids["classifier_07_id"]
    if ids.get("labeled_id"):
        drive["art_sales_labeled"] = ids["labeled_id"]
        nested["art_sales_labeled"] = ids["labeled_id"]
    if ids.get("xlsx_id"):
        drive["personal_income_2025"] = ids["xlsx_id"]
        nested["personal_income_2025"] = ids["xlsx_id"]
    if ids.get("xlsx_07_id"):
        drive["personal_income_2025_07"] = ids["xlsx_07_id"]
        nested["personal_income_2025_07"] = ids["xlsx_07_id"]
    if ids.get("xlsx_root_id"):
        drive["personal_income_2025_root"] = ids["xlsx_root_id"]
        nested["personal_income_2025_root"] = ids["xlsx_root_id"]
    if ids.get("start_here_id"):
        drive["start_here"] = ids["start_here_id"]
        nested["start_here_all_tabs"] = ids["start_here_id"]
    if ids.get("start_here_mercury_id"):
        drive["start_here_eoeb"] = ids["start_here_mercury_id"]
        nested["start_here_eoeb"] = ids["start_here_mercury_id"]
        packet["start_here_mercury"] = ids["start_here_mercury_id"]
    if ids.get("lock_id"):
        drive["lock_two_deals"] = ids["lock_id"]
        nested["lock_two_deals"] = ids["lock_id"]
    PACKET.write_text(json.dumps(packet, indent=2) + "\n", encoding="utf-8")
    MERCURY_JSON.write_text(json.dumps(mercury, indent=2) + "\n", encoding="utf-8")

    if DRIVE_MAP.exists():
        dmap = json.loads(DRIVE_MAP.read_text(encoding="utf-8"))
        dmap["updated"] = "2026-09-21"
        key = dmap.setdefault("key_files", {})
        if ids.get("xlsx_id"):
            key["like_last_year_income_properties"] = ids["xlsx_id"]
            key["like_last_year_all_tabs"] = ids["xlsx_id"]
            key["personal_income_2025"] = ids["xlsx_id"]
        if ids.get("xlsx_07_id"):
            key["like_last_year_all_tabs_07"] = ids["xlsx_07_id"]
            key["personal_income_2025_07"] = ids["xlsx_07_id"]
        if ids.get("xlsx_root_id"):
            key["like_last_year_all_tabs_root"] = ids["xlsx_root_id"]
            key["personal_income_2025_root"] = ids["xlsx_root_id"]
        if ids.get("ask_id"):
            key["ask_mercury"] = ids["ask_id"]
        if ids.get("classifier_id"):
            key["deal_classifier"] = ids["classifier_id"]
        if ids.get("labeled_id"):
            key["art_sales_labeled"] = ids["labeled_id"]
        if ids.get("start_here_id"):
            key["start_here_xlsx"] = ids["start_here_id"]
            key["start_here_all_tabs"] = ids["start_here_id"]
        if ids.get("start_here_mercury_id"):
            key["start_here_mercury"] = ids["start_here_mercury_id"]
        if ids.get("start_here_07_id"):
            key["start_here_all_tabs_07"] = ids["start_here_07_id"]
        if ids.get("start_here_root_id"):
            key["start_here_all_tabs_root"] = ids["start_here_root_id"]
        notes = dmap.setdefault("notes", [])
        notes.append(
            "2026-09-21 FEE lock: Q10 $50k/$40k/$15k; L5 leftover FEE; Canosan $1k FEE; "
            "NBerk $21k FEE. I8 $24,055.06. Art Sales $236,000. Consultant $87,991.32. "
            "EOEB remainder $108,270 UNALLOCATED. Not tax advice."
        )
        DRIVE_MAP.write_text(json.dumps(dmap, indent=2) + "\n", encoding="utf-8")


def main() -> None:
    patch_xlsx()
    patch_ledger()
    patch_classifier()
    patch_ask()
    patch_lock_q()
    write_start_here()
    patch_json()
    print("locked L5 leftover FEE", L5_FEES)
    print("locked Q10 mosaics SALE", MOSAICS_SALE, "FEE", MOSAICS_FEE)
    print("locked Canosan FEE", CANOSAN_FEE, "NBerk FEE", NBERK_FEE)
    print("I8", I8, "Feb Art", FEB_ART, "year Art", ART_GROSS)
    print("Consultant", CONSULTANT, "EOEB remainder UNALLOCATED", EOEB_REMAINING)


if __name__ == "__main__":
    main()
