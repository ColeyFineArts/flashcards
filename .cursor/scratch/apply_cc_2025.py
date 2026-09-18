#!/usr/bin/env python3
"""Apply folder-07 credit-card statements onto Personal Income 2025 Tax Turbo.

Starting-point allocations (cash basis, transaction date month):
  EPGC  <- Coinbase 6108 (none) + BoA 4469 (where prior Excel categories match)
           + Sapphire 5423 AT&T cellphone (user) + Squarespace Tech (2024 pattern)
           + Prime Abebooks / Canva / WW Norton
  524   <- Sapphire HOME + Prime utilities/Grove/yard; Nicor split by 2024 Unit 2
           gas pattern (smaller bill = Unit 2).
Not tax advice.

Second pass (user answers 2026-09-18) is apply_cc_answers_2025.py — do not re-run this
file on an already-filled workbook (it would insert a duplicate Art Library row).
"""
from __future__ import annotations

import csv
import json
import shutil
from collections import defaultdict
from copy import copy
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path("/workspace/.cursor/scratch")
XLSX = ROOT / "Personal_Income_2025_Tax_Turbo.xlsx"
DELIVERABLE = ROOT / "tax_turbo_deliverable" / "Personal_Income_2025_Tax_Turbo.xlsx"
PARTS = ROOT / "tax_turbo_parts"
LEDGER_JSON = ROOT / "cc_starting_point_2025.json"
CSV_DIR = ROOT / "cc_fill_csv"
CSV_DIR.mkdir(exist_ok=True)

MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
PROP_COLS = {m: i + 3 for i, m in enumerate(MONTHS)}  # C=3 .. N=14
EPGC_COLS = {m: i + 2 for i, m in enumerate(MONTHS)}  # B=2 .. M=13

GREEN = PatternFill("solid", fgColor="C6EFCE")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
BLUE = PatternFill("solid", fgColor="DDEBF7")
GRAY = PatternFill("solid", fgColor="F2F2F2")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def D(x) -> Decimal:
    return Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def money(x: Decimal) -> float:
    return float(D(x))


def split_half(amount: Decimal) -> tuple[Decimal, Decimal]:
    """Give Unit 2 the extra penny on odd cents (rental)."""
    a = D(amount)
    u2 = (a / 2).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    u1 = a - u2
    return u2, u1


# --- allocations: month-index 0-11 -> Decimal ---
def zeros():
    return [Decimal("0.00")] * 12


def add(arr, month: int, amt):
    arr[month] += D(amt)


# EPGC
epgc_cell = zeros()
epgc_soft = zeros()
epgc_tech = zeros()
epgc_ad = zeros()
epgc_lib = zeros()

# Sapphire AT&T cellphone -> EPGC Cellphone (user instruction)
for m, amt in [
    (0, "51.12"), (1, "93.26"), (2, "98.26"), (3, "98.26"),
    (4, "98.26"), (5, "98.26"), (6, "98.23"), (7, "98.34"),
    (8, "98.34"), (9, "98.39"), (10, "98.39"), (11, "97.91"),
]:
    add(epgc_cell, m, amt)

# ChatGPT $20 cash-basis by transaction date (BoA 4469)
for m, amt in [
    (0, 20), (1, 20), (2, 20), (3, 20), (4, 20), (5, 20), (6, 20),
    (8, 20),  # Sep 1
    (8, 20),  # Sep 30
    (9, 20),  # Oct 31
    (11, 20),  # Dec 1 (Nov period, Dec cash)
    (11, 20),  # Dec 31
]:
    add(epgc_soft, m, amt)

# Google One $2.17 x 12
for m in range(12):
    add(epgc_soft, m, "2.17")

# Canva Prime $15 Sep-Dec -> Software Fees (starting point; flagged ASK)
for m in (8, 9, 10, 11):
    add(epgc_soft, m, "15.00")

# Squarespace -> Tech (matches 2024 Tech $20 Apr + $192 May)
add(epgc_tech, 3, "20.00")
add(epgc_tech, 4, "192.00")

# Tailor Brands -> Advertising
add(epgc_ad, 2, "300.00")
add(epgc_ad, 2, "227.99")
add(epgc_ad, 2, "227.99")
add(epgc_ad, 4, "199.00")

# Art Library
add(epgc_lib, 0, "113.84")  # Prime Abebooks Jan 31
add(epgc_lib, 0, "39.60")   # Prime Abebooks Jan 31
add(epgc_lib, 1, "31.34")   # BoA Abebooks Feb 25
add(epgc_lib, 1, "47.13")   # BoA Blurb Feb 7
add(epgc_lib, 3, "13.70")   # BoA Abebooks Apr 2
add(epgc_lib, 3, "49.99")   # Sapphire WW Norton Apr 14

# 524 gas
u2_gas, u1_gas = zeros(), zeros()
for m, amt in [
    (0, "39.70"), (1, "31.37"), (2, "32.09"), (3, "26.31"),
    (4, "25.41"), (5, "24.17"), (6, "23.53"), (7, "22.70"),
    (8, "22.71"), (9, "22.72"), (11, "23.48"), (11, "30.23"),
]:
    add(u2_gas, m, amt)
for m, amt in [
    (0, "98.57"), (1, "133.33"), (2, "115.82"), (3, "77.80"),
    (4, "79.95"), (5, "48.08"), (7, "42.65"), (7, "36.68"),
    (8, "36.25"), (9, "38.22"), (11, "49.80"), (11, "89.55"),
]:
    add(u1_gas, m, amt)

# Prime ATT internet -> Unit 2 (2024 Unit 2 was $50/mo)
u2_inet = zeros()
for m, amt in [
    (0, "45.00"), (1, "45.00"), (2, "45.00"), (3, "45.00"),
    (4, "50.00"), (5, "50.00"), (6, "50.00"), (7, "50.00"),
    (8, "50.00"), (9, "50.00"), (10, "50.00"),
]:
    add(u2_inet, m, amt)

# Village water 50/50
u2_water, u1_water = zeros(), zeros()
for m, amt in [(1, "201.11"), (3, "181.35"), (5, "225.67"), (7, "275.27"), (9, "265.53")]:
    a, b = split_half(D(amt))
    add(u2_water, m, a)
    add(u1_water, m, b)

# ComEd from Monarch checking (not CC) — Jacob -> U2, Megan -> U1
u2_elec, u1_elec = zeros(), zeros()
for m, amt in [
    (0, "59.04"), (1, "57.47"), (2, "49.15"), (3, "54.52"),
    (4, "49.73"), (5, "57.54"), (6, "109.00"), (7, "123.54"),
    (8, "106.93"), (9, "62.88"), (10, "46.47"), (11, "52.88"),
]:
    add(u2_elec, m, amt)
for m, amt in [
    (0, "132.94"), (1, "216.44"), (2, "184.09"), (3, "60.95"),
    (4, "83.76"), (5, "69.69"), (6, "164.88"), (7, "203.43"),
    (8, "187.63"), (9, "139.78"), (10, "71.29"), (11, "50.38"),
]:
    add(u1_elec, m, amt)

# State Farm — user 2026-09-18: bundled car + home. Previously home=Travelers, car=Geico.
# Feb 11 Prime: two unbundled policies $157.53 + $145.25 = $302.78.
# $145.25 ≈ leftover Prime SF ~$147 in 2026 after 524 sold → AUTO (personal).
# $157.53 = the other policy → HOME (replaces Travelers) → 524 INSURANCE 50/50 U1/U2.
# Mar–Jul combined $302.75; Aug $308.26; none Sep–Dec 2025 on Prime.
# Combined months: home = round(bill * 157.53/302.78); auto = remainder.
# Yellow in the workbook: ratio is a proxy until declarations pages. Not tax advice.
HOME_UNBUNDLED = D("157.53")
AUTO_UNBUNDLED = D("145.25")
SF_RATIO_BASE = HOME_UNBUNDLED + AUTO_UNBUNDLED  # 302.78
GEICO_AUTO_CREDIT = D("540.89")
SF_BILLS = [
    (1, D("302.78"), "Feb 11 two policies $157.53 + $145.25"),
    (2, D("302.75"), "Mar combined"),
    (3, D("302.75"), "Apr combined"),
    (4, D("302.75"), "May combined"),
    (5, D("302.75"), "Jun combined"),
    (6, D("302.75"), "Jul combined"),
    (7, D("308.26"), "Aug combined"),
]


def sf_home_auto(bill: Decimal) -> tuple[Decimal, Decimal]:
    home = D(D(bill) * HOME_UNBUNDLED / SF_RATIO_BASE)
    auto = D(bill) - home
    return home, auto


u2_ins, u1_ins = zeros(), zeros()
sf_auto = zeros()
sf_split_rows = []
for m, billed, note in SF_BILLS:
    home, auto = sf_home_auto(billed)
    u2_p, u1_p = split_half(home)
    add(u2_ins, m, u2_p)
    add(u1_ins, m, u1_p)
    add(sf_auto, m, auto)
    sf_split_rows.append(
        {
            "month": MONTHS[m],
            "billed": billed,
            "home": home,
            "auto": auto,
            "u2": u2_p,
            "u1": u1_p,
            "note": note,
        }
    )
_sf_billed = sum((r["billed"] for r in sf_split_rows), Decimal("0.00"))
assert sum(u2_ins) + sum(u1_ins) + sum(sf_auto) == _sf_billed
assert _sf_billed == D("2124.79")

# Grove Collaborative (household/STR supplies) -> 524 Unit 2 SUPPLIES
# Not 827 N Grove the house. User 2026-09-18: “grove orders” = Grove Collaborative.
u2_sup = zeros()
for m, amt in [
    (0, "32.98"), (0, "77.82"), (2, "92.96"), (3, "105.25"),
    (5, "68.73"), (6, "128.37"), (7, "176.17"), (8, "108.93"),
]:
    add(u2_sup, m, amt)

# Interior: Schauer + Rubio Monocoat
u2_int = zeros()
add(u2_int, 0, "7.46")
add(u2_int, 2, "52.78")
add(u2_int, 3, "85.78")
add(u2_int, 3, "45.62")
add(u2_int, 9, "45.60")
add(u2_int, 10, "2.75")
add(u2_int, 10, "9.88")

# Exterior/yard: TruGreen + Alsip + Good Earth
u2_ext = zeros()
for m, amt in [
    (2, "60.95"), (3, "60.95"), (4, "60.95"), (5, "60.95"),
    (6, "60.95"), (7, "60.95"), (8, "249.00"), (9, "60.95"),
]:
    add(u2_ext, m, amt)
add(u2_ext, 4, "384.25")  # Alsip Nursery
add(u2_ext, 2, "12.15")
add(u2_ext, 6, "7.76")
add(u2_ext, 6, "91.31")
add(u2_ext, 7, "16.54")
add(u2_ext, 9, "7.76")

# Aquasana
u2_equip = zeros()
add(u2_equip, 2, "73.24")
add(u2_equip, 8, "82.27")

LEDGER = [
    # card, date, merchant, amount, bucket, tab, status, note
    ("BoA 4469", "2025-01-31", "OPENAI CHATGPT", 20.00, "Software Fees", "EPGC LLC", "APPLIED", "Matches 2024 Software Fees pattern"),
    ("BoA 4469", "2025-02-28", "OPENAI CHATGPT", 20.00, "Software Fees", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-03-31", "OPENAI CHATGPT", 20.00, "Software Fees", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-04-30", "OPENAI CHATGPT", 20.00, "Software Fees", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-05-31", "OPENAI CHATGPT", 20.00, "Software Fees", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-06-30", "OPENAI CHATGPT", 20.00, "Software Fees", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-07-31", "OPENAI CHATGPT", 20.00, "Software Fees", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-09-01", "OPENAI CHATGPT", 20.00, "Software Fees", "EPGC LLC", "APPLIED", "Cash Sep (Aug period)"),
    ("BoA 4469", "2025-09-30", "OPENAI CHATGPT", 20.00, "Software Fees", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-10-31", "OPENAI CHATGPT", 20.00, "Software Fees", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-12-01", "OPENAI CHATGPT", 20.00, "Software Fees", "EPGC LLC", "APPLIED", "Cash Dec (Nov period)"),
    ("BoA 4469", "2025-12-31", "OPENAI CHATGPT", 20.00, "Software Fees", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-01-13", "GOOGLE One", 2.17, "Software Fees", "EPGC LLC", "APPLIED", "x12 months"),
    ("BoA 4469", "2025-03-18", "Tailor Brands", 300.00, "Advertising", "EPGC LLC", "APPLIED", "Brand/logo — prior Advertising"),
    ("BoA 4469", "2025-03-28", "Tailor Brands", 227.99, "Advertising", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-03-28", "Tailor Brands", 227.99, "Advertising", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-05-08", "Tailor Brands", 199.00, "Advertising", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-02-07", "Blurb", 47.13, "Art Library", "EPGC LLC", "APPLIED", "Prior Art Library / Abebooks pattern"),
    ("BoA 4469", "2025-02-25", "Abebooks", 31.34, "Art Library", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-04-02", "Abebooks", 13.70, "Art Library", "EPGC LLC", "APPLIED", ""),
    ("BoA 4469", "2025-03-27", "TruGreen", 60.95, "EXTERIOR REPAIRS / yard", "524 Unit 2", "APPLIED", "NOT EPGC — 2024 Unit 2 YARD SERVICE"),
    ("BoA 4469", "2025-04-25", "TruGreen", 60.95, "EXTERIOR REPAIRS / yard", "524 Unit 2", "APPLIED", ""),
    ("BoA 4469", "2025-05-30", "TruGreen", 60.95, "EXTERIOR REPAIRS / yard", "524 Unit 2", "APPLIED", ""),
    ("BoA 4469", "2025-06-27", "TruGreen", 60.95, "EXTERIOR REPAIRS / yard", "524 Unit 2", "APPLIED", ""),
    ("BoA 4469", "2025-07-21", "TruGreen", 60.95, "EXTERIOR REPAIRS / yard", "524 Unit 2", "APPLIED", ""),
    ("BoA 4469", "2025-08-14", "TruGreen", 60.95, "EXTERIOR REPAIRS / yard", "524 Unit 2", "APPLIED", ""),
    ("BoA 4469", "2025-09-27", "TruGreen", 249.00, "EXTERIOR REPAIRS / yard", "524 Unit 2", "APPLIED", "Larger fall treatment"),
    ("BoA 4469", "2025-10-10", "TruGreen", 60.95, "EXTERIOR REPAIRS / yard", "524 Unit 2", "APPLIED", ""),
    ("BoA 4469", "various", "DoorDash / dining / RH restaurant / Salt Shed / barber / Pete's / Monarch Money", 0, "personal", "—", "EXCLUDED", "BoA dining $3,780.82 + personal services not EPGC"),
    ("BoA 4469", "various", "PARK CHICAGO MOBILE x5", 100.00, "Travel?", "EPGC LLC", "HOLD", "Could be gallery/auction parking — confirm"),
    ("Coinbase 6108", "2025-11-29", "Parennial Golf", 1856.24, "personal", "—", "EXCLUDED", "CSV has almost no EPGC spend"),
    ("Coinbase 6108", "2025-11-01", "Field Museum + Soldier Field parking", 110.00, "personal", "—", "EXCLUDED", ""),
    ("Coinbase 6108", "2025-09-30", "Coinbase", 49.99, "personal", "—", "EXCLUDED", ""),
    ("Coinbase 6108", "2025-12-15", "Freeman's $1 voided x3", 0, "void", "—", "EXCLUDED", "Voids / payments not expenses"),
    ("Sapphire 5423", "2025-01..12", "ATT* BILL PAYMENT (cellphone)", float(sum(epgc_cell)), "Cellphone and Internet", "EPGC LLC", "APPLIED", "User: AT&T cellphone can be charged to EPGC"),
    ("Sapphire 5423", "2025-04-27", "Squarespace INV179625580", 20.00, "Tech", "EPGC LLC", "APPLIED", "Matches 2024 Tech $20"),
    ("Sapphire 5423", "2025-05-08", "Squarespace INV181162020", 192.00, "Tech", "EPGC LLC", "APPLIED", "Matches 2024 Tech $192"),
    ("Sapphire 5423", "2025-04-14", "WW Norton", 49.99, "Art Library", "EPGC LLC", "APPLIED", "Art/reference book"),
    ("Sapphire 5423", "2025-03-22", "Schauer Hardware", 52.78, "INTERIOR MAINTENANCE", "524 Unit 2", "APPLIED", "Forest Park hardware — 524"),
    ("Sapphire 5423", "2025-04-10", "Rubio Monocoat", 85.78, "INTERIOR MAINTENANCE", "524 Unit 2", "APPLIED", "Wood finish"),
    ("Sapphire 5423", "2025-04-26", "Schauer Hardware", 45.62, "INTERIOR MAINTENANCE", "524 Unit 2", "APPLIED", ""),
    ("Sapphire 5423", "2025-05-04", "Alsip Nursery", 384.25, "EXTERIOR REPAIRS / yard", "524 Unit 2", "APPLIED", "Yard — 2024 Unit 2 YARD SERVICE"),
    ("Sapphire 5423", "2025-07-24", "Los Angeles Modern Auctions", 6821.06, "Inventory COGS", "Art Sales", "HOLD-COGS", "Purchase — not EPGC P&L; add object when identified"),
    ("Sapphire 5423", "2025-04-03", "Hindman LLC", 282.24, "Inventory COGS", "Art Sales", "HOLD-COGS", ""),
    ("Sapphire 5423", "2025-07-29", "Hindman LLC", 403.51, "Inventory COGS", "Art Sales", "HOLD-COGS", ""),
    ("Sapphire 5423", "2025-02-16", "The Great Frame Up", 536.41, "framing?", "EPGC / personal", "HOLD", "Could be inventory framing or personal"),
    ("Prime 2351", "2025-01..11", "ATT*BILL PAYMENT $45 then $50", float(sum(u2_inet)), "INTERNET", "524 Unit 2", "APPLIED", "2024 Unit 2 internet was $50/mo; one house bill"),
    ("Prime 2351", "2025 Nicor smaller", "NICOR GAS BILL (smaller of pair)", float(sum(u2_gas)), "GAS", "524 Unit 2", "APPLIED", "Matches 2024 Unit 2 gas $21–62"),
    ("Prime 2351", "2025 Nicor larger", "NICOR GAS BILL (larger of pair)", float(sum(u1_gas)), "GAS", "524 Unit 1", "APPLIED", "Seasonal heating — residence meter"),
    ("Prime 2351", "bi-monthly", "VILLAGE OF FOREST PARK", 1148.93, "VILLAGE WATER AND REFUSE", "U1 50% / U2 50%", "APPLIED", "One building bill; 2023 sat on Unit 1, 2024 year-total on Unit 2"),
    ("Prime 2351", "Jan/Mar/Apr/Jun/Jul/Aug/Sep", "GROVE COLLABORATIVE", float(sum(u2_sup)), "SUPPLIES", "524 Unit 2", "APPLIED", "User 2026-09-18: “grove orders” = Grove Collaborative, not 827 N Grove the house. 8 Prime charges → 524 Unit 2 SUPPLIES."),
    ("Prime 2351", "2025-01-03", "Schauer Hardware", 7.46, "INTERIOR MAINTENANCE", "524 Unit 2", "APPLIED", ""),
    ("Prime 2351", "2025-10-18", "Schauer Hardware", 45.60, "INTERIOR MAINTENANCE", "524 Unit 2", "APPLIED", ""),
    ("Prime 2351", "2025-11-08", "Schauer Hardware", 2.75, "INTERIOR MAINTENANCE", "524 Unit 2", "APPLIED", ""),
    ("Prime 2351", "2025-11-13", "Schauer Hardware", 9.88, "INTERIOR MAINTENANCE", "524 Unit 2", "APPLIED", ""),
    ("Prime 2351", "2025-03-15", "Aquasana Water Filters", 73.24, "EQUIPMENT", "524 Unit 2", "APPLIED", "STR water filter"),
    ("Prime 2351", "2025-09-19", "Aquasana Water Filters", 82.27, "EQUIPMENT", "524 Unit 2", "APPLIED", ""),
    ("Prime 2351", "2025-01-31", "Abebooks x2", 153.44, "Art Library", "EPGC LLC", "APPLIED", "On Prime not BoA — still EPGC library"),
    ("Prime 2351", "2025-09..12", "Canva $15 x4", 60.00, "Software Fees", "EPGC LLC", "APPLIED", "Starting point — confirm not personal"),
    ("Prime 2351", "Feb-Aug", "STATE FARM INSURANCE (home share)", float(sum(u2_ins) + sum(u1_ins)), "INSURANCE", "U1 50% / U2 50%", "APPLIED", "User 2026-09-18: bundled home+auto. Only home share on 524. Feb unbundled $157.53 home / $145.25 auto used as the ratio. Yellow until declarations. None Sep–Dec on Prime."),
    ("Prime 2351", "Feb-Aug", "STATE FARM INSURANCE (auto share)", float(sum(sf_auto)), "personal auto", "—", "EXCLUDED", "Bundled car. Previously Geico. Not 524 Sch E."),
    ("Prime 2351", "2025-02-11", "GEICO *AUTO credit", -float(GEICO_AUTO_CREDIT), "personal auto", "—", "EXCLUDED", "Unused Geico auto premium when switching to State Farm auto. Not 524 income."),
    ("—", "2025", "Travelers (prior home carrier)", 0, "INSURANCE", "—", "EXCLUDED", "No 2025 Travelers on Prime/Sapphire/BoA/Monarch. Home moved to State Farm in February."),
    ("Prime 2351", "HOME", "Home Depot / Lowe's / IKEA / Extra Space", 3732.32, "repairs vs personal", "524?", "HOLD", "Store HD/Lowes/IKEA not dumped into P&L until confirmed"),
    ("Checking 0203", "monthly", "ComEd Jacob Coley", float(sum(u2_elec)), "ELECTRIC", "524 Unit 2", "APPLIED", "Not on CC — Monarch Adv Plus; smaller account"),
    ("Checking 0203", "monthly", "ComEd Megan Gerrard", float(sum(u1_elec)), "ELECTRIC", "524 Unit 1", "APPLIED", "Not on CC — larger account"),
]

ASKS = [
    ("Nicor meters", "Smaller Nicor (~$23–40) applied to Unit 2 because 2024 Unit 2 gas was $21–62. Larger seasonal bill applied to Unit 1. Swap if meters are reversed."),
    ("Prime AT&T $45–50", "Applied 100% to Unit 2 INTERNET (2024 Unit 2 was $50/mo). 2023 Unit 1 was $45/mo — this may be house-wide. Split if both units used it."),
    ("Village water", "One Forest Park bill, 50/50 U1/U2. 2023 monthly sat on Unit 1; 2024 year-total $1,769 sat on Unit 2."),
    ("State Farm", "Bundled home+auto (was Travelers home / Geico auto). Only home share on 524 INSURANCE 50/50 U1/U2; auto + Geico $540.89 credit personal. Feb $157.53/$145.25 proxy. No Sep–Dec on Prime."),
    ("Home Depot / Lowe's / IKEA", "Prime HOME $3,732 includes HD.com $1,185.80 (Mar), HD $410 (Oct), Lowe's $209 (Sep), IKEA $365. Held out of P&L."),
    ("Canva $15 Sep–Dec", "Applied to EPGC Software Fees as a starting point. Move to personal if Megan/kids."),
    ("Park Chicago $100", "Five $20 BoA charges. Held — EPGC travel vs personal."),
    ("Great Frame Up $536.41", "Sapphire Feb. Held — inventory framing vs personal."),
    ("LAMA $6,821.06 + Hindman $685.75", "Inventory purchases on Sapphire. Not EPGC operating expense; identify objects for Art Sales COGS."),
    ("Electric", "ComEd is on checking, not these cards. Jacob → Unit 2, Megan → Unit 1. Confirm meters."),
    ("Coinbase 6108", "CSV is payments/voids/golf/museum. No EPGC operating expenses applied."),
    ("TruGreen / Alsip / Good Earth", "All on Unit 2 exterior/yard (2024 YARD SERVICE was on Unit 2). Split with Unit 1 if shared grounds should follow §121."),
]


def write_month_row(ws: Worksheet, row: int, values, colmap: dict, fill):
    for i, m in enumerate(MONTHS):
        cell = ws.cell(row, colmap[m])
        cell.value = money(values[i])
        if values[i] != 0:
            cell.fill = fill
            cell.number_format = "0.00"
        else:
            cell.number_format = "0.00"


def copy_style(src, dst):
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)


def insert_labeled_row(ws: Worksheet, at: int, label_col: int, label: str, colmap: dict, year_col: int):
    """Insert a row and put label + monthly 0s + SUM formula, copying style from row below after insert."""
    proto = at  # style from current row that will shift down
    styles = [copy(ws.cell(proto, c)) for c in range(1, year_col + 1)]
    ws.insert_rows(at)
    lab = ws.cell(at, label_col, label)
    lab.font = Font(bold=False)
    first = get_column_letter(min(colmap.values()))
    last = get_column_letter(max(colmap.values()))
    for m, c in colmap.items():
        cell = ws.cell(at, c, 0)
        cell.number_format = "0.00"
    year = ws.cell(at, year_col, f"=SUM({first}{at}:{last}{at})")
    year.number_format = "0.00"


def style_header_row(ws, row, start, end):
    for c in range(start, end + 1):
        cell = ws.cell(row, c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT


def main():
    wb = load_workbook(XLSX)

    # ----- EPGC: insert Art Library at row 20 -----
    epgc = wb["EPGC LLC"]
    insert_labeled_row(epgc, 20, 1, "Art Library", EPGC_COLS, 14)
    # After insert: TOTAL was 24 -> 25. Fix SUM range to include new row.
    # Old TOTAL formula SUM(B8:B23) became SUM(B8:B24) if Excel-adjusted; openpyxl often does NOT adjust.
    # Find TOTAL
    total_row = None
    ytd_row = None
    for r in range(1, 40):
        v = epgc.cell(r, 1).value
        if v == "TOTAL EXPENSES":
            total_row = r
        if v and "YEAR TO DATE" in str(v):
            ytd_row = r
    if total_row is None:
        raise SystemExit("EPGC TOTAL row not found")
    for col in range(2, 15):
        letter = get_column_letter(col)
        epgc.cell(total_row, col).value = f"=SUM({letter}8:{letter}{total_row - 1})"
        epgc.cell(total_row, col).number_format = "0.00"
    if ytd_row:
        for col in range(2, 15):
            letter = get_column_letter(col)
            # income total is row 5
            epgc.cell(ytd_row, col).value = f"={letter}5-{letter}{total_row}"
            epgc.cell(ytd_row, col).number_format = "0.00"

    # Map labels
    label_to_row = {}
    for r in range(1, 40):
        v = epgc.cell(r, 1).value
        if v:
            label_to_row[str(v).strip()] = r

    write_month_row(epgc, label_to_row["Cellphone and Internet"], epgc_cell, EPGC_COLS, GREEN)
    write_month_row(epgc, label_to_row["Software Fees"], epgc_soft, EPGC_COLS, GREEN)
    write_month_row(epgc, label_to_row["Tech"], epgc_tech, EPGC_COLS, GREEN)
    write_month_row(epgc, label_to_row["Advertising"], epgc_ad, EPGC_COLS, GREEN)
    write_month_row(epgc, label_to_row["Art Library"], epgc_lib, EPGC_COLS, GREEN)

    note_row = 28 if epgc.cell(28, 1).value else 29
    # find last used
    last = epgc.max_row
    epgc.cell(last + 2, 1, (
        "CC starting point 2025-09-18: BoA 4469 ChatGPT/Google One/Tailor Brands/Abebooks/Blurb; "
        "Sapphire AT&T cellphone (user) + Squarespace Tech + WW Norton; Prime Abebooks + Canva. "
        "Coinbase 6108: no EPGC operating charges (golf/museum/voids). "
        "TruGreen on BoA went to 524 Unit 2 yard, not EPGC. Yellow ASK on CC_LEDGER tab. Not tax advice."
    ))
    epgc.cell(last + 2, 1).alignment = Alignment(wrap_text=True)
    epgc.merge_cells(start_row=last + 2, start_column=1, end_row=last + 3, end_column=14)

    # ----- Unit 2 -----
    u2 = wb["524 Ferdinand Ave, Unit 2"]
    u2_map = {}
    for r in range(1, 40):
        v = u2.cell(r, 2).value
        if v:
            u2_map[str(v).strip()] = r
    write_month_row(u2, u2_map["GAS"], u2_gas, PROP_COLS, GREEN)
    write_month_row(u2, u2_map["ELECTRIC"], u2_elec, PROP_COLS, GREEN)
    write_month_row(u2, u2_map["INTERNET"], u2_inet, PROP_COLS, GREEN)
    write_month_row(u2, u2_map["VILLAGE WATER AND REFUSE"], u2_water, PROP_COLS, GREEN)
    write_month_row(u2, u2_map["INSURANCE"], u2_ins, PROP_COLS, GREEN)
    write_month_row(u2, u2_map["SUPPLIES"], u2_sup, PROP_COLS, GREEN)
    write_month_row(u2, u2_map["INTERIOR MAINTENANCE"], u2_int, PROP_COLS, GREEN)
    write_month_row(u2, u2_map["EXTERIOR REPAIRS"], u2_ext, PROP_COLS, GREEN)
    write_month_row(u2, u2_map["EQUIPMENT/APPLIANCE PURCHASES"], u2_equip, PROP_COLS, GREEN)
    u2["B30"] = (
        "CC starting point: Prime Nicor SMALLER bill = Unit 2 gas (2024 U2 was $21–62); "
        "Prime ATT $45–50 = internet (2024 U2 $50); Grove Collaborative = supplies (not 827 N Grove); Schauer/Rubio = interior; "
        "TruGreen+Alsip+Good Earth = exterior/yard; Aquasana = equipment; water 50/50 with Unit 1; "
        "State Farm HOME share 50/50 (auto excluded); ComEd Jacob Coley from checking (not CC). "
        "Home Depot/Lowe's/IKEA HELD — see CC_LEDGER. Income green = Monarch STR deposits vs PLATFORM net $17,086.94."
    )
    u2["B30"].alignment = Alignment(wrap_text=True)
    u2.row_dimensions[30].height = 60

    # ----- Unit 1 -----
    u1 = wb["524 Ferdinand Ave, Unit 1"]
    u1_map = {}
    for r in range(1, 40):
        v = u1.cell(r, 2).value
        if v:
            u1_map[str(v).strip()] = r
    write_month_row(u1, u1_map["GAS"], u1_gas, PROP_COLS, GREEN)
    write_month_row(u1, u1_map["ELECTRIC"], u1_elec, PROP_COLS, GREEN)
    write_month_row(u1, u1_map["VILLAGE WATER AND REFUSE"], u1_water, PROP_COLS, GREEN)
    write_month_row(u1, u1_map["INSURANCE"], u1_ins, PROP_COLS, GREEN)
    u1["B30"] = (
        "§121 residence through sale 2025-12-18. Starting point from Prime: Nicor LARGER seasonal bill = Unit 1 gas; "
        "ComEd Megan Gerrard = electric (checking, not CC); Village water 50/50; State Farm HOME share 50/50 (auto excluded). "
        "Internet left at $0 — Prime ATT $45–50 put on Unit 2 to match 2024 STR $50 (confirm if house-wide). "
        "Shared building costs may need CPA allocation vs Unit 2."
    )
    u1["B30"].alignment = Alignment(wrap_text=True)
    u1.row_dimensions[30].height = 60

    # ----- SOURCE notes -----
    src = wb["_SOURCE_2025"]
    src["A16"] = "CC starting point (folder 07)"
    src["B16"] = (
        "Coinbase 6108 + BoA 4469 → EPGC where prior Excel categories match. "
        "Sapphire 5423 AT&T cellphone → EPGC. Prime 2351 + Sapphire HOME → 524 U1/U2. "
        "Green cells = applied. HOLD/ASK on CC_LEDGER. Not all EPGC is on Coinbase/BoA."
    )

    # ----- Art Sales note for HOLD-COGS -----
    art = wb["Art Sales and Purchases"]
    r = art.max_row + 2
    art.cell(r, 1, "2025 CC purchases not yet tied to an object (HOLD — not in Cost total)")
    art.cell(r, 1).fill = YELLOW
    art.cell(r + 1, 1, "Los Angeles Modern Auctions")
    art.cell(r + 1, 2, "Sapphire 5423")
    art.cell(r + 1, 3, 6821.06)
    art.cell(r + 1, 4, "2025-07-24")
    art.cell(r + 1, 3).number_format = "0.00"
    art.cell(r + 1, 3).fill = YELLOW
    art.cell(r + 2, 1, "Hindman LLC")
    art.cell(r + 2, 2, "Sapphire 5423")
    art.cell(r + 2, 3, 282.24)
    art.cell(r + 2, 4, "2025-04-03")
    art.cell(r + 2, 3).number_format = "0.00"
    art.cell(r + 2, 3).fill = YELLOW
    art.cell(r + 3, 1, "Hindman LLC")
    art.cell(r + 3, 2, "Sapphire 5423")
    art.cell(r + 3, 3, 403.51)
    art.cell(r + 3, 4, "2025-07-29")
    art.cell(r + 3, 3).number_format = "0.00"
    art.cell(r + 3, 3).fill = YELLOW
    art.cell(r + 4, 1, "Identify objects before adding into Cost (COGS) total — avoid double count vs Mercury inventory.")
    art.cell(r + 4, 1).fill = YELLOW

    # ----- CC_LEDGER -----
    if "CC_LEDGER" in wb.sheetnames:
        del wb["CC_LEDGER"]
    led = wb.create_sheet("CC_LEDGER", 1)
    led["A1"] = "2025 credit-card starting-point ledger (folder 07) — not tax advice"
    led["A1"].font = Font(bold=True, size=14)
    led.merge_cells("A1:H1")
    headers = ["Card", "Date", "Merchant / item", "Amount", "Prior-Excel category", "Tab", "Status", "Note"]
    for i, h in enumerate(headers, 1):
        cell = led.cell(3, i, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = THIN
    status_fill = {
        "APPLIED": GREEN,
        "EXCLUDED": GRAY,
        "HOLD": YELLOW,
        "HOLD-COGS": YELLOW,
    }
    for i, row in enumerate(LEDGER, 4):
        for c, val in enumerate(row, 1):
            cell = led.cell(i, c, val)
            cell.border = THIN
            cell.fill = status_fill.get(row[6], BLUE)
            if c == 4 and isinstance(val, (int, float)) and val:
                cell.number_format = "0.00"
    last_led = 3 + len(LEDGER)
    led.cell(last_led + 2, 1, "STATUS: APPLIED = in monthly P&L (green). EXCLUDED = personal / void. HOLD = confirm before P&L. HOLD-COGS = inventory not operating expense.")
    led.merge_cells(start_row=last_led + 2, start_column=1, end_row=last_led + 2, end_column=8)
    widths = [18, 16, 55, 12, 28, 18, 12, 70]
    for i, w in enumerate(widths, 1):
        led.column_dimensions[get_column_letter(i)].width = w
    led.auto_filter.ref = f"A3:H{last_led}"
    led.freeze_panes = "A4"

    # ----- ASK -----
    if "ASK" in wb.sheetnames:
        del wb["ASK"]
    ask = wb.create_sheet("ASK", 2)
    ask["A1"] = "Starting-point flags — confirm or leave for CPA (do not block the packet)"
    ask["A1"].font = Font(bold=True, size=14)
    ask.merge_cells("A1:B1")
    ask["A3"] = "Item"
    ask["B3"] = "What I applied / what to confirm"
    ask["A3"].fill = HEADER_FILL
    ask["B3"].fill = HEADER_FILL
    ask["A3"].font = HEADER_FONT
    ask["B3"].font = HEADER_FONT
    for i, (item, detail) in enumerate(ASKS, 4):
        ask.cell(i, 1, item).fill = YELLOW
        ask.cell(i, 2, detail).fill = YELLOW
        ask.cell(i, 1).alignment = Alignment(wrap_text=True, vertical="top")
        ask.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top")
        ask.row_dimensions[i].height = 48
        ask.cell(i, 1).border = THIN
        ask.cell(i, 2).border = THIN
    ask.column_dimensions["A"].width = 28
    ask.column_dimensions["B"].width = 110

    # Totals summary on ASK
    trow = 4 + len(ASKS) + 2
    ask.cell(trow, 1, "Applied totals (this fill)").font = Font(bold=True)
    summary = [
        ("EPGC Cellphone (Sapphire AT&T)", sum(epgc_cell)),
        ("EPGC Software (ChatGPT + Google One + Canva)", sum(epgc_soft)),
        ("EPGC Tech (Squarespace)", sum(epgc_tech)),
        ("EPGC Advertising (Tailor Brands)", sum(epgc_ad)),
        ("EPGC Art Library", sum(epgc_lib)),
        ("Unit 2 GAS (Nicor smaller)", sum(u2_gas)),
        ("Unit 1 GAS (Nicor larger)", sum(u1_gas)),
        ("Unit 2 ELECTRIC (ComEd Jacob)", sum(u2_elec)),
        ("Unit 1 ELECTRIC (ComEd Megan)", sum(u1_elec)),
        ("Unit 2 INTERNET (Prime ATT)", sum(u2_inet)),
        ("Unit 2 WATER 50%", sum(u2_water)),
        ("Unit 1 WATER 50%", sum(u1_water)),
        ("Unit 2 INSURANCE", sum(u2_ins)),
        ("Unit 1 INSURANCE", sum(u1_ins)),
        ("Unit 2 SUPPLIES (Grove Collaborative)", sum(u2_sup)),
        ("Unit 2 INTERIOR (Schauer/Rubio)", sum(u2_int)),
        ("Unit 2 EXTERIOR/YARD (TruGreen/Alsip/Good Earth)", sum(u2_ext)),
        ("Unit 2 EQUIPMENT (Aquasana)", sum(u2_equip)),
    ]
    ask.cell(trow, 1, "Applied totals")
    ask.cell(trow, 1).font = Font(bold=True)
    ask.cell(trow, 1).fill = HEADER_FILL
    ask.cell(trow, 1).font = HEADER_FONT
    ask.cell(trow, 2).fill = HEADER_FILL
    for i, (lab, val) in enumerate(summary):
        ask.cell(trow + 1 + i, 1, lab)
        cell = ask.cell(trow + 1 + i, 2, money(val))
        cell.number_format = '"$"#,##0.00'
        ask.cell(trow + 1 + i, 1).fill = GREEN
        cell.fill = GREEN

    wb.save(XLSX)
    shutil.copy2(XLSX, DELIVERABLE)

    # Export CSVs for Drive (formulas evaluated by writing values + formula text)
    def sheet_to_csv(ws, path: Path, data_only=False):
        with path.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            for row in ws.iter_rows(max_row=ws.max_row, max_col=ws.max_column):
                w.writerow([("" if c.value is None else c.value) for c in row])

    # Reload to export
    for name, fname in [
        ("EPGC LLC", "EPGC_LLC.csv"),
        ("524 Ferdinand Ave, Unit 2", "524_Unit_2.csv"),
        ("524 Ferdinand Ave, Unit 1", "524_Unit_1.csv"),
        ("CC_LEDGER", "CC_LEDGER.csv"),
        ("ASK", "ASK.csv"),
        ("Art Sales and Purchases", "Art_Sales.csv"),
        ("_SOURCE_2025", "SOURCE_2025.csv"),
    ]:
        sheet_to_csv(wb[name], CSV_DIR / fname)

    payload = {
        "updated": "2026-09-18",
        "method": "cash basis, transaction date month, prior Personal Income.xlsx categories",
        "epgc": {
            "cellphone": float(sum(epgc_cell)),
            "software": float(sum(epgc_soft)),
            "tech": float(sum(epgc_tech)),
            "advertising": float(sum(epgc_ad)),
            "art_library": float(sum(epgc_lib)),
        },
        "unit2": {
            "gas": float(sum(u2_gas)),
            "electric": float(sum(u2_elec)),
            "internet": float(sum(u2_inet)),
            "water": float(sum(u2_water)),
            "insurance": float(sum(u2_ins)),
            "supplies_grove": float(sum(u2_sup)),
            "interior": float(sum(u2_int)),
            "exterior_yard": float(sum(u2_ext)),
            "equipment": float(sum(u2_equip)),
        },
        "unit1": {
            "gas": float(sum(u1_gas)),
            "electric": float(sum(u1_elec)),
            "water": float(sum(u1_water)),
            "insurance": float(sum(u1_ins)),
        },
        "asks": [{"item": a, "detail": b} for a, b in ASKS],
    }
    LEDGER_JSON.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    print("EPGC cellphone", sum(epgc_cell))
    print("EPGC software", sum(epgc_soft))
    print("EPGC tech", sum(epgc_tech))
    print("EPGC ad", sum(epgc_ad))
    print("EPGC lib", sum(epgc_lib))
    print("U2 gas", sum(u2_gas), "U1 gas", sum(u1_gas))
    print("U2 elec", sum(u2_elec), "U1 elec", sum(u1_elec))
    print("U2 inet", sum(u2_inet))
    print("U2 water", sum(u2_water), "U1 water", sum(u1_water), "sum", sum(u2_water) + sum(u1_water))
    print("U2 ins", sum(u2_ins), "U1 ins", sum(u1_ins))
    print("U2 grove", sum(u2_sup))
    print("U2 interior", sum(u2_int))
    print("U2 exterior", sum(u2_ext))
    print("U2 equip", sum(u2_equip))
    print("saved", XLSX, "bytes", XLSX.stat().st_size)
    print("Art Library row", label_to_row.get("Art Library"), "TOTAL", total_row)


if __name__ == "__main__":
    main()
