#!/usr/bin/env python3
"""Lock mosaics Cost funding: Jamal $90k from personal BoA ($50k + $40k), profit $15k.

User 2026-09-21: The $90k owed to Jamal for the mosaics had to be sent from
personal BoA. Transferred $90k (thinks $50k and $40k). $15k was the profit.

Bank on 8291:
  2/21 EOEB +$105,000 sale (already LOCKED).
  2/24 Mercury → BoA 9922 −$90,000 (same-day funding of personal).
  2/25 Plutus/Antiquarium (Jamal DBA) −$50,000.

Do NOT add the $90k BoA transfer as extra Cost — Art Sales Cost is already $90,000
and I8 already includes mosaics net $15,000. Split the 2/24 $90k: $40,000 mosaics
Cost funding (not owner draw) + $50,000 owner draw.

Monarch 9922 has no labeled $40k outgoing to Jamal in February (only $90k IN from
EPGC). Soft note, not a P&L blocker.

I8 / EPGC Art Sales gross unchanged. Not tax advice.
"""
from __future__ import annotations

import csv
import json
import shutil
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

GREEN = PatternFill("solid", fgColor="C6EFCE")
CG = Font(name="Century Gothic", size=10)
WRAP = Alignment(wrap_text=True, vertical="top")

MOSAIC_NOTE = (
    "LOCKED 2026-09-21 Cost $90,000 to Jamal Rifai / DBA Antiquarium. User: the $90k "
    "owed to Jamal had to be sent from personal BoA 9922 (thinks $50k + $40k). "
    "Profit $15,000 ($105,000 EOEB 2/21 − $90,000 Cost). Payment: Mercury 8291 → "
    "Plutus & Mnemosyne (DBA Antiquarium) 2/24–25 $50,000 + personal BoA $40,000. "
    "Mercury 2/24 $90,000 EPGC→9922 is same-day funding of personal — do NOT add as "
    "a second $90k of Cost. $40,000 of that transfer is mosaics Cost funding (not "
    "owner draw); $50,000 remains owner draw. Monarch 9922 has no labeled $40k "
    "outgoing to Jamal in Feb (only the $90k IN from EPGC) — confirm on 9922 stmt "
    "if needed. ASK 10 CLOSED. I8 already has mosaics net $15,000. Not tax advice."
)
ART_FOOTER = (
    "Mercury 8291 MATCHED/LOCKED: Berk $14,000 (1/10 seals) + $30,000 (2/7 lots). "
    "EOEB $105,000 (2/21 mosaics) / Cost $90,000 to Jamal (Plutus $50,000 on 8291 + "
    "personal BoA $40,000) / profit $15,000. Fortuna $13,000 (1/8 seals Cost). "
    "Roman Gold Belt SALE to L5: Cost $45,000 / Sale $50,000 / net $5,000. "
    "Venus joint SALE EPGC share: Cost $20,000 / proceeds $26,000 / net $6,000. "
    "August Fortuna joint SALE EPGC share: Cost $20,000 / proceeds $25,000 / net $5,000 "
    "(Sale Price = remittance, not full hammer; object unnamed). "
    "Newstar $23,581 jewelry fabrication COGS parked below. "
    "David Aaron $13,595 LOCKED consultant (EPGC Consultant June — not a sale). "
    "Aquinas Hobor $1,000 LOCKED book sale (Feb cash; Cost TBD, not in I8). "
    "Do not dump EOEB remainder $130,270 (see Deal Classifier per invoice) or other "
    "L5 $37,396.32. Not tax advice."
)
INCOME_NOTE = (
    "2025 Actual: I4 Monarch paycheck cash ≠ W-2 Box 1. I5 216 LTR cash $19,500. "
    "I6 524 #2 STR platform net LOCKED. I7 GCM 1099-NEC $21,500. "
    "I8 Art net $39,055.06 = prior MATCHED nets $23,055.06 (Berk lots + seals "
    "$1,000 + mosaics $15,000) + Belt net $5,000 + Venus EPGC-share net $6,000 "
    "+ August Fortuna joint net $5,000. Mosaics Cost $90,000 LOCKED (Jamal via "
    "Plutus $50k on Mercury + personal BoA $40k); profit $15,000. "
    "Excludes Sale 6428 $11,000 until Cost. Excludes Aquinas books $1,000 until Cost. "
    "David Aaron $13,595 is EPGC Consultant June (not I8). "
    "EOEB remainder $130,270 stays UNALLOCATED (per-invoice Deal Classifier; proposed "
    "FEE, no matching Cost — not dumped). Other L5 $37,396.32 / Erdal 11/7 $6,000 "
    "/ Aysel $50,000 still ASK. Not tax advice."
)
EPGC_NOTE = (
    "2025 Mercury 2026-09-21: Art Sales gross cash by month = Jan $14,000 + Feb "
    "$136,000 MATCHED + July $76,000 (Belt L5 $50,000 + Venus remittance $26,000) "
    "+ October $25,000 (August Fortuna joint remittance) = $251,000. "
    "Mosaics Cost $90,000 to Jamal LOCKED (Plutus $50k on 8291 + personal BoA $40k); "
    "profit $15,000 already in I8. Consultant June $13,595 = David Aaron Limited "
    "LOCKED (not a sale). Consultant Fees July $334.17 + October $658.63 = Wise "
    "expertise write-ups LOCKED ($992.80). EOEB remainder $130,270 is per-invoice "
    "UNALLOCATED on the Deal Classifier (proposed FEE) — not dumped. Other L5 "
    "$37,396.32 MIXED — not dumped. Erdal 11/7 $6,000 FORTUNA PAYMENT still ASK. "
    "Aysel Monarch +$50k vs native Failed −$50k still ASK (not the belt). "
    "Koziol/Ariadne $150,000 LOCKED pass-through — not P&L. Coinbase is Investments. "
    "Newstar $23,581 is Art Sales COGS. Dec Wise $565.75 reimbursed. Aquinas books "
    "Cost TBD. Joint Sale Price = EPGC proceeds, not full hammer. GCM 1099 is "
    "personal. Not tax advice."
)
LEDGER_BANNER = (
    "Mercury Checking 8291 — EPGC LLC 2025. 64 unique cash txns. "
    "Deal Classifier: SALE vs FEE vs PASS vs UNALLOCATED. "
    "Mosaics Cost $90,000 to Jamal LOCKED (Plutus $50k on 8291 + personal BoA $40k); "
    "profit $15,000. 2/24 $90k EPGC→9922 = $40k mosaics funding + $50k owner draw "
    "(not extra Cost). Roman Gold Belt = SALE. Venus = joint SALE. August Fortuna = "
    "joint SALE. Remaining EOEB $130,270 per-invoice UNALLOCATED (proposed FEE — not "
    "dumped). Other L5 $37,396.32 UNALLOCATED. David Aaron = FEE. Koziol/Ariadne = PASS. "
    "11/7 $6,000 FORTUNA PAYMENT still ASK. Not tax advice."
)
START_HERE_MERCURY = """START HERE — Mercury Bank 8291 (2026-09-21 mosaics BoA Cost lock)

Not tax advice. Same Personal Income.xlsx tabs as last year. Organizational packet only.

Mercury Checking 8291 is the EPGC LLC operating account (Choice Financial).
64 unique 2025 cash transactions from Monarch IDs.

LOCKED this pass
- Three Ancient Mosaics SALE: EOEB 2/21 $105,000 / Cost $90,000 to Jamal Rifai (DBA Antiquarium) / profit $15,000.
  User 2026-09-21: the $90k owed to Jamal had to be sent from personal BoA (thinks $50k + $40k).
  Payment: Mercury → Plutus 2/24–25 $50,000 + personal BoA $40,000.
  Mercury 2/24 $90,000 EPGC→9922 is same-day funding of personal — not a second $90k of Cost.
  $40,000 of that transfer = mosaics Cost funding (not owner draw). ASK 10 CLOSED.
- Roman Gold Belt SALE: Fortuna Cost $45,000 / L5 Sale $50,000 / net $5,000 → EPGC Art Sales July + I8.
- Venus joint SALE: Fortuna Cost $20,000 / Erdal proceeds $26,000 / net $6,000 → EPGC Art Sales July + I8.
- August Fortuna joint SALE: Fortuna Cost $20,000 / Erdal proceeds $25,000 / net $5,000 → EPGC Art Sales October + I8.
- David Aaron Limited $13,595 (6/24) → EPGC Consultant June. FEE, not a sale.
- Koziol $150,000 (10/20) / Ariadne $150,000 (10/21) → PASS. Not P&L.
- Wise expertise write-ups $992.80 → EPGC Consultant Fees.
- EOEB 12/2 $565.75 reimbursed — not income.

EPGC 2025 Art Sales LOCKED cash $251,000 (Jan $14,000 / Feb $136,000 / July $76,000 / Oct $25,000).
EPGC Consultant June is $13,595 (David Aaron only).
Income I8 Art net is $39,055.06 (includes mosaics profit $15,000). I10 EBITDA =SUM(I4:I9) → $167,760.
BoA 9922 draws $221,172 after pulling $40,000 mosaics funding out of the 2/24 $90,000.

ASK remaining — do not dump
1. EOEB remainder $130,270.00 — per-invoice UNALLOCATED (proposed FEE, no matching Cost). Jan 16 $1,000 Canosan-horse OPEN. 5/1 $21,000 is NOT Venus.
2. Other L5 $37,396.32 after removing $50k belt.
3. Erdal 11/7 $6,000 FORTUNA PAYMENT — not August, not Venus.
4. Aysel Dere — Monarch 7/18 +$50,000 IN vs native Failed 7/17 −$50,000 OUT. Not the belt.
10. CLOSED — mosaics $40k Cost is personal BoA. Soft: no labeled $40k OUT to Jamal on Monarch 9922 in Feb (only $90k IN from EPGC). Confirm on 9922 statement if you have it. P&L already correct.

Sale 6428 $11,000 and Aquinas books $1,000 wait on Cost (not in I8).
"""
START_HERE_ALL = """START HERE — last year's spreadsheet, 2025 numbers filled
Not tax advice. Organizational packet only.

Open this ONE spreadsheet (same 11 tab names as last year, 2023 | 2024 | 2025 year blocks, 2025 numbers, formulas that calculate, Art Sales has 2022–2024 lots + 2025 Belt, Venus, August joint; mosaics Cost $90k to Jamal / profit $15k):
https://docs.google.com/spreadsheets/d/1Dw4eUPQD6N0cJVMnYq49GrhbnU8CXiqnk9WtS9c1Q-o/edit

Tabs in last year's order (all inside that one file)
1. Income — 2022–2025. 2025 Actual: Hindman $70,618 / 216 $19,500 / 524 #2 $17,086.94 / GCM $21,500 / Art $39,055.06. I10 EBITDA =SUM(I4:I9) → $167,760.
2. 524 Ferdinand Ave, Unit 2
3. 216 N. Oak Park Ave
4. EPGC LLC — 2025 Art Sales $14,000 + $136,000 + July $76,000 + October $25,000 = $251,000; Consultant June $13,595 (David Aaron only).
5. Refrence Library
6. Art Sales and Purchases — 2022–2024 lots, then Sales 2025 including mosaics (Jamal Cost $90k / profit $15k), Roman Gold Belt, Venus joint, August Fortuna joint. Do not dump EOEB $130,270.
7. GCM
8. Hindman W2
9. Megan T4
10. Investments — Coinbase ACH net -$51,000.
11. 524 Ferdinand Ave, Unit 1

Mosaics Cost $90,000 LOCKED (user 2026-09-21): owed to Jamal, sent from personal BoA as $50k + $40k. Mercury Plutus $50k + BoA $40k. Profit $15,000. Do not double-count the 2/24 $90k EPGC→9922 transfer as extra Cost.

Deal Classifier: SALE vs FEE vs PASS vs UNALLOCATED.
Belt = SALE. Venus = SALE (joint). August Fortuna = SALE (joint). Mosaics = SALE. EOEB remainder = UNALLOCATED per invoice (proposed FEE). David Aaron = FEE.

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


def patch_xlsx() -> None:
    wb = load_workbook(XLSX)
    art = wb["Art Sales and Purchases"]
    if art["A74"].value is None or "Mosaic" not in str(art["A74"].value):
        raise SystemExit(f"Art A74 is {art['A74'].value!r}")
    paint(art["I74"], MOSAIC_NOTE, fill=GREEN)
    art["I74"].alignment = WRAP
    art.row_dimensions[74].height = 72
    for r in range(78, 90):
        a = str(art.cell(r, 1).value or "")
        if a.startswith("Mercury 8291 MATCHED") or "Plutus $50,000 of mosaics" in a:
            paint(art.cell(r, 1), ART_FOOTER, fill=GREEN)
            art.cell(r, 1).alignment = WRAP
            art.row_dimensions[r].height = 72
            break

    inc = wb["Income"]
    paint(inc["A12"], INCOME_NOTE)
    inc["A12"].alignment = WRAP

    ep = wb["EPGC LLC"]
    for r in range(74, 90):
        v = str(ep.cell(r, 1).value or "")
        if "2025 Mercury" in v or "Art Sales gross" in v:
            paint(ep.cell(r, 1), EPGC_NOTE)
            ep.cell(r, 1).alignment = WRAP
            break

    ds = wb["2025 Data sources"]
    blob_row = None
    for r in range(1, 40):
        v = str(ds.cell(r, 1).value or "")
        if "mosaics" in v.lower() or "HOW THIS FILE" in v or "ASK 10" in v:
            blob_row = r
        if v.startswith("Mosaics Cost") or "40,000 of Cost" in v:
            paint(
                ds.cell(r, 1),
                "Mosaics Cost $90,000 LOCKED 2026-09-21: Jamal / DBA Antiquarium. "
                "Plutus $50,000 on Mercury 8291 + personal BoA $40,000. Profit $15,000. "
                "2/24 $90k EPGC→9922 is funding, not extra Cost.",
                fill=GREEN,
            )
    if blob_row is None:
        ds.cell(ds.max_row + 1, 1, "Mosaics Cost $90,000 to Jamal LOCKED (BoA $40k + Plutus $50k). Not tax advice.")

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
    for r in rows[3:]:
        if len(r) <= notes_i or not r[date_i]:
            continue
        date, who, amt = r[date_i], r[who_i], r[amt_i]
        if date == "2025-02-24" and "9922" in who and amt.startswith("-90000"):
            r[bkt_i] = "TRANSFER / MOSAICS_FUNDING"
            r[st_i] = "LOCKED"
            r[deal_i] = "Mosaics — $40,000 Cost funding to 9922 + $50,000 owner draw"
            r[notes_i] = (
                "User 2026-09-21: $90k owed to Jamal sent from personal BoA (thinks $50k+$40k). "
                "This $90,000 EPGC→9922 is same-day funding of personal, NOT extra Cost. "
                "Split: $40,000 mosaics Cost funding (Jamal paid from 9922) + $50,000 owner draw. "
                "Art Sales Cost stays $90,000 = Plutus $50k + BoA $40k. Profit $15,000. "
                "Do not double-count. BoA 9922 draws $261,172 − $40,000 = $221,172. Not tax advice."
            )
        if date == "2025-02-25" and "Plutus" in who:
            r[deal_i] = "Three Ancient Mosaics — $50,000 of $90,000 Cost (Jamal DBA Antiquarium)"
            r[notes_i] = (
                "Plutus & Mnemosyne LLC / DBA Antiquarium (Jamal Rifai). With the 2/24 $10 test "
                "this is $50,000 of the $90,000 mosaics Cost. Other $40,000 paid from personal "
                "BoA 9922 (user 2026-09-21). Profit $15,000. ASK 10 CLOSED. Do not add the 2/24 "
                "$90k EPGC→9922 as a second Cost. Not tax advice."
            )
        if date == "2025-02-21" and who == "EOEB LLC" and amt.startswith("105000"):
            r[notes_i] = (
                "EOEB LLC MAKE A PAYMENT. Matches Art Sales to Jonathan Yantis $105,000. "
                "Cost $90,000 to Jamal LOCKED (Plutus $50k + personal BoA $40k). Profit $15,000."
            )
    with path.open("w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)


def patch_classifier() -> None:
    path = CSV_DIR / "DEAL_CLASSIFIER.csv"
    rows = list(csv.reader(path.open(encoding="utf-8")))
    for r in rows:
        if len(r) >= 7 and r[0] == "2025-02-21" and "Mosaic" in r[2]:
            r[3] = (
                "LOCKED Cost $90,000 to Jamal: Plutus/Antiquarium 2/24–25 $50,000 on 8291 "
                "+ personal BoA $40,000. 2/24 $90k EPGC→9922 is funding, not extra Cost. Profit $15,000."
            )
            r[5] = (
                "Matches Art Sales to Jonathan Yantis $105,000. User 2026-09-21: $90k owed to "
                "Jamal sent from personal BoA (thinks $50k+$40k). ASK 10 CLOSED."
            )
            r[6] = "LOCKED / MATCHED (Cost $90k / profit $15k)"
    with path.open("w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)


def patch_ask() -> None:
    path = CSV_DIR / "ASK_Mercury.csv"
    rows = list(csv.reader(path.open(encoding="utf-8")))
    out = []
    for r in rows:
        line = ",".join(r)
        if r and str(r[0]).startswith("10. Mosaics"):
            out.append(
                [
                    "10. Mosaics Cost $90,000 to Jamal — LOCKED (personal BoA $40k + Plutus $50k)",
                    "User 2026-09-21: $90k owed to Jamal had to be sent from personal BoA (thinks $50k + $40k). $15k profit. Mercury Plutus 2/24–25 $50,000 + BoA $40,000. 2/24 $90k EPGC→9922 is funding, not extra Cost. I8 already has mosaics net $15,000. Soft: no labeled $40k OUT to Jamal on Monarch 9922 in Feb (only $90k IN from EPGC).",
                    "SALE (Cost locked)",
                ]
            )
            continue
        if r and r[0] == "Mosaics" and "Plutus $50,000 of $90,000" in line:
            out.append(
                [
                    "Mosaics",
                    "EOEB $105,000 sale (2/21) / Cost $90,000 to Jamal (Plutus $50k + BoA $40k) / profit $15,000",
                    "SALE",
                ]
            )
            continue
        out.append(r)
    with path.open("w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(out)


def patch_lock_q() -> None:
    path = CSV_DIR / "LOCK_Mercury_Q6_Q9.csv"
    rows = list(csv.reader(path.open(encoding="utf-8")))
    rows.append(
        [
            "10 mosaics Cost",
            "LOCKED",
            "Jamal $90,000 (Plutus $50k on 8291 + personal BoA $40k). Profit $15,000.",
            "Art Sales Cost already $90,000. I8 mosaics net $15,000. 2/24 $90k EPGC→9922 = $40k funding + $50k draw. Not extra Cost.",
        ]
    )
    with path.open("w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)


def patch_json() -> None:
    packet = json.loads(PACKET.read_text(encoding="utf-8"))
    packet["updated"] = "2026-09-21"
    m = packet.setdefault("mercury_8291", {})
    m["updated"] = "2026-09-21"
    locked = m.setdefault("locked", {})
    locked["mosaics_sale"] = 105000.0
    locked["mosaics_cost"] = 90000.0
    locked["mosaics_profit"] = 15000.0
    locked["mosaics_plutus_8291"] = 50000.0
    locked["mosaics_boa_9922"] = 40000.0
    locked["boa_9922_draws"] = 221172.0
    locked["boa_9922_draws_note"] = (
        "Was $261,172. Pull $40,000 of 2/24 $90,000 EPGC→9922 out as mosaics Cost funding. "
        "Remaining 2/24 $50,000 is still an owner draw."
    )
    ask = m.setdefault("ask", {})
    ask.pop("mosaics_cost_missing_on_mercury", None)
    ask["mosaics_boa_40k_wire_label"] = (
        "Soft: Monarch 9922 has no labeled $40,000 outgoing to Jamal in Feb "
        "(only $90,000 IN from EPGC 2/24). Cost/profit locked. Confirm on 9922 statement if needed."
    )
    dc = packet.setdefault("deal_classifier", {})
    dc["eoeb_mosaics"] = "SALE (Cost $90k Jamal / profit $15k LOCKED)"
    PACKET.write_text(json.dumps(packet, indent=2) + "\n", encoding="utf-8")

    mercury = json.loads(MERCURY_JSON.read_text(encoding="utf-8"))
    mercury["updated"] = "2026-09-21"
    mercury.setdefault("locked", {}).update(locked)
    mercury.setdefault("ask", {}).update(ask)
    mercury["ask"].pop("mosaics_cost_missing_on_mercury", None)
    mercury.setdefault("deal_classifier", {}).update({"eoeb_mosaics": dc["eoeb_mosaics"]})
    MERCURY_JSON.write_text(json.dumps(mercury, indent=2) + "\n", encoding="utf-8")


def main() -> None:
    patch_xlsx()
    patch_ledger()
    patch_classifier()
    patch_ask()
    patch_lock_q()
    (CSV_DIR / "START_HERE_mercury.txt").write_text(START_HERE_MERCURY, encoding="utf-8")
    (CSV_DIR / "START_HERE_all_tabs.txt").write_text(START_HERE_ALL, encoding="utf-8")
    patch_json()
    print("locked mosaics BoA Cost $90k / profit $15k")
    print("boa draws 221172 (was 261172)")
    print("I8 unchanged 39055.06")


if __name__ == "__main__":
    main()
