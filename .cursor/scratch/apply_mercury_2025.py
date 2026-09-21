#!/usr/bin/env python3
"""Classify 2025 Mercury Checking 8291 (EPGC LLC / Choice Financial).

User (2026-09-20): some are art purchases/sales, some art advisory fees,
some investments (Coinbase), some COGS (Newstar Jewelers = jewelry from
intaglios / engraved gems / scarabs).

LOCKED this pass:
  Coinbase ACH = investment transfer (not EPGC operating, not Art Sales)
  Newstar Jewelers = jewelry-fabrication COGS
  Berk $14k + $30k, Fortuna $13k seals, EOEB $105k mosaics, Plutus $50k
    mosaics partial cost, BoA 9922 draws, $10 test wires, Dec Wise/EOEB
    $565.75 reimbursement
  David Aaron Limited $13,595 (6/24) = EPGC Consultant / art-advisory fee
    (user 2026-09-20: "6. Consultant fee" — not a sale)
  Jack Koziol & Tracy Hoffman $150,000 IN (10/20) /
    Ariadne Demirjian LLC $150,000 OUT (10/21) = pass-through, not P&L
    (user 2026-09-20: "8. Pass through")
  Wise $334.17 (7/9) + $658.63 (10/30) = EPGC Consultant Fees expense
    (user 2026-09-20: "7. Business expenses (expertise write ups)").
    Distinct from Dec Wise $565.75 reimbursed by EOEB (not P&L).
  Aquinas Hobor $1,000 IN (2/4) = books sold (ART_SALE LOCKED)
    (user 2026-09-20: "9. Books sold to Aquinas"). Hits EPGC Art Sales
    February. Cost / titles TBD — not in Income I8 until Cost.
  EOEB remainder and remaining L5 mixed art + advisory CONFIRMED.
    Do NOT dump EOEB $130,270 or other L5 $37,396.32 onto Consultant
    or Art Sales. Jan 16 EOEB $1,000 Canosan-horse still open.
  User 2026-09-21: Roman Gold Belt = SALE to L5 (Fortuna Cost 7/21
    −$45,000 / L5 7/15 +$50,000). Venus = joint SALE EPGC share
    (Fortuna Cost 5/9 −$20,000 / Erdal remittance 7/24 +$26,000).
    Not unsold inventory. Aug Fortuna $20,000 still unnamed.

Belt + Venus SALE lock for the last-year 11-tab workbook lives in
lock_belt_venus_2025.py. Do not re-run this script to recast those deals.

Do NOT dump unclassified Mercury into EPGC Art Sales.
Do not re-run apply_checking_answers_2025.py.
Not tax advice.
"""
from __future__ import annotations

import csv
import json
import shutil
from copy import copy
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path("/workspace/.cursor/scratch")
MONARCH = ROOT / "monarch_ty2025.csv"
NATIVE_CSV = ROOT / "mercury_8291_native_2025.csv"
NATIVE_DRIVE_ID = "1pJL5IFjAeQxkzgKHEh_xkcGjhbAxGevU"
XLSX = ROOT / "Personal_Income_2025_Tax_Turbo.xlsx"
DELIVERABLE = ROOT / "tax_turbo_deliverable" / "Personal_Income_2025_Tax_Turbo.xlsx"
PARTS = ROOT / "tax_turbo_parts"
CSV_DIR = ROOT / "cc_fill_csv"
JSON_PATH = ROOT / "mercury_2025.json"
PACKET = ROOT / "cpa_packet_v1.json"

MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
PEACH = PatternFill("solid", fgColor="F7CAAC")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GREEN = PatternFill("solid", fgColor="C6EFCE")
BLUE = PatternFill("solid", fgColor="DDEBF7")
GRAY = PatternFill("solid", fgColor="F2F2F2")
ORANGE = PatternFill("solid", fgColor="F8CBAD")
NAVY = PatternFill("solid", fgColor="1F4E79")
WHITE = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
CG = Font(name="Century Gothic", size=10)
CG_B = Font(name="Century Gothic", size=11, bold=True)
ACCT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
WRAP = Alignment(wrap_text=True, vertical="top")


def D(x) -> Decimal:
    return Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def money(x) -> float:
    return float(D(x))


def zeros():
    return [Decimal("0.00")] * 12


def add(arr, month: int, amt):
    arr[month] += D(amt)


MIXED_EOEB = (
    "User 2026-09-21 Deal Classifier: leftover EOEB after mosaics has no matching "
    "dealer Cost on 8291 (Fortuna Venus/Belt/Aug joints are NOT EOEB Cost) and looks "
    "like FEE — still UNALLOCATED. Do NOT dump onto EPGC Consultant or Art Sales. "
    "5/1 $21,000 is NOT Venus. Jan 16 $1,000 Canosan-horse still OPEN. Not tax advice."
)
MIXED_L5 = (
    "User 2026-09-20: L5 LLC mixed art sales AND advisory/consultant fees. "
    "Pattern confirmed; remaining invoices still unallocated after locking "
    "7/15 $50,000 as Roman Gold Belt SALE. Do NOT dump L5 remainder "
    "$37,396.32 onto Consultant or Art Sales. Invoice-level split still ASK. "
    "Do NOT attach remaining L5 invoices to the belt."
)

# (date, amount) -> classification. Amounts signed as Mercury cash (IN +, OUT -).
# bucket: ART_SALE | ART_PURCHASE | ADVISORY | INVESTMENT | COGS_JEWELRY | TRANSFER
#         | PASS_THROUGH | REIMBURSE | TEST | ASK | EPGC_EXPENSE
# status: LOCKED | MATCHED | PROPOSED | ASK | HOLD | NOT_PL
RULES = {
    ("2025-01-08", D("-13000.00")): (
        "ART_PURCHASE",
        "MATCHED",
        "A Collection of Seals — COGS",
        "Fortuna / Erdal Dere. Mercury native note: J.K. New York Collection. "
        "Matches Art Sales Cost $13,000; Berk sold it 1/10 for $14,000. "
        "Same-day Cancelled ACH to Erdal Dere (closed) is not cash.",
    ),
    ("2025-01-10", D("14000.00")): (
        "ART_SALE",
        "MATCHED",
        "A Collection of Seals — Sale Price",
        "Harlan J. Berk Ltd wire. Matches Art Sales Sale Price $14,000.",
    ),
    ("2025-01-16", D("1000.00")): (
        "ASK",
        "ASK",
        "EOEB $1,000 — Canosan horse vs advisory (STILL OPEN)",
        "STILL OPEN. Art Sales register horse is to Erdal $1,000; this cash is EOEB. "
        "Do not lock as the horse sale. Classifier: UNALLOCATED / OPEN. Not tax advice.",
    ),
    ("2025-02-04", D("1000.00")): (
        "ART_SALE",
        "LOCKED",
        "Books (titles TBD) — Sale Price",
        "User 2026-09-20: book SALE to Aquinas Hobor. Hits EPGC Art Sales February +$1,000. "
        "Art Sales register: Books (titles TBD), Cost TBD. Not in Income I8 until Cost is known (same as Sale 6428).",
    ),
    ("2025-02-04", D("-25000.00")): (
        "TRANSFER",
        "LOCKED",
        "Owner transfer to BoA 9922",
        "Not EPGC P&L. Draw / transfer to Jake Adv Plus 9922.",
    ),
    ("2025-02-07", D("30000.00")): (
        "ART_SALE",
        "MATCHED",
        "Berk lots lump Sale Price $30,000",
        "Harlan J. Berk Ltd PER ARB. Matches Art Sales F71 hardcoded $30,000 covering rows 51–70. Per-object invoices still yellow.",
    ),
    ("2025-02-21", D("105000.00")): (
        "ART_SALE",
        "MATCHED",
        "Three Ancient Mosaics — Sale Price",
        "EOEB LLC MAKE A PAYMENT. Matches Art Sales to Jonathan Yantis $105,000.",
    ),
    ("2025-02-21", D("10.00")): (
        "TEST",
        "LOCKED",
        "$10 test wire in",
        "Not P&L.",
    ),
    ("2025-02-21", D("-10.00")): (
        "TEST",
        "LOCKED",
        "$10 test wire to Jamal M. Rifai",
        "Not P&L. Related to mosaics counterparty test.",
    ),
    ("2025-02-24", D("-90000.00")): (
        "TRANSFER",
        "LOCKED",
        "Owner transfer to BoA 9922",
        "Not EPGC P&L.",
    ),
    ("2025-02-24", D("10.00")): (
        "TEST",
        "LOCKED",
        "$10 test wire EPGC",
        "Not P&L.",
    ),
    ("2025-02-25", D("-49990.00")): (
        "ART_PURCHASE",
        "MATCHED",
        "Three Ancient Mosaics — partial COGS",
        "Plutus & Mnemosyne LLC. With the 2/24 $10 test this is $50,000 of the $90,000 mosaics Cost. $40,000 of Cost is not on Mercury 8291.",
    ),
    ("2025-02-24", D("-10.00")): (
        "TEST",
        "LOCKED",
        "$10 test (Jamal or Plutus)",
        "Not P&L. If this is the Plutus $10, it is the first $10 of the $50,000 mosaics cost; the $49,990 follows 2/25.",
    ),
    ("2025-03-10", D("14000.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-05-01", D("21000.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-05-09", D("-20000.00")): (
        "ART_PURCHASE",
        "LOCKED",
        "Venus — EPGC Cost (joint with Erdal)",
        "User 2026-09-21: Erdal and Jacob bought together. Not unsold inventory. "
        "Paired with Erdal 7/24 +$26,000 remittance (investment + $6,000 earnings). "
        "I8 net $6,000. Classifier: SALE (joint, EPGC share). Not tax advice.",
    ),
    ("2025-05-13", D("16200.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-06-12", D("25000.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-06-12", D("3300.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-06-24", D("13595.00")): (
        "ADVISORY",
        "LOCKED",
        "David Aaron Limited — consultant fee",
        "User 2026-09-20: consultant fee, not a sale. Hits EPGC Consultant June $13,595. Not Art Sales.",
    ),
    ("2025-07-03", D("-40000.00")): (
        "TRANSFER",
        "LOCKED",
        "Owner transfer to BoA 9922",
        "Not EPGC P&L.",
    ),
    ("2025-07-08", D("2000.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-07-09", D("-334.17")): (
        "EPGC_EXPENSE",
        "LOCKED",
        "Wise — expertise write-ups (Consultant Fees)",
        "User 2026-09-20: business expense (expertise write-ups). Hits EPGC Consultant Fees July $334.17. Not COGS, not reimbursed. Distinct from Dec Wise $565.75 (EOEB reimburse).",
    ),
    ("2025-07-11", D("8534.79")): (
        "ADVISORY",
        "ASK",
        "L5 — mixed unallocated (sale vs advisory)",
        MIXED_L5,
    ),
    ("2025-07-15", D("50000.00")): (
        "ART_SALE",
        "LOCKED",
        "Roman Gold Belt — Sale Price (L5)",
        "User 2026-09-21: sold to L5. Only L5 payment large enough. "
        "Hits EPGC Art Sales July +$50,000. L5 paid 6 days before Fortuna Cost — "
        "back-to-back dealer cash. Do NOT attach remaining L5 invoices. "
        "Classifier: SALE. Not tax advice.",
    ),
    ("2025-07-15", D("2500.00")): (
        "ADVISORY",
        "ASK",
        "L5 — mixed unallocated (sale vs advisory)",
        MIXED_L5,
    ),
    ("2025-07-18", D("50000.00")): (
        "ASK",
        "ASK",
        "Aysel Dere $50,000 IN on Monarch — NOT on Mercury native Sent",
        "Monarch 7/18 +$50,000 IN. Mercury native 2025 CSV (user upload 2026-09-20) has no Sent Aysel cash. "
        "Native has Failed $50,000 OUT 7/17 'Recipient account does not exist'. "
        "Keep on ASK — do not dump onto Art Sales or Consultant. Is the Monarch IN a phantom of the failed send?",
    ),
    ("2025-07-21", D("-45000.00")): (
        "ART_PURCHASE",
        "LOCKED",
        "Roman Gold Belt — Cost (Fortuna)",
        "User 2026-09-21: Cost of SALE to L5, not unsold inventory. Native memo Roman Gold Belt. "
        "Paired with L5 7/15 +$50,000. Net $5,000 on I8. Classifier: SALE. Not tax advice.",
    ),
    ("2025-07-24", D("26000.00")): (
        "ART_SALE",
        "LOCKED",
        "Venus — EPGC proceeds (Erdal remittance)",
        "User 2026-09-21: Erdal sold it and paid back Jacob’s $20,000 + $6,000 earnings. "
        "Art Sales Sale Price = EPGC proceeds $26,000, not full hammer. "
        "Hits EPGC Art Sales July. Other Erdal IN still ASK. "
        "Classifier: SALE (joint, EPGC share). Not tax advice.",
    ),
    ("2025-07-28", D("-60000.00")): (
        "TRANSFER",
        "LOCKED",
        "Owner transfer to BoA 9922",
        "Not EPGC P&L.",
    ),
    ("2025-08-15", D("-20000.00")): (
        "ART_PURCHASE",
        "ASK",
        "Fortuna / Erdal — object still unnamed",
        "Same $20,000 OUT as Mercury native initiated 2025-08-18 / Monarch posted 2025-08-15. "
        "No object name in the native memo ('From EPGC LLC via mercury.com'). "
        "Inventory vs other? Do not add to sold Cost until named. Not a second unique txn.",
    ),
    ("2025-08-29", D("10000.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-09-02", D("5042.00")): (
        "ADVISORY",
        "ASK",
        "L5 — mixed unallocated (sale vs advisory)",
        MIXED_L5,
    ),
    ("2025-09-04", D("-3000.00")): (
        "INVESTMENT",
        "LOCKED",
        "Coinbase ACH funding",
        "Investment transfer off Mercury 8291. Not EPGC operating. Distinct from Coinbase 6108 credit-card spend.",
    ),
    ("2025-09-19", D("5000.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-09-23", D("9119.27")): (
        "ADVISORY",
        "ASK",
        "L5 — mixed unallocated (sale vs advisory)",
        MIXED_L5,
    ),
    ("2025-09-25", D("-40000.00")): (
        "TRANSFER",
        "LOCKED",
        "Owner transfer to BoA 9922",
        "Not EPGC P&L.",
    ),
    ("2025-09-25", D("3100.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-09-29", D("-6000.00")): (
        "INVESTMENT",
        "LOCKED",
        "Coinbase ACH funding",
        "Investment transfer. Not EPGC operating.",
    ),
    ("2025-10-01", D("10000.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-10-01", D("-10000.00")): (
        "INVESTMENT",
        "LOCKED",
        "Coinbase ACH funding",
        "Investment transfer. Not EPGC operating.",
    ),
    ("2025-10-09", D("-6172.00")): (
        "TRANSFER",
        "LOCKED",
        "Owner transfer to BoA 9922",
        "Not EPGC P&L.",
    ),
    ("2025-10-09", D("25000.00")): (
        "ART_SALE",
        "ASK",
        "Erdal Dere $25,000 IN — sale vs other",
        "Incoming wire ERDAL DERE. Sale of inventory to Erdal or something else?",
    ),
    ("2025-10-14", D("-8000.00")): (
        "COGS_JEWELRY",
        "LOCKED",
        "Newstar Jewelers — jewelry fabrication",
        "Fees for making jewelry out of intaglios, engraved gems, and scarabs. COGS, not EPGC operating. Not added to sold-lot Cost (Berk lots already sold in February).",
    ),
    ("2025-10-15", D("16000.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-10-20", D("150000.00")): (
        "PASS_THROUGH",
        "LOCKED",
        "Jack Koziol & Tracy Hoffman $150,000 IN",
        "User 2026-09-20: pass-through, not a sale. Incoming wire /BNF/Per your request. Next-day $150,000 OUT to Ariadne Demirjian LLC. Not P&L.",
    ),
    ("2025-10-21", D("7500.00")): (
        "ADVISORY",
        "ASK",
        "L5 — mixed unallocated (sale vs advisory)",
        MIXED_L5,
    ),
    ("2025-10-21", D("-150000.00")): (
        "PASS_THROUGH",
        "LOCKED",
        "Ariadne Demirjian LLC $150,000 OUT",
        "User 2026-09-20: pass-through, not a purchase. Same window as Koziol $150,000 IN. Not P&L.",
    ),
    ("2025-10-24", D("1500.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-10-30", D("-658.63")): (
        "EPGC_EXPENSE",
        "LOCKED",
        "Wise — expertise write-ups (Consultant Fees)",
        "User 2026-09-20: business expense (expertise write-ups). Hits EPGC Consultant Fees October $658.63. Not COGS, not reimbursed. Distinct from Dec Wise $565.75 (EOEB reimburse).",
    ),
    ("2025-10-31", D("2170.00")): (
        "ADVISORY",
        "ASK",
        "EOEB remainder — UNALLOCATED (proposed FEE, no matching Cost)",
        MIXED_EOEB,
    ),
    ("2025-11-06", D("-10000.00")): (
        "INVESTMENT",
        "LOCKED",
        "Coinbase ACH funding",
        "Investment transfer. Not EPGC operating.",
    ),
    ("2025-11-07", D("6000.00")): (
        "ART_SALE",
        "ASK",
        "Erdal Dere Fortuna PAYMENT $6,000 IN",
        "Wire memo ERDAL DERE FORTUNA PAYMENT. Sale, refund of a Fortuna purchase, or other?",
    ),
    ("2025-11-12", D("-8000.00")): (
        "COGS_JEWELRY",
        "LOCKED",
        "Newstar Jewelers — jewelry fabrication",
        "Intaglios / engraved gems / scarabs. COGS, not EPGC operating.",
    ),
    ("2025-11-14", D("-10000.00")): (
        "INVESTMENT",
        "LOCKED",
        "Coinbase ACH funding",
        "Investment transfer. Not EPGC operating.",
    ),
    ("2025-11-17", D("-10000.00")): (
        "INVESTMENT",
        "LOCKED",
        "Coinbase ACH funding",
        "Investment transfer. Not EPGC operating.",
    ),
    ("2025-11-24", D("-5000.00")): (
        "INVESTMENT",
        "LOCKED",
        "Coinbase ACH funding",
        "Investment transfer. Not EPGC operating.",
    ),
    ("2025-12-02", D("565.75")): (
        "REIMBURSE",
        "LOCKED",
        "EOEB reimburses Wise $565.75",
        "Monarch tag Reimburse. Nets with same-day Wise OUT. Not income.",
    ),
    ("2025-12-02", D("-565.75")): (
        "REIMBURSE",
        "LOCKED",
        "Wise $565.75 reimbursed by EOEB",
        "Same-day EOEB IN $565.75. Net $0. Not EPGC expense.",
    ),
    ("2025-12-04", D("4700.26")): (
        "ADVISORY",
        "ASK",
        "L5 — mixed unallocated (sale vs advisory)",
        MIXED_L5,
    ),
    ("2025-12-11", D("-5000.00")): (
        "INVESTMENT",
        "LOCKED",
        "Coinbase ACH funding",
        "Investment transfer. Not EPGC operating.",
    ),
    ("2025-12-22", D("8000.00")): (
        "INVESTMENT",
        "LOCKED",
        "Coinbase ACH back to Mercury",
        "Investment withdrawal onto 8291. Not Art Sales / not advisory income.",
    ),
    ("2025-12-24", D("-7581.00")): (
        "COGS_JEWELRY",
        "LOCKED",
        "Newstar Jewelers — jewelry fabrication (3 of 3)",
        "Mercury.com ‘3 of 3’. Intaglios / engraved gems / scarabs. COGS, not EPGC operating.",
    ),
}


def merchant_from_statement(stmt: str, monarch_merchant: str) -> str:
    s = stmt or ""
    if "Merchant name:" in s:
        return s.split("Merchant name:", 1)[1].strip()
    return monarch_merchant


def _native_iso(mdy: str) -> str:
    # MM-DD-YYYY → YYYY-MM-DD
    mm, dd, yyyy = mdy.split("-")
    return f"{yyyy}-{mm}-{dd}"


def load_native_rows() -> list[dict]:
    """Mercury native 2025 CSV (user Drive upload). Account 8291 only."""
    if not NATIVE_CSV.exists():
        return []
    out = []
    with NATIVE_CSV.open(newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            acct = row.get("Source Account") or ""
            if "8291" not in acct:
                continue
            date = _native_iso(row["Date (UTC)"])
            amt = D(row["Amount"])
            out.append(
                {
                    "date": date,
                    "amount": amt,
                    "status": (row.get("Status") or "").strip(),
                    "description": (row.get("Description") or "").strip(),
                    "reference": (row.get("Reference") or "").strip(),
                    "note": (row.get("Note") or "").strip(),
                    "bank": (row.get("Bank Description") or "").strip(),
                    "fail": (row.get("Failure Reason") or "").strip(),
                    "tracking": (row.get("Tracking ID") or "").strip(),
                }
            )
    return out


def native_ingest_stats(native: list[dict]) -> dict:
    sent = [r for r in native if r["status"] == "Sent"]
    noncash = [r for r in native if r["status"] != "Sent"]
    sent_keys = {(r["date"], r["amount"]) for r in sent}
    rules_keys = set(RULES)
    # 8/18 native Fortuna $20k is the 8/15 Monarch key, not a new unique.
    alias = {("2025-08-18", D("-20000.00")): ("2025-08-15", D("-20000.00"))}
    mapped = {alias.get(k, k) for k in sent_keys}
    new_unique = sorted(mapped - rules_keys)
    missing_from_native = sorted(rules_keys - mapped)
    return {
        "drive_id": NATIVE_DRIVE_ID,
        "title": "epgc-llc-transactions-2025-jan-01-to-2025-dec-31.csv",
        "account": "Mercury Checking xx8291",
        "n_rows": len(native),
        "sent_cash_rows": len(sent),
        "sent_unique_keys": len(sent_keys),
        "new_unique_vs_64": len(new_unique),
        "new_unique_keys": [f"{d} {a}" for d, a in new_unique],
        "missing_from_native_sent": [f"{d} {a}" for d, a in missing_from_native],
        "date_shift": "Fortuna $20,000 Monarch 2025-08-15 = Mercury native 2025-08-18 (same cash)",
        "non_cash": [
            {
                "date": r["date"],
                "amount": float(r["amount"]),
                "status": r["status"],
                "who": r["description"],
                "reason": r["fail"] or r["note"] or r["reference"],
            }
            for r in noncash
        ],
    }


def overlay_native(rows: list[dict], native: list[dict]) -> None:
    """Attach native memo / initiated date. Do not add Failed/Cancelled as cash."""
    sent = [r for r in native if r["status"] == "Sent"]
    by_key: dict[tuple, list] = {}
    for n in sent:
        key = (n["date"], n["amount"])
        if key == ("2025-08-18", D("-20000.00")):
            key = ("2025-08-15", D("-20000.00"))
        by_key.setdefault(key, []).append(n)
    used = set()
    for r in rows:
        cands = by_key.get((r["date"], r["amount"])) or []
        pick = None
        for i, n in enumerate(cands):
            tag = (n["date"], n["amount"], n["tracking"] or n["description"], i)
            if tag in used:
                continue
            pick = n
            used.add(tag)
            break
        if not pick and cands:
            pick = cands[0]
        if not pick:
            continue
        extra = " | ".join(x for x in (pick["reference"], pick["note"], pick["bank"]) if x)
        r["native_date"] = pick["date"]
        r["native_status"] = pick["status"]
        r["native_memo"] = extra
        if extra and extra not in (r["statement"] or ""):
            r["statement"] = f"{r['statement']}; Mercury native: {extra}"


def load_mercury_rows() -> list[dict]:
    seen = set()
    out = []
    with MONARCH.open(newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if "8291" not in (row.get("Account") or ""):
                continue
            if not (row.get("Date") or "").startswith("2025"):
                continue
            tid = row["Id"]
            if tid in seen:
                continue
            seen.add(tid)
            amt = D(row["Amount"])
            date = row["Date"]
            key = (date, amt)
            if key not in RULES:
                raise SystemExit(f"unclassified Mercury txn {date} {amt} {row['Merchant']} {row['Original Statement']}")
            bucket, status, deal, note = RULES[key]
            month = int(date[5:7]) - 1
            out.append(
                {
                    "date": date,
                    "month": month,
                    "counterparty": merchant_from_statement(row["Original Statement"], row["Merchant"]),
                    "monarch_merchant": row["Merchant"],
                    "statement": row["Original Statement"],
                    "monarch_category": row["Category"],
                    "amount": amt,
                    "bucket": bucket,
                    "status": status,
                    "deal": deal,
                    "note": note,
                    "id": tid,
                }
            )
    out.sort(key=lambda r: (r["date"], r["amount"], r["counterparty"]))
    if len(out) != 64:
        raise SystemExit(f"expected 64 unique 2025 Mercury 8291 txns, got {len(out)}")
    native = load_native_rows()
    if native:
        overlay_native(out, native)
    return out


def monthly_sum(rows, pred) -> list[Decimal]:
    arr = zeros()
    for r in rows:
        if pred(r):
            add(arr, r["month"], r["amount"])
    return arr


def write_months(ws: Worksheet, row: int, start_col: int, values, fill, zero_fill=None):
    for i, v in enumerate(values):
        cell = ws.cell(row, start_col + i)
        cell.value = money(v)
        cell.number_format = ACCT
        cell.font = CG
        if D(v) != 0:
            cell.fill = fill
        elif zero_fill is not None:
            cell.fill = zero_fill


def style_header_row(ws, row, cols, titles):
    for i, t in enumerate(titles, start=1):
        cell = ws.cell(row, i, t)
        cell.fill = NAVY
        cell.font = WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 22


def fill_status(cell, status: str):
    cell.value = status
    cell.font = CG
    cell.fill = {
        "LOCKED": GREEN,
        "MATCHED": GREEN,
        "PROPOSED": YELLOW,
        "ASK": YELLOW,
        "HOLD": BLUE,
        "NOT_PL": GRAY,
    }.get(status, YELLOW)


def build_ledger_sheet(wb, rows: list[dict]) -> Worksheet:
    if "MERCURY 8291" in wb.sheetnames:
        del wb["MERCURY 8291"]
    ws = wb.create_sheet("MERCURY 8291", 0)
    ws.sheet_properties.tabColor = "1F4E79"
    ws["A1"] = (
        "Mercury Checking 8291 — EPGC LLC (Choice Financial) 2025. "
        "64 unique cash transactions (Monarch IDs). Native Mercury CSV 2026-09-20 is the same 8291 year — "
        "0 new unique Sent cash. Coinbase ACH = investment. "
        "Newstar = jewelry COGS on intaglios/gems/scarabs. "
        "Fortuna Venus $20k (5/9) + Roman Gold Belt $45k (7/21) LOCKED inventory (not sold Cost). "
        "BoA 9922 draws and Koziol/Ariadne $150k are LOCKED not-P&L (pass-through / owner transfer). "
        "David Aaron $13,595 is LOCKED EPGC Consultant (June). "
        "Wise $334.17 + $658.63 are LOCKED EPGC Consultant Fees (expertise write-ups). "
        "Aquinas Hobor $1,000 (2/4) is LOCKED books sold (EPGC Art Sales February; Cost TBD). "
        "Not tax advice."
    )
    ws["A1"].font = CG_B
    ws["A1"].alignment = WRAP
    ws.merge_cells("A1:L1")
    ws.row_dimensions[1].height = 48

    headers = [
        "Date",
        "Counterparty",
        "Amount",
        "Direction",
        "Bucket",
        "Status",
        "Deal / object",
        "Hits EPGC Art Sales?",
        "Hits EPGC Consultant?",
        "Monarch category",
        "Original statement",
        "Notes",
    ]
    style_header_row(ws, 3, 12, headers)

    for i, r in enumerate(rows, start=4):
        amt = r["amount"]
        direction = "IN" if amt >= 0 else "OUT"
        art = "YES" if r["bucket"] == "ART_SALE" and r["status"] in ("LOCKED", "MATCHED") else "no"
        if r["bucket"] == "ADVISORY" and r["status"] == "LOCKED":
            adv = "YES"
        elif r["bucket"] == "ADVISORY":
            adv = "no — ASK first"
        else:
            adv = "no"
        ws.cell(i, 1, r["date"]).font = CG
        ws.cell(i, 2, r["counterparty"]).font = CG
        c = ws.cell(i, 3, money(amt))
        c.number_format = ACCT
        c.font = CG
        ws.cell(i, 4, direction).font = CG
        ws.cell(i, 5, r["bucket"]).font = CG
        fill_status(ws.cell(i, 6), r["status"])
        ws.cell(i, 7, r["deal"]).font = CG
        ws.cell(i, 8, art).font = CG
        ws.cell(i, 9, adv).font = CG
        ws.cell(i, 10, r["monarch_category"]).font = CG
        ws.cell(i, 11, r["statement"]).font = CG
        n = ws.cell(i, 12, r["note"])
        n.font = CG
        n.alignment = WRAP
        bucket_fill = {
            "ART_SALE": PEACH,
            "ART_PURCHASE": ORANGE,
            "COGS_JEWELRY": ORANGE,
            "INVESTMENT": BLUE,
            "TRANSFER": GRAY,
            "PASS_THROUGH": BLUE,
            "REIMBURSE": GRAY,
            "TEST": GRAY,
            "EPGC_EXPENSE": PEACH,
            "ADVISORY": YELLOW,
            "ASK": YELLOW,
        }.get(r["bucket"], YELLOW)
        ws.cell(i, 5).fill = GREEN if (r["bucket"] == "ADVISORY" and r["status"] == "LOCKED") else bucket_fill
        if r["status"] in ("LOCKED", "MATCHED") and r["bucket"] in (
            "ART_SALE",
            "COGS_JEWELRY",
            "INVESTMENT",
            "ADVISORY",
            "EPGC_EXPENSE",
        ):
            ws.cell(i, 3).fill = GREEN
        elif r["status"] in ("ASK", "PROPOSED", "HOLD"):
            ws.cell(i, 3).fill = YELLOW
        else:
            ws.cell(i, 3).fill = GRAY
        ws.row_dimensions[i].height = 36

    last = 3 + len(rows)
    ws.auto_filter.ref = f"A3:L{last}"
    ws.freeze_panes = "A4"
    widths = [12, 28, 14, 10, 16, 12, 42, 22, 22, 22, 55, 70]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def build_ask_sheet(wb, rows: list[dict], sums: dict) -> Worksheet:
    if "ASK Mercury" in wb.sheetnames:
        del wb["ASK Mercury"]
    ws = wb.create_sheet("ASK Mercury", 1)
    ws.sheet_properties.tabColor = "FFC000"
    ws["A1"] = (
        "ASK — remaining Mercury 8291 questions. "
        "Q6 consultant, Q7 Wise expertise, Q8 pass-through, Q9 Aquinas books LOCKED (green). "
        "Q1/Q2 mixed pattern CONFIRMED (blue) — remainder dollars unallocated until invoice split. Not tax advice."
    )
    ws["A1"].font = CG_B
    ws.merge_cells("A1:C1")

    questions = [
        (
            "1. EOEB LLC remainder after mosaics — MIXED PATTERN CONFIRMED, dollars unallocated",
            f"${sums['eoeb_remainder']:,.2f} of EOEB IN is not mosaics $105,000 and not Wise reimburse $565.75. "
            "User 2026-09-20: EOEB is mixed art sales AND advisory/consultant fees. "
            "Do NOT dump this remainder onto Consultant or Art Sales. Mark each invoice sale vs advisory: "
            "1/16 $1,000 (Canosan-horse STILL OPEN — Art Sales horse is to Erdal, this cash is EOEB); "
            "3/10 $14,000; 5/1 $21,000; 5/13 $16,200; 6/12 $25,000 + $3,300; 7/8 $2,000; "
            "8/29 $10,000; 9/19 $5,000; 9/25 $3,100; 10/1 $10,000; 10/15 $16,000; 10/24 $1,500; 10/31 $2,170.",
        ),
        (
            "2. L5 (all seven invoices) — MIXED PATTERN CONFIRMED, dollars unallocated",
            f"${sums['l5']:,.2f}. User 2026-09-20: L5 is mixed art sales AND advisory/consultant fees. "
            "Do NOT dump onto Consultant or Art Sales. Mark each invoice sale vs advisory: "
            "7/11 $8,534.79; 7/15 $50,000 + $2,500 (same week as Aysel $50k / Fortuna $45k); "
            "9/2 $5,042; 9/23 $9,119.27; 10/21 $7,500; 12/4 $4,700.26.",
        ),
        (
            "3. Fortuna / Erdal $20,000 Aug — object still unnamed",
            f"${sums['fortuna_remainder_out']:,.2f}. Monarch posted 8/15 = Mercury native initiated 8/18 "
            "(same cash, not a second unique txn). Native memo has no object name. "
            "Venus $20,000 (5/9) and Roman Gold Belt $45,000 (7/21) are LOCKED inventory below — not this question.",
        ),
        (
            "4. Erdal Dere IN",
            f"${sums['erdal_in']:,.2f} (7/24 $26k, 10/9 $25k, 11/7 $6k ‘FORTUNA PAYMENT’). "
            "Sales to Erdal, refunds of Fortuna purchases, or other?",
        ),
        (
            "5. Aysel Dere $50,000 — Monarch IN vs native Failed OUT",
            "Monarch 7/18 +$50,000 IN is in the 64. Mercury native CSV has NO Sent Aysel cash — "
            "only Failed $50,000 OUT 7/17 'Recipient account does not exist'. "
            "Is the Monarch IN a phantom of the failed send, or a real incoming the export omitted? "
            "Do not dump onto Art Sales or Consultant until confirmed.",
        ),
        (
            "6. David Aaron Limited $13,595 IN (6/24) — LOCKED consultant fee",
            "CONFIRMED 2026-09-20: consultant fee, not a sale. Booked EPGC Consultant June $13,595 (peach). Not Art Sales. Not Income I8.",
        ),
        (
            "7. Wise $334.17 (7/9) and $658.63 (10/30) — LOCKED expertise write-ups",
            "CONFIRMED 2026-09-20: business expenses (expertise write-ups). "
            "Booked EPGC Consultant Fees July $334.17 + October $658.63 = $992.80. "
            "Not COGS. Distinct from Dec Wise $565.75 reimbursed by EOEB (still not P&L).",
        ),
        (
            "8. Jack Koziol & Tracy Hoffman $150,000 IN (10/20) / Ariadne Demirjian $150,000 OUT (10/21) — LOCKED pass-through",
            "CONFIRMED 2026-09-20: pass-through, not P&L. Not a sale, not a purchase, not COGS, not Consultant.",
        ),
        (
            "9. Aquinas Hobor $1,000 IN (2/4) — LOCKED book sale",
            "CONFIRMED 2026-09-20: book SALE. Hits EPGC Art Sales February +$1,000 "
            "(Feb $136,000; matched cash $150,000). Art Sales register: Books (titles TBD) to Aquinas Hobor, "
            "Sale Price $1,000, Sale Date 2025-02-04, Cost TBD. Not in Income I8 until Cost (same as Sale 6428).",
        ),
        (
            "10. Mosaics Cost $90,000 vs Mercury $50,000",
            "Plutus & Mnemosyne paid $50,000 in February. Where is the other $40,000 of Cost (other account / 2024 / still owed)?",
        ),
    ]

    ws["A3"] = "Question"
    ws["B3"] = "What I need"
    ws["A3"].fill = NAVY
    ws["B3"].fill = NAVY
    ws["A3"].font = WHITE
    ws["B3"].font = WHITE
    confirmed_idx = {6, 7, 8, 9}  # 1-based question numbers
    mixed_idx = {1, 2}  # pattern confirmed; dollars still unallocated
    for i, (q, detail) in enumerate(questions, start=4):
        qnum = i - 3
        if qnum in confirmed_idx:
            fill = GREEN
        elif qnum in mixed_idx:
            fill = BLUE
        else:
            fill = YELLOW
        ws.cell(i, 1, q).font = CG_B
        ws.cell(i, 1).fill = fill
        ws.cell(i, 1).alignment = WRAP
        ws.cell(i, 2, detail).font = CG
        ws.cell(i, 2).fill = fill
        ws.cell(i, 2).alignment = WRAP
        ws.row_dimensions[i].height = 48

    r = 15
    ws.cell(r, 1, "LOCKED / MATCHED this pass (do not recast without saying so)").font = WHITE
    ws.cell(r, 1).fill = NAVY
    ws.merge_cells("A15:B15")
    locked = [
        ("Coinbase ACH net", f"${sums['coinbase_net']:,.2f}  (OUT ${sums['coinbase_out']:,.2f} / IN ${sums['coinbase_in']:,.2f}) — Investments tab, not EPGC"),
        ("Newstar Jewelers", f"${sums['newstar']:,.2f} jewelry fabrication COGS (10/14 $8,000 + 11/12 $8,000 + 12/24 $7,581)"),
        ("Berk sales", "$14,000 seals (1/10) + $30,000 lots (2/7) = $44,000 Art Sales"),
        ("Mosaics", "EOEB $105,000 sale (2/21) / Plutus $50,000 of $90,000 Cost"),
        ("BoA 9922 draws", f"${sums['boa_out']:,.2f} owner transfer — not P&L"),
        ("Wise/EOEB 12/2", "$565.75 reimbursed — not income, not expense"),
        ("$10 test wires", "Jamal Rifai / EPGC / Plutus tests — not P&L"),
        ("David Aaron Limited", "$13,595 (6/24) consultant fee → EPGC Consultant June. Not a sale."),
        ("Koziol / Ariadne", "$150,000 IN 10/20 + $150,000 OUT 10/21 LOCKED pass-through — not P&L"),
        ("Wise expertise write-ups", "$334.17 (7/9) + $658.63 (10/30) = $992.80 → EPGC Consultant Fees expense. Dec $565.75 still reimbursed / not P&L."),
        ("Aquinas Hobor books", "$1,000 IN 2/4 LOCKED book sale → EPGC Art Sales February +$1,000 (Feb $136,000; matched cash $150,000). Cost / titles TBD (not in I8 until Cost, same as Sale 6428)."),
        ("EOEB / L5 mixed pattern", "CONFIRMED mixed art sales AND advisory/consultant. Remainder $130,270 / $87,396.32 unallocated — not on P&L until invoice split. Jan 16 $1,000 Canosan-horse still open."),
        ("Fortuna Venus inventory", "$20,000 OUT 5/9 LOCKED. Mercury native note: bronze head of a goddess, likely Venus. Parked Cost — not sold 2025, not dumped onto Art Sales/Consultant."),
        ("Fortuna Roman Gold Belt", "$45,000 OUT 7/21 LOCKED. Mercury native note: Roman Gold Belt. Parked Cost — not sold 2025, not dumped onto Art Sales/Consultant."),
        ("Native CSV 2026-09-20", "epgc-llc-transactions-2025-jan-01-to-2025-dec-31.csv is account 8291, same 2025 year already classified. 0 new unique Sent cash vs the 64. Failed Aysel 7/17 and Cancelled Erdal 1/8 are not cash."),
    ]
    for i, (k, v) in enumerate(locked, start=16):
        ws.cell(i, 1, k).font = CG
        ws.cell(i, 1).fill = GREEN
        ws.cell(i, 2, v).font = CG
        ws.cell(i, 2).fill = GREEN

    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 110
    ws.column_dimensions["C"].width = 20
    return ws


def patch_epgc(ws: Worksheet, art_sales_months, consultant_months, consultant_fees_months, note: str) -> None:
    # 2025 Art Sales is row 53, Consultant row 54 (from rebuild_like_prior_2025).
    # Consultant Fees expense is row 69 (distinct from Consultant income).
    write_months(ws, 53, 2, art_sales_months, PEACH, zero_fill=GRAY)
    write_months(ws, 54, 2, consultant_months, PEACH, zero_fill=GRAY)
    n53 = ws.cell(53, 14)
    n53.value = "=SUM(B53:M53)"
    n53.number_format = ACCT
    n54 = ws.cell(54, 14)
    n54.value = "=SUM(B54:M54)"
    n54.number_format = ACCT
    # Wipe 2024 values that copy_row_block left on the 2025 expense block.
    for row in (58, 59, 60, 61, 62, 65, 69, 71, 72):
        for col in range(2, 14):
            cell = ws.cell(row, col)
            if isinstance(cell.value, (int, float, Decimal)):
                cell.value = 0
                cell.fill = GRAY
                cell.number_format = ACCT
    write_months(ws, 69, 2, consultant_fees_months, PEACH, zero_fill=GRAY)
    n69 = ws.cell(69, 14)
    n69.value = "=SUM(B69:M69)"
    n69.number_format = ACCT
    ws["A75"] = note
    ws["A75"].font = CG
    ws["A75"].alignment = WRAP
    try:
        ws.merge_cells("A75:N76")
    except Exception:
        pass
    ws.row_dimensions[75].height = 72


def patch_investments(ws: Worksheet, coin_months, coin_in, coin_out, coin_net) -> None:
    ws["A1"] = "Fidelity"
    ws["B1"] = "SEP IRA"
    ws["A2"] = "ML"
    ws["B2"] = "UTMA EMMA COLEY"
    ws["A3"] = "ML"
    ws["B3"] = "UTMA PHOEBE COLEY"
    ws["A5"] = "Coinbase (ACH from Mercury 8291)"
    ws["A5"].font = Font(name="Century Gothic", size=12, bold=True)
    ws["B5"] = "Investment transfers — not EPGC Art Sales, not advisory. Distinct from Coinbase 6108 card."
    ws["B5"].font = CG
    headers = ["", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC", "2025 NET"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(6, i, h)
        cell.fill = NAVY
        cell.font = WHITE
    ws["A7"] = "To Coinbase (funding)"
    ws["A8"] = "From Coinbase (back to 8291)"
    ws["A9"] = "Net transfer"
    for i, v in enumerate(coin_months):
        # coin_months is signed net by month
        pass
    # rebuild in/out months from signed net isn't enough — pass arrays
    ws["A11"] = (
        f"2025: funded ${money(coin_out) * -1:,.2f} / returned ${money(coin_in):,.2f} / net ${money(coin_net):,.2f} "
        "off the EPGC operating account. Not income, not an expense. Not tax advice."
    )
    ws["A11"].alignment = WRAP
    ws.merge_cells("A11:N11")
    ws.row_dimensions[11].height = 32
    for col in range(1, 15):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions["A"].width = 32


def write_coinbase_months(ws, out_m, in_m, net_m):
    write_months(ws, 7, 2, out_m, BLUE, zero_fill=GRAY)
    write_months(ws, 8, 2, in_m, BLUE, zero_fill=GRAY)
    write_months(ws, 9, 2, net_m, GREEN, zero_fill=GRAY)
    for row, formula in ((7, "=SUM(B7:M7)"), (8, "=SUM(B8:M8)"), (9, "=SUM(B9:M9)")):
        c = ws.cell(row, 14)
        c.value = formula
        c.number_format = ACCT
        c.font = CG


def _already_has(ws: Worksheet, needle: str) -> bool:
    for r in range(1, (ws.max_row or 1) + 1):
        v = ws.cell(r, 1).value
        if v and needle in str(v):
            return True
    return False


def reclass_david_aaron(ws: Worksheet) -> None:
    """Always convert the proposed David Aaron sale row off Art Sales.

    patch_art_sales returns early once Newstar rows exist, so this must run
    on every apply — user 2026-09-20 locked it as a consultant fee, not a sale.
    """
    for r in range(1, (ws.max_row or 1) + 1):
        a = str(ws.cell(r, 1).value or "")
        b = str(ws.cell(r, 2).value or "")
        blob = a + " " + b
        if "proposed sale / ASK cash" in a:
            ws.cell(r, 1, "2025 Mercury 8291 — David Aaron Limited is LOCKED consultant fee (not Art Sales)")
            ws.cell(r, 1).font = CG_B
            ws.cell(r, 1).fill = GREEN
            continue
        # Only the dedicated sale row (not the A78 summary note, which is merged).
        if not (
            a.startswith("Proposed:")
            or a.startswith("NOT a 2025 art sale")
            or (b == "David Aaron Limited" and "MATCHED" not in a)
        ):
            continue
        ws.cell(r, 1, "NOT a 2025 art sale — David Aaron Limited consultant fee (EPGC Consultant June)")
        ws.cell(r, 2, "David Aaron Limited")
        for col in range(1, 10):
            ws.cell(r, col).fill = GREEN
            ws.cell(r, col).font = CG
        ws.cell(r, 3).value = None
        ws.cell(r, 4).value = None
        ws.cell(r, 5).value = None
        ws.cell(r, 6).value = None
        ws.cell(r, 7).value = None
        ws.cell(r, 8).value = None
        ws.cell(
            r,
            9,
            "LOCKED 2026-09-20: consultant fee, not a sale. $13,595 is on EPGC Consultant June. Not tax advice.",
        )


def _write_aquinas_register_row(ws: Worksheet, r: int) -> None:
    """Sale 6428 pattern: Sale Price locked, Cost TBD yellow, net off I8 until Cost."""
    ws.cell(r, 1, "Books (titles TBD)").fill = GREEN
    ws.cell(r, 1).font = CG
    ws.cell(r, 2, "Cost TBD").fill = YELLOW
    ws.cell(r, 2).font = CG
    ws.cell(r, 3).value = None
    ws.cell(r, 3).fill = YELLOW
    ws.cell(r, 4).value = None
    ws.cell(r, 5, "Aquinas Hobor").fill = GREEN
    ws.cell(r, 5).font = CG
    f = ws.cell(r, 6, 1000)
    f.number_format = ACCT
    f.fill = GREEN
    f.font = CG
    ws.cell(r, 7, datetime(2025, 2, 4))
    ws.cell(r, 7).number_format = "YYYY-MM-DD"
    ws.cell(r, 7).fill = GREEN
    ws.cell(r, 7).font = CG
    ws.cell(r, 8, f'=IF(OR(F{r}="",C{r}=""),"",F{r}-C{r})')
    ws.cell(r, 8).font = CG
    n = ws.cell(
        r,
        9,
        "LOCKED 2026-09-20: book SALE $1,000 on Mercury 8291. Cost TBD — not in Income I8 yet (same as Sale 6428). Titles TBD. Not tax advice.",
    )
    n.fill = YELLOW
    n.font = CG
    n.alignment = WRAP


def _clear_parked_aquinas(ws: Worksheet, rows: list[int]) -> None:
    for r in sorted(rows, reverse=True):
        header = str(ws.cell(r - 1, 1).value or "") if r > 1 else ""
        if "Books sold to Aquinas" in header and header.startswith("2025 Mercury"):
            ws.delete_rows(r - 1, 2)
        else:
            ws.delete_rows(r, 1)


def lock_aquinas_books(ws: Worksheet) -> None:
    """Always add/refresh the Aquinas books sale next to Sale 6428. Cost TBD yellow."""
    existing = []
    sale_row = None
    for r in range(1, (ws.max_row or 1) + 1):
        a = str(ws.cell(r, 1).value or "")
        if a.startswith("Sale 6428"):
            sale_row = r
        if ("Books (titles TBD)" in a or "Books sold to Aquinas" in a) and not a.startswith("2025 Mercury"):
            existing.append(r)
    register = [r for r in existing if r <= 80]
    parked = [r for r in existing if r > 80]
    if register:
        _write_aquinas_register_row(ws, register[0])
        _clear_parked_aquinas(ws, parked)
        return
    insert_at = (sale_row + 1) if sale_row else 77
    to_unmerge = [str(rng) for rng in list(ws.merged_cells.ranges) if rng.min_row <= insert_at <= rng.max_row]
    for rng in to_unmerge:
        try:
            ws.unmerge_cells(rng)
        except Exception:
            pass
    ws.insert_rows(insert_at)
    _write_aquinas_register_row(ws, insert_at)
    shifted = [r + 1 if r >= insert_at else r for r in parked]
    _clear_parked_aquinas(ws, shifted)


def _fix_art_sales_footer(ws: Worksheet, note: str) -> None:
    """Keep Sale 6428 / Aquinas / Total / one MATCHED note in that order."""
    sale_row = aquinas_row = total_row = inv_row = None
    matched_rows = []
    for r in range(70, min((ws.max_row or 80) + 1, 90)):
        a = str(ws.cell(r, 1).value or "")
        if a.startswith("Sale 6428"):
            sale_row = r
        elif a == "Books (titles TBD)":
            aquinas_row = r
        elif a == "Total":
            total_row = r
        elif a.startswith("Mercury 8291 MATCHED"):
            matched_rows.append(r)
        elif a.startswith("Art Purchases"):
            inv_row = r
    for r in sorted(matched_rows[1:], reverse=True):
        ws.delete_rows(r)
        if total_row and total_row > r:
            total_row -= 1
        if inv_row and inv_row > r:
            inv_row -= 1
        matched_rows = [x if x < r else x - 1 for x in matched_rows if x != r]
    if total_row is None:
        insert_at = (aquinas_row or sale_row or 76) + 1
        to_unmerge = [str(rng) for rng in list(ws.merged_cells.ranges) if rng.min_row <= insert_at <= rng.max_row]
        for rng in to_unmerge:
            try:
                ws.unmerge_cells(rng)
            except Exception:
                pass
        ws.insert_rows(insert_at)
        ws.cell(insert_at, 1, "Total").font = CG
        h = ws.cell(insert_at, 8, "=SUM(H51:H75)")
        h.number_format = ACCT
        h.font = CG
        total_row = insert_at
        matched_rows = [x + 1 if x >= insert_at else x for x in matched_rows]
        if inv_row and inv_row >= insert_at:
            inv_row += 1
    note_row = matched_rows[0] if matched_rows else (total_row + 1 if total_row else 79)
    # If MATCHED landed on Total or Aquinas, push it after Total.
    a_note = str(ws.cell(note_row, 1).value or "")
    if note_row == total_row or a_note == "Total" or a_note == "Books (titles TBD)":
        note_row = total_row + 1
    ws.cell(note_row, 1, note)
    ws.cell(note_row, 1).fill = GREEN
    ws.cell(note_row, 1).alignment = WRAP
    try:
        ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=9)
    except Exception:
        pass
    ws.row_dimensions[note_row].height = 48


def lock_fortuna_named_inventory(ws: Worksheet) -> None:
    """Always refresh Fortuna inventory rows from native CSV object names.

    Venus $20k and Roman Gold Belt $45k LOCKED parked Cost. Aug $20k stays ASK.
    Do not add named unsold inventory to sold-lot Cost.
    """
    header_row = None
    by_amt_date = {}
    remaining_row = None
    for r in range(1, (ws.max_row or 1) + 1):
        a = str(ws.cell(r, 1).value or "")
        dt = str(ws.cell(r, 4).value or "")
        cost = ws.cell(r, 3).value
        if "Fortuna inventory" in a or (a.startswith("2025 Mercury ASK") and "Fortuna" in a):
            header_row = r
        if a.startswith("Fortuna remaining 2025"):
            remaining_row = r
        if a.startswith("Fortuna / Erdal") or "Venus" in a or "Roman Gold Belt" in a:
            by_amt_date[(str(cost), dt[:10] if dt else dt)] = r
    targets = [
        (
            ("20000", "2025-05-09"),
            "SOLD 2025 — Venus joint (see Sales 2025). Not unsold inventory.",
            GREEN,
            "LOCKED 2026-09-21 joint SALE. EPGC proceeds $26,000 / Cost $20,000 / net $6,000. Sale Price = remittance, not full hammer.",
        ),
        (
            ("45000", "2025-07-21"),
            "SOLD 2025 — Roman Gold Belt to L5 (see Sales 2025). Not unsold inventory.",
            GREEN,
            "LOCKED 2026-09-21 SALE. L5 $50,000 / Cost $45,000 / net $5,000. Remaining L5 not attached.",
        ),
        (
            ("20000", "2025-08-15"),
            "Fortuna / Erdal Dere — object still unnamed (native 8/18 = Monarch 8/15)",
            YELLOW,
            "ASK: same $20,000 as Mercury native 2025-08-18. No object name. Do not add to sold Cost. Not a second unique txn.",
        ),
    ]
    # also match float costs
    def find_row(cost_s, date_s):
        for key, r in by_amt_date.items():
            c, d = key
            try:
                same_cost = D(c) == D(cost_s)
            except Exception:
                same_cost = str(c).startswith(cost_s)
            if same_cost and date_s in d:
                return r
        return None

    for (cost_s, date_s), title, fill, note in targets:
        r = find_row(cost_s, date_s)
        if r is None:
            continue
        ws.cell(r, 1, title).fill = fill
        ws.cell(r, 1).font = CG
        ws.cell(r, 2, "Mercury 8291").fill = fill
        ws.cell(r, 2).font = CG
        c = ws.cell(r, 3)
        try:
            c.value = float(cost_s)
        except Exception:
            pass
        c.number_format = ACCT
        c.fill = fill
        ws.cell(r, 4, date_s).fill = fill
        ws.cell(r, 9, note).fill = fill
        ws.cell(r, 9).font = CG
        ws.cell(r, 9).alignment = WRAP
    if header_row:
        ws.cell(
            header_row,
            1,
            "2025 Mercury Fortuna inventory — Venus + Roman Gold Belt LOCKED; Aug $20k unnamed ASK",
        )
        ws.cell(header_row, 1).font = CG_B
        ws.cell(header_row, 1).fill = GREEN
    if remaining_row:
        ws.cell(remaining_row, 1, "Fortuna unnamed remaining 2025 (excl. seals $13,000 + Venus $20,000 + Belt $45,000)")
        ws.cell(remaining_row, 3, 20000)
        ws.cell(remaining_row, 3).number_format = ACCT
        ws.cell(remaining_row, 3).fill = YELLOW


def patch_art_sales(ws: Worksheet) -> None:
    # Restore net formulas on Berk lots 51–70; F71 stays $30,000 lump.
    for r in range(51, 71):
        ws.cell(r, 8).value = f'=IF(OR(F{r}="",C{r}=""),"",F{r}-C{r})'
        ws.cell(r, 9).value = f'=IF(OR(H{r}="",C{r}=0),"",H{r}/C{r})'
    ws["F71"] = 30000
    ws["F71"].number_format = ACCT
    ws["F71"].fill = GREEN
    ws["G71"] = datetime(2025, 2, 7)
    ws["G71"].number_format = "YYYY-MM-DD"
    ws["G71"].fill = GREEN
    ws["H71"] = "=F71-C71"
    ws["C73"].fill = GREEN
    ws["F73"].fill = GREEN
    ws["D73"] = datetime(2025, 1, 8)
    ws["D73"].number_format = "YYYY-MM-DD"
    ws["G73"] = datetime(2025, 1, 10)
    ws["G73"].number_format = "YYYY-MM-DD"
    ws["C74"].fill = GREEN
    ws["F74"].fill = GREEN
    matched_note = (
        "Mercury 8291 MATCHED: Berk $14,000 (1/10 seals) + $30,000 (2/7 lots). "
        "EOEB $105,000 (2/21 mosaics). Fortuna $13,000 (1/8 seals Cost). "
        "Plutus $50,000 of mosaics $90,000 Cost — $40,000 of Cost not on 8291 (ASK). "
        "Newstar $23,581 is jewelry fabrication COGS, parked below — not in sold-lot Cost "
        "(Berk gem/scarab lots already sold in February). "
        "David Aaron $13,595 (6/24) is LOCKED consultant fee — not a sale (EPGC Consultant June). "
        "Aquinas Hobor $1,000 (2/4) is LOCKED book sale — EPGC Art Sales February; Cost TBD (not in I8). Not tax advice."
    )

    reclass_david_aaron(ws)
    lock_aquinas_books(ws)
    _fix_art_sales_footer(ws, matched_note)
    lock_fortuna_named_inventory(ws)

    # Insert Newstar / Fortuna ASK rows once. Do not insert David Aaron as a sale.
    if _already_has(ws, "Newstar Jewelers — jewelry from intaglios"):
        return
    last = 1
    for r in range(1, ws.max_row + 1):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, 10)):
            last = r
    start = last + 2
    ws.cell(start, 1, "2025 Mercury 8291 — David Aaron Limited is LOCKED consultant fee (not Art Sales)")
    ws.cell(start, 1).font = CG_B
    ws.cell(start, 1).fill = GREEN
    ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=9)
    r = start + 1
    ws.cell(r, 1, "NOT a 2025 art sale — David Aaron Limited consultant fee (EPGC Consultant June)").fill = GREEN
    ws.cell(r, 2, "David Aaron Limited").fill = GREEN
    ws.cell(
        r,
        9,
        "LOCKED 2026-09-20: consultant fee, not a sale. $13,595 is on EPGC Consultant June. Not tax advice.",
    ).fill = GREEN

    r = start + 3
    ws.cell(r, 1, "2025 Mercury COGS — Newstar Jewelers (LOCKED jewelry fabrication)")
    ws.cell(r, 1).font = CG_B
    ws.cell(r, 1).fill = GREEN
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    newstar = [
        ("Newstar Jewelers — jewelry from intaglios / engraved gems / scarabs", "Mercury 8291", 8000.00, "2025-10-14"),
        ("Newstar Jewelers — jewelry from intaglios / engraved gems / scarabs", "Mercury 8291", 8000.00, "2025-11-12"),
        ("Newstar Jewelers — jewelry from intaglios / engraved gems / scarabs (3 of 3)", "Mercury 8291", 7581.00, "2025-12-24"),
    ]
    for i, (obj, src, cost, dt) in enumerate(newstar, start=r + 1):
        ws.cell(i, 1, obj).fill = GREEN
        ws.cell(i, 2, src).fill = GREEN
        c = ws.cell(i, 3, cost)
        c.number_format = ACCT
        c.fill = GREEN
        ws.cell(i, 4, dt).fill = GREEN
        ws.cell(i, 9, "LOCKED COGS fabrication. Not in sold-lot Cost total (Feb Berk lots already sold). Identify which gems/intaglios/scarabs.")
    tot_row = r + 4
    ws.cell(tot_row, 1, "Newstar 2025 total")
    ws.cell(tot_row, 3, f"=SUM(C{r+1}:C{r+3})")
    ws.cell(tot_row, 3).number_format = ACCT
    ws.cell(tot_row, 3).fill = GREEN

    r = tot_row + 2
    ws.cell(r, 1, "2025 Mercury Fortuna inventory — Venus + Roman Gold Belt LOCKED; Aug $20k unnamed ASK")
    ws.cell(r, 1).font = CG_B
    ws.cell(r, 1).fill = GREEN
    fortuna = [
        ("Bronze head of a goddess, likely Venus — Fortuna inventory", "Mercury 8291", 20000.00, "2025-05-09"),
        ("Roman Gold Belt — Fortuna inventory", "Mercury 8291", 45000.00, "2025-07-21"),
        ("Fortuna / Erdal Dere — object still unnamed (native 8/18 = Monarch 8/15)", "Mercury 8291", 20000.00, "2025-08-15"),
    ]
    for i, (obj, src, cost, dt) in enumerate(fortuna, start=r + 1):
        named = "unnamed" not in obj
        fill = GREEN if named else YELLOW
        ws.cell(i, 1, obj).fill = fill
        ws.cell(i, 2, src).fill = fill
        c = ws.cell(i, 3, cost)
        c.number_format = ACCT
        c.fill = fill
        ws.cell(i, 4, dt).fill = fill
        ws.cell(
            i,
            9,
            "LOCKED parked Cost. Not sold 2025 — not in sold-lot Cost."
            if named
            else "ASK: same $20,000 as Mercury native 2025-08-18. No object name. Do not add to sold Cost.",
        )
    ws.cell(r + 4, 1, "Fortuna unnamed remaining 2025 (excl. seals $13,000 + Venus $20,000 + Belt $45,000)")
    ws.cell(r + 4, 3, 20000)
    ws.cell(r + 4, 3).number_format = ACCT
    ws.cell(r + 4, 3).fill = YELLOW


def patch_income(ws: Worksheet, art_net: Decimal) -> None:
    ws["I8"] = money(art_net)
    ws["I8"].number_format = ACCT
    ws["I8"].fill = YELLOW
    ws["I8"].font = CG
    ws["A12"] = (
        "2025 Actual: I4 Monarch paycheck cash ≠ W-2 Box 1. I5 216 LTR cash $19,500. "
        "I6 524 #2 STR platform net LOCKED. I7 GCM 1099-NEC $21,500. "
        f"I8 Art net ${money(art_net):,.2f} = MATCHED Mercury deals only "
        "(Berk lots $30,000 − $22,944.94 Cost) + seals $1,000 + mosaics $15,000. "
        "Excludes Sale 6428 $11,000 until Cost. Excludes Aquinas books $1,000 until Cost. "
        "David Aaron $13,595 is EPGC Consultant June "
        "(not I8). Erdal/Aysel/EOEB/L5 ASK. Not tax advice."
    )
    ws["A12"].alignment = WRAP
    ws.row_dimensions[12].height = 48


def summarize(rows: list[dict]) -> dict:
    def s(pred):
        return sum((r["amount"] for r in rows if pred(r)), D(0))

    def sabs_out(pred):
        return sum((-r["amount"] for r in rows if pred(r) and r["amount"] < 0), D(0))

    eoeb = [r for r in rows if "EOEB" in r["counterparty"].upper() or "EOEB" in r["statement"].upper()]
    l5 = [r for r in rows if r["counterparty"].strip() == "L5" or "Merchant name: L5" in r["statement"]]
    fortuna = [r for r in rows if "Fortuna" in r["statement"] or "Fortuna" in r["counterparty"]]
    erdal_in = [
        r
        for r in rows
        if r["amount"] > 0
        and "ERDAL" in (r["counterparty"] + " " + r["statement"]).upper()
        and r["bucket"] != "REIMBURSE"
    ]
    # 11/7 is Erdal Fortuna PAYMENT — include in erdal_in
    coin = [r for r in rows if r["bucket"] == "INVESTMENT"]
    newstar = [r for r in rows if r["bucket"] == "COGS_JEWELRY"]
    boa = [r for r in rows if r["bucket"] == "TRANSFER"]
    mosaics_sale = D("105000")
    reimb = D("565.75")
    eoeb_in = sum((r["amount"] for r in eoeb if r["amount"] > 0), D(0))
    fortuna_out = sum((-r["amount"] for r in fortuna if r["amount"] < 0), D(0))
    fortuna_named_unsold = sum(
        (
            -r["amount"]
            for r in rows
            if r["bucket"] == "ART_PURCHASE"
            and r["status"] == "LOCKED"
            and ("Venus" in r["deal"] or "Gold Belt" in r["deal"] or "Venus" in r["note"] or "Gold Belt" in r["note"])
        ),
        D(0),
    )
    fortuna_ask = sum(
        (-r["amount"] for r in rows if r["bucket"] == "ART_PURCHASE" and r["status"] == "ASK" and ("Fortuna" in r["counterparty"] or "Fortuna" in r["deal"] or "Erdal" in r["counterparty"])),
        D(0),
    )
    return {
        "n": len(rows),
        "eoeb_in": eoeb_in,
        "eoeb_remainder": eoeb_in - mosaics_sale - reimb,
        "l5": sum(
            (
                r["amount"]
                for r in l5
                if not (r["bucket"] == "ART_SALE" and r["status"] == "LOCKED")
            ),
            D(0),
        ),
        "fortuna_out": fortuna_out,
        "fortuna_remainder_out": fortuna_ask,
        "fortuna_named_inventory": D(0),
        "erdal_in": sum(
            (r["amount"] for r in erdal_in if r["status"] != "LOCKED"),
            D(0),
        ),
        "aysel": s(lambda r: "Aysel" in r["counterparty"] or "Aysel" in r["statement"]),
        "coinbase_in": sum((r["amount"] for r in coin if r["amount"] > 0), D(0)),
        "coinbase_out": sum((r["amount"] for r in coin if r["amount"] < 0), D(0)),
        "coinbase_net": sum((r["amount"] for r in coin), D(0)),
        "newstar": sum((-r["amount"] for r in newstar), D(0)),
        "boa_out": sum((-r["amount"] for r in boa), D(0)),
        "david_aaron": D("13595.00"),
        "wise_expertise": D("334.17") + D("658.63"),
        "aquinas_books": D("1000.00"),
        "art_sale_locked_gross": D("14000")
        + D("30000")
        + D("105000")
        + D("1000")
        + D("50000")
        + D("26000"),
        "art_net_locked": (D("30000") - D("22944.94"))
        + D("1000")
        + D("15000")
        + D("5000")
        + D("6000"),
        "koziol": D("150000"),
    }


def sheet_to_csv(ws: Worksheet, path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        for row in ws.iter_rows(max_row=ws.max_row, max_col=max(ws.max_column or 1, 1)):
            w.writerow([("" if c.value is None else c.value) for c in row])


def export_sheet(wb, name: str, path: Path) -> None:
    out = Workbook()
    out.remove(out.active)
    src = wb[name]
    ws = out.create_sheet(name)
    for r in src.iter_rows(min_row=1, max_row=src.max_row, max_col=src.max_column):
        for cell in r:
            d = ws.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                d.font = copy(cell.font)
                d.fill = copy(cell.fill)
                d.border = copy(cell.border)
                d.alignment = copy(cell.alignment)
                d.number_format = cell.number_format
    for col, dim in src.column_dimensions.items():
        ws.column_dimensions[col].width = dim.width
    out.save(path)


def write_start_here(
    path: Path,
    sums: dict,
    mercury_sheet_url: str,
    epgc_url: str,
    income_url: str,
    lock_url: str = "(Drive URL filled after upload)",
) -> None:
    text = f"""START HERE — Mercury Bank 8291 (2026-09-20)

Not tax advice. Same Personal Income.xlsx tabs as last year.

Mercury Checking 8291 is the EPGC LLC operating account (Choice Financial).
64 unique 2025 cash transactions from Monarch IDs.

User uploaded native Mercury CSV 2026-09-20 15:39 UTC:
epgc-llc-transactions-2025-jan-01-to-2025-dec-31.csv (Drive {NATIVE_DRIVE_ID}).
Account 8291. Same 2025 year already classified — 0 new unique Sent cash vs the 64.
Failed Aysel $50k OUT 7/17 and Cancelled Erdal $13k 1/8 are not cash.
Fortuna $20k Monarch 8/15 = native 8/18 (date shift, not a second txn).

Open ASK (Q6–Q9 locked green; Q1/Q2 mixed pattern confirmed, dollars unallocated):
{mercury_sheet_url}

EPGC LLC 2025 (Art Sales cash $150,000 Jan $14,000 / Feb $136,000; Consultant June $13,595; Consultant Fees $992.80):
{epgc_url}

Lock sheet Q6 consultant / Q7 write-ups / Q8 pass-through / Q9 Aquinas books + Venus/Belt inventory:
{lock_url}

Income + properties (I8 still $23,055.06 — Aquinas Cost TBD):
{income_url}

LOCKED this pass
- Coinbase ACH net ${money(sums['coinbase_net']):,.2f} (funded ${money(-sums['coinbase_out']):,.2f} / back ${money(sums['coinbase_in']):,.2f}) → Investments. Not EPGC. Not the Coinbase 6108 card.
- Newstar Jewelers ${money(sums['newstar']):,.2f} → jewelry fabrication COGS (intaglios, engraved gems, scarabs). Not EPGC operating.
- Art sales matched to last year’s Art Sales tab: Berk seals $14,000 (1/10) + Berk lots $30,000 (2/7) + mosaics $105,000 (EOEB 2/21) + Aquinas books $1,000 (2/4).
- Seals Cost $13,000 = Fortuna 1/8 (native note: J.K. New York Collection). Mosaics Cost $90,000 of which Plutus paid $50,000 on Mercury ($40,000 still ASK).
- Fortuna Venus $20,000 (5/9) + Roman Gold Belt $45,000 (7/21) LOCKED inventory from native memos. Parked Cost — not sold 2025, not on P&L.
- BoA 9922 transfers ${money(sums['boa_out']):,.2f} = owner draws, not P&L.
- Dec 2 Wise $565.75 reimbursed by EOEB — not income.
- David Aaron Limited $13,595 (6/24) → EPGC Consultant June. User: consultant fee, not a sale.
- Koziol $150,000 (10/20) / Ariadne $150,000 (10/21) → LOCKED pass-through. Not P&L.
- Wise $334.17 (7/9) + $658.63 (10/30) → EPGC Consultant Fees (expertise write-ups). Dec $565.75 still reimbursed.
- Aquinas Hobor $1,000 (2/4) → books sold. EPGC Art Sales February. Cost / titles TBD (not in I8 until Cost).

EPGC 2025 Art Sales is MATCHED/LOCKED cash ($150,000: Jan $14,000 / Feb $136,000).
EPGC Consultant June is $13,595 (David Aaron).
EOEB remainder / L5 MIXED PATTERN CONFIRMED (art sales + advisory/consultant) — dollars unallocated; not on Consultant or Art Sales.

ANSWERED
6. David Aaron Limited $13,595 (6/24) — LOCKED consultant fee (EPGC Consultant June).
7. Wise $334.17 + $658.63 — LOCKED business expenses (expertise write-ups) on EPGC Consultant Fees.
8. Koziol $150,000 / Ariadne $150,000 — LOCKED pass-through (not P&L).
9. Aquinas Hobor $1,000 (2/4) — LOCKED book sale (EPGC Art Sales February +$1,000; Cost TBD, not in I8).
Fortuna Venus $20,000 (5/9) + Roman Gold Belt $45,000 (7/21) — LOCKED inventory (native CSV notes).

ASK remaining
1. EOEB remainder ${money(sums['eoeb_remainder']):,.2f} — MIXED PATTERN CONFIRMED; invoice-level split ASK. Do not dump onto Consultant or Art Sales. Jan 16 $1,000 Canosan-horse still open.
2. L5 ${money(sums['l5']):,.2f} — MIXED PATTERN CONFIRMED; invoice-level split ASK. Do not dump.
3. Fortuna unnamed ${money(sums['fortuna_remainder_out']):,.2f} — Aug $20,000 (Monarch 8/15 = native 8/18). Object name?
4. Erdal IN ${money(sums['erdal_in']):,.2f} — sales to Erdal?
5. Aysel Dere — Monarch 7/18 +$50,000 IN vs native Failed 7/17 −$50,000 OUT (recipient account does not exist). Phantom or omitted incoming?
10. Mosaics Cost $40,000 missing on Mercury 8291.

Income I8 Art net is ${money(sums['art_net_locked']):,.2f} from MATCHED deals with Cost only (Berk lots net + seals $1,000 + mosaics $15,000). Sale 6428 $11,000 and Aquinas books $1,000 wait on Cost. Consultant is EPGC, not I8.
"""
    path.write_text(text, encoding="utf-8")


def write_compact_lock(path: Path, sums: dict) -> None:
    rows = [
        ["Q", "Status", "What", "Where it hits"],
        [
            "6",
            "LOCKED",
            "David Aaron Limited $13,595 IN (6/24) consultant fee",
            "EPGC Consultant June $13,595. Not Art Sales. Not I8.",
        ],
        [
            "7",
            "LOCKED",
            "Wise $334.17 (7/9) + $658.63 (10/30) expertise write-ups",
            "EPGC Consultant Fees $992.80 (Jul $334.17 / Oct $658.63). Dec Wise $565.75 still REIMBURSE / not P&L.",
        ],
        [
            "8",
            "LOCKED",
            "Koziol $150,000 IN (10/20) / Ariadne $150,000 OUT (10/21)",
            "Pass-through — not P&L.",
        ],
        [
            "9",
            "LOCKED",
            "Aquinas Hobor $1,000 IN (2/4) book SALE",
            "EPGC Art Sales February +$1,000 (Feb $136,000; matched cash $150,000). Register: Books (titles TBD), Cost TBD. NOT in I8 until Cost (same as Sale 6428).",
        ],
        [
            "1",
            "PATTERN CONFIRMED / dollars unallocated",
            f"EOEB remainder ${money(sums['eoeb_remainder']):,.2f} mixed art sales + advisory/consultant",
            "Do NOT dump onto Consultant or Art Sales. Invoice-level split ASK. Jan 16 $1,000 Canosan-horse still open.",
        ],
        [
            "2",
            "PATTERN CONFIRMED / dollars unallocated",
            f"L5 ${money(sums['l5']):,.2f} mixed art sales + advisory/consultant",
            "Do NOT dump onto Consultant or Art Sales. Invoice-level split ASK.",
        ],
        [
            "ASK remaining",
            "ASK",
            f"Fortuna unnamed ${money(sums['fortuna_remainder_out']):,.2f} / Erdal IN ${money(sums['erdal_in']):,.2f} / Aysel Monarch IN vs native Failed OUT / mosaics Cost $40,000",
            "Aug $20k object name; Aysel phantom vs omitted incoming; invoice split still needed.",
        ],
        [
            "Fortuna Venus + Belt",
            "LOCKED",
            "Venus $20,000 (5/9) + Roman Gold Belt $45,000 (7/21) inventory from native CSV notes",
            "Parked Cost. Not sold 2025. Not dumped onto Art Sales or Consultant.",
        ],
        [
            "Native CSV 2026-09-20",
            "INGESTED",
            "epgc-llc-transactions-2025-jan-01-to-2025-dec-31.csv account 8291",
            "0 new unique Sent cash vs the 64. Failed Aysel 7/17 and Cancelled Erdal 1/8 not cash. Fortuna $20k date 8/15 Monarch = 8/18 native.",
        ],
        [
            "LOCKED totals",
            "LOCKED",
            "Consultant Fees Jul+Oct $992.80; Consultant June $13,595; Art Sales Feb $136,000; I8 $23,055.06",
            "Dec Wise $565.75 REIMBURSE. EOEB remainder / L5 not on P&L.",
        ],
    ]
    with path.open("w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)


def main() -> None:
    rows = load_mercury_rows()
    sums = summarize(rows)
    native_stats = native_ingest_stats(load_native_rows())

    art_sales_months = monthly_sum(
        rows,
        lambda r: r["bucket"] == "ART_SALE" and r["status"] in ("LOCKED", "MATCHED"),
    )
    consultant_months = monthly_sum(
        rows,
        lambda r: r["bucket"] == "ADVISORY" and r["status"] == "LOCKED",
    )
    consultant_fees_months = zeros()
    for r in rows:
        if r["bucket"] == "EPGC_EXPENSE" and r["status"] == "LOCKED" and r["amount"] < 0:
            add(consultant_fees_months, r["month"], -r["amount"])
    coin_net_m = monthly_sum(rows, lambda r: r["bucket"] == "INVESTMENT")
    coin_out_m = monthly_sum(rows, lambda r: r["bucket"] == "INVESTMENT" and r["amount"] < 0)
    coin_in_m = monthly_sum(rows, lambda r: r["bucket"] == "INVESTMENT" and r["amount"] > 0)

    wb = load_workbook(XLSX)
    build_ledger_sheet(wb, rows)
    build_ask_sheet(wb, rows, sums)
    preferred = [
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
        "524 Home Sale",
        "827 Grove CapEx",
        "Childcare 2441",
        "MERCURY 8291",
        "ASK Mercury",
    ]
    existing = [n for n in preferred if n in wb.sheetnames]
    extras = [n for n in wb.sheetnames if n not in existing]
    wb._sheets[:] = [wb[n] for n in existing + extras]
    patch_epgc(
        wb["EPGC LLC"],
        art_sales_months,
        consultant_months,
        consultant_fees_months,
        (
            "2025 Mercury 2026-09-20: Art Sales = MATCHED/LOCKED cash "
            "(Berk $14,000 Jan + Berk $30,000 / mosaics $105,000 / Aquinas books $1,000 Feb = $150,000). "
            "Consultant June $13,595 = David Aaron Limited LOCKED (user: consultant fee, not a sale). "
            "Consultant Fees July $334.17 + October $658.63 = Wise expertise write-ups LOCKED "
            f"(${money(sums['wise_expertise']):,.2f}). "
            f"EOEB remainder ${money(sums['eoeb_remainder']):,.2f} and L5 ${money(sums['l5']):,.2f} "
            "MIXED PATTERN CONFIRMED (art sales + advisory/consultant) — invoice-level split still ASK; "
            "not dumped onto Consultant or Art Sales. "
            "Koziol/Ariadne $150,000 LOCKED pass-through — not P&L. "
            "Coinbase is Investments, not this tab. Newstar $23,581 is Art Sales COGS, not operating expense. "
            "Dec Wise $565.75 reimbursed — not on Consultant Fees. "
            "Aquinas books Cost TBD — cash on Art Sales, not in I8. "
            "GCM 1099 is personal. Not tax advice."
        ),
    )
    inv = wb["Investments"]
    patch_investments(inv, coin_net_m, sums["coinbase_in"], sums["coinbase_out"], sums["coinbase_net"])
    write_coinbase_months(inv, coin_out_m, coin_in_m, coin_net_m)
    patch_art_sales(wb["Art Sales and Purchases"])
    patch_income(wb["Income"], sums["art_net_locked"])

    wb.save(XLSX)
    DELIVERABLE.parent.mkdir(exist_ok=True)
    shutil.copy2(XLSX, DELIVERABLE)
    PARTS.mkdir(exist_ok=True)
    CSV_DIR.mkdir(exist_ok=True)

    export_names = [
        ("MERCURY 8291", "MERCURY_LEDGER.csv"),
        ("ASK Mercury", "ASK_Mercury.csv"),
        ("EPGC LLC", "EPGC_LLC.csv"),
        ("Art Sales and Purchases", "Art_Sales.csv"),
        ("Investments", "Investments.csv"),
        ("Income", "Income.csv"),
    ]
    for name, fname in export_names:
        sheet_to_csv(wb[name], CSV_DIR / fname)
        safe = name.replace(",", "").replace(".", "")[:40]
        export_sheet(wb, name, PARTS / f"{safe}.xlsx")

    # Multi-tab Mercury pack for Drive conversion
    pack = Workbook()
    pack.remove(pack.active)
    for name in ("ASK Mercury", "MERCURY 8291", "EPGC LLC", "Art Sales and Purchases", "Investments", "Income"):
        src = wb[name]
        ws = pack.create_sheet(name)
        for r in src.iter_rows(min_row=1, max_row=src.max_row, max_col=min(src.max_column or 1, 14)):
            for cell in r:
                d = ws.cell(cell.row, cell.column, cell.value)
                if cell.has_style:
                    d.font = copy(cell.font)
                    d.fill = copy(cell.fill)
                    d.alignment = copy(cell.alignment)
                    d.number_format = cell.number_format
        for col, dim in src.column_dimensions.items():
            ws.column_dimensions[col].width = dim.width
        if src.sheet_properties.tabColor:
            ws.sheet_properties.tabColor = src.sheet_properties.tabColor.rgb
    pack_path = PARTS / "Personal Income 2025 — Mercury 8291.xlsx"
    pack.save(pack_path)

    epgc_pack = Workbook()
    epgc_pack.remove(epgc_pack.active)
    for name in ("EPGC LLC", "Art Sales and Purchases", "GCM", "Investments", "ASK Mercury", "MERCURY 8291"):
        src = wb[name]
        ws = epgc_pack.create_sheet(name)
        for r in src.iter_rows(min_row=1, max_row=src.max_row, max_col=min(src.max_column or 1, 14)):
            for cell in r:
                d = ws.cell(cell.row, cell.column, cell.value)
                if cell.has_style:
                    d.font = copy(cell.font)
                    d.fill = copy(cell.fill)
                    d.alignment = copy(cell.alignment)
                    d.number_format = cell.number_format
        for col, dim in src.column_dimensions.items():
            ws.column_dimensions[col].width = dim.width
    epgc_pack_path = PARTS / "Personal Income 2025 — like last year (EPGC+Art+GCM).xlsx"
    epgc_pack.save(epgc_pack_path)

    payload = {
        "updated": "2026-09-20",
        "disclaimer": "Not tax advice. Organizational packet only.",
        "account": "Mercury Checking 8291 EPGC LLC",
        "n_txns": sums["n"],
        "locked": {
            "coinbase_net": float(sums["coinbase_net"]),
            "coinbase_out": float(sums["coinbase_out"]),
            "coinbase_in": float(sums["coinbase_in"]),
            "newstar_cogs": float(sums["newstar"]),
            "art_sale_locked_gross": float(sums["art_sale_locked_gross"]),
            "art_net_locked": float(sums["art_net_locked"]),
            "boa_9922_draws": float(sums["boa_out"]),
            "david_aaron_consultant": float(sums["david_aaron"]),
            "koziol_ariadne_passthrough": float(sums["koziol"]),
            "wise_expertise_writeups": float(sums["wise_expertise"]),
            "aquinas_books_sale": float(sums["aquinas_books"]),
            "roman_gold_belt_sale": 50000.0,
            "roman_gold_belt_cost": 45000.0,
            "venus_epgc_proceeds": 26000.0,
            "venus_epgc_cost": 20000.0,
        },
        "ask": {
            "eoeb_remainder": float(sums["eoeb_remainder"]),
            "l5": float(sums["l5"]),
            "eoeb_l5_pattern": "mixed art sales + advisory/consultant CONFIRMED; invoice-level split ASK; do not dump remainder",
            "fortuna_remainder_out": float(sums["fortuna_remainder_out"]),
            "fortuna_named_inventory": float(sums["fortuna_named_inventory"]),
            "erdal_in": float(sums["erdal_in"]),
            "aysel": float(sums["aysel"]),
            "aysel_native": "Monarch 7/18 +$50,000 IN vs native Failed 7/17 −$50,000 OUT; not dumped",
            "mosaics_cost_missing_on_mercury": 40000.0,
            "eoeb_jan16_canosan_horse": 1000.0,
            "aquinas_books_cost_tbd": True,
        },
        "epgc_art_sales_2025_monthly": [float(x) for x in art_sales_months],
        "epgc_consultant_2025_monthly": [float(x) for x in consultant_months],
        "epgc_consultant": float(sum(consultant_months)),
        "epgc_consultant_fees_2025_monthly": [float(x) for x in consultant_fees_months],
        "epgc_consultant_fees": float(sum(consultant_fees_months)),
        "native_csv": native_stats,
    }
    if JSON_PATH.exists():
        try:
            old = json.loads(JSON_PATH.read_text(encoding="utf-8"))
            if isinstance(old.get("drive"), dict):
                payload["drive"] = old["drive"]
        except Exception:
            pass
    JSON_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    drive = payload.get("drive") or {}

    def gsheet(key: str) -> str:
        fid = drive.get(key)
        return f"https://docs.google.com/spreadsheets/d/{fid}/edit" if fid else "(Drive URL filled after upload)"

    write_start_here(
        CSV_DIR / "START_HERE_mercury.txt",
        sums,
        mercury_sheet_url=gsheet("ask_mercury"),
        epgc_url=gsheet("epgc_2025_matched"),
        income_url="https://docs.google.com/spreadsheets/d/1XkpQn0ztm6reE5GzIhRqNJaA2mZNp6lPsfz2MuI523k/edit",
        lock_url=gsheet("mercury_ledger_q6_q9"),
    )
    write_compact_lock(CSV_DIR / "LOCK_Mercury_Q6_Q9.csv", sums)

    if PACKET.exists():
        pkt = json.loads(PACKET.read_text(encoding="utf-8"))
        pkt["updated"] = "2026-09-20"
        pkt["mercury_8291"] = payload
        pkt["xlsx_note"] = (
            "2026-09-20 native Mercury CSV ingest: 0 new unique Sent cash vs the 64. "
            "Fortuna Venus $20k (5/9) + Roman Gold Belt $45k (7/21) LOCKED inventory (parked Cost). "
            "Aug Fortuna $20k unnamed (8/15 Monarch = 8/18 native). "
            "Aysel Monarch IN vs native Failed OUT still ASK. "
            "Q6–Q9 + EOEB/L5 mixed unallocated unchanged. I8 $23,055.06."
        )
        PACKET.write_text(json.dumps(pkt, indent=2), encoding="utf-8")

    print("n", sums["n"])
    print("art sales months", [float(x) for x in art_sales_months], "year", float(sum(art_sales_months)))
    print("coinbase net", float(sums["coinbase_net"]))
    print("newstar", float(sums["newstar"]))
    print("eoeb remainder", float(sums["eoeb_remainder"]))
    print("l5", float(sums["l5"]))
    print("fortuna remainder", float(sums["fortuna_remainder_out"]))
    print("fortuna named inventory", float(sums["fortuna_named_inventory"]))
    print("native new unique vs 64", native_stats.get("new_unique_vs_64"))
    print("native missing from sent", native_stats.get("missing_from_native_sent"))
    print("native non-cash", native_stats.get("non_cash"))
    print("erdal in", float(sums["erdal_in"]))
    print("aysel", float(sums["aysel"]))
    print("boa", float(sums["boa_out"]))
    print("consultant months", [float(x) for x in consultant_months], "year", float(sum(consultant_months)))
    print("consultant fees months", [float(x) for x in consultant_fees_months], "year", float(sum(consultant_fees_months)))
    print("art net locked", float(sums["art_net_locked"]))
    print("aquinas books", float(sums["aquinas_books"]))
    print("art sale locked gross", float(sums["art_sale_locked_gross"]))
    print("saved", XLSX, XLSX.stat().st_size)
    print("pack", pack_path, pack_path.stat().st_size)


if __name__ == "__main__":
    main()
