#!/usr/bin/env python3
"""Lean themed Personal Income 2025 for Drive MCP convert.

Prior converts either dropped Art Sales (sparse OOXML) or dropped column-A
labels (numeric-only payload). This file keeps:
  - Income labels + I8 $39,055.06 + I10 =SUM(I4:I9)
  - Art Sales rows 1–81 with 2022–2024 lots + Belt / Venus / August
  - EPGC 2025 July $76,000 / October $25,000 / Consultant June $13,595
  - 216 Oak Park 2025 + AO5 formula
and stubs the other last-year tabs so the 11-tab order survives.

Must include xl/theme (Google rejects no-theme custom OOXML). Keep the zip
small enough for Drive create_file (~20KB / ≲30k base64).

Not tax advice. Organizational packet only.
"""
from __future__ import annotations

import base64
import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

ROOT = Path("/workspace/.cursor/scratch")
SRC = ROOT / "tax_turbo_parts" / "Personal Income 2025.xlsx"
OUT = ROOT / "tax_turbo_parts" / "Personal Income 2025.labeled.xlsx"
B64 = Path("/tmp/drive_labeled.b64")

ORDER = [
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
    "2025 Data sources",
]
CG = Font(name="Century Gothic", size=10)
CGB = Font(name="Century Gothic", size=10, bold=True)
MONEY = "$#,##0.00"


def put(ws, r, c, val, bold=False, is_f=False):
    if val is None or val == "":
        return
    cell = ws.cell(r, c, val)
    cell.font = CGB if bold else CG
    if is_f or isinstance(val, (int, float)):
        cell.number_format = MONEY


def copy(src, wb, name, coords):
    ws_s = src[name]
    ws = wb[name] if name in wb.sheetnames else wb.create_sheet(name)
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 28
    for r, c in coords:
        sc = ws_s.cell(r, c)
        val = sc.value
        if val is None or val == "":
            continue
        is_f = sc.data_type == "f" and isinstance(val, str) and "[" not in val
        if is_f and len(val) > 48 and "IF(" in val:
            continue
        if isinstance(val, str) and len(val) > 70:
            val = val[:67] + "..."
        put(ws, r, c, val, bold=bool(sc.font and sc.font.bold), is_f=is_f)
    return ws


def prop_nonzero(src, name, last, year_cols):
    ws = src[name]
    coords = []
    for r in range(1, last + 1):
        coords.append((r, 1))
        coords.append((r, 2))
        for c in year_cols:
            cell = ws.cell(r, c)
            v = cell.value
            if cell.data_type == "f" or (v not in (None, "", 0, 0.0)):
                coords.append((r, c))
    return coords


def stub_prop(src, name, last, year_cols):
    ws = src[name]
    coords = []
    for r in range(1, min(last, 8) + 1):
        coords += [(r, 1), (r, 2)]
    for r in range(1, last + 1):
        if str(ws.cell(r, 2).value or "").strip():
            coords.append((r, 2))
        for c in year_cols:
            if ws.cell(r, c).data_type == "f":
                coords.append((r, c))
    return coords


def write_book(src_path: Path, dest: Path) -> None:
    src = load_workbook(src_path, data_only=False)
    wb = Workbook()
    wb.remove(wb.active)
    copy(src, wb, "Income", [(r, c) for r in range(1, 11) for c in range(1, 10)])
    copy(
        src,
        wb,
        "524 Ferdinand Ave, Unit 2",
        stub_prop(src, "524 Ferdinand Ave, Unit 2", 37, range(29, 42)),
    )
    copy(
        src,
        wb,
        "216 N. Oak Park Ave",
        prop_nonzero(src, "216 N. Oak Park Ave", 44, range(29, 42)),
    )
    copy(src, wb, "EPGC LLC", [(r, c) for r in range(51, 74) for c in range(1, 15)])
    copy(src, wb, "Refrence Library", [(1, 1), (74, 1), (75, 1), (75, 2)])
    copy(
        src,
        wb,
        "Art Sales and Purchases",
        [(r, c) for r in range(1, 82) for c in (1, 3, 6, 8, 11)],
    )
    copy(src, wb, "GCM", [(r, c) for r in range(20, 27) for c in range(1, 6)])
    copy(src, wb, "Hindman W2", [(r, c) for r in range(1, 10) for c in range(1, 6)])
    copy(src, wb, "Megan T4", [(r, c) for r in range(1, 10) for c in range(1, 6)])
    copy(src, wb, "Investments", [(r, c) for r in range(1, 8) for c in range(1, 15)])
    copy(
        src,
        wb,
        "524 Ferdinand Ave, Unit 1",
        stub_prop(src, "524 Ferdinand Ave, Unit 1", 43, range(16, 29)),
    )
    copy(src, wb, "2025 Data sources", [(r, c) for r in range(1, 7) for c in range(1, 4)])
    for i, name in enumerate(ORDER):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest)
    sys.path.insert(0, str(ROOT))
    from sanitize_xlsx_for_sheets import sanitize_xlsx

    sanitize_xlsx(dest)


def main() -> None:
    write_book(SRC, OUT)
    raw = OUT.read_bytes()
    B64.write_text(base64.b64encode(raw).decode("ascii"))
    names = zipfile.ZipFile(OUT).namelist()
    print("labeled", OUT, OUT.stat().st_size, "bytes", "b64", B64.stat().st_size)
    print("theme", any(n.startswith("xl/theme/") for n in names))
    sys.path.insert(0, str(ROOT))
    from sanitize_xlsx_for_sheets import assert_formulas_are_real

    assert_formulas_are_real(OUT)
    wb = load_workbook(OUT, data_only=False)
    art = wb["Art Sales and Purchases"]
    inc = wb["Income"]
    ep = wb["EPGC LLC"]
    assert inc["A4"].value == "Hindman "
    assert inc["I8"].value == 39055.06
    assert inc["I10"].value == "=SUM(I4:I9)"
    assert art["A1"].value == "Art Sales and Purchases"
    assert art["A4"].value == "An Ashanti Wood Fertility Figure"
    assert art["A78"].value == "Roman Gold Belt"
    assert art["C78"].value == 45000 and art["F78"].value == 50000
    assert "Venus" in str(art["A79"].value)
    assert art["F79"].value == 26000
    assert "August Fortuna" in str(art["A80"].value)
    assert art["F80"].value == 25000
    assert ep["H53"].value == 76000
    assert ep["K53"].value == 25000
    assert ep["G54"].value == 13595
    print("LOCK cells OK")


if __name__ == "__main__":
    main()
