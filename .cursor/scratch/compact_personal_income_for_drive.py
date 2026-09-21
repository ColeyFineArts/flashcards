#!/usr/bin/env python3
"""Compact Personal Income 2025.xlsx for Google Sheets conversion.

Keeps every value, formula, merge, and fill. Drops empty font-only cells
that bloat sheet XML so the Drive create_file payload stays small enough.
Never goes through ODS (that stores =SUM as text).

Not tax advice.
"""
from __future__ import annotations

import sys
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path("/workspace/.cursor/scratch")
SRC = ROOT / "tax_turbo_parts" / "Personal Income 2025.xlsx"
OUT = ROOT / "tax_turbo_parts" / "Personal Income 2025.drive.xlsx"


def cell_has_payload(cell) -> bool:
    if cell.value not in (None, ""):
        return True
    fill = cell.fill
    if fill is None or fill.fill_type in (None, "none"):
        return False
    fg = getattr(fill, "fgColor", None)
    rgb = getattr(fg, "rgb", None) if fg is not None else None
    if rgb in (None, "00000000", "0"):
        return False
    return True


def copy_sheet(src: Worksheet, dest: Worksheet) -> None:
    dest.sheet_state = src.sheet_state
    if src.sheet_properties.tabColor:
        dest.sheet_properties.tabColor = src.sheet_properties.tabColor.rgb
    dest.freeze_panes = src.freeze_panes
    max_r = src.max_row or 1
    max_c = src.max_column or 1
    for r in range(1, max_r + 1):
        rd = src.row_dimensions[r]
        if rd.height:
            dest.row_dimensions[r].height = rd.height
        for c in range(1, max_c + 1):
            s = src.cell(r, c)
            if not cell_has_payload(s):
                continue
            d = dest.cell(r, c, s.value)
            if s.has_style:
                d.font = copy(s.font)
                d.fill = copy(s.fill)
                d.border = copy(s.border)
                d.alignment = copy(s.alignment)
                d.number_format = s.number_format
    for col, dim in src.column_dimensions.items():
        if dim.width:
            dest.column_dimensions[col].width = dim.width
    for rng in src.merged_cells.ranges:
        dest.merge_cells(str(rng))
        # merged title cells must keep the anchor value
        min_col, min_row, max_col, max_row = rng.bounds
        anchor = src.cell(min_row, min_col)
        d = dest.cell(min_row, min_col, anchor.value)
        if anchor.has_style:
            d.font = copy(anchor.font)
            d.fill = copy(anchor.fill)
            d.alignment = copy(anchor.alignment)


def compact(src_path: Path, dest_path: Path) -> None:
    src = load_workbook(src_path)
    out = Workbook()
    out.remove(out.active)
    for name in src.sheetnames:
        copy_sheet(src[name], out.create_sheet(name))
    try:
        out.defined_names.clear()
    except Exception:
        pass
    out._external_links = []
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    out.save(dest_path)


if __name__ == "__main__":
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else SRC
    dest = Path(sys.argv[2]) if len(sys.argv) > 2 else OUT
    compact(src, dest)
    sys.path.insert(0, str(ROOT))
    from assert_personal_income_2025 import assert_workbook
    from inject_cached_formula_values import inject
    from sanitize_xlsx_for_sheets import assert_formulas_are_real, sanitize_xlsx

    sanitize_xlsx(dest)
    n = inject(dest)
    sanitize_xlsx(dest)
    assert_workbook(dest)
    assert_formulas_are_real(dest)
    print("drive xlsx", dest, dest.stat().st_size, "bytes", f"cached={n}")
