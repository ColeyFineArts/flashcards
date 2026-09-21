#!/usr/bin/env python3
"""Build a Drive-sized clone of last year's workbook with 2025 numbers.

Canonical source of truth:
  tax_turbo_parts/Personal Income 2025.xlsx
  (copy of Personal_Income_prior.xlsx + 2025 filled)

This does NOT rebuild from CSVs. It only shrinks empty-cell XML so Google
Drive can convert the xlsx to a Google Sheet. Run assert_personal_income_2025.py
on the canonical file first.

Not tax advice.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path("/workspace/.cursor/scratch")
CANON = ROOT / "tax_turbo_parts" / "Personal Income 2025.xlsx"


def main() -> None:
    sys.path.insert(0, str(ROOT))
    from assert_personal_income_2025 import assert_workbook

    assert_workbook(CANON)
    print(
        "LOCK OK. Drive upload must convert this xlsx (or a compact clone of it). "
        "Never upload a CSV rebuild. Never upload an 8-tab subset."
    )


if __name__ == "__main__":
    main()
