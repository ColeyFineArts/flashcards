#!/usr/bin/env python3
"""Fail loud when a Google Drive create_file result looks truncated/corrupt.

Cloud agents MUST run this after every create_file that claims to upload
a spreadsheet or workbook. Exit 0 = ok; exit 1 = trash and retry as text Sheet.

Usage:
  python .cursor/tools/drive_create_verify.py \\
    --got-filesize 1526 --expected-min-bytes 10000 --title "Personal Income.xlsx"

  python .cursor/tools/drive_create_verify.py --self-test
"""
from __future__ import annotations

import argparse
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# Known corrupt stub sizes from 2026-09-17 tax-turbo Drive thrash
KNOWN_STUB_SIZES = {1, 8, 1526, 7500}
DEFAULT_MIN_SPREADSHEET = 2048


def verify(
    got_filesize: int,
    expected_min_bytes: int,
    expected_exact: int | None = None,
    title: str = "",
) -> tuple[bool, str]:
    if got_filesize < 0:
        return False, f"invalid fileSize={got_filesize}"
    if got_filesize in KNOWN_STUB_SIZES:
        return (
            False,
            f"CORRUPT stub fileSize={got_filesize} (known truncation sizes "
            f"{sorted(KNOWN_STUB_SIZES)}); title={title!r} — trash and retry as Sheet/textContent",
        )
    if expected_exact is not None and got_filesize != expected_exact:
        return (
            False,
            f"SIZE MISMATCH got={got_filesize} expected_exact={expected_exact}; "
            f"title={title!r} — treat as truncated",
        )
    if got_filesize < expected_min_bytes:
        return (
            False,
            f"TOO SMALL got={got_filesize} < min={expected_min_bytes}; "
            f"title={title!r} — trash and retry as Sheet/textContent",
        )
    return True, f"OK fileSize={got_filesize} title={title!r}"


def self_test() -> None:
    cases = [
        # (got, min, exact, expect_ok)
        (1, 1000, None, False),
        (8, 1000, None, False),
        (1526, 1000, None, False),
        (7500, 10000, None, False),
        (27104, 20000, 27104, True),
        (27104, 20000, 34030, False),
        (5000, 2048, None, True),
        (500, 2048, None, False),
    ]
    failed = 0
    for got, mn, ex, want_ok in cases:
        ok, msg = verify(got, mn, ex, title="test")
        if ok != want_ok:
            failed += 1
            print(f"FAIL case got={got} min={mn} exact={ex}: want_ok={want_ok} got_ok={ok} ({msg})")
        else:
            print(f"PASS got={got} min={mn} exact={ex} -> {ok}")
    if failed:
        raise SystemExit(f"self-test: {failed} failure(s)")
    print("self-test: OK")


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--self-test", action="store_true")
    p.add_argument("--got-filesize", type=int)
    p.add_argument(
        "--expected-min-bytes",
        type=int,
        default=DEFAULT_MIN_SPREADSHEET,
        help=f"default {DEFAULT_MIN_SPREADSHEET} for spreadsheets",
    )
    p.add_argument(
        "--expected-exact",
        type=int,
        default=None,
        help="when uploading a known local file, pass its byte size",
    )
    p.add_argument("--title", default="")
    args = p.parse_args()

    if args.self_test:
        self_test()
        return

    if args.got_filesize is None:
        p.error("--got-filesize is required unless --self-test")

    ok, msg = verify(
        args.got_filesize,
        args.expected_min_bytes,
        args.expected_exact,
        args.title,
    )
    print(msg)
    raise SystemExit(0 if ok else 1)


if __name__ == "__main__":
    main()
