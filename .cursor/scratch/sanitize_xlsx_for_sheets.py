#!/usr/bin/env python3
"""Make Personal Income 2025.xlsx convert as a Google Sheet with real formulas.

Last Drive upload went through LibreOffice ODS. That stored every =SUM as a
shared string (t="s"), so Sheets showed the formula text instead of a total.
Google also rejects xlsx that still carry last year's external workbook links
and #REF! defined names ("Invalid conversion requested").

This rewrites the zip in place:
  - drop xl/externalLinks
  - drop definedNames (#REF! and the leftover name 'a')
  - use relative relationship targets (openpyxl writes /xl/...)
  - unhide 524 Unit 1 so 2025 sale-year numbers are visible
  - keep <f>SUM(...) cells (never stringify formulas)

Not tax advice.
"""
from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

ROOT = Path("/workspace/.cursor/scratch")
DEFAULT = ROOT / "tax_turbo_parts" / "Personal Income 2025.xlsx"


def _strip_xml_tag_block(xml: str, tag: str) -> str:
    return re.sub(rf"<{tag}\b[^>]*>.*?</{tag}>", "", xml, flags=re.S)


def sanitize_xlsx(path: Path) -> None:
    src = zipfile.ZipFile(path)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as out:
        for item in src.infolist():
            name = item.filename
            if name.startswith("xl/externalLinks/"):
                continue
            data = src.read(name)
            if name == "xl/_rels/workbook.xml.rels":
                text = data.decode("utf-8")
                text = re.sub(
                    r"<Relationship[^>]*externalLink[^/]*/>",
                    "",
                    text,
                )
                text = text.replace('Target="/xl/worksheets/', 'Target="worksheets/')
                text = text.replace('Target="/xl/externalLinks/', 'Target="externalLinks/')
                text = text.replace('Target="/xl/theme/', 'Target="theme/')
                data = text.encode("utf-8")
            elif name == "xl/workbook.xml":
                text = data.decode("utf-8")
                text = _strip_xml_tag_block(text, "externalReferences")
                text = _strip_xml_tag_block(text, "definedNames")
                text = text.replace(
                    'name="524 Ferdinand Ave, Unit 1" sheetId="11" state="hidden"',
                    'name="524 Ferdinand Ave, Unit 1" sheetId="11" state="visible"',
                )
                text = text.replace('state="hidden"', 'state="visible"')
                data = text.encode("utf-8")
            elif name == "[Content_Types].xml":
                text = data.decode("utf-8")
                text = re.sub(r"<Override[^>]*externalLink[^/]*/>", "", text)
                data = text.encode("utf-8")
            out.writestr(name, data)
    path.write_bytes(buf.getvalue())


def assert_formulas_are_real(path: Path) -> None:
    z = zipfile.ZipFile(path)
    wb = z.read("xl/workbook.xml").decode("utf-8")
    if "externalReferences" in wb:
        raise SystemExit("LOCK FAIL: xlsx still has externalReferences (Google will not convert)")
    if "definedNames" in wb:
        raise SystemExit("LOCK FAIL: xlsx still has definedNames (includes #REF!)")
    if 'name="524 Ferdinand Ave, Unit 1"' in wb and "state=\"hidden\"" in wb:
        raise SystemExit("LOCK FAIL: Unit 1 is still hidden")
    s1 = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    if "<f>SUM(I4:I9)</f>" not in s1:
        raise SystemExit("LOCK FAIL: Income I10 is not an Excel <f> formula (Drive will show =SUM as text)")
    s3 = z.read("xl/worksheets/sheet3.xml").decode("utf-8")
    if "<f>SUM(AC5:AN5)</f>" not in s3:
        raise SystemExit("LOCK FAIL: 216 AO5 is not an Excel <f> formula")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    if "externalLink" in rels:
        raise SystemExit("LOCK FAIL: workbook rels still point at externalLink")
    if 'Target="/xl/worksheets/' in rels:
        raise SystemExit("LOCK FAIL: absolute /xl/ relationship targets (Google conversion rejects these)")
    print(f"SHEETS-SAFE OK {path} ({path.stat().st_size} bytes) real <f> formulas, no external links")


if __name__ == "__main__":
    import sys

    target = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
    sanitize_xlsx(target)
    assert_formulas_are_real(target)
