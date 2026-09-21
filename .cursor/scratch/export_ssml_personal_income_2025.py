#!/usr/bin/env python3
"""Compact SpreadsheetML 2003 of last year's workbook + 2025.

Google Drive converts this XML to a Sheet that actually calculates.

Google's SpreadsheetML importer treats ss:Formula as Excel R1C1. A1
formulas that look like R1C1 (SUM(C5:N5), SUM(C4:C9)) become #ERROR! or
sum the wrong columns. Always emit R1C1 relative refs.

Also emit every month cell C–N (empty placeholders for vacant months).
Google ignores Cell ss:Index, so skipped Oct/Nov shifted Dec into the
total column and SUM became circular.

Not tax advice.
"""
from __future__ import annotations

import html
import re
from datetime import datetime, date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

ROOT = Path("/workspace/.cursor/scratch")
SRC = ROOT / "tax_turbo_parts" / "Personal Income 2025.xlsx"
OUT = ROOT / "tax_turbo_parts" / "Personal_Income_2025_ssml.xml"

PEACH, GREEN, YELLOW = "F7CAAC", "C6EFCE", "FFF2CC"
SUM_RE = re.compile(r"^SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$")
SUM_ONE_RE = re.compile(r"^SUM\(([A-Z]+)(\d+)\)$")
MINUS_RE = re.compile(r"^([A-Z]+)(\d+)-([A-Z]+)(\d+)$")
SUM_MINUS_RE = re.compile(r"^SUM\(([A-Z]+\d+)-([A-Z]+\d+)\)$")
A1_TOKEN = re.compile(r"(?<![A-Z])(\$?)([A-Z]+)(\$?)(\d+)")
YEAR_TOTAL_COLS = {15, 28, 41}  # O, AB, AO
EPGC_FMLA_ROWS = {5, 22, 24, 30, 48, 50, 55, 73}

PROPERTY_TABS = {
    "524 Ferdinand Ave, Unit 2",
    "216 N. Oak Park Ave",
    "524 Ferdinand Ave, Unit 1",
}

LAST_YEAR_TABS = [
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
]
PACKET_TABS = LAST_YEAR_TABS + ["2025 Data sources"]


def year_start_for_total_col(c: int) -> int:
    if c == 15:
        return 3
    if c == 28:
        return 16
    if c == 41:
        return 29
    return 0


def row_has_month_values(ws, r: int, start: int) -> bool:
    if start <= 0:
        return False
    for c in range(start, start + 12):
        v = ws.cell(r, c).value
        if isinstance(v, (int, float)) and not isinstance(v, bool) and v != 0:
            return True
    return False


def src_2025_start(name: str) -> int:
    # Unit 1 2025 is P–AB (16); 216 and Unit 2 2025 is AC–AO (29).
    return 16 if name == "524 Ferdinand Ave, Unit 1" else 29


def property_dest_col(name: str, src_c: int) -> int | None:
    """Keep last year's layout: 2023 C–O, 2024 P–AB, 2025 AC–AO (Unit 1: 2023 C–O, 2025 P–AB)."""
    if src_c <= 2:
        return src_c
    limit = 28 if name == "524 Ferdinand Ave, Unit 1" else 41
    if 3 <= src_c <= limit:
        return src_c
    return None


def rewrite_formula_cols(fml, col_map: dict[int, int]) -> str:
    def repl(m):
        col, row = m.group(1), m.group(2)
        src = column_index_from_string(col)
        if src in col_map:
            return f"{get_column_letter(col_map[src])}{row}"
        return m.group(0)

    return re.sub(r"([A-Z]+)(\d+)", repl, fml)


def a1_to_r1c1_cell(col: int, row: int, dest_c: int, dest_r: int, abs_col: bool, abs_row: bool) -> str:
    if abs_row:
        rpart = f"R{row}"
    elif row == dest_r:
        rpart = "R"
    else:
        rpart = f"R[{row - dest_r}]"
    if abs_col:
        cpart = f"C{col}"
    elif col == dest_c:
        cpart = "C"
    else:
        cpart = f"C[{col - dest_c}]"
    return rpart + cpart


def formula_to_r1c1(fml: str, dest_r: int, dest_c: int) -> str:
    """SpreadsheetML formulas are R1C1. C5:N5 looks like R1C1 and #ERROR!s."""
    f = fml[1:] if fml.startswith("=") else fml

    def repl(m):
        abs_col = m.group(1) == "$"
        col = column_index_from_string(m.group(2))
        abs_row = m.group(3) == "$"
        row = int(m.group(4))
        return a1_to_r1c1_cell(col, row, dest_c, dest_r, abs_col, abs_row)

    return "=" + A1_TOKEN.sub(repl, f)


def keep_formula(ws, r, c, fml: str) -> bool:
    name = ws.title
    if "[" in fml or "![" in fml:
        return False
    if name == "Income":
        return True
    if name in PROPERTY_TABS:
        # Live year-total formulas for every year block (O / AB / AO).
        return c in YEAR_TOTAL_COLS and c <= max_col_for(name, ws)
    if name == "GCM":
        # Same-row R1C1 SUMs calculate; GCM's =R[-1]C / SUM(R[-1]C[-3]:R[-1]C)
        # recalculate to 0 in Sheets. Store computed 2022–2025 totals.
        return False
    if name in {"Hindman W2", "Megan T4"}:
        return False
    if name == "Investments":
        return False
    if name == "EPGC LLC":
        # 2025 Total / TOTAL EXPENSES stay live; 2023/2024 store calculated numbers.
        return r in {55, 73}
    if name == "Art Sales and Purchases":
        return r >= 49 and "SUM(" in fml.upper()
    if name == "2025 Data sources":
        return True
    return False


def row_label(ws, r: int) -> str:
    for c in (1, 2):
        v = ws.cell(r, c).value
        if isinstance(v, str) and v.strip():
            return v.strip()
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return str(v)
    return ""


def dest_row_index(name: str, r: int) -> int:
    """Keep last year's row order (oldest year first, 2025 appended)."""
    return r


def keep_row(name: str, r: int, ws=None) -> bool:
    if name == "Income":
        return r <= 10
    if name == "2025 Data sources":
        return r <= 12
    if name == "EPGC LLC":
        # Last year stacked 2023 then 2024 then 2025. Keep all three P&L blocks.
        return r <= 73
    if name in PROPERTY_TABS and ws is not None:
        # Last-year P&L grid through YTD (skip closing-cost notes underneath).
        last_pnl = {
            "524 Ferdinand Ave, Unit 2": 37,
            "524 Ferdinand Ave, Unit 1": 43,
            "216 N. Oak Park Ave": 44,
        }[name]
        if r > last_pnl:
            return False
        if r <= 2:
            return True
        if row_label(ws, r):
            return True
        if name == "524 Ferdinand Ave, Unit 1":
            return row_has_month_values(ws, r, 3) or row_has_month_values(ws, r, 16)
        return (
            row_has_month_values(ws, r, 3)
            or row_has_month_values(ws, r, 16)
            or row_has_month_values(ws, r, 29)
        )
    if name == "Art Sales and Purchases":
        # Include row 1 and 2022–2024 lots. Sparse tiny OOXML skipped these
        # and Google convert dropped Art cells. Keep 2025 Belt/Venus rows too.
        if r == 1:
            return True
        return 2 <= r <= 82
    if name == "Refrence Library":
        return r == 1 or r >= 74  # column headers + 2025 Jewelry Books
    if name == "GCM":
        return r <= 26
    if name in {"Hindman W2", "Megan T4"}:
        return True
    return True


def rgb(color) -> str | None:
    if color is None:
        return None
    val = getattr(color, "rgb", None)
    if not val or val in ("00000000", "0"):
        return None
    s = str(val).upper()
    if len(s) == 8 and s.endswith(("000000", "00000000")):
        return None
    if len(s) >= 6:
        s = s[-6:]
    if s in ("000000", "FFFFFF"):
        return None
    if len(s) == 6:
        return s
    return None


def keep_value(v, cell, ws) -> bool:
    if isinstance(v, str) and v != "":
        return True
    if v not in (None, "", 0, 0.0):
        return True
    bg = None
    if cell.fill and cell.fill.fill_type not in (None, "none"):
        bg = rgb(getattr(cell.fill, "fgColor", None))
    if bg == GREEN:
        return True
    if bg == YELLOW and v not in (0, 0.0, None, ""):
        return True
    # Peach/yellow zeros are empty month placeholders. Last year left those
    # cells blank; skipping them keeps the SpreadsheetML uploadable.
    if bg in {PEACH, YELLOW} and v in (0, 0.0, None, ""):
        return False
    if v in (0, 0.0) and ws.title == "Income" and cell.column >= 8:
        return True
    return False


def compute_formulas(ws) -> dict[tuple[int, int], float]:
    values: dict[tuple[int, int], float] = {}
    formulas: list[tuple[int, int, str]] = []
    for r in range(1, (ws.max_row or 1) + 1):
        for c in range(1, (ws.max_column or 1) + 1):
            cell = ws.cell(r, c)
            if cell.data_type == "f" and isinstance(cell.value, str):
                fml = cell.value[1:] if cell.value.startswith("=") else cell.value
                formulas.append((r, c, fml))
            elif isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                values[(r, c)] = float(cell.value)

    def lookup(rr, cc):
        return values.get((rr, cc), 0.0)

    def eval_one(fml: str) -> float | None:
        m = SUM_RE.fullmatch(fml)
        if m:
            c1, r1, c2, r2 = (
                column_index_from_string(m.group(1)),
                int(m.group(2)),
                column_index_from_string(m.group(3)),
                int(m.group(4)),
            )
            total = 0.0
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    total += lookup(rr, cc)
            return total
        m = SUM_ONE_RE.fullmatch(fml)
        if m:
            return lookup(int(m.group(2)), column_index_from_string(m.group(1)))
        m = SUM_MINUS_RE.fullmatch(fml)
        if m:
            fml = f"{m.group(1)}-{m.group(2)}"
        m = MINUS_RE.fullmatch(fml)
        if not m:
            return None
        a = lookup(int(m.group(2)), column_index_from_string(m.group(1)))
        b = lookup(int(m.group(4)), column_index_from_string(m.group(3)))
        return a - b

    # Dependent totals (GCM Grand Total = SUM of Total row) need a second pass.
    for _ in range(4):
        for r, c, fml in formulas:
            got = eval_one(fml)
            if got is not None:
                values[(r, c)] = got
    return values


def style_of(cell, ws, r, c, write_f: bool) -> str:
    bg = None
    if cell.fill and cell.fill.fill_type not in (None, "none"):
        bg = rgb(getattr(cell.fill, "fgColor", None))
    if bg == PEACH:
        return "p"
    if bg == GREEN:
        return "g"
    if bg == YELLOW:
        return "y"
    if cell.font and cell.font.bold:
        return "b"
    if write_f:
        return "n"
    return ""


def merge_map(ws) -> dict[tuple[int, int], tuple[int, int]]:
    out = {}
    for rng in ws.merged_cells.ranges:
        across = rng.max_col - rng.min_col
        down = rng.max_row - rng.min_row
        out[(rng.min_row, rng.min_col)] = (across, down)
    return out


def merged_covered(ws) -> set[tuple[int, int]]:
    cov = set()
    for rng in ws.merged_cells.ranges:
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if r == rng.min_row and c == rng.min_col:
                    continue
                cov.add((r, c))
    return cov


def esc(s: str) -> str:
    return html.escape(s, quote=True)


def num_str(v: float) -> str:
    if abs(v - round(v)) < 1e-9:
        return str(int(round(v)))
    return f"{v:.12g}"


def sheets_formula(fml: str) -> str | None:
    f = fml[1:] if fml.startswith("=") else fml
    if "[" in f:
        return None
    m = SUM_MINUS_RE.fullmatch(f)
    if m:
        return f"={m.group(1)}-{m.group(2)}"
    # SUM(B24) becomes #REF!/0 in Sheets R1C1. Use the cell itself.
    m = SUM_ONE_RE.fullmatch(f)
    if m:
        return f"={m.group(1)}{m.group(2)}"
    return f"={f}"


def max_col_for(name: str, ws) -> int:
    if name == "Income":
        return 9
    if name == "524 Ferdinand Ave, Unit 1":
        return 28
    if name in PROPERTY_TABS:
        return 41
    if name in {"GCM", "Hindman W2", "Megan T4"}:
        return 5
    if name == "Investments":
        return 14
    if name == "EPGC LLC":
        return 14
    if name == "Art Sales and Purchases":
        return 11
    if name == "Refrence Library":
        return 5
    if name == "2025 Data sources":
        return 5
    return min(ws.max_column or 1, 41)


def cell_text(val, name: str) -> str:
    if isinstance(val, datetime):
        return val.date().isoformat()
    if isinstance(val, date):
        return val.isoformat()
    text = str(val) if val is not None else ""
    limit = 80 if name == "2025 Data sources" else 36
    if len(text) > limit:
        text = text[: limit - 3] + "..."
    return text


HEADER = [
    '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>',
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">',
    "<Styles>",
    '<Style ss:ID="Default" ss:Name="Normal"><Font ss:FontName="Century Gothic" ss:Size="10"/></Style>',
    '<Style ss:ID="n"><Font ss:FontName="Century Gothic" ss:Size="10"/><NumberFormat ss:Format="$#,##0.00"/></Style>',
    '<Style ss:ID="b"><Font ss:FontName="Century Gothic" ss:Size="11" ss:Bold="1"/></Style>',
    f'<Style ss:ID="p"><Font ss:FontName="Century Gothic" ss:Size="11" ss:Bold="1"/><Interior ss:Color="#{PEACH}" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>',
    f'<Style ss:ID="g"><Font ss:FontName="Century Gothic" ss:Size="10"/><Interior ss:Color="#{GREEN}" ss:Pattern="Solid"/><NumberFormat ss:Format="$#,##0.00"/></Style>',
    f'<Style ss:ID="y"><Font ss:FontName="Century Gothic" ss:Size="10"/><Interior ss:Color="#{YELLOW}" ss:Pattern="Solid"/><NumberFormat ss:Format="$#,##0.00"/></Style>',
    "</Styles>",
]


def cell_xml(dest_c: int, last: int, sid: str, merge: str, body: str) -> str:
    idx = f' ss:Index="{dest_c}"' if dest_c != last + 1 else ""
    st = f' ss:StyleID="{sid}"' if sid else ""
    return f"<Cell{idx}{st}{merge}>{body}</Cell>" if body else f"<Cell{idx}{st}{merge}/>"


def emit_cells(pieces: dict[int, str], skip: set[int] | None = None) -> str:
    """Emit dest columns in order, filling holes so Google does not shift cells.

    skip = columns covered by MergeAcross/MergeDown — do not insert empty
    cells inside a merge, or the following year-block shifts right.
    """
    if not pieces:
        return ""
    skip = skip or set()
    last = 0
    bits: list[str] = []
    lo, hi = min(pieces), max(pieces)
    for dest_c in range(lo, hi + 1):
        if dest_c in skip and dest_c not in pieces:
            continue
        if dest_c in pieces:
            xml = pieces[dest_c]
            if dest_c != last + 1 and 'ss:Index="' not in xml:
                xml = xml.replace("<Cell", f'<Cell ss:Index="{dest_c}"', 1)
            bits.append(xml)
            last = dest_c
        else:
            if dest_c == last + 1:
                bits.append("<Cell/>")
            else:
                bits.append(f'<Cell ss:Index="{dest_c}"/>')
            last = dest_c
    return "".join(bits)


def worksheet_xml(ws) -> tuple[str, int, int]:
    name = ws.title
    cached = compute_formulas(ws)
    merges = merge_map(ws)
    covered = merged_covered(ws)
    bits: list[str] = [f'<Worksheet ss:Name="{esc(name)}"><Table>']
    n_cells = 0
    n_f = 0
    max_r = ws.max_row or 1
    max_c = max_col_for(name, ws)
    prop = name in PROPERTY_TABS
    col_map: dict[int, int] = {}
    if prop:
        # Identity map — 2023/2024/2025 stay in last year's columns.
        col_map = {c: c for c in range(1, max_c + 1)}
    rows: dict[int, str] = {}
    for r in range(1, max_r + 1):
        if not keep_row(name, r, ws):
            continue
        pieces: dict[int, str] = {}
        for c in range(1, max_c + 1):
            dest_c = property_dest_col(name, c) if prop else c
            if dest_c is None:
                continue
            if (r, c) in covered:
                continue
            cell = ws.cell(r, c)
            is_f = cell.data_type == "f" and isinstance(cell.value, str)
            fml = ""
            raw_f = ""
            if is_f:
                raw_f = sheets_formula(cell.value) or ""
                fml = rewrite_formula_cols(raw_f, col_map) if prop and raw_f else raw_f
                if fml:
                    fml = formula_to_r1c1(fml, r, dest_c)
            write_f = bool(is_f and fml and keep_formula(ws, r, c, raw_f[1:] if raw_f else ""))
            if write_f and prop and cached.get((r, c), 0.0) == 0.0:
                year_start = year_start_for_total_col(c)
                if year_start and not row_has_month_values(ws, r, year_start):
                    write_f = False
            val = cell.value
            if name == "Income" and r == 1 and isinstance(val, (int, float)) and not isinstance(val, bool):
                val = str(int(val))
            if is_f and not write_f:
                val = cached.get((r, c), 0.0)
                if not keep_value(val, cell, ws) and (r, c) not in merges:
                    continue
            elif not is_f:
                if not keep_value(val, cell, ws) and (r, c) not in merges:
                    continue
            n_cells += 1
            sid = style_of(cell, ws, r, c, write_f)
            merge = ""
            if (r, c) in merges:
                across, down = merges[(r, c)]
                if across:
                    merge += f' ss:MergeAcross="{across}"'
                if down:
                    merge += f' ss:MergeDown="{down}"'
            st = f' ss:StyleID="{sid}"' if sid else ""
            if write_f:
                n_f += 1
                cached_v = cached.get((r, c), 0.0)
                pieces[dest_c] = (
                    f'<Cell{st}{merge} ss:Formula="{esc(fml)}">'
                    f'<Data ss:Type="Number">{num_str(cached_v)}</Data></Cell>'
                )
            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                # Skip literal zeros — emit_cells fills the hole with <Cell/>
                # so JAN–DEC stay aligned without Number-0 bloat.
                if val == 0 and (r, c) not in merges and dest_c not in YEAR_TOTAL_COLS:
                    continue
                pieces[dest_c] = (
                    f"<Cell{st}{merge}><Data ss:Type=\"Number\">{num_str(float(val))}</Data></Cell>"
                )
            else:
                text = cell_text(val, name)
                if r == 2 and dest_c in YEAR_TOTAL_COLS and prop and "YR TOTAL" in str(text):
                    if dest_c == 28 and name == "524 Ferdinand Ave, Unit 1":
                        text = "2025 YR TOTAL"
                    elif dest_c == 15:
                        text = "2023 YR TOTAL"
                    elif dest_c == 28:
                        text = "2024 YR TOTAL"
                    elif dest_c == 41:
                        text = "2025 YR TOTAL"
                if not text.strip() and (r, c) not in merges:
                    continue
                pieces[dest_c] = (
                    f"<Cell{st}{merge}><Data ss:Type=\"String\">{esc(text)}</Data></Cell>"
                )
        if not pieces:
            continue
        row_skip: set[int] = set()
        for (mr, mc), (across, down) in merges.items():
            if mr != r:
                continue
            dest_m = property_dest_col(name, mc) if prop else mc
            if dest_m is None:
                continue
            for dc in range(dest_m + 1, dest_m + across + 1):
                row_skip.add(dc)
        rows[dest_row_index(name, r)] = emit_cells(pieces, row_skip)
    last_r = max(rows) if rows else 0
    # Google's converter ignores ss:Index when rows are omitted, which shifts
    # R1C1 year-totals onto the wrong line. Emit every row through last_r.
    fill_empty = prop or name in {"GCM", "EPGC LLC", "Art Sales and Purchases"}
    for r in range(1, last_r + 1):
        if r in rows:
            bits.append(f'<Row ss:Index="{r}">' + rows[r] + "</Row>")
        elif fill_empty:
            bits.append(f'<Row ss:Index="{r}"/>')
    # Freeze A–B like a working P&L; scroll property tabs to the 2025 block.
    opts = ['<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">']
    if name == "EPGC LLC":
        opts.append("<TopRowVisible>50</TopRowVisible>")
    elif name == "Art Sales and Purchases":
        opts.append("<TopRowVisible>48</TopRowVisible>")
    elif name == "GCM":
        opts.append("<TopRowVisible>21</TopRowVisible>")
    elif prop:
        # Last year had no freeze panes. Open scrolled to the 2025 year block.
        left = 16 if name == "524 Ferdinand Ave, Unit 1" else 29
        opts.append(f"<LeftColumnVisible>{left}</LeftColumnVisible>")
    opts.append("</WorksheetOptions>")
    bits.append("</Table>" + "".join(opts) + "</Worksheet>")
    return "".join(bits), n_cells, n_f


def wrap(sheet_xmls: list[str]) -> str:
    return "".join(HEADER) + "".join(sheet_xmls) + "</Workbook>"


def emit(src: Path, dest: Path) -> dict[str, str]:
    wb = load_workbook(src, data_only=False)
    sheets: dict[str, str] = {}
    stats: list[tuple[str, int, int, int]] = []
    for name in PACKET_TABS:
        xml, n_cells, n_f = worksheet_xml(wb[name])
        sheets[name] = xml
        stats.append((name, len(xml), n_cells, n_f))
    full = wrap([sheets[n] for n in PACKET_TABS])
    dest.write_text(full, encoding="utf-8")
    print(f"SSML {dest} {dest.stat().st_size} bytes")
    for name, nbytes, n_cells, n_f in stats:
        print(f"  {name:40} {nbytes:7} bytes cells={n_cells:4} formulas={n_f:3}")
    a1_hits = re.findall(r'ss:Formula="=SUM\([A-Z]+\d+', full)
    if a1_hits:
        print(f"WARN A1-looking formulas remain: {a1_hits[:8]}")
    else:
        print("R1C1 OK — no A1 SUM(A1) formulas")
    r1c1_n = len(re.findall(r"ss:Formula=\"=SUM\(R", full))
    print(f"R1C1 SUM formulas: {r1c1_n}")

    groups = {
        "all": PACKET_TABS,
        "income": ["Income"],
        "216": ["216 N. Oak Park Ave"],
        "u2": ["524 Ferdinand Ave, Unit 2"],
        "u1": ["524 Ferdinand Ave, Unit 1"],
        "gcm": ["GCM"],
        "w2": ["Hindman W2"],
        "t4": ["Megan T4"],
        "inv": ["Investments"],
        "epgc": ["EPGC LLC"],
        "art": ["Art Sales and Purchases"],
        "ref": ["Refrence Library"],
        "ds": ["2025 Data sources"],
        "book": LAST_YEAR_TABS,
        "inc216": ["Income", "216 N. Oak Park Ave"],
        "core": [
            "Income",
            "216 N. Oak Park Ave",
            "GCM",
        ],
        "props": [
            "Income",
            "524 Ferdinand Ave, Unit 2",
            "216 N. Oak Park Ave",
            "524 Ferdinand Ave, Unit 1",
        ],
        "main": [
            "Income",
            "524 Ferdinand Ave, Unit 2",
            "216 N. Oak Park Ave",
            "524 Ferdinand Ave, Unit 1",
            "GCM",
            "Hindman W2",
            "Megan T4",
            "Investments",
        ],
        "biz": [
            "EPGC LLC",
            "Art Sales and Purchases",
            "GCM",
            "Refrence Library",
            "Hindman W2",
            "Megan T4",
            "Investments",
            "2025 Data sources",
        ],
    }
    out_files = {"all": str(dest)}
    for key, names in groups.items():
        if key == "all":
            continue
        xml = wrap([sheets[n] for n in names])
        path = dest.with_name(f"Personal_Income_2025_ssml_{key}.xml")
        path.write_text(xml, encoding="utf-8")
        out_files[key] = str(path)
        print(f"  GROUP {key:8} {len(xml):7} bytes -> {path.name}")
    return out_files


if __name__ == "__main__":
    emit(SRC, OUT)
