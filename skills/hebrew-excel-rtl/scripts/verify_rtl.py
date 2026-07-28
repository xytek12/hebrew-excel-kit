# -*- coding: utf-8 -*-
"""Self-check a Hebrew RTL Excel workbook, sheet by sheet.

    python verify_rtl.py workbook.xlsx
    python verify_rtl.py workbook.xlsx --com        # also open it through real Excel
    python verify_rtl.py workbook.xlsx --strict     # warnings count as failures

Exit code 0 = clean, 1 = at least one FAIL.

Every sheet gets its own report block. A workbook is only correct if EVERY sheet is
correct — the classic Hebrew-Excel defect is a file where tab 1 is perfect and tab 4 is
backwards, and a single whole-workbook verdict hides exactly that.

Two families of check:

  RTL correctness — does it render right in Excel at all
  Readability     — can a manager who did not build it actually read it
"""
from __future__ import annotations

import argparse
import os
import re
import sys

import openpyxl

HEBREW = re.compile(r"[֐-׿]")
DATE_STR = re.compile(r"^\s*\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}\s*$")
ILLEGAL_SHEET = set(r':\/?*[]')
DEFAULT_WIDTH = 13.0

OK, FAIL, WARN, SKIP = "OK", "FAIL", "WARN", "SKIP"
MARK = {OK: "OK  ", FAIL: "FAIL", WARN: "warn", SKIP: " -  "}


class Report:
    def __init__(self) -> None:
        self.sheets: dict[str, list[tuple[str, str, str]]] = {}
        self.book: list[tuple[str, str, str]] = []

    def sheet(self, name: str, check: str, level: str, detail: str = "") -> None:
        self.sheets.setdefault(name, []).append((check, level, detail))

    def workbook(self, check: str, level: str, detail: str = "") -> None:
        self.book.append((check, level, detail))

    def counts(self) -> tuple[int, int, int]:
        rows = self.book + [r for v in self.sheets.values() for r in v]
        return (sum(1 for _, l, _ in rows if l == OK),
                sum(1 for _, l, _ in rows if l == WARN),
                sum(1 for _, l, _ in rows if l == FAIL))


rep = Report()


# ---------------------------------------------------------------------------
# per-sheet checks
# ---------------------------------------------------------------------------
def check_sheet(ws) -> None:
    name = ws.title
    cells = [c for row in ws.iter_rows() for c in row if c.value is not None]
    heb_cells = [c for c in cells
                 if isinstance(c.value, str) and HEBREW.search(c.value)]
    has_hebrew = bool(heb_cells)

    # --- RTL correctness ---------------------------------------------------
    if ws.sheet_view.rightToLeft:
        rep.sheet(name, "RTL (rightToLeft)", OK)
    elif has_hebrew:
        rep.sheet(name, "RTL (rightToLeft)", FAIL,
                  "sheet holds Hebrew but is left-to-right")
    else:
        rep.sheet(name, "RTL (rightToLeft)", SKIP, "no Hebrew on this sheet")

    if heb_cells:
        bad = [c.coordinate for c in heb_cells
               if not c.alignment.readingOrder or int(c.alignment.readingOrder) != 2]
        if bad:
            shown = ", ".join(bad[:6])
            more = f" +{len(bad) - 6}" if len(bad) > 6 else ""
            rep.sheet(name, "Reading order", FAIL,
                      f"{len(bad)}/{len(heb_cells)} Hebrew cells missing readingOrder=2: "
                      f"{shown}{more}")
        else:
            rep.sheet(name, "Reading order", OK, f"{len(heb_cells)}/{len(heb_cells)} cells")
    else:
        rep.sheet(name, "Reading order", SKIP)

    if len(name) > 31:
        rep.sheet(name, "Sheet name", FAIL, "longer than 31 characters")
    elif ILLEGAL_SHEET.intersection(name):
        rep.sheet(name, "Sheet name", FAIL,
                  f"illegal characters {sorted(ILLEGAL_SHEET.intersection(name))}")
    else:
        rep.sheet(name, "Sheet name", OK)

    # --- data typing -------------------------------------------------------
    text_dates = [c.coordinate for c in cells
                  if isinstance(c.value, str) and DATE_STR.match(c.value)]
    if text_dates:
        rep.sheet(name, "Dates typed", FAIL,
                  f"{len(text_dates)} date(s) stored as text: {', '.join(text_dates[:6])}")
    else:
        rep.sheet(name, "Dates typed", OK)

    money_text = [c.coordinate for c in cells
                  if isinstance(c.value, str) and "₪" in c.value
                  and not c.value.strip().startswith("=")
                  and any(ch.isdigit() for ch in c.value)]
    if money_text:
        rep.sheet(name, "Currency typed", WARN,
                  f"{len(money_text)} amount(s) written as text: {', '.join(money_text[:5])}")
    else:
        rep.sheet(name, "Currency typed", OK)

    # --- readability -------------------------------------------------------
    merged = list(ws.merged_cells.ranges)
    if merged:
        rep.sheet(name, "Merged cells", WARN,
                  f"{len(merged)} merged range(s) — breaks sort, filter and pivots; "
                  f"use centerContinuous")
    else:
        rep.sheet(name, "Merged cells", OK, "none")

    # Excel groups equal-width columns into one ranged <col min max> record on save,
    # which openpyxl parses as a single ColumnDimension. Count the span, not the keys,
    # or a workbook that has been saved by real Excel under-reports its sized columns.
    sized_cols: set[int] = set()
    for key, d in ws.column_dimensions.items():
        if d.width:
            lo = d.min if d.min else openpyxl.utils.column_index_from_string(key)
            hi = d.max if d.max else lo
            sized_cols.update(range(lo, hi + 1))
    sized = len(sized_cols)
    used_cols = max((c.column for c in cells), default=0)
    if used_cols == 0:
        rep.sheet(name, "Column widths", SKIP, "empty sheet")
    elif sized == 0:
        rep.sheet(name, "Column widths", FAIL,
                  f"no width set on any of {used_cols} used columns — Hebrew will clip")
    elif sized < used_cols:
        rep.sheet(name, "Column widths", WARN,
                  f"{sized}/{used_cols} columns sized; the rest use the default")
    else:
        rep.sheet(name, "Column widths", OK, f"{sized} columns")

    fills = {c.fill.fgColor.rgb for c in cells
             if c.fill and c.fill.fill_type == "solid"
             and c.fill.fgColor.rgb not in (None, "00000000")}
    if len(fills) > 6:
        rep.sheet(name, "Fill palette", WARN,
                  f"{len(fills)} distinct fill colours — hard on the eye, aim for <= 5")
    else:
        rep.sheet(name, "Fill palette", OK, f"{len(fills)} distinct")

    if ws.freeze_panes:
        rep.sheet(name, "Frozen header", OK, str(ws.freeze_panes))
    elif len(cells) > 120:
        rep.sheet(name, "Frozen header", WARN,
                  "long sheet with no frozen header — headings scroll away")
    else:
        rep.sheet(name, "Frozen header", SKIP, "short sheet")

    if ws.sheet_view.showGridLines is False:
        rep.sheet(name, "Gridlines", OK, "hidden")
    else:
        rep.sheet(name, "Gridlines", WARN,
                  "default gridlines still on — they fight with your own borders")


def check_workbook(path: str) -> None:
    wb = openpyxl.load_workbook(path)
    for ws in wb.worksheets:
        check_sheet(ws)

    seen: dict[str, str] = {}
    problems = []
    for ws in wb.worksheets:
        for tname in ws.tables:
            if " " in tname:
                problems.append(f"{tname!r} on {ws.title!r} contains a space")
            if tname in seen:
                problems.append(f"{tname!r} duplicated on {ws.title!r} and {seen[tname]!r}")
            seen[tname] = ws.title
    if problems:
        rep.workbook("Table names", FAIL, "; ".join(problems))
    elif seen:
        rep.workbook("Table names", OK, f"{len(seen)} table(s), unique and space-free")


# ---------------------------------------------------------------------------
# real Excel
# ---------------------------------------------------------------------------
def check_com(path: str) -> None:
    if sys.platform != "win32":
        rep.workbook("Real Excel", SKIP, "not on Windows")
        return
    try:
        import win32com.client as win32  # type: ignore
    except ImportError:
        rep.workbook("Real Excel", SKIP, "pywin32 not installed")
        return

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path), 0, True)
        rep.workbook("Opens in Excel", OK, "no repair prompt")

        total_ok = total_bad = 0
        for ws in wb.Worksheets:
            name = ws.Name
            rep.sheet(name, "Excel says RTL", OK if ws.DisplayRightToLeft else WARN,
                      "" if ws.DisplayRightToLeft else "Excel renders this sheet LTR")

            good = bad = 0
            used = ws.UsedRange
            for r in range(1, min(used.Rows.Count, 300) + 1):
                for c in range(1, min(used.Columns.Count, 60) + 1):
                    cell = used.Cells(r, c)
                    f = cell.Formula
                    if isinstance(f, str) and f.startswith("="):
                        v = cell.Value2
                        if v is None or (isinstance(v, str) and v.startswith("#")):
                            bad += 1
                        else:
                            good += 1
            total_ok += good
            total_bad += bad
            if bad:
                rep.sheet(name, "Formulas evaluate", FAIL, f"{bad} broken, {good} fine")
            elif good:
                rep.sheet(name, "Formulas evaluate", OK, f"{good} formulas")
            else:
                rep.sheet(name, "Formulas evaluate", SKIP, "no formulas")

        rep.workbook("Formulas (whole book)", FAIL if total_bad else OK,
                     f"{total_ok} evaluated, {total_bad} broken")
        wb.Close(False)
    except Exception as e:                      # noqa: BLE001
        rep.workbook("Opens in Excel", FAIL, str(e))
    finally:
        xl.Quit()


# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description="Self-check a Hebrew RTL Excel workbook")
    ap.add_argument("path")
    ap.add_argument("--com", action="store_true",
                    help="also verify through real Excel (Windows only)")
    ap.add_argument("--strict", action="store_true",
                    help="treat warnings as failures")
    args = ap.parse_args()

    if not os.path.exists(args.path):
        print(f"no such file: {args.path}")
        return 1

    check_workbook(args.path)
    if args.com:
        check_com(args.path)

    print(f"\n{'=' * 72}\n  {os.path.basename(args.path)}\n{'=' * 72}")
    for name, rows in rep.sheets.items():
        worst = FAIL if any(l == FAIL for _, l, _ in rows) else (
            WARN if any(l == WARN for _, l, _ in rows) else OK)
        flag = {OK: "clean", WARN: "has warnings", FAIL: "HAS FAILURES"}[worst]
        print(f"\n  גיליון: {name}   [{flag}]")
        print(f"  {'-' * 68}")
        for check, level, detail in rows:
            print(f"    [{MARK[level]}] {check:<22} {detail}")

    if rep.book:
        print(f"\n  Workbook-level\n  {'-' * 68}")
        for check, level, detail in rep.book:
            print(f"    [{MARK[level]}] {check:<22} {detail}")

    passed, warns, fails = rep.counts()
    print(f"\n{'=' * 72}")
    print(f"  {len(rep.sheets)} sheets — {passed} passed, {warns} warnings, {fails} failed")
    print(f"{'=' * 72}")
    return 1 if (fails or (args.strict and warns)) else 0


if __name__ == "__main__":
    raise SystemExit(main())
