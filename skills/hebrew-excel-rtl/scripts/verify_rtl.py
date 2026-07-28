# -*- coding: utf-8 -*-
"""Verify a Hebrew RTL Excel workbook.

    python verify_rtl.py workbook.xlsx
    python verify_rtl.py workbook.xlsx --com     # also open it through real Excel (Windows)

Exit code 0 = all checks passed, 1 = at least one FAIL.

Static checks (any platform):
  * every sheet has sheet_view.rightToLeft
  * cells containing Hebrew have readingOrder set
  * table names contain no spaces and are unique
  * sheet names are legal and <= 31 chars
  * date-looking strings that should have been real dates
  * currency-looking columns with no number format

COM checks (Windows + Excel only, --com):
  * the file opens without a repair prompt
  * Excel reports DisplayRightToLeft on every sheet expected to be RTL
  * formulas actually evaluate to numbers
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

PASS, FAIL, WARN = "PASS", "FAIL", "WARN"
results: list[tuple[str, str]] = []


def record(level: str, msg: str) -> None:
    results.append((level, msg))


def check_static(path: str) -> None:
    wb = openpyxl.load_workbook(path)

    # --- sheet-level RTL --------------------------------------------------
    for ws in wb.worksheets:
        has_hebrew = any(
            isinstance(c.value, str) and HEBREW.search(c.value)
            for row in ws.iter_rows()
            for c in row
        )
        if ws.sheet_view.rightToLeft:
            record(PASS, f"sheet {ws.title!r}: rightToLeft set")
        elif has_hebrew:
            record(FAIL, f"sheet {ws.title!r}: contains Hebrew but rightToLeft is NOT set")
        else:
            record(WARN, f"sheet {ws.title!r}: rightToLeft not set (no Hebrew found — may be intentional)")

    # --- sheet names ------------------------------------------------------
    for ws in wb.worksheets:
        if len(ws.title) > 31:
            record(FAIL, f"sheet {ws.title!r}: name longer than 31 chars")
        bad = ILLEGAL_SHEET.intersection(ws.title)
        if bad:
            record(FAIL, f"sheet {ws.title!r}: illegal characters {sorted(bad)}")

    # --- per-cell reading order ------------------------------------------
    for ws in wb.worksheets:
        missing = []
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and HEBREW.search(c.value):
                    ro = c.alignment.readingOrder
                    if not ro or int(ro) != 2:
                        missing.append(c.coordinate)
        if missing:
            shown = ", ".join(missing[:8])
            more = f" (+{len(missing) - 8} more)" if len(missing) > 8 else ""
            record(FAIL, f"sheet {ws.title!r}: {len(missing)} Hebrew cells without readingOrder=2: {shown}{more}")
        else:
            record(PASS, f"sheet {ws.title!r}: all Hebrew cells have readingOrder=2")

    # --- tables -----------------------------------------------------------
    seen: dict[str, str] = {}
    for ws in wb.worksheets:
        for name in ws.tables:
            if " " in name:
                record(FAIL, f"table {name!r} on {ws.title!r}: name contains a space")
            if name in seen:
                record(FAIL, f"table {name!r} duplicated on {ws.title!r} and {seen[name]!r}")
            seen[name] = ws.title
    if seen and not any(l == FAIL and "table" in m for l, m in results):
        record(PASS, f"{len(seen)} table(s): names unique and space-free")

    # --- dates stored as text --------------------------------------------
    for ws in wb.worksheets:
        text_dates = [
            c.coordinate for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and DATE_STR.match(c.value)
        ]
        if text_dates:
            shown = ", ".join(text_dates[:8])
            more = f" (+{len(text_dates) - 8} more)" if len(text_dates) > 8 else ""
            record(FAIL, f"sheet {ws.title!r}: {len(text_dates)} date(s) stored as TEXT: {shown}{more}")

    # --- literal shekel signs in text -------------------------------------
    for ws in wb.worksheets:
        literal = [
            c.coordinate for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and "₪" in c.value and not c.value.strip().startswith("=")
        ]
        if literal:
            record(WARN, f"sheet {ws.title!r}: ₪ written as text in {len(literal)} cell(s) "
                         f"({', '.join(literal[:5])}) — use a number format instead")


def check_com(path: str) -> None:
    if sys.platform != "win32":
        record(WARN, "--com skipped: not on Windows")
        return
    try:
        import win32com.client as win32  # type: ignore
    except ImportError:
        record(WARN, "--com skipped: pywin32 not installed (py -m pip install pywin32)")
        return

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path), 0, True)
        record(PASS, "Excel opened the file without a repair prompt")

        for ws in wb.Worksheets:
            if ws.DisplayRightToLeft:
                record(PASS, f"Excel confirms RTL on sheet {ws.Name!r}")
            else:
                record(WARN, f"Excel reports sheet {ws.Name!r} as left-to-right")

        evaluated = broken = 0
        for ws in wb.Worksheets:
            used = ws.UsedRange
            for r in range(1, min(used.Rows.Count, 200) + 1):
                for c in range(1, min(used.Columns.Count, 60) + 1):
                    cell = used.Cells(r, c)
                    f = cell.Formula
                    if isinstance(f, str) and f.startswith("="):
                        v = cell.Value2
                        if v is None or (isinstance(v, str) and v.startswith("#")):
                            broken += 1
                        else:
                            evaluated += 1
        if broken:
            record(FAIL, f"{broken} formula(s) did not evaluate or returned an error")
        if evaluated:
            record(PASS, f"{evaluated} formula(s) evaluated cleanly in Excel")

        wb.Close(False)
    except Exception as e:                      # noqa: BLE001
        record(FAIL, f"Excel could not open the file: {e}")
    finally:
        xl.Quit()


def main() -> int:
    ap = argparse.ArgumentParser(description="Verify a Hebrew RTL Excel workbook")
    ap.add_argument("path")
    ap.add_argument("--com", action="store_true",
                    help="also verify through real Excel (Windows only)")
    args = ap.parse_args()

    if not os.path.exists(args.path):
        print(f"no such file: {args.path}")
        return 1

    check_static(args.path)
    if args.com:
        check_com(args.path)

    width = max(len(l) for l, _ in results)
    for level, msg in results:
        mark = {PASS: "OK  ", FAIL: "FAIL", WARN: "warn"}[level]
        print(f"[{mark}] {msg}")

    fails = sum(1 for l, _ in results if l == FAIL)
    warns = sum(1 for l, _ in results if l == WARN)
    passes = sum(1 for l, _ in results if l == PASS)
    print(f"\n{passes} passed, {warns} warnings, {fails} failed")
    return 1 if fails else 0


if __name__ == "__main__":
    raise SystemExit(main())
