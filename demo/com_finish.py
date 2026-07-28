# -*- coding: utf-8 -*-
"""Finish the demo workbook through real Excel (Windows + Excel required).

    py com_finish.py

openpyxl cannot do two things this workbook needs:

  1. **Real PivotTables.** openpyxl can only preserve existing pivots, not create
     them. The pivot sheet here is built by Excel itself over the inventory table.
  2. **Cached formula values.** openpyxl writes formulas with no results, so pandas,
     Firecrawl parse, and every other cached-value reader sees blanks until Excel
     has recalculated and saved once. This script does that recalc-and-save.

ORDER MATTERS: this runs AFTER build_demo.py, and nothing may touch the file with
openpyxl afterwards — an openpyxl load/save round-trip silently deletes the charts
and the PivotTables. Verify with verify_rtl.py (read-only) as the final step.
"""
from __future__ import annotations

import os
import sys

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(HERE, "דוח-כספי-לדוגמה.xlsx")

PIVOT_SHEET = "ניתוח Pivot"
AFTER_SHEET = "חיפוש מהיר"

# Excel constants (hardcoded so this works without makepy type generation)
xlDatabase, xlRowField, xlColumnField, xlSum = 1, 1, 2, -4157
xlRTL = -5004

NAVY = 0x5C3A1F      # RGB 1F3A5C in COM's BGR ordering
MUTED = 0x80726B     # RGB 6B7280


def _title(ws, cell: str, text: str, size: int, bold: bool, colour: int) -> None:
    c = ws.Range(cell)
    c.Value = text
    c.Font.Name = "Arial"
    c.Font.Size = size
    c.Font.Bold = bold
    c.Font.Color = colour


def main() -> int:
    if sys.platform != "win32":
        print("Windows + Excel required"); return 1
    if not os.path.exists(PATH):
        print(f"run build_demo.py first — missing {PATH}"); return 1

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(PATH))

        # 1. Recalculate FIRST — the pivot cache snapshots the table's values, and
        #    straight after build_demo.py every formula cell is still empty.
        xl.CalculateFullRebuild()

        # 2. Fresh pivot sheet (idempotent: drop a leftover one from a previous run)
        for ws in wb.Worksheets:
            if ws.Name == PIVOT_SHEET:
                ws.Delete()
                break
        ws = wb.Worksheets.Add(After=wb.Worksheets(AFTER_SHEET))
        ws.Name = PIVOT_SHEET

        _title(ws, "A1", "ניתוח Pivot — מי מוכר מה, ואיפה הכסף תקוע", 18, True, NAVY)
        _title(ws, "A2", "טבלאות מסתובבות אמיתיות. אפשר לגרור שדות ולפרוס אחרת — הנתונים מהמלאי.",
               10, False, MUTED)

        # 3. Pivot 1 — sales: category rows × store columns
        cache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData="טבלת_מלאי")
        _title(ws, "A4", "מכירות חודשיות (₪) — קטגוריה × חנות", 12, True, NAVY)
        pt = cache.CreatePivotTable(TableDestination=ws.Range("A5"),
                                    TableName="Pivot_מכירות")
        pt.PivotFields("קטגוריה").Orientation = xlRowField
        pt.PivotFields("חנות").Orientation = xlColumnField
        df = pt.AddDataField(pt.PivotFields("מכירות בחודש (₪)"), "סה״כ מכירות (₪)", xlSum)
        df.NumberFormat = "#,##0"
        pt.TableStyle2 = "PivotStyleLight16"

        # 4. Pivot 2 — where money is stuck: status rows × store columns
        first_end = pt.TableRange2.Row + pt.TableRange2.Rows.Count
        r2 = first_end + 3
        _title(ws, f"A{r2}", "שווי מלאי (₪) לפי סטטוס — איפה הכסף תקוע", 12, True, NAVY)
        pt2 = cache.CreatePivotTable(TableDestination=ws.Range(f"A{r2 + 1}"),
                                     TableName="Pivot_מלאי")
        pt2.PivotFields("סטטוס").Orientation = xlRowField
        pt2.PivotFields("חנות").Orientation = xlColumnField
        df2 = pt2.AddDataField(pt2.PivotFields("שווי מלאי (₪)"), "סה״כ שווי (₪)", xlSum)
        df2.NumberFormat = "#,##0"
        pt2.TableStyle2 = "PivotStyleLight16"

        note_row = pt2.TableRange2.Row + pt2.TableRange2.Rows.Count + 2
        _title(ws, f"A{note_row}",
               "שורת 'עודף' למטה = כסף ששוכב על המדף. עמודה גבוהה בפיבוט העליון = החנות שמוכרת.",
               9, False, MUTED)

        # 5. RTL + readability on the sheet Excel just built — Excel does not set
        #    these on its own, which is exactly the point this kit keeps making.
        ws.DisplayRightToLeft = True
        ws.UsedRange.ReadingOrder = xlRTL
        ws.UsedRange.Columns.AutoFit()
        for col in ws.UsedRange.Columns:
            if col.ColumnWidth < 11:
                col.ColumnWidth = 11
        ws.Activate()
        xl.ActiveWindow.DisplayGridlines = False
        # Keep the page title visible while scrolling the two stacked pivots.
        xl.ActiveWindow.SplitRow = 2
        xl.ActiveWindow.FreezePanes = True

        # 6. Recalculate again (the pivots are new) and save cached values.
        xl.CalculateFullRebuild()
        wb.Worksheets("לוח בקרה").Activate()   # the file should open on the dashboard
        wb.Save()
        wb.Close(False)
        print(f"pivots + recalc done: {PATH}")
        return 0
    finally:
        xl.Quit()


if __name__ == "__main__":
    raise SystemExit(main())
