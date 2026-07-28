# -*- coding: utf-8 -*-
"""Build the Hebrew RTL demo workbook.

Fictional company: אופנת גלים בע״מ — an Israeli clothing retailer.
All figures are invented for demonstration.

    py build_demo.py
"""
from __future__ import annotations

import datetime
import os
import sys

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "skills", "hebrew-excel-rtl", "scripts"))

from rtl_helpers import (  # noqa: E402
    BORDER, FMT, FONT_NAME, GREY, LIGHT, NAVY, RTL, RTL_CENTER,
    add_table, autofit, banner, freeze_header, rtl_alignment,
    rtl_workbook, style_header, suppress_text_number_warning, write_row,
)

OUT = os.path.join(HERE, "דוח-כספי-לדוגמה.xlsx")

# Financial-model colour convention: blue = hardcoded input, black = formula,
# green = link to another sheet. Same convention the davila7 spreadsheet skill uses.
INPUT_FONT = Font(name=FONT_NAME, size=11, color="0000CC")
LINK_FONT = Font(name=FONT_NAME, size=11, color="008000")

MONTHS = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני",
          "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"]
ACTUAL_MONTHS = 6          # ינואר–יוני actual, יולי–דצמבר forecast

wb = openpyxl.Workbook()


# ---------------------------------------------------------------------------
# 1. הנחות ופרמטרים — assumptions. Every hardcoded number in the model lives here.
# ---------------------------------------------------------------------------
ws = wb.active
ws.title = "הנחות ופרמטרים"
banner(ws, 1, "אופנת גלים בע״מ — הנחות ופרמטרים", 4)
ws["A2"] = "כל המספרים בגיליון הזה הם קלט. שאר הגיליונות מחושבים מהם בנוסחאות."
ws["A2"].font = Font(name=FONT_NAME, size=10, italic=True, color="666666")
ws["A2"].alignment = RTL

style_header(ws, 4, ["פרמטר", "ערך", "יחידה", "הערה"])
assumptions = [
    ("שיעור מע״מ",              0.18,  "אחוז",  "18% מ־1 בינואר 2025"),
    ("שיעור עלות המכר",         0.42,  "אחוז",  "אחוז מההכנסה ברוטו"),
    ("צמיחה חודשית בתחזית",     0.035, "אחוז",  "יולי–דצמבר"),
    ("ימי אשראי לקוחות",        45,    "ימים",  "ממוצע גבייה"),
    ("ימי אשראי ספקים",         60,    "ימים",  "ממוצע תשלום"),
    ("מלאי בטחון",              0.20,  "אחוז",  "מהצריכה החודשית"),
    ("זמן אספקה מספק",          21,    "ימים",  "ממוצע משוקלל"),
    ("שיעור החזרות",            0.065, "אחוז",  "מהמכירות באתר"),
]
fmt_map = {"אחוז": "pct", "ימים": "int"}
for i, (name, val, unit, note) in enumerate(assumptions, start=5):
    write_row(ws, i, [name, val, unit, note], [None, fmt_map[unit], None, None])
    ws.cell(row=i, column=2).font = INPUT_FONT
add_table(ws, "הנחות", f"A4:D{4 + len(assumptions)}")
freeze_header(ws, 5)
autofit(ws)

VAT = "'הנחות ופרמטרים'!$B$5"
COGS_RATE = "'הנחות ופרמטרים'!$B$6"
GROWTH = "'הנחות ופרמטרים'!$B$7"


# ---------------------------------------------------------------------------
# 2. הכנסות — revenue by channel
# ---------------------------------------------------------------------------
ws = wb.create_sheet("הכנסות")
banner(ws, 1, "הכנסות לפי ערוץ מכירה — 2026", 15)

channels = [
    ("חנות רחוב — חיפה",   142_000),
    ("חנות קניון — קריון", 198_500),
    ("אתר אינטרנט",        126_400),
    ("מרקטפלייס",           74_800),
    ("מכירה סיטונאית",      95_200),
]

style_header(ws, 3, ["ערוץ מכירה"] + MONTHS + ["סה״כ שנתי"])
for r, (channel, base) in enumerate(channels, start=4):
    ws.cell(row=r, column=1, value=channel).alignment = RTL
    ws.cell(row=r, column=1).font = Font(name=FONT_NAME, size=11)
    ws.cell(row=r, column=1).border = BORDER
    for m in range(12):
        col = 2 + m
        c = ws.cell(row=r, column=col)
        if m < ACTUAL_MONTHS:
            # invented but plausible seasonality
            season = [1.00, 0.88, 1.05, 0.97, 1.12, 1.21][m]
            c.value = round(base * season, -2)
            c.font = INPUT_FONT
        else:
            prev = get_column_letter(col - 1)
            c.value = f"={prev}{r}*(1+{GROWTH})"
            c.font = Font(name=FONT_NAME, size=11)
        c.number_format = FMT["shekel_whole"]
        c.alignment = RTL
        c.border = BORDER
    t = ws.cell(row=r, column=14, value=f"=SUM(B{r}:M{r})")
    t.number_format = FMT["shekel_whole"]
    t.alignment = RTL
    t.font = Font(name=FONT_NAME, size=11, bold=True)
    t.border = BORDER

total_row = 4 + len(channels)
ws.cell(row=total_row, column=1, value="סה״כ הכנסות").alignment = RTL
for col in range(1, 15):
    c = ws.cell(row=total_row, column=col)
    if col > 1:
        letter = get_column_letter(col)
        c.value = f"=SUM({letter}4:{letter}{total_row - 1})"
        c.number_format = FMT["shekel_whole"]
    c.font = Font(name=FONT_NAME, bold=True, size=11, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = RTL
    c.border = BORDER

ws.cell(row=total_row + 2, column=1, value="מזה מע״מ").alignment = RTL
v = ws.cell(row=total_row + 2, column=14, value=f"=N{total_row}*{VAT}/(1+{VAT})")
v.number_format, v.alignment, v.font = FMT["shekel_whole"], RTL, LINK_FONT

ws.conditional_formatting.add(
    f"B4:M{total_row - 1}",
    ColorScaleRule(start_type="min", start_color="FFF5F5",
                   end_type="max", end_color="9CC3E5"),
)
freeze_header(ws, 4)
autofit(ws, min_w=13)


# ---------------------------------------------------------------------------
# 3. הוצאות — operating expenses
# ---------------------------------------------------------------------------
ws = wb.create_sheet("הוצאות")
banner(ws, 1, "הוצאות תפעוליות — 2026", 15)

expenses = [
    ("שכר ונלוות",        168_000),
    ("שכר דירה וארנונה",   62_000),
    ("שיווק ופרסום",       41_500),
    ("חשמל, מים ותקשורת",  11_800),
    ("הוצאות משרד",         7_400),
    ("עמלות סליקה",        14_200),
    ("הובלה ושילוח",       19_600),
    ("ביטוח",               8_300),
    ("פחת והפחתות",        12_000),
]
style_header(ws, 3, ["סעיף הוצאה"] + MONTHS + ["סה״כ שנתי"])
for r, (name, base) in enumerate(expenses, start=4):
    ws.cell(row=r, column=1, value=name).alignment = RTL
    ws.cell(row=r, column=1).font = Font(name=FONT_NAME, size=11)
    ws.cell(row=r, column=1).border = BORDER
    for m in range(12):
        col = 2 + m
        c = ws.cell(row=r, column=col)
        drift = [1.00, 0.99, 1.02, 1.01, 1.04, 1.06, 1.05, 1.07, 1.06, 1.08, 1.10, 1.15][m]
        c.value = round(base * drift, -2)
        c.font = INPUT_FONT if m < ACTUAL_MONTHS else Font(name=FONT_NAME, size=11)
        c.number_format = FMT["shekel_whole"]
        c.alignment = RTL
        c.border = BORDER
    t = ws.cell(row=r, column=14, value=f"=SUM(B{r}:M{r})")
    t.number_format, t.alignment, t.border = FMT["shekel_whole"], RTL, BORDER
    t.font = Font(name=FONT_NAME, size=11, bold=True)

exp_total_row = 4 + len(expenses)
ws.cell(row=exp_total_row, column=1, value="סה״כ הוצאות").alignment = RTL
for col in range(1, 15):
    c = ws.cell(row=exp_total_row, column=col)
    if col > 1:
        letter = get_column_letter(col)
        c.value = f"=SUM({letter}4:{letter}{exp_total_row - 1})"
        c.number_format = FMT["shekel_whole"]
    c.font = Font(name=FONT_NAME, bold=True, size=11, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="C00000")
    c.alignment = RTL
    c.border = BORDER
freeze_header(ws, 4)
autofit(ws, min_w=13)


# ---------------------------------------------------------------------------
# 4. רווח והפסד — P&L, entirely formulas linking to the sheets above
# ---------------------------------------------------------------------------
ws = wb.create_sheet("רווח והפסד")
banner(ws, 1, "דוח רווח והפסד — 2026", 15)
style_header(ws, 3, ["סעיף"] + MONTHS + ["סה״כ שנתי"])

REV = f"'הכנסות'!{{col}}{total_row}"
EXP = f"'הוצאות'!{{col}}{exp_total_row}"

pl_rows = [
    ("הכנסות ברוטו",   lambda col: f"={REV.format(col=col)}",                       False),
    ("מע״מ",           lambda col: f"=-{REV.format(col=col)}*{VAT}/(1+{VAT})",      False),
    ("הכנסות נטו",     lambda col: f"={col}4+{col}5",                                True),
    ("עלות המכר",      lambda col: f"=-{col}6*{COGS_RATE}",                          False),
    ("רווח גולמי",     lambda col: f"={col}6+{col}7",                                True),
    ("הוצאות תפעוליות", lambda col: f"=-{EXP.format(col=col)}",                       False),
    ("רווח תפעולי",    lambda col: f"={col}8+{col}9",                                True),
]
for i, (label, f, is_subtotal) in enumerate(pl_rows, start=4):
    ws.cell(row=i, column=1, value=label).alignment = RTL
    ws.cell(row=i, column=1).font = Font(name=FONT_NAME, size=11, bold=is_subtotal)
    ws.cell(row=i, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2 + m)
        c = ws.cell(row=i, column=2 + m, value=f(col))
        c.number_format = FMT["shekel_neg"]
        c.alignment = RTL
        c.border = BORDER
        c.font = Font(name=FONT_NAME, size=11, bold=is_subtotal)
        if is_subtotal:
            c.fill = PatternFill("solid", fgColor=LIGHT)
    t = ws.cell(row=i, column=14, value=f"=SUM(B{i}:M{i})")
    t.number_format, t.alignment, t.border = FMT["shekel_neg"], RTL, BORDER
    t.font = Font(name=FONT_NAME, size=11, bold=True)
    if is_subtotal:
        t.fill = PatternFill("solid", fgColor=LIGHT)

ws.cell(row=12, column=1, value="שיעור רווח גולמי").alignment = RTL
ws.cell(row=13, column=1, value="שיעור רווח תפעולי").alignment = RTL
for row_out, num_row in ((12, 8), (13, 10)):
    ws.cell(row=row_out, column=1).font = Font(name=FONT_NAME, size=11, italic=True)
    for m in range(12):
        col = get_column_letter(2 + m)
        c = ws.cell(row=row_out, column=2 + m, value=f"=IFERROR({col}{num_row}/{col}6,0)")
        c.number_format = FMT["pct"]
        c.alignment = RTL
        c.font = Font(name=FONT_NAME, size=11, italic=True)
    t = ws.cell(row=row_out, column=14, value=f"=IFERROR(N{num_row}/N6,0)")
    t.number_format, t.alignment = FMT["pct"], RTL
    t.font = Font(name=FONT_NAME, size=11, italic=True, bold=True)

ws.conditional_formatting.add(
    "B10:M10",
    CellIsRule(operator="lessThan", formula=["0"],
               font=Font(color="C00000", bold=True),
               fill=PatternFill("solid", fgColor="FFC7CE")),
)
freeze_header(ws, 4)
autofit(ws, min_w=13)


# ---------------------------------------------------------------------------
# 5. מלאי — inventory with reorder logic
# ---------------------------------------------------------------------------
ws = wb.create_sheet("מלאי")
banner(ws, 1, "ניהול מלאי — פריטים פעילים", 11)

style_header(ws, 3, ["מק״ט", "תיאור פריט", "קטגוריה", "מידה", "צבע",
                     "מלאי נוכחי", "צריכה חודשית", "מלאי בטחון",
                     "נקודת הזמנה", "עלות ליחידה", "שווי מלאי"])
items = [
    ("TS-1041", "חולצת טי כותנה",      "חולצות",  "M",   "לבן",   340,  120, 24.90),
    ("TS-1042", "חולצת טי כותנה",      "חולצות",  "L",   "שחור",   92,  145, 24.90),
    ("SH-2210", "חולצה מכופתרת פשתן",  "חולצות",  "L",   "תכלת",  156,   48, 68.00),
    ("JN-3305", "מכנסי ג׳ינס סלים",    "מכנסיים", "32",  "כחול",   78,   62, 96.50),
    ("JN-3307", "מכנסי ג׳ינס ישר",     "מכנסיים", "34",  "שחור",  210,   55, 96.50),
    ("DR-4102", "שמלת קיץ פרחונית",    "שמלות",   "S",   "ורוד",   45,   38, 112.00),
    ("DR-4108", "שמלת ערב",            "שמלות",   "M",   "שחור",   31,   14, 245.00),
    ("JK-5501", "מעיל חורף מרופד",     "מעילים",  "XL",  "אפור",  118,   22, 189.00),
    ("JK-5504", "ג׳קט ג׳ינס",          "מעילים",  "M",   "כחול",   64,   31, 154.00),
    ("AC-6001", "חגורת עור",           "אביזרים", "יחיד", "חום",   205,   40, 42.00),
    ("AC-6014", "צעיף צמר",            "אביזרים", "יחיד", "בורדו",  88,   18, 38.50),
    ("SK-7203", "חצאית מידי",          "חצאיות",  "S",   "ירוק",   52,   27, 78.00),
]
SAFETY = "'הנחות ופרמטרים'!$B$10"
LEADTIME = "'הנחות ופרמטרים'!$B$11"
for r, (sku, desc, cat, size, colour, stock, usage, cost) in enumerate(items, start=4):
    write_row(ws, r,
              [sku, desc, cat, size, colour, stock, usage, None, None, cost, None],
              [None, None, None, None, None, "int", "int", "int", "int", "shekel", "shekel"])
    ws.cell(row=r, column=8, value=f"=ROUND(G{r}*{SAFETY},0)")
    ws.cell(row=r, column=9, value=f"=ROUND(G{r}/30*{LEADTIME}+H{r},0)")
    ws.cell(row=r, column=11, value=f"=F{r}*J{r}")
    for col in (6, 7):
        ws.cell(row=r, column=col).font = INPUT_FONT
    ws.cell(row=r, column=10).font = INPUT_FONT

inv_last = 3 + len(items)
add_table(ws, "טבלת_מלאי", f"A3:K{inv_last}")

# red where current stock has fallen below the reorder point
ws.conditional_formatting.add(
    f"F4:F{inv_last}",
    CellIsRule(operator="lessThan", formula=[f"$I4"],
               font=Font(color="9C0006", bold=True),
               fill=PatternFill("solid", fgColor="FFC7CE")),
)
ws.conditional_formatting.add(
    f"K4:K{inv_last}",
    DataBarRule(start_type="min", end_type="max", color="638EC6"),
)

ws.cell(row=inv_last + 2, column=1, value="סה״כ שווי מלאי").alignment = RTL
ws.cell(row=inv_last + 2, column=1).font = Font(name=FONT_NAME, bold=True, size=11)
tot = ws.cell(row=inv_last + 2, column=11, value=f"=SUM(K4:K{inv_last})")
tot.number_format, tot.alignment = FMT["shekel"], RTL
tot.font = Font(name=FONT_NAME, bold=True, size=12, color=NAVY)

ws.cell(row=inv_last + 3, column=1, value="פריטים מתחת לנקודת הזמנה").alignment = RTL
ws.cell(row=inv_last + 3, column=1).font = Font(name=FONT_NAME, bold=True, size=11)
cnt = ws.cell(row=inv_last + 3, column=11,
              value=f"=SUMPRODUCT(--(F4:F{inv_last}<I4:I{inv_last}))")
cnt.number_format, cnt.alignment = FMT["int"], RTL
cnt.font = Font(name=FONT_NAME, bold=True, size=12, color="C00000")

freeze_header(ws, 4)
autofit(ws, min_w=11)


# ---------------------------------------------------------------------------
# 6. תזרים מזומנים — cash flow, with real date objects
# ---------------------------------------------------------------------------
ws = wb.create_sheet("תזרים מזומנים")
banner(ws, 1, "תזרים מזומנים חודשי — 2026", 6)
style_header(ws, 3, ["חודש", "תאריך סוף חודש", "תקבולים", "תשלומים",
                     "תזרים נטו", "יתרת מזומן מצטברת"])

opening = 285_000
for m in range(12):
    r = 4 + m
    eom = datetime.date(2026, m + 1, 1)
    eom = (eom.replace(day=28) + datetime.timedelta(days=4)).replace(day=1) - datetime.timedelta(days=1)
    ws.cell(row=r, column=1, value=MONTHS[m]).alignment = RTL
    ws.cell(row=r, column=1).font = Font(name=FONT_NAME, size=11)
    d = ws.cell(row=r, column=2, value=eom)          # real date object, not a string
    d.number_format, d.alignment = FMT["date"], RTL
    col = get_column_letter(2 + m)
    inflow = ws.cell(row=r, column=3, value=f"='הכנסות'!{col}{total_row}")
    outflow = ws.cell(row=r, column=4,
                      value=f"=-'הוצאות'!{col}{exp_total_row}-'הכנסות'!{col}{total_row}*{COGS_RATE}")
    net = ws.cell(row=r, column=5, value=f"=C{r}+D{r}")
    bal = ws.cell(row=r, column=6,
                  value=(f"={opening}+E{r}" if m == 0 else f"=F{r - 1}+E{r}"))
    for c in (inflow, outflow, net, bal):
        c.number_format, c.alignment, c.border = FMT["shekel_neg"], RTL, BORDER
        c.font = Font(name=FONT_NAME, size=11)
    inflow.font = LINK_FONT
    net.font = Font(name=FONT_NAME, size=11, bold=True)
    bal.font = Font(name=FONT_NAME, size=11, bold=True)

ws.conditional_formatting.add(
    "F4:F15",
    CellIsRule(operator="lessThan", formula=["0"],
               font=Font(color="9C0006", bold=True),
               fill=PatternFill("solid", fgColor="FFC7CE")),
)
add_table(ws, "טבלת_תזרים", "A3:F15")
freeze_header(ws, 4)
autofit(ws, min_w=15)


# ---------------------------------------------------------------------------
# 7. לוח בקרה — dashboard, moved to the front
# ---------------------------------------------------------------------------
ws = wb.create_sheet("לוח בקרה")
wb.move_sheet("לוח בקרה", offset=-(len(wb.sheetnames) - 1))

banner(ws, 1, "אופנת גלים בע״מ — לוח בקרה 2026", 6)
ws["A2"] = "נתוני דמה. ינואר–יוני בפועל, יולי–דצמבר תחזית."
ws["A2"].font = Font(name=FONT_NAME, size=10, italic=True, color="666666")
ws["A2"].alignment = RTL

kpis = [
    ("הכנסות ברוטו",        "='רווח והפסד'!N4",                     "shekel_whole"),
    ("רווח גולמי",          "='רווח והפסד'!N8",                     "shekel_whole"),
    ("רווח תפעולי",         "='רווח והפסד'!N10",                    "shekel_whole"),
    ("שיעור רווח גולמי",    "='רווח והפסד'!N12",                    "pct"),
    ("שווי מלאי",           "='מלאי'!K17",                          "shekel_whole"),
    ("יתרת מזומן בסוף שנה", "='תזרים מזומנים'!F15",                 "shekel_whole"),
]
style_header(ws, 4, ["מדד", "ערך"])
for i, (label, formula, fmt) in enumerate(kpis, start=5):
    ws.cell(row=i, column=1, value=label).alignment = RTL
    ws.cell(row=i, column=1).font = Font(name=FONT_NAME, size=11, bold=True)
    ws.cell(row=i, column=1).border = BORDER
    c = ws.cell(row=i, column=2, value=formula)
    c.number_format = FMT[fmt]
    c.alignment = RTL
    c.font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
    c.border = BORDER
    c.fill = PatternFill("solid", fgColor=GREY)

rev_ws = wb["הכנסות"]
bar = BarChart()
bar.title = "הכנסות לפי ערוץ — סה״כ שנתי"
bar.y_axis.title = "₪"
bar.add_data(Reference(rev_ws, min_col=14, min_row=3, max_row=3 + len(channels)),
             titles_from_data=True)
bar.set_categories(Reference(rev_ws, min_col=1, min_row=4, max_row=3 + len(channels)))
bar.height, bar.width = 8, 16
ws.add_chart(bar, "D4")

pl_ws = wb["רווח והפסד"]
line = LineChart()
line.title = "רווח תפעולי חודשי"
line.add_data(Reference(pl_ws, min_col=2, max_col=13, min_row=10, max_row=10),
              from_rows=True)
line.set_categories(Reference(pl_ws, min_col=2, max_col=13, min_row=3, max_row=3))
line.height, line.width = 8, 16
ws.add_chart(line, "D22")

autofit(ws, min_w=22)


# ---------------------------------------------------------------------------
# Flip everything to RTL, last, so no sheet is missed.
# ---------------------------------------------------------------------------
rtl_workbook(wb)
wb.save(OUT)

# Post-save pass. מידה holds labels ("32", "L", "יחיד"), not quantities — Excel flags the
# numeric-looking ones with a green triangle. Suppress the warning rather than corrupting
# the column into numbers. Must run after save; see the helper's docstring for why.
suppress_text_number_warning(OUT, {"מלאי": f"D4:D{inv_last}"})

print(f"saved: {OUT}")
print(f"sheets: {', '.join(wb.sheetnames)}")
