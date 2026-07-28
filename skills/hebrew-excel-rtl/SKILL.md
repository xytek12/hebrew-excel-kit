---
name: hebrew-excel-rtl
description: Build Hebrew right-to-left Excel workbooks that render correctly in real Excel. Use when creating, editing, or fixing .xlsx/.xlsm files that contain Hebrew, when the user asks for a Hebrew spreadsheet, "אקסל בעברית", "גיליון בעברית", "דוח כספי באקסל", an Israeli financial report, or a Hebrew BI/dashboard workbook. ALSO use for the symptom where a Hebrew spreadsheet looks fine in code or in a viewer but opens in Excel with columns in the wrong order, headers on the wrong side, or English/numbers/punctuation mangled inside Hebrew cells. Covers per-sheet RTL, per-cell reading order, shekel and Israeli date number formats, RTL tables, charts, conditional formatting, and a verification script. Do NOT use for Hebrew in web pages or CSS (use hebrew-rtl-best-practices) or Hebrew Word/PDF documents (use hebrew-document-generator).
license: MIT
compatibility: Claude Code, Claude Desktop, Codex, Cursor. Requires Python + openpyxl. No network required.
---

# Hebrew RTL Excel

## The one thing that breaks everything

Right-to-left in Excel is **not** a document setting and **not** a cell alignment. It is two
independent switches that must both be set:

| Switch | Scope | What it controls |
|---|---|---|
| `sheet_view.rightToLeft` | **per worksheet** | Column order. Whether `A1` sits top-**right** and columns run right→left. |
| `Alignment(readingOrder=2)` | **per cell** | Bidi resolution *inside* the cell. Where numbers, Latin words, and punctuation land within a Hebrew string. |

Setting only the first gives you a mirrored grid full of scrambled cells. Setting only the
second gives you correct cells in a backwards grid. Both, or the workbook is wrong.

`rightToLeft` is per **sheet** — a workbook with 6 sheets needs it set 6 times. This is the
single most common defect; the first sheet looks right and nobody scrolls further.

## Instructions

### Step 1 — Flip every sheet

```python
for ws in wb.worksheets:
    ws.sheet_view.rightToLeft = True
```

Write it as a loop over `wb.worksheets`, never per-sheet by hand. If a sheet must stay LTR
(raw English export, a data dump feeding a formula), leave it and say so in the handoff —
it is a legitimate choice, but it should be deliberate.

### Step 2 — Set reading order on every cell holding text

`readingOrder=2` is RTL. `readingOrder=1` is LTR. `0`/unset is "context" — Excel guesses
from the first strong character, which is exactly the guessing you are trying to eliminate.

```python
from openpyxl.styles import Alignment

RTL = Alignment(horizontal="right", vertical="center", readingOrder=2, wrap_text=True)
cell.alignment = RTL
```

Use `readingOrder=2` on **every** cell, including the numeric ones. A bare number is
unaffected, but the moment someone types `12 יח'` into it the cell resolves correctly.

Verified mapping — openpyxl `readingOrder=2` becomes Excel's `xlRTL` (`-5004`) when read
back through the COM API. Unset becomes `xlContext` (`-5002`).

### Step 3 — Use Israeli number formats

Never write `₪` as a literal prefix into the string. Use a number format so the value stays
numeric and stays sortable.

```python
SHEKEL   = '#,##0.00\\ [$₪-40D]'
SHEKEL_N = '#,##0.00\\ [$₪-40D];[Red]\\(#,##0.00\\ [$₪-40D]\\)'   # negatives red, in parens
DATE_IL  = 'DD/MM/YYYY'
PERCENT  = '0.0%'
```

`[$₪-40D]` is the currency token with the Hebrew (Israel) locale id `40D` (hex for 1037).
Without the locale id Excel may re-interpret the format on a machine with different regional
settings. Verified rendering in Excel: `39.90 ₪`, and `(2,396.00 ₪)` in red for negatives.

`DD/MM/YYYY` matters more than it looks. Write a date as a `datetime` object, never a
string — a string date is text, sorts alphabetically, and breaks every formula downstream.

See `reference/number-formats.md` for the full table.

### Step 4 — Write formulas as strings and let Excel compute

openpyxl does **not** evaluate formulas. `cell.value = "=B2*C2"` stores the formula; the
cached result is empty until Excel opens the file. This is normal and correct.

Verified: a workbook written by openpyxl with `=B2*C2` opens in Excel showing `4788`.

Two consequences:
- **Never** read a computed value back with openpyxl expecting a number. You get `None` or
  the formula string. If you need the value, either compute it in Python as well, or open
  the file through Excel/LibreOffice once to populate the cache.
- Prefer real formulas over Python-computed constants for anything the user will maintain.
  A hardcoded total is a landmine; `=SUM(D2:D40)` survives them adding a row.

Hebrew sheet names inside formulas must be quoted: `='דוח מכירות'!D2`.

### Step 5 — Tables, charts, panes

All three survive the round trip and all three respect the sheet-level RTL flag.

```python
from openpyxl.worksheet.table import Table, TableStyleInfo
tbl = Table(displayName="טבלת_מכירות", ref="A1:E40")   # underscores, no spaces
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
ws.add_table(tbl)
```

`displayName` accepts Hebrew but **not spaces** — openpyxl raises
`ValueError: Table names cannot have spaces` at build time. Use underscores. The name must
also be unique across the workbook.

Freeze the header with `ws.freeze_panes = "A2"`. In an RTL sheet Excel mirrors this
automatically; do not try to compensate by freezing a different cell.

For charts, set `chart.title` in Hebrew directly. Category axis labels come from the cells
and inherit their reading order.

### Step 6 — Verify, do not assume

Run the bundled checker before you hand the file over:

```bash
python scripts/verify_rtl.py path/to/workbook.xlsx
```

It fails loudly on: a sheet missing `rightToLeft`, Hebrew text cells with no `readingOrder`,
currency columns without a number format, date columns stored as strings, and table names
containing spaces.

On Windows with Excel installed, add `--com` to additionally open the file through Excel and
assert `DisplayRightToLeft` is true on every sheet and that formulas actually evaluate. That
is the only check that proves the file is genuinely correct rather than merely well-formed.

## Typography

Hebrew Excel needs a font with real Hebrew coverage and Latin coverage that does not clash:

| Font | Use | Notes |
|---|---|---|
| **Arial** | safest default | Ships with every Office install on every platform. Boring, universal. |
| **Calibri** | Office 2007+ default | Good Hebrew. Safe on Windows and Mac Office. |
| **David** | formal/legal Hebrew | Windows-only. Falls back badly on Mac. |
| **Heebo**, **Rubik**, **Assistant** | branded reports | Google Fonts, must be installed on the reader's machine. Do not use for a file you are emailing to someone. |

Pick **Arial** unless you know what the recipient has installed. A missing font silently
substitutes and can change every column width in the workbook.

Set an explicit font on every styled cell. Do not rely on the theme — the workbook theme
carries Latin and East-Asian font slots but the Hebrew fallback is resolved by the OS.

## Column widths

Hebrew renders wider than the same character count in English, and openpyxl cannot measure
text. Auto-fit does not exist in the file format — Excel computes it at render time and
openpyxl has no equivalent.

Estimate instead, and be generous:

```python
width = max(12, min(50, int(max(len(str(v)) for v in column_values) * 1.3) + 4))
```

The `1.3` factor is for Hebrew; a 1.0 factor produces visibly cramped columns. Cap at 50 so
one long free-text cell does not blow out the layout.

## Anti-patterns

**Do not** write `dir` attributes, CSS, or HTML anywhere near this. That is the web RTL
problem and shares nothing with the Excel one but the name.

**Do not** reverse Hebrew strings in Python to "fix" display. If text looks backwards, the
reading order is wrong — reversing the characters corrupts the actual data and breaks search,
sort, and copy-paste. This is the single worst thing you can do to a Hebrew file.

**Do not** put the `₪` sign in the cell text. Use the number format.

**Do not** merge cells for layout. Merged cells break sorting, filtering, and pivot tables —
the three things a BI workbook exists to do. Use `Alignment(horizontal="centerContinuous")`
across a range if you need a visually centered banner without the merge. (Note the exact
camelCase spelling — `center_continuous` raises `ValueError` in openpyxl.)

**Do not** name a sheet with characters Excel forbids: `: \ / ? * [ ]`. Hebrew letters,
spaces, and quotes are fine. Cap at 31 characters.

## Workflow for a messy inherited workbook

1. Read it without touching it — `openpyxl.load_workbook(path, data_only=True)` to see cached
   values, then again with `data_only=False` to see the formulas. You need both views.
2. Map what exists before changing anything: sheet names, used ranges, which columns are
   text vs numbers vs dates, where the formulas point, what is hardcoded.
3. Report the defects you found and what you intend to do, then act. Inherited financial
   workbooks encode undocumented business rules; a "cleanup" that silently drops one is worse
   than the mess.
4. Build the clean version as a **new file**. Never overwrite the source.
5. Run `verify_rtl.py --com` on the result.

## Reference

- `reference/number-formats.md` — full Israeli number, currency, date, and percent formats
- `reference/pitfalls.md` — the failure modes with symptoms, causes, and fixes
- `scripts/rtl_helpers.py` — importable helpers (`apply_rtl`, `style_header`, `autofit`)
- `scripts/verify_rtl.py` — the checker, with optional real-Excel COM verification
