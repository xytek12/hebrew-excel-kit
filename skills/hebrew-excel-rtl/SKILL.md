---
name: hebrew-excel-rtl
description: Build Hebrew right-to-left Excel workbooks that render correctly in real Excel and that a non-financial manager can actually read. Use when creating, editing, or fixing .xlsx/.xlsm files containing Hebrew, when the user asks for a Hebrew spreadsheet, "אקסל בעברית", "גיליון בעברית", "דוח כספי באקסל", "לוח בקרה", an Israeli financial report, an inventory or BI workbook, or wants a messy Hebrew workbook rebuilt cleanly. ALSO use for the symptom where a Hebrew spreadsheet looks fine in code or a viewer but opens in Excel with columns in the wrong order, headings on the wrong side, or numbers and English mangled inside Hebrew cells. Covers per-sheet RTL, per-cell reading order, Israeli number and date formats, Hebrew wording and typography, dashboard and report design for managers, and a per-sheet self-check. Do NOT use for Hebrew in web pages or CSS (use hebrew-rtl-best-practices) or Hebrew Word/PDF documents (use hebrew-document-generator).
license: MIT
compatibility: Claude Code, Claude Desktop, Codex, Cursor. Requires Python + openpyxl. No network required.
---

# Hebrew RTL Excel

Two jobs, and both must be done or the file fails:

1. **Make it render right** — RTL is two independent switches, not one.
2. **Make it readable** — a Hebrew workbook nobody can interpret is a failed deliverable
   even when every flag is correct.

## Part 1 — The two switches

| Switch | Scope | What it controls |
|---|---|---|
| `sheet_view.rightToLeft` | **per worksheet** | Column order. Whether `A1` sits top-**right** and columns run right→left. |
| `Alignment(readingOrder=2)` | **per cell** | Bidi resolution *inside* the cell. Where numbers, Latin words and punctuation land within a Hebrew string. |

Set only the first and you get a mirrored grid full of scrambled cells. Set only the second
and you get correct cells in a backwards grid.

`rightToLeft` is per **sheet**. A workbook with 8 sheets needs it 8 times. This is the single
most common defect in the wild — the author only ever looked at the first tab.

```python
for ws in wb.worksheets:            # always a loop, never per-sheet by hand
    ws.sheet_view.rightToLeft = True
```

```python
from openpyxl.styles import Alignment
RTL = Alignment(horizontal="right", vertical="center", readingOrder=2)
cell.alignment = RTL
```

`readingOrder=2` is RTL, `1` is LTR, unset is "context" — Excel guesses from the first strong
character, which is the guessing you are trying to eliminate. Apply it to **every** cell
holding text, including numeric ones: a bare number is unaffected, but the moment someone
types `12 יח׳` into it the cell resolves correctly.

Verified: openpyxl `readingOrder=2` reads back through Excel's COM API as `xlRTL` (`-5004`);
unset reads back as `xlContext` (`-5002`).

**Never reverse a Hebrew string in Python to "fix" display.** The characters are already in
the correct logical order; only the display resolution is wrong. Reversing corrupts the data
and breaks search, sort and copy-paste.

## Part 2 — Israeli formats

```python
SHEKEL  = '#,##0.00\\ [$₪-40D]'                                     # 39.90 ₪
SHEKEL_N = '#,##0.00\\ [$₪-40D];[Red]\\(#,##0.00\\ [$₪-40D]\\)'     # (2,396.00 ₪) in red
DATE_IL = 'DD/MM/YYYY'                                              # 15/01/2026
PERCENT = '0.0%'
```

`[$₪-40D]` is the currency token with Hebrew (Israel) locale id `40D`. Without the locale id
a machine with different regional settings may re-interpret the format.

Never put `₪` in the cell text — the value must stay numeric and sortable. Never write a date
as a string — a text date sorts alphabetically, fails `SUMIFS`, and cannot be grouped in a
pivot. Write `datetime.date` objects and format them.

Full table, all 21 formats rendered through real Excel: `reference/number-formats.md`.

## Part 3 — Formulas

openpyxl does **not** evaluate formulas. `cell.value = "=B2*C2"` stores the formula; the
cached result stays empty until Excel opens the file. This is correct behaviour, not a bug.

- Reading a formula cell back in Python gives you the formula string or `None`, never the
  number. If you need the value in Python, compute it in Python too.
- Prefer real formulas over Python-computed constants for anything the user will maintain.
  A hardcoded total is a landmine; `=SUM(D2:D40)` survives them adding a row.
- Quote Hebrew sheet names: `='רווח והפסד'!N12`.

**Ties break `MATCH`.** `SMALL(range,k)` with `MATCH(...,0)` returns the *first* row holding
that value, so a "top 10 worst" list silently repeats one item whenever values tie. Add a
unique sort key column and match on that:

```python
cell.value = f"=H{r}+ROW()/1000000"     # ranking unchanged, every value now distinct
```

## Part 4 — Design for the reader

The audience is a manager who did not build the file and will look at it for ninety seconds.

- **One accent colour.** Colour on a number means "needs attention" and nothing else.
- **Do not colour inputs blue and formulas black.** That is a modeller's debugging aid; to a
  manager it looks like the file contains two kinds of money. Put every input on one named
  assumptions sheet instead.
- **Turn Excel's gridlines off** on any formatted sheet — they fight your own borders.
- **Never merge inside data.** Use `Alignment(horizontal="centerContinuous")` — same look,
  and sorting, filtering and pivots keep working. Note the camelCase; `center_continuous`
  raises `ValueError`.
- **Dashboard order:** title → 4–6 KPI tiles → plain-Hebrew insight lines → action table →
  charts. Someone who reads only the insight lines should be able to run the meeting.
- **Insight lines are computed, never hardcoded.** Build with `&` and `TEXT()`.
- **Status columns carry the word, then the colour** — readable with colour stripped out.
- **Every number gets its unit** in the header: `מכירות בחודש (₪)`, `מלאי (יח׳)`.

Full guidance: `reference/design-for-managers.md`. Ready-made helpers: `scripts/excel_design.py`.

## Part 5 — Hebrew wording

A cell label is UX copy: a few words read fast under pressure. Business register, not
literary Hebrew. Geresh `׳` and gershayim `״`, not ASCII quotes — `יח׳`, `מק״ט`, `סה״כ`,
`ש״ח`. Ktiv maleh, consistently. Arial unless you know the reader's fonts.

Full guidance, condensed from `hebrew-content-writer`, `hebrew-i18n` and
`israeli-ui-design-system`: `reference/hebrew-for-spreadsheets.md`.

## Part 6 — Nothing else sets RTL for you

**No Excel-writing tool sets RTL.** Not the Excel MCP server, not `pandas.to_excel`, not
most export buttons, not an agent that has not read this skill. They all produce a
left-to-right sheet with Hebrew sitting in it — the exact defect this skill exists to fix.

Measured against `excel-mcp-server` 1.29.0: a sheet written through `write_data_to_excel`
came out with `rightToLeft` unset, `readingOrder` missing on all 11 Hebrew cells, and no
column widths. Excel confirmed it rendered left-to-right.

That is not a fault in the MCP — it is a data tool and it does its job well. It just means
the division of labour is:

| Tool | Job |
|---|---|
| Excel MCP server | hands — read, write, format, chart, pivot |
| this skill | correctness — RTL, formats, wording, design |
| `verify_rtl.py` | proof |
| `fix_rtl.py` | repair |

**Always run the repair after any tool other than your own code touches a Hebrew workbook:**

```bash
python scripts/fix_rtl.py workbook.xlsx --in-place
python scripts/verify_rtl.py workbook.xlsx --com
```

`fix_rtl.py` flips every Hebrew-bearing sheet, adds `readingOrder=2` while preserving
existing alignment, estimates column widths, and turns gridlines off on formatted sheets.
It does not invent design — layout, colour and wording still need Part 4.

### A trap inside the repair

`ws.column_dimensions` is a defaultdict. Merely **reading**
`ws.column_dimensions['A'].width` creates the entry and hands back the phantom default
`13.0`, which is then serialised as a real `<col>` record on save. So the obvious test —
`if not ws.column_dimensions[letter].width:` — both lies (13.0 is truthy, so it never fires)
and silently stamps a uniform useless width onto every column. `customWidth` is `True` on the
phantom too, so it is no help either.

Test membership first, before touching the key:

```python
already_sized = {k for k, d in ws.column_dimensions.items() if d.width}
if letter not in already_sized:
    ws.column_dimensions[letter].width = ...
```

## Part 7 — Self-check every sheet, every time

Never hand over a workbook you have not run the checker against. A whole-workbook verdict
hides the exact defect this skill exists to prevent, so the report is **per sheet**.

```bash
python scripts/verify_rtl.py workbook.xlsx --com
```

Per sheet it checks: RTL flag · reading order on every Hebrew cell · sheet name legality ·
dates typed as dates · currency not written as text · merged cells · column widths set ·
fill-palette size · frozen header · gridlines. With `--com` on Windows it additionally opens
the file through real Excel and asserts, per sheet, that Excel reports `DisplayRightToLeft`
and that the formulas evaluate.

`--strict` makes warnings fail the run. Exit code 0 = clean.

Fix everything it reports before delivering. If you deliberately leave a warning — an LTR
sheet holding a raw English export, say — say so explicitly in the handoff.

## Workflow for a messy inherited workbook

1. Read it twice without touching it: `load_workbook(path, data_only=True)` for cached
   values, then `data_only=False` for the formulas. You need both views.
2. Map before changing: sheet names, used ranges, which columns are text vs numbers vs
   dates, where formulas point, what is hardcoded.
3. Report the defects and your intended plan, then act. Inherited financial workbooks encode
   undocumented business rules; a cleanup that silently drops one is worse than the mess.
4. Build the clean version as a **new file**. Never overwrite the source.
5. Run `verify_rtl.py --com` and fix everything before handing over.

## Reference

| File | What |
|---|---|
| `reference/number-formats.md` | 21 Israeli formats, each verified against real Excel |
| `reference/pitfalls.md` | Failure modes: symptom → cause → fix |
| `reference/design-for-managers.md` | Colour, layout, KPI tiles, insight lines |
| `reference/hebrew-for-spreadsheets.md` | Register, punctuation, typography, plurals |
| `scripts/rtl_helpers.py` | RTL mechanics, formats, tables, ignored-errors patch |
| `scripts/excel_design.py` | The visual system — titles, KPI tiles, status colours |
| `scripts/verify_rtl.py` | The per-sheet self-check |
| `scripts/fix_rtl.py` | Repair RTL in a workbook some other tool wrote |
