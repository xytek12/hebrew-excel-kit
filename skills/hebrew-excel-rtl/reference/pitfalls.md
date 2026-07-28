# Hebrew Excel pitfalls

Each entry: what the user sees, what actually caused it, how to fix it.

---

## Columns run the wrong way / header is on the left

**Cause:** `sheet_view.rightToLeft` not set on that sheet.

**Fix:** `ws.sheet_view.rightToLeft = True` — and check *every* sheet, not just the first.
This flag is per worksheet. A 6-sheet workbook needs it 6 times.

```python
for ws in wb.worksheets:
    ws.sheet_view.rightToLeft = True
```

---

## Sheet 1 looks perfect, sheet 4 is backwards

Same cause as above. This is the most common real-world version of the bug because the
author only ever looked at the first tab. Always loop.

---

## Numbers or English words jump to the wrong end of a Hebrew cell

Symptom: `דגם ABC-12` displays as `ABC-12 דגם`, or a phone number lands before the label.

**Cause:** cell reading order is unset, so Excel resolves bidi from the first strong
character in the string. Mixed-script cells resolve inconsistently.

**Fix:** `Alignment(readingOrder=2)` on the cell. Not on the sheet — per cell.

**Do NOT fix this by reversing the string in Python.** That corrupts the stored data. The
characters are already in the correct logical order; only the *display* resolution is wrong.

---

## Parentheses or quotes point the wrong way

Symptom: `(סה"כ)` renders as `)סה"כ(`.

**Cause:** same as above — brackets are neutral characters and take direction from context.

**Fix:** `readingOrder=2`. If a single label mixes a Hebrew phrase with a bracketed Latin
term and still resolves badly, split it into two cells rather than fighting bidi.

---

## Dates sort wrongly, or `SUMIFS` over a date range returns 0

**Cause:** dates stored as strings. `"15/01/2026"` is text, not a date.

**Fix:** write `datetime.date(2026, 1, 15)` and set `number_format = 'DD/MM/YYYY'`.
To detect it in an inherited file: a date column where `isinstance(cell.value, str)` is true.

---

## Dates are off by a month — 03/04 became April 3rd instead of March 4th

**Cause:** a string date was parsed by Excel on import using US month-first order.

**Fix:** never let Excel parse date strings. Convert in Python with an explicit format
(`datetime.strptime(s, "%d/%m/%Y")`) and write real date objects.

---

## Totals show 0 or blank when read back in Python

**Cause:** openpyxl does not evaluate formulas. The cached value is empty until Excel opens
the file and recalculates.

**Fix:** this is expected, not a bug. If you need the number in Python, compute it in Python
too. If you need the file to carry cached values, open it once in Excel/LibreOffice and save.
`load_workbook(path, data_only=True)` returns `None` for formula cells in a file Excel has
never opened.

---

## Excel shows "We found a problem with some content..."

**Causes, in order of likelihood:**

1. A defined name or table name containing a space or an illegal character.
2. Two tables sharing a `displayName`.
3. A sheet name over 31 characters or containing `: \ / ? * [ ]`.
4. A chart referencing a range on a deleted sheet.
5. A conditional-formatting range that does not match the data range.

**Fix:** these are all build-time errors you can catch before shipping. `verify_rtl.py`
checks 1–3.

---

## Hebrew renders as boxes or question marks

**Cause:** the font on those cells has no Hebrew coverage, and the substitution failed.

**Fix:** set an explicit font with Hebrew coverage on every styled cell — Arial is the safe
universal choice. Do not rely on the workbook theme; the Hebrew fallback slot is resolved by
the OS, not the file.

---

## Columns are too narrow, everything shows `########`

**Cause:** there is no auto-fit in the xlsx format. Excel computes fit at render time;
openpyxl cannot measure text and does not set widths for you.

**Fix:** estimate width explicitly. Hebrew needs roughly a 1.3× factor over character count:

```python
width = max(12, min(50, int(max_len * 1.3) + 4))
```

Note that `########` on a *number* cell means width, not a broken format. Widen before
concluding the number format is wrong.

---

## Pivot table refuses to group, or filter drops rows

**Cause:** merged cells, or a header row that is not a single contiguous row, or blank rows
inside the data range.

**Fix:** never merge inside a data range. Use `Alignment(horizontal="centerContinuous")` for
a visually centered banner. Keep exactly one header row and no blank rows.

---

## Conditional formatting colours look wrong or missing

**Cause:** the number format already carries a colour (e.g. `[Red]`), and it wins over the
conditional format.

**Fix:** pick one system. For cells under a data bar or colour scale, keep the number format
colour-neutral.

---

## The file is correct in LibreOffice but wrong in Excel (or vice versa)

**Cause:** LibreOffice is more forgiving about reading order and locale tokens.

**Fix:** treat real Excel as the only authority for a file that will be opened in Excel.
`verify_rtl.py --com` on Windows is the check that matters.
