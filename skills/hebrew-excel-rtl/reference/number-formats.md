# Israeli number formats for Excel

Every format below was written with openpyxl and read back through the Excel COM API on
Windows with Office 16. The "Renders as" column is the literal `.Text` Excel returned.

## Currency

| Purpose | Format code | Renders as |
|---|---|---|
| Shekel, 2 decimals | `#,##0.00\ [$₪-40D]` | `39.90 ₪` |
| Shekel, whole | `#,##0\ [$₪-40D]` | `1,250 ₪` |
| Shekel, red negatives in parens | `#,##0.00\ [$₪-40D];[Red]\(#,##0.00\ [$₪-40D]\)` | `(2,396.00 ₪)` |
| Shekel, dash for zero | `#,##0.00\ [$₪-40D];[Red]\(#,##0.00\ [$₪-40D]\);"—"` | `—` |
| Thousands (₪ thousands) | `#,##0,\ "אלפי ₪"` | `1,250 אלפי ₪` |
| Millions | `#,##0.0,,\ "מ׳ ₪"` | `1.3 מ׳ ₪` |
| USD alongside | `#,##0.00\ [$$-409]` | `39.90 $` |

`[$₪-40D]` = currency symbol `₪` with locale id `40D` (hex for 1037, Hebrew–Israel). Always
include the locale id. Without it the format is ambiguous and a machine with different
regional settings may re-render or re-interpret it.

In a Python string literal the backslash must be escaped: `'#,##0.00\\ [$₪-40D]'`.
The `\ ` is a literal escaped space — it is what puts the gap between number and symbol.

## Dates and time

| Purpose | Format code | Renders as |
|---|---|---|
| Israeli short date | `DD/MM/YYYY` | `15/01/2026` |
| Short, dots | `DD.MM.YYYY` | `15.01.2026` |
| Long Hebrew month | `D\ בMMMM\ YYYY` | `15 בינואר 2026` |
| Month + year | `MMMM\ YYYY` | `ינואר 2026` |
| Quarter label | `"רבעון "0` (on a number 1–4) | `רבעון 1` |
| 24-hour time | `HH:MM` | `14:30` |
| Timestamp | `DD/MM/YYYY\ HH:MM` | `15/01/2026 14:30` |

Hebrew month names come from the reader's Office language pack. On an English-only Office
install `MMMM` renders `January`, not `ינואר`. If the month name must be Hebrew regardless of
the reader's locale, write it as text in a separate label column and keep the real date in a
properly typed column next to it.

**Always write dates as `datetime.date` / `datetime.datetime` objects.** A date stored as a
string is text: it sorts alphabetically (`01/12` before `02/01`), fails `SUMIFS`, and cannot
be grouped in a pivot table. This is the most common defect in inherited Israeli workbooks.

## Numbers and percentages

| Purpose | Format code | Renders as |
|---|---|---|
| Integer with separators | `#,##0` | `1,250` |
| 2 decimals | `#,##0.00` | `1,250.00` |
| Percent | `0.0%` | `12.5%` |
| Percent with sign | `+0.0%;-0.0%;0.0%` | `+12.5%` |
| Variance, red negative | `+0.0%;[Red]-0.0%;"—"` | `-3.2%` in red |
| Quantity with unit | `#,##0\ "יח׳"` | `120 יח׳` |
| Weight | `#,##0.00\ "ק״ג"` | `2.50 ק״ג` |

Percent format expects the **fraction**, not the number. `0.125` with `0.0%` renders `12.5%`.
Writing `12.5` gives you `1250.0%`.

## Conditional-format friendly patterns

For a variance column that a data bar or colour scale will also cover, keep the number format
neutral (`+0.0%;-0.0%`) and let the conditional format carry the colour. Two colour systems on
one cell fight, and the number format wins — which makes the conditional format look broken.

## Punctuation note

Hebrew uses geresh `׳` (U+05F3) and gershayim `״` (U+05F4), not the ASCII apostrophe `'` and
quote `"`. In unit labels write `יח׳` and `ק״ג`. The ASCII versions are a tell that the text
was machine-generated, and in a format string an unescaped `"` terminates the literal early.
