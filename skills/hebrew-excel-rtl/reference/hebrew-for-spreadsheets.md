# Hebrew writing for spreadsheets

Condensed from the `hebrew-content-writer`, `hebrew-i18n` and `israeli-ui-design-system`
skills, keeping only what applies to a workbook. A spreadsheet label is UX copy: a handful
of words, read fast, under pressure, by someone who did not build the file.

---

## Register

A business workbook is **business register** — clear and professional, not formal-literary
and not slangy. The most common failure is an agent defaulting to `safa gvoha`
(literary Hebrew), which reads like a government form.

| Too formal | Use |
|---|---|
| סך כל ההכנסות ברוטו | סה״כ הכנסות |
| יתרת מזומנים לתום התקופה | מזומן בסוף השנה |
| פריטים אשר מלאים נמוך מנקודת ההזמנה | פריטים שצריך להזמין |
| בהתאם לנתונים המוצגים לעיל | לפי הנתונים למעלה |

Column headers take the **imperative/nominal** style, not sentences. `מלאי (יח׳)`, not
`כמות המלאי הקיימת ביחידות`.

## Gendered language

For headers, statuses and totals, prefer **gender-neutral rewording** over slash notation —
`יש להזמין` rather than `הזמן/הזמיני`. Slashes inside a narrow column wrap badly and read
as noise.

## Ktiv maleh

Use full spelling consistently: `תוכנה` not `תכנה`, `שירות` not `שרות`, `הזמנה` not `הזמנה`
mixed forms. Pick one and never mix within a workbook — a header reading `מלאי` on one sheet
and `מלאיים` on another looks careless.

## Punctuation — geresh and gershayim

Hebrew abbreviations use **geresh** `׳` (U+05F3) and **gershayim** `״` (U+05F4), not the
ASCII `'` and `"`.

| Wrong | Right | Meaning |
|---|---|---|
| יח' | יח׳ | units |
| מק"ט | מק״ט | SKU |
| סה"כ | סה״כ | total |
| ש"ח | ש״ח | shekels |
| ק"ג | ק״ג | kilograms |

Two reasons this matters beyond looking right. ASCII quotes are a tell that text was
machine-generated. And an unescaped `"` inside an Excel number-format string or a formula
literal terminates the string early — `"סה"כ"` is a syntax error, `"סה״כ"` is fine.

In a Python source file both characters are just ordinary Unicode; no escaping needed.

## Numbers and dates

- Israeli date order is **day before month**, always. `DD/MM/YYYY`.
  `Intl.DateTimeFormat('he-IL')` actually renders dots (`15.01.2026`); both read as native,
  slashes are more common in finance. Pick one per workbook.
- 24-hour clock. `14:30`, never `2:30 PM`.
- Currency renders **after** the number: `1,234.50 ₪`.
- Thousands separator is a comma, decimal a point — same as English.
- In labels, prefer digits: `3 ימים`, not `שלושה ימים`.

### The number–gender trap

Hebrew numbers 1–10 take the **opposite** gender form to the noun they count. Agents get
this wrong constantly:

- `שלושה ימים` — three days (יום is masculine, so the number takes the feminine-looking form)
- `שלוש שנים` — three years (שנה is feminine, so the number takes the masculine-looking form)

This only bites in written sentences — the live insight lines on a dashboard. Column headers
rarely count anything, so they are safe.

### Plurals

Modern Hebrew has three CLDR categories: `one`, `two`, `other`. There is no `many` — it was
removed in CLDR 42. For a dashboard sentence built with `TEXT()`, the safe construction
avoids the problem entirely:

```
="נמצאו "&TEXT(n,"0")&" פריטים"
```

`פריטים` works for 0, 2, and 5. Do not try to branch on count inside a formula; it makes the
formula unreadable for a gain nobody notices.

## Mixed Hebrew and English

Israeli business writing mixes freely. Decide per term:

- **Keep English** — SKUs, model codes, brand names, sizes (`XL`, `M`), `API`, `SaaS`
- **Transliterate** — words people actually say that way: `אימייל`, `סטארטאפ`, `באג`
- **Translate** — where a normal Hebrew word exists: `משתמש`, `הורדה`, `עדכון`, `מלאי`

English nouns take Hebrew grammar naturally — `הדשבורד`, `לינקים`. Do not "correct" these.

**In a cell**, any mixed-script string needs `readingOrder=2` or the Latin run lands in the
wrong place. See the main SKILL.md.

## Typography in Excel

| Font | Verdict |
|---|---|
| **Arial** | The default choice. Ships with every Office install, Windows and Mac. |
| **Calibri** | Fine on Office 2007+. Slightly softer. |
| **David**, **Narkisim**, **FrankRuehl**, **Guttman \*** | Windows-only. Substitute badly elsewhere. |
| **Heebo**, **Rubik**, **Assistant** | Best looking. Google Fonts — **must be installed on the reader's machine.** Never in a file you email. |

A missing font substitutes silently and can change every column width in the workbook. If
you do not know what the recipient has, use Arial.

Hebrew renders visually larger than Latin at the same point size, and there are no
ascenders/descenders to give the eye a hook. Two consequences:

- Column widths need roughly a **1.3×** factor over character count.
- Body text at 11pt is right; 9pt Hebrew is genuinely hard to read where 9pt English is fine.

## Naming sheets

- Max 31 characters, and Excel forbids `: \ / ? * [ ]`
- Hebrew letters, spaces, geresh and gershayim are all fine
- Name by **what a manager wants**, not by what the data is:
  `מלאי לפי חנות` beats `נתוני מלאי`; `מה חשוב לדעת` beats `סיכום`
- Quote Hebrew sheet names in formulas: `='רווח והפסד'!N12`

## Writing the insight lines

The single highest-value thing on a Hebrew dashboard is a few sentences of plain Hebrew that
compute themselves. Rules:

1. **Lead with the answer.** `החנות המובילה: עזריאלי תל אביב` — not
   `בבחינת נתוני המכירות עולה כי...`
2. **One fact per line.** If it needs a comma-and-then, split it.
3. **Give the number and its unit.** `170,721 ש״ח בחודש`, not `הכי הרבה`.
4. **Say what it means when you can.** `זה שולי — כל ירידה במכירות תמחק אותו` is worth ten
   rows of numbers.
5. **Never hardcode.** Build with `&` and `TEXT()` so the sentence cannot go stale.

```python
f'="החנות המובילה: "&INDEX({names},MATCH(MAX({sales}),{sales},0))'
f'&" — "&TEXT(MAX({sales}),"#,##0")&" ש""ח בחודש."'
```

Note `""` inside the Excel string literal — that is how you get one `"` into a formula. In
Python it is just a doubled quote inside the f-string.
