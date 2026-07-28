# Hebrew Excel Kit

Everything an AI coding agent needs to build **Hebrew right-to-left Excel workbooks that
actually render correctly in Excel** — plus the finance/BI tooling to make them useful.

Built and tested on Windows 11 against real Microsoft Excel (Office 16, Hebrew locale).
Every technical claim in this repo was verified by round-tripping a file through Excel's COM
API, not inferred from documentation.

---

## What's in here

| Path | What it is |
|---|---|
| `skills/hebrew-excel-rtl/` | **The main deliverable.** An agent skill for Hebrew RTL Excel. Nothing equivalent exists on any skills marketplace. |
| `demo/build_demo.py` | Generator for a 7-sheet Hebrew financial model (fictional Israeli clothing retailer) |
| `demo/דוח-כספי-לדוגמה.xlsx` | The generated workbook — 280 live formulas, all RTL |
| `docs/INSTALL-claude-code.md` | Install for Claude Code |
| `docs/INSTALL-codex.md` | Install for OpenAI Codex |
| `docs/INSTALL-claude-desktop.md` | Install for Claude Desktop |
| `docs/SECURITY-REVIEW.md` | Source audit of every third-party component, done before install |
| `install.ps1` / `install.sh` | One-shot installer |

## The problem this solves

RTL in Excel is **two independent switches**, and almost everyone sets only one:

| Switch | Scope | Controls |
|---|---|---|
| `sheet_view.rightToLeft` | per **worksheet** | Column order — whether `A1` sits top-right |
| `Alignment(readingOrder=2)` | per **cell** | Bidi *inside* the cell — where numbers and Latin text land |

Set only the first and you get a mirrored grid full of scrambled cells. Set only the second
and you get correct cells in a backwards grid. And because `rightToLeft` is per *sheet*, the
usual failure is a workbook where tab 1 looks perfect and tab 4 is backwards.

The `hebrew-excel-rtl` skill encodes this, the Israeli number formats, and a verifier that
proves a finished file is correct instead of assuming it.

## Quick start

```bash
git clone https://github.com/xytek12/hebrew-excel-kit.git
```
```bash
cd hebrew-excel-kit && pwsh -File install.ps1
```

Then rebuild the demo to check your setup:

```bash
py demo/build_demo.py
```
```bash
py skills/hebrew-excel-rtl/scripts/verify_rtl.py "demo/דוח-כספי-לדוגמה.xlsx" --com
```

Expected: `24 passed, 0 warnings, 0 failed`.

## The demo workbook

Fictional company **אופנת גלים בע״מ**, an Israeli clothing retailer. All figures invented.

| Sheet | Contents |
|---|---|
| לוח בקרה | KPI dashboard, 2 charts, all values linked by formula |
| הנחות ופרמטרים | Every hardcoded input in the model, in one place |
| הכנסות | Revenue by 5 sales channels × 12 months |
| הוצאות | 9 operating expense lines × 12 months |
| רווח והפסד | P&L, 100% formulas, gross and operating margin |
| מלאי | 12 SKUs with reorder points, safety stock, live below-ROP flagging |
| תזרים מזומנים | Monthly cash flow with running balance |

It demonstrates: per-sheet RTL, per-cell reading order, shekel/date/percent formats, Excel
tables with Hebrew names, cross-sheet formulas with quoted Hebrew sheet names, conditional
formatting (data bars, colour scales, threshold rules), frozen panes, charts with Hebrew
titles, and the financial-modelling colour convention (blue = input, black = formula,
green = cross-sheet link).

Verified in real Excel: opens with no repair prompt, all 7 sheets report
`DisplayRightToLeft = True`, and all 280 formulas evaluate.

## What gets installed

Reviewed before install — see [`docs/SECURITY-REVIEW.md`](docs/SECURITY-REVIEW.md).

**MCP server** — [`haris-musa/excel-mcp-server`](https://github.com/haris-musa/excel-mcp-server)
(4.1k★, MIT). 25 tools: read/write cells, formulas, pivot tables, charts, formatting,
conditional formatting, data validation, tables, sheet management. Does not require Excel
to be installed.

**Skills**
- `hebrew-excel-rtl` — this repo
- [`spreadsheet`](https://github.com/davila7/claude-code-templates) — openpyxl/pandas conventions, investment-banking formatting
- [`inventory-manager`](https://github.com/jmsktm/claude-settings) — reorder points, demand forecasting
- [`audit-xls`](https://github.com/anthropics/financial-services) — formula tracing, hardcode detection, balance checks (Anthropic, first-party)

**Python** — `openpyxl`, `pandas`, `xlsxwriter`, and `pywin32` for the Excel COM verifier.

## Known limits

- `verify_rtl.py --com` needs Windows with Excel installed. The static checks run anywhere.
- openpyxl does not evaluate formulas. Reading a formula cell back in Python gives you the
  formula string, not the number, until Excel has opened and saved the file once. This is
  expected — see the skill for what to do about it.
- Hebrew month names (`MMMM`) come from the reader's Office language pack, not the file.
- `excel-mcp-server`'s HTTP transport defaults to binding `0.0.0.0`. Use stdio, which is what
  the installers configure.

## License

MIT. The demo data is fictional.
