# Hebrew Excel Kit

Everything an AI coding agent needs to build **Hebrew right-to-left Excel workbooks that
render correctly in Excel and that a manager can actually read.**

Built and tested on Windows 11 against real Microsoft Excel (Office 16, Hebrew locale).
Every technical claim here was verified by round-tripping files through Excel's COM API.

---

## Install everything with one command

Point your agent at this repo and say **"install everything in here"** — [`INSTALL.md`](INSTALL.md)
is written as an instruction set for the agent, with [`manifest.json`](manifest.json) as the
machine-readable version.

Or do it yourself:

```bash
git clone https://github.com/xytek12/hebrew-excel-kit.git
```
```powershell
cd hebrew-excel-kit ; powershell -ExecutionPolicy Bypass -File .\install.ps1 -Agent both
```
```bash
cd hebrew-excel-kit && bash install.sh both
```

That installs 11 skills into `~/.claude/skills` and `~/.codex/skills`, registers both Excel
MCP servers, installs the Python packages, rebuilds the demo (with real PivotTables when
Excel is present), and verifies it. Re-running is safe.

## What's in here

| Path | What |
|---|---|
| `skills/hebrew-excel-rtl/` | **The main deliverable** — a Hebrew RTL Excel skill. Nothing equivalent exists on any marketplace. |
| `demo/build_demo.py` | Generator for a 10-sheet Hebrew retail-chain model |
| `demo/com_finish.py` | The COM pass: real PivotTables + cached formula values through Excel itself |
| `demo/דוח-כספי-לדוגמה.xlsx` | The generated workbook — 851 live formulas, 2 PivotTables, XLOOKUP cards |
| `INSTALL.md` | Agent-readable install instructions |
| `docs/INSTALL-*.md` | Per-platform detail: Claude Code, Codex, Claude Desktop |
| `docs/SECURITY-REVIEW.md` | Source audit of every third-party component, done before install |

## The problem this solves

**RTL in Excel is two independent switches**, and almost everyone sets only one:

| Switch | Scope | Controls |
|---|---|---|
| `sheet_view.rightToLeft` | per **worksheet** | Column order — whether `A1` sits top-right |
| `Alignment(readingOrder=2)` | per **cell** | Bidi *inside* the cell — where numbers and Latin text land |

Because `rightToLeft` is per *sheet*, the usual failure is a workbook where tab 1 looks
perfect and tab 4 is backwards — the author only ever checked the first one.

**And a correct workbook nobody can read is still a failed deliverable.** The skill covers
both: the RTL mechanics, and the design rules that make a Hebrew report legible to a manager
who did not build it.

## The skill

`skills/hebrew-excel-rtl/` — six parts:

1. **The two switches** — per-sheet RTL, per-cell reading order, and why reversing a Hebrew
   string in Python is the worst possible "fix"
2. **Israeli formats** — 21 number, currency and date formats, each rendered through real Excel
3. **Formulas** — openpyxl never evaluates them; Hebrew sheet names need quoting; ties break
   `MATCH` and silently duplicate rows in a "top 10 worst" list; XLOOKUP with a Hebrew
   `if_not_found`, compound keys for two-condition lookups, same-sheet dropdown lists
4. **Design for the reader** — one accent colour, no blue/black input convention, no merges,
   dashboard ordering, live insight sentences
5. **Hebrew wording** — register, geresh/gershayim, ktiv maleh, typography, the number–gender
   trap. Condensed from `hebrew-content-writer`, `hebrew-i18n` and `israeli-ui-design-system`
6. **Self-check every sheet** — see below

### The self-check

```bash
py skills/hebrew-excel-rtl/scripts/verify_rtl.py workbook.xlsx --com
```

Reports **per sheet**, because a whole-workbook verdict hides exactly the defect this skill
exists to prevent:

```
  גיליון: מלאי לפי חנות   [clean]
  --------------------------------------------------------------------
    [OK  ] RTL (rightToLeft)
    [OK  ] Reading order          316/316 cells
    [OK  ] Dates typed
    [OK  ] Currency typed
    [OK  ] Merged cells           none
    [OK  ] Column widths          14 columns
    [OK  ] Fill palette           2 distinct
    [OK  ] Frozen header          A5
    [OK  ] Gridlines              hidden
    [OK  ] Excel says RTL
    [OK  ] Formulas evaluate      360 formulas
```

`--com` on Windows additionally opens the file through real Excel and asserts, per sheet,
that Excel reports `DisplayRightToLeft` and that formulas evaluate. `--strict` fails on
warnings. Current demo: **10 sheets — 119 passed, 0 warnings, 0 failed, 851 formulas
evaluated, 0 broken.**

## The MCP server does not set RTL — and that is the point

Tested end to end against `excel-mcp-server` 1.29.0 with the tools live in a Claude Code
session. Reads are flawless: Hebrew sheet names, Hebrew data and cross-sheet formulas with
quoted Hebrew names all round-trip correctly.

**But a sheet written through `write_data_to_excel` came out left-to-right** — `rightToLeft`
unset, `readingOrder` missing on all 11 Hebrew cells, no column widths. Excel confirmed it
rendered LTR. That is not a fault in the MCP; it is a data tool, not a localisation tool.
It just means these are complementary, not redundant:

| Tool | Job |
|---|---|
| Excel MCP server | hands — read, write, format, chart, pivot |
| `hebrew-excel-rtl` skill | correctness — RTL, formats, wording, design |
| `verify_rtl.py` | proof |
| `fix_rtl.py` | repair |

The verifier caught the MCP's own output as broken, and the repair tool fixed it:

```bash
py skills/hebrew-excel-rtl/scripts/fix_rtl.py workbook.xlsx --in-place
```
```
  sheets scanned      9
  sheets flipped RTL  1
  cells given RTL     12
  columns sized       5
```

After repair: **9 sheets — 108 passed, 0 warnings, 0 failed**, and Excel reports RTL on the
previously-broken sheet.

Run `fix_rtl.py` after **any** tool other than your own code touches a Hebrew workbook —
that includes `pandas.to_excel` and most export buttons.

## The demo workbook

**אופנת גלים בע״מ** — a fictional Israeli clothing chain with **6 branches**. All figures
invented. ₪130k/month per store, 8.6% operating margin — deliberately realistic retail
numbers rather than flattering ones.

Built in two passes: `build_demo.py` (openpyxl — structure, formulas, design, tables,
dropdowns) then `com_finish.py` (real Excel via COM — PivotTables, full recalculation,
cached values). The order is one-way: an openpyxl re-save after the COM pass would delete
the charts and pivots.

| Sheet | What a manager gets from it |
|---|---|
| לוח בקרה | 6 KPI tiles, 6 plain-Hebrew sentences computed live, top-10 urgent items, 2 charts |
| חיפוש מהיר | Interactive: pick a store or an item from a dropdown, XLOOKUP fills the card live |
| ניתוח Pivot | Two **real PivotTables** — sales by category × store, stuck stock by status × store |
| הנחות ופרמטרים | Every changeable number in the model, in one place, each with "why it matters" |
| מלאי לפי חנות | 72 rows (12 SKUs × 6 branches) with weeks-of-cover and a status word |
| סיכום חנויות | Which branch sells most, which has stock stuck, ranked |
| הכנסות / הוצאות | By branch and by line, 12 months |
| רווח והפסד | P&L, all formulas, gross and operating margin |
| תזרים מזומנים | Monthly cash flow with running balance |

The dashboard answers the questions directly, in Hebrew, and updates itself:

> 1. החנות המובילה: עזריאלי תל אביב — 170,721 ש״ח בחודש.
> 3. הפריט שנגמר הכי מהר: מכנסי ג׳ינס סלים מידה 32 בקניון הקריון — נשארו 0.5 שבועות מלאי.
> 6. הרווח התפעולי הוא 8.6% מההכנסות. זה בריא.

## What gets installed

All reviewed before install — [`docs/SECURITY-REVIEW.md`](docs/SECURITY-REVIEW.md). The
stack covers the whole pipeline: **intake → structure → render → prove**.

**MCP servers**
- [`haris-musa/excel-mcp-server`](https://github.com/haris-musa/excel-mcp-server) (4.1k★, MIT)
  — file-level: 25 tools, works **without** Excel installed
- [`sbroenne/mcp-server-excel`](https://github.com/sbroenne/mcp-server-excel) (MIT, Windows+Excel)
  — drives the **real Excel engine** via COM: live recalculation, PivotTables, Power Query,
  DAX, range screenshots. Carries opt-out anonymous usage telemetry (no file content —
  disclosed in the security review)

**Skills**
- `hebrew-excel-rtl` — this repo: RTL correctness, design, verify + repair
- [`xlsx`](https://github.com/anthropics/skills) — Anthropic's official spreadsheet skill: the
  zero-formula-errors gate, messy-file intake (Windows: clone with `core.longpaths=true`)
- [`excel-hygiene`](https://github.com/BayramAnnakov/excel-hygiene) — independent error
  checker + the Panko–Halverson spreadsheet-error taxonomy
- [`clean-data-xls`, `audit-xls`, `3-statement-model`, `dcf-model`, `comps-analysis`,
  `xlsx-author`](https://github.com/anthropics/financial-services) — the CFO layer (Anthropic,
  first-party). Installed as **copied skills, not as the plugin** — the plugin form would
  register 12 remote data connectors
- [`spreadsheet`](https://github.com/davila7/claude-code-templates) — openpyxl/pandas conventions
- [`inventory-manager`](https://github.com/jmsktm/claude-settings) — reorder points, forecasting

**Deliberately rejected:** `xlsx-for-ai` (hosted API — spreadsheet data would leave the
machine), `excel-automation` (macOS-only). Reasons in the security review.

## Known limits

- `--com` verification needs Windows with Excel installed. Static checks run anywhere.
- openpyxl does not evaluate formulas — reading a formula cell in Python gives the formula
  string until Excel has opened and saved the file once.
- Hebrew month names (`MMMM`) come from the reader's Office language pack, not the file.
- Claude Desktop has no shell, so it cannot run the verifier. Use Claude Code or Codex to
  generate and verify; Desktop is fine for reading and editing through the MCP tools.
- Heebo/Rubik/Assistant look better than Arial but must be installed on the reader's machine.
  The skill defaults to Arial for anything you might email.
- Windows PowerShell 5.1 reads `.ps1` as ANSI without a BOM — `install.ps1` ships with one.

## License

MIT. The demo data is fictional.
