# Install on Claude Code

## 1. Python prerequisites

```bash
py -m pip install openpyxl pandas xlsxwriter pywin32
```

macOS/Linux: `python3 -m pip install openpyxl pandas xlsxwriter` (skip `pywin32`).

## 2. Excel MCP server

```bash
claude mcp add excel --scope user -- uvx excel-mcp-server stdio
```

`--scope user` makes it available in every project. Drop it to scope to the current project.

**Warm the cache first.** On a cold `uvx` the initial launch downloads and builds the package,
and the 30-second MCP handshake times out. Run it once manually:

```bash
uvx excel-mcp-server stdio
```

Let it start, Ctrl-C, then check:

```bash
claude mcp list
```

You want `excel: uvx excel-mcp-server stdio - ✓ Connected`.

**The tools appear only after a restart.** MCP servers connect at session start, so a server
added mid-session shows as connected but its tools are not callable until you start a new
Claude Code session.

## 3. The hebrew-excel-rtl skill

**Personal** (all projects):

```bash
cp -r ./skills/hebrew-excel-rtl ~/.claude/skills/
```

PowerShell:

```powershell
Copy-Item -Recurse -Force ".\skills\hebrew-excel-rtl" "$env:USERPROFILE\.claude\skills\"
```

**Or project-scoped** — copy to `.claude/skills/` inside the project instead.

## 4. The supporting skills

```bash
npx -y skills add davila7/claude-code-templates --skill spreadsheet --agent claude-code
```
```bash
npx -y skills add jmsktm/claude-settings --skill "Inventory Manager" --agent claude-code
```

These install into `.claude/skills/` of the **current project**. Run them from the project you
want them in.

Anthropic's finance plugin marketplace, for `audit-xls` (formula tracing, hardcode detection,
balance checks):

```bash
claude plugin marketplace add anthropics/financial-services
```
```bash
claude plugin install model-builder@claude-for-financial-services
```

Other useful plugins from the same marketplace: `gl-reconciler`, `month-end-closer`,
`unit-economics`.

## 5. Verify

```bash
py demo/build_demo.py
```
```bash
py skills/hebrew-excel-rtl/scripts/verify_rtl.py "demo/דוח-כספי-לדוגמה.xlsx" --com
```

Expected: `24 passed, 0 warnings, 0 failed`.

## Already installed on this machine

If you are Gal, on this Windows box, all of the above is already done:

- `excel` MCP registered at user scope and connected
- `spreadsheet` + `inventory-manager` in `hebrew-excel-kit/.claude/skills/`
- `openpyxl 3.1.5`, `pandas 3.0.5`, `xlsxwriter 3.2.9`, `pywin32` installed

Outstanding: copy `hebrew-excel-rtl` to `~/.claude/skills/` if you want it in every project
rather than just this one, and restart Claude Code to pick up the MCP tools.

## Usage

Once installed, just describe what you want:

> תבנה לי דוח כספי בעברית עם גיליון מלאי ותזרים

The skill fires on Hebrew spreadsheet work automatically. To force it: *"use the
hebrew-excel-rtl skill"*.

For a messy inherited workbook:

> קח את הקובץ הזה, תנתח מה יש בו, ותבנה גרסה נקייה בעברית עם RTL נכון

The skill's workflow section tells the agent to map the source before changing anything, to
report defects before acting, and to build the clean version as a new file rather than
overwriting the original.
