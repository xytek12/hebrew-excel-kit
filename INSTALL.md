# INSTALL — instructions for the agent

**If you are Claude Code, Codex, or another coding agent and the user pointed you at this
repository and said "install everything": this file is your instruction set. Follow it top to
bottom. You do not need to ask which parts to install — install all of it.**

Machine-readable version of the same thing: [`manifest.json`](manifest.json).
Source audit of every third-party component: [`docs/SECURITY-REVIEW.md`](docs/SECURITY-REVIEW.md).

---

## The short version

Everything below is automated. Run one of these from the repo root and stop reading:

**Windows**
```powershell
pwsh -File install.ps1 -Agent both
```

**macOS / Linux**
```bash
bash install.sh both
```

`-Agent` / the positional argument takes `claude-code`, `codex`, or `both`.

If the script fails, or you are on a platform it does not cover, do the steps manually below.

---

## Step 1 — Python packages

```bash
py -m pip install openpyxl pandas xlsxwriter pywin32
```

macOS/Linux: use `python3 -m pip` and drop `pywin32` — it is Windows-only and only powers the
optional real-Excel verification.

## Step 2 — the hebrew-excel-rtl skill

Copy `skills/hebrew-excel-rtl` into the agent's skills directory.

| Agent | Directory |
|---|---|
| Claude Code | `~/.claude/skills/` |
| Codex | `~/.codex/skills/` (or `$CODEX_HOME/skills/`) |
| Copilot CLI / Gemini CLI | `~/.agents/skills/` |
| Claude Desktop | ZIP it, then Settings → Capabilities → Skills → **+** → Upload a skill |

```bash
mkdir -p ~/.claude/skills && cp -r skills/hebrew-excel-rtl ~/.claude/skills/
```

```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.claude\skills" | Out-Null
Copy-Item -Recurse -Force ".\skills\hebrew-excel-rtl" "$env:USERPROFILE\.claude\skills\"
```

## Step 3 — the ten supporting skills

Not vendored here — they carry their own licences. Fetch from source and copy into the
same skills directory as step 2. All ten are listed with paths in `manifest.json`; each
was security-reviewed before being listed (`docs/SECURITY-REVIEW.md`).

| Skill | From | What it adds |
|---|---|---|
| `xlsx` | `anthropics/skills` | Official spreadsheet skill: zero-formula-errors gate, messy-file intake |
| `excel-hygiene` | `BayramAnnakov/excel-hygiene` | Independent error checker + spreadsheet-error taxonomy |
| `audit-xls` | `anthropics/financial-services` | Formula tracing, hardcode detection |
| `clean-data-xls` | `anthropics/financial-services` | Normalize messy raw data before analysis |
| `3-statement-model` | `anthropics/financial-services` | IS/BS/CF model structure |
| `dcf-model` | `anthropics/financial-services` | DCF valuation + validator script |
| `comps-analysis` | `anthropics/financial-services` | Peer benchmarking with multiples |
| `xlsx-author` | `anthropics/financial-services` | Headless .xlsx production discipline |
| `spreadsheet` | `davila7/claude-code-templates` | openpyxl/pandas conventions |
| `inventory-manager` | `jmsktm/claude-settings` | Reorder points, forecasting, ABC |

Clone each repo **once** and copy the paths from `manifest.json`. Two things to know:

- **Windows:** clone `anthropics/skills` with `git -c core.longpaths=true` — it holds
  schema files whose paths overflow the 260-character limit and break the clone.
- **Do NOT install `financial-analysis` as a plugin.** The plugin form registers 12
  remote MCP connectors (FactSet, Moody's, PitchBook…). Copy the six skills instead —
  same content, no external endpoints.

```bash
git -c core.longpaths=true clone --depth 1 https://github.com/anthropics/skills /tmp/ansk && cp -r /tmp/ansk/skills/xlsx ~/.claude/skills/
```
```bash
git clone --depth 1 https://github.com/BayramAnnakov/excel-hygiene ~/.claude/skills/excel-hygiene && rm -rf ~/.claude/skills/excel-hygiene/.git
```
```bash
git clone --depth 1 https://github.com/anthropics/financial-services /tmp/afs && for s in audit-xls clean-data-xls 3-statement-model dcf-model comps-analysis xlsx-author; do cp -r "/tmp/afs/plugins/vertical-plugins/financial-analysis/skills/$s" ~/.claude/skills/; done
```
```bash
git clone --depth 1 https://github.com/davila7/claude-code-templates /tmp/cct && cp -r /tmp/cct/cli-tool/components/skills/document-processing/spreadsheet ~/.claude/skills/
```
```bash
git clone --depth 1 https://github.com/jmsktm/claude-settings /tmp/jms && cp -r /tmp/jms/skills/inventory-manager ~/.claude/skills/
```

For Codex repeat with `~/.codex/skills/`.

## Step 4 — the Excel MCP server

`haris-musa/excel-mcp-server` — 25 tools, does **not** require Microsoft Excel to be
installed. Needs [uv](https://docs.astral.sh/uv/).

**Warm the cache first.** On a cold `uvx` the first launch downloads and builds the package
and the 30-second MCP handshake times out:

```bash
uvx excel-mcp-server stdio
```

Let it start, then Ctrl-C.

**Claude Code:**
```bash
claude mcp add excel --scope user -- uvx excel-mcp-server stdio
```

**Codex** — add to `~/.codex/config.toml`:
```toml
[mcp_servers.excel]
command = "uvx"
args = ["excel-mcp-server", "stdio"]
```

**Claude Desktop** — add to `claude_desktop_config.json`
(`%APPDATA%\Claude\` on Windows, `~/Library/Application Support/Claude/` on macOS).
On Windows give `uvx` its absolute path; Desktop does not inherit your shell PATH:
```json
{
  "mcpServers": {
    "excel": {
      "command": "C:\\Users\\YOURNAME\\.local\\bin\\uvx.exe",
      "args": ["excel-mcp-server", "stdio"]
    }
  }
}
```

**Restart the agent afterwards.** MCP servers connect at session start — a server added
mid-session reports as connected but its tools are not callable until a new session.

## Step 4b — the live-Excel MCP server (Windows + Excel only)

`sbroenne/mcp-server-excel` drives the **real Excel engine** via COM — live recalculation,
PivotTables, Power Query, DAX, and range screenshots the agent can look at. It covers the
two things the file-level server cannot: making formulas actually evaluate, and proving
what Excel really renders.

Skip this step on macOS/Linux or when Excel is not installed — everything else still works.

Download `ExcelMcp-MCP-Server-<version>-windows.zip` from
[the latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest), extract
`mcp-excel.exe` to `%USERPROFILE%\tools\excel-mcp\`, then:

**Claude Code:**
```powershell
claude mcp add excel-live --scope user -- "$env:USERPROFILE\tools\excel-mcp\mcp-excel.exe"
```

**Codex** — add to `~/.codex/config.toml`:
```toml
[mcp_servers.excel-live]
command = "C:\\Users\\YOURNAME\\tools\\excel-mcp\\mcp-excel.exe"
```

**Claude Desktop** — download the `.mcpb` file from the same release page and double-click it.

> Disclosure (from the security review): release builds send opt-out anonymous telemetry —
> tool names, duration, a hashed machine id. Never file paths, cell values, or formulas.
> Close all Excel files before using it; it needs exclusive access during automation.

## Step 5 — verify the install

```bash
py demo/build_demo.py
```

On Windows with Excel, also run the finishing pass — it builds the real PivotTables and
saves cached formula values (openpyxl can do neither):

```bash
py demo/com_finish.py
```

```bash
py skills/hebrew-excel-rtl/scripts/verify_rtl.py "demo/דוח-כספי-לדוגמה.xlsx" --com
```

Expected after both passes: `10 sheets — 119 passed, 0 warnings, 0 failed`.
Without Excel (no `com_finish.py`, no `--com`): 9 sheets, static checks only.

**Never re-save the finished workbook with openpyxl** — an openpyxl load/save round-trip
silently deletes the charts and the PivotTables.

If `verify_rtl.py` reports failures on a freshly cloned repo, something in the install is
wrong — do not ignore it.

## Step 6 — tell the user what to do next

Report which of the eleven skills installed, which of the two MCP servers connected, and
that they must restart the agent for the MCP tools to appear. Then they can just ask, in Hebrew:

> תבנה לי דוח כספי בעברית עם מלאי לפי חנות ולוח בקרה

or point you at a messy workbook:

> קח את הקובץ הזה ותבנה ממנו גרסה נקייה בעברית עם RTL נכון
