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

## Step 3 — the three supporting skills

Not vendored here — they carry their own licences. Fetch from source.

**Claude Code** (installs into the current project's `.claude/skills/`):

```bash
npx -y skills add davila7/claude-code-templates --skill spreadsheet --agent claude-code
```
```bash
npx -y skills add jmsktm/claude-settings --skill "Inventory Manager" --agent claude-code
```
```bash
claude plugin marketplace add anthropics/financial-services
```
```bash
claude plugin install model-builder@claude-for-financial-services
```

**Codex, or any agent without the `skills` CLI** — clone and copy:

```bash
git clone --depth 1 https://github.com/davila7/claude-code-templates /tmp/cct && cp -r /tmp/cct/cli-tool/components/skills/document-processing/spreadsheet ~/.codex/skills/
```
```bash
git clone --depth 1 https://github.com/jmsktm/claude-settings /tmp/jms && cp -r /tmp/jms/skills/inventory-manager ~/.codex/skills/
```
```bash
git clone --depth 1 https://github.com/anthropics/financial-services /tmp/afs && cp -r /tmp/afs/plugins/agent-plugins/model-builder/skills/audit-xls ~/.codex/skills/
```

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

## Step 5 — verify the install

```bash
py demo/build_demo.py
```
```bash
py skills/hebrew-excel-rtl/scripts/verify_rtl.py "demo/דוח-כספי-לדוגמה.xlsx" --com
```

Expected: a per-sheet report ending in `8 sheets — 96 passed, 0 warnings, 0 failed`.
Drop `--com` off Windows; the static checks still run.

If `verify_rtl.py` reports failures on a freshly cloned repo, something in the install is
wrong — do not ignore it.

## Step 6 — tell the user what to do next

Report which of the four skills installed, whether the MCP server connected, and that they
must restart the agent for the MCP tools to appear. Then they can just ask, in Hebrew:

> תבנה לי דוח כספי בעברית עם מלאי לפי חנות ולוח בקרה

or point you at a messy workbook:

> קח את הקובץ הזה ותבנה ממנו גרסה נקייה בעברית עם RTL נכון
