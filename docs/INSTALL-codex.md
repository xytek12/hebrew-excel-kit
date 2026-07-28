# Install on OpenAI Codex

Codex loads skills natively — no plugin system, no marketplace. A skill is a directory with a
`SKILL.md` in the right folder.

## 1. Python prerequisites

```bash
py -m pip install openpyxl pandas xlsxwriter pywin32
```

On macOS or Linux use `python3 -m pip ...` and drop `pywin32` (it is Windows-only and only
powers the optional `--com` verification).

## 2. Install the skills

User-level Codex skills live in `$CODEX_HOME/skills/`, default `~/.codex/skills/`. Codex also
reads the cross-runtime path `~/.agents/skills/`, shared with Copilot CLI and Gemini CLI.

**Windows (PowerShell):**

```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.codex\skills" | Out-Null
Copy-Item -Recurse -Force ".\skills\hebrew-excel-rtl" "$env:USERPROFILE\.codex\skills\"
```

**macOS / Linux:**

```bash
mkdir -p ~/.codex/skills && cp -r ./skills/hebrew-excel-rtl ~/.codex/skills/
```

For the other three skills, clone and copy the same way:

```bash
git clone --depth 1 https://github.com/davila7/claude-code-templates /tmp/cct && cp -r /tmp/cct/cli-tool/components/skills/document-processing/spreadsheet ~/.codex/skills/
```
```bash
git clone --depth 1 https://github.com/jmsktm/claude-settings /tmp/jms && cp -r /tmp/jms/skills/inventory-manager ~/.codex/skills/
```
```bash
git clone --depth 1 https://github.com/anthropics/financial-services /tmp/afs && cp -r /tmp/afs/plugins/agent-plugins/model-builder/skills/audit-xls ~/.codex/skills/
```

## 3. Add the Excel MCP server

Edit `~/.codex/config.toml` (Windows: `%USERPROFILE%\.codex\config.toml`):

```toml
[mcp_servers.excel]
command = "uvx"
args = ["excel-mcp-server", "stdio"]
```

Keys are `mcp_servers.<id>.command` and `mcp_servers.<id>.args` per the Codex config
reference. Newer Codex builds also accept:

```bash
codex mcp add excel -- uvx excel-mcp-server stdio
```

If `uvx` is not on PATH, give the absolute path (`C:\Users\<you>\.local\bin\uvx.exe`) or
install uv from https://docs.astral.sh/uv/.

**First run is slow.** `uvx` downloads and builds the package on first launch and the MCP
handshake can time out. Warm the cache once before relying on it:

```bash
uvx excel-mcp-server stdio
```

Let it start, then Ctrl-C. Subsequent launches are instant.

## 4. Add project context

Codex reads `AGENTS.md` from the project root (and `~/.codex/AGENTS.md` globally). Add:

```markdown
## Hebrew Excel

Any .xlsx containing Hebrew must follow the `hebrew-excel-rtl` skill:
- `sheet_view.rightToLeft = True` on EVERY sheet, not just the first
- `Alignment(readingOrder=2)` on every cell holding text
- Shekel: `#,##0.00\ [$₪-40D]`   Dates: `DD/MM/YYYY`, written as real date objects
- Verify before delivering: `py skills/hebrew-excel-rtl/scripts/verify_rtl.py <file> --com`
```

## 5. Verify

```bash
py demo/build_demo.py
```
```bash
py skills/hebrew-excel-rtl/scripts/verify_rtl.py "demo/דוח-כספי-לדוגמה.xlsx"
```

Expect `0 failed`. Add `--com` on Windows with Excel installed for the full check.

## Codex-specific notes

- Codex reads files through `shell` (`cat`, `rg`) and edits through `apply_patch`. The skill's
  scripts are plain Python and work unchanged.
- Skills load automatically from their description. To force one, name it: *"use the
  hebrew-excel-rtl skill"*.
- If you also want parallel subagents, enable them in `~/.codex/config.toml`:

  ```toml
  [features]
  multi_agent = true
  ```
