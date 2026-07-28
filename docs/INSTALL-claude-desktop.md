# Install on Claude Desktop

Claude Desktop handles skills and MCP servers differently from Claude Code. Skills are
uploaded through the UI as ZIP files; MCP servers go in a JSON config file.

## 1. Excel MCP server

Open the config file:

**Windows** — `%APPDATA%\Roaming\Claude\claude_desktop_config.json`

```powershell
notepad "$env:APPDATA\Claude\claude_desktop_config.json"
```

**macOS** — `~/Library/Application Support/Claude/claude_desktop_config.json`

```bash
open -e ~/Library/Application\ Support/Claude/claude_desktop_config.json
```

You can also reach it from **Settings → Developer → Edit Config**.

Add the server. If the file already has an `mcpServers` block, add `excel` inside it rather
than replacing the block:

```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["excel-mcp-server", "stdio"]
    }
  }
}
```

**Windows usually needs the absolute path** — Claude Desktop does not always inherit your
shell PATH:

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

Note the doubled backslashes — required in JSON.

**Warm the cache before first launch**, or the server will time out and show as failed:

```bash
uvx excel-mcp-server stdio
```

Let it start, then Ctrl-C.

**Restart Claude Desktop completely** — quit from the tray/menu bar, not just close the
window. Then check the tools icon in the chat box; you should see the Excel tools listed.

## 2. Skills

Claude Desktop takes skills as ZIP uploads, one per skill.

**Zip it (Windows PowerShell):**

```powershell
Compress-Archive -Path ".\skills\hebrew-excel-rtl\*" -DestinationPath ".\hebrew-excel-rtl.zip" -Force
```

**macOS / Linux:**

```bash
cd skills && zip -r ../hebrew-excel-rtl.zip hebrew-excel-rtl && cd ..
```

Then in Claude Desktop:

1. **Settings → Capabilities → Skills**
2. Click **+** and choose **Upload a skill**
3. Select `hebrew-excel-rtl.zip`
4. Start a **new conversation** — skills do not load into an existing one

The ZIP must contain `SKILL.md` at its root alongside `reference/` and `scripts/`. If you zip
the parent folder instead of its contents on Windows, the upload fails validation.

## 3. Important limitation

Claude Desktop cannot run the Python scripts in this skill. It has no shell.

That means `verify_rtl.py` and `build_demo.py` will not run there. Desktop can use the skill's
*knowledge* — formats, rules, pitfalls — and can drive the Excel MCP server to read and write
workbooks. But the verification step needs Claude Code or Codex.

**Practical split:**
- **Claude Desktop** — good for asking questions about a workbook, reading data, making
  targeted edits through the MCP tools.
- **Claude Code / Codex** — required for generating a workbook from a script and verifying it.

For your friend's actual job (rebuilding a messy workbook properly), use Claude Code or Codex.

## 4. Verify

Start a new conversation and ask:

> What Excel tools do you have available?

You should get the 25 tools from `excel-mcp-server` — `read_data_from_excel`,
`write_data_to_excel`, `create_pivot_table`, `create_chart`, `format_range`, and so on.

Then point it at the demo file:

> Read the מלאי sheet from <path>\demo\דוח-כספי-לדוגמה.xlsx and tell me which items are below their reorder point

Correct answer: one item, `TS-1042` (חולצת טי כותנה, L, שחור) — stock 92 against a reorder
point of 131.
