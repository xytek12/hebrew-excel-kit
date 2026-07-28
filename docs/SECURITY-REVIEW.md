# Security review

Every third-party component was reviewed **before** it was installed. Sources were downloaded
without executing them, then read and pattern-scanned.

Reviewed 2026-07-28.

---

## excel-mcp-server (haris-musa) — PASS

- PyPI `excel-mcp-server` 0.1.8, runtime reports `excel-mcp 1.29.0`
- Author email in package metadata matches the GitHub owner
- MIT, 4.1k★, ~2,900 lines of Python across 13 modules

**Method:** `pip download --no-deps --no-binary :all:` to get the sdist without running any
install hooks, then extracted and audited the source.

**Dependencies** — all mainstream, no unexpected transitive pulls:
`mcp[cli]>=1.10.1`, `fastmcp>=2.0.0,<3.0.0`, `openpyxl>=3.1.5`, `typer>=0.16.0`

**Scanned for** `subprocess`, `os.system`, `eval(`, `exec(`, `__import__`, `pickle`,
`marshal`, `base64`, `requests.`, `urllib`, `httpx`, `socket`, `urlopen`, `Popen`,
`shutil.rmtree`, `os.remove`, `os.unlink`, `getenv`, `environ`, `telemetry`, `analytics`,
`sentry`.

**Total hits: 4**, all benign — three `os.environ.get` calls for `FASTMCP_HOST`,
`FASTMCP_PORT`, and `EXCEL_FILES_PATH`. No network egress, no process spawning, no dynamic
code execution, no telemetry, no filesystem deletion.

**Path traversal:** guarded. In SSE and HTTP modes `get_excel_path()` resolves with
`os.path.realpath` and rejects any candidate that escapes `EXCEL_FILES_PATH`.

**Two notes, neither blocking:**

1. In **stdio** mode `EXCEL_FILES_PATH` is unset and the server takes absolute paths — full
   filesystem access. That is by design and matches every other local agent tool. It is the
   mode the installers configure.
2. The **HTTP** transport defaults to `host=0.0.0.0`, binding all interfaces. Do not run the
   HTTP or SSE transports on an untrusted network. Not an issue for stdio.

---

## spreadsheet skill (davila7/claude-code-templates) — PASS

- Source: `cli-tool/components/skills/document-processing/spreadsheet/SKILL.md`
- 211 lines, prose instructions only, no bundled executables
- Frontmatter declares `author: openai` — it is a port of OpenAI's spreadsheet skill, which
  also makes it a natural fit for Codex

**Findings:** no network calls, no credential access, no prompt-injection patterns, no
instructions to read `.env` or exfiltrate anything.

Only flagged lines are documented dependency installs — `pip install openpyxl pandas`,
`pip install matplotlib`, and a Linux-only `apt-get install libreoffice` for the optional PDF
render step. All are visible, conventional, and irrelevant on Windows.

---

## inventory-manager skill (jmsktm/claude-settings) — PASS

- Source: `skills/inventory-manager/SKILL.md`
- 457 lines, prose only, no scripts
- Declares `author: "ID8Labs"`, version 1.0.0

**Findings: zero hits** on the entire scan pattern. No URLs, no commands, no credential
references. Content is domain guidance — ABC classification, demand forecasting, reorder
points, safety stock, supplier management, KPIs.

---

## audit-xls skill (anthropics/financial-services) — PASS

- First-party Anthropic repository
- Source: `plugins/agent-plugins/model-builder/skills/audit-xls/SKILL.md`, 248 lines
- **Findings: zero hits.** Prose only.

---

---

# Round 2 — reviewed 2026-07-29

Four components added after a marketplace-wide research sweep. Same method: source fetched
without executing, read and pattern-scanned before install.

## xlsx skill (anthropics/skills) — PASS

- First-party Anthropic, 53 files (SKILL.md, recalc.py, office/ validators + ISO schemas)
- Scanned hits: `subprocess` only ever launches **local** `soffice` (LibreOffice) for
  recalculation; the `socket` code in `office/soffice.py` is a Linux sandbox shim for
  LibreOffice's local Unix socket — no network egress anywhere
- On Windows without LibreOffice `recalc.py` fails gracefully (caught `FileNotFoundError`);
  use `demo/com_finish.py`-style COM recalculation instead

## excel-hygiene (BayramAnnakov) — PASS

- Apache-2.0, 8 files total; `check_excel.py` imports only `sys`, `re`, `openpyxl`
- `requirements.txt`: `openpyxl>=3.1` and nothing else
- No network, no subprocess, no dynamic code, no env access. The bundled
  `examples/broken.xlsx` is a test fixture (xlsx cannot carry macros)

## financial-analysis skills (anthropics/financial-services) — PASS with a decision

- First-party Anthropic. The six copied skills (`audit-xls`, `clean-data-xls`,
  `3-statement-model`, `dcf-model`, `comps-analysis`, `xlsx-author`) are markdown plus one
  Python script, `validate_dcf.py`, which imports only stdlib + openpyxl
- `dcf-model/requirements.txt` lists `requests` but the script never imports it — do not
  install from that file
- **Decision: skills copied individually; the plugin form was NOT installed.** The plugin's
  `.mcp.json` registers 12 remote MCP connectors (FactSet, Moody's, PitchBook, Egnyte, Box…).
  Copying the skills delivers the same content with zero external endpoints
- The plugin's `hooks/hooks.json` is empty — no event-triggered code

## mcp-server-excel (sbroenne) — PASS with disclosure

- MIT, 408★, C#/.NET, actively maintained (pushed the day before review)
- Distributed as a signed GitHub release binary (`mcp-excel.exe`); all Excel work happens
  locally through the COM API — workbook content never leaves the machine
- **Telemetry disclosure:** release builds send opt-out anonymous usage telemetry to Azure
  Application Insights — tool name, action, duration, success/failure, a SHA256-hashed
  machine id, and redacted exception traces. The documented and code-verified exclusions:
  no file paths, no file names, no cell values, no formulas, no credentials
  (`SensitiveDataRedactor.cs` strips paths, connection strings, and emails). The connection
  string is compiled in at build time; there is **no runtime kill-switch** short of
  building from source or blocking the endpoint at the firewall
- One outbound HTTPS call to `api.github.com` checks for a newer release (version banner
  only)
- Accepted because the hard requirement — spreadsheet data stays local — holds. If the
  usage telemetry itself is unacceptable, skip `install.ps1`'s excel-live step

## Rejected during the research sweep

- **xlsx-for-ai** (senoff) — 39 tools, free tier, but it is a thin npm client over a hosted
  API (`api.xlsx-for-ai.dev`): every spreadsheet operation executes server-side. Financial
  data leaving the machine is disqualifying, regardless of convenience
- **excel-automation** (daymade) — clean skill, but built around macOS AppleScript; wrong
  platform for this kit

## Scan pattern used

```
curl|wget|http://|https://|api_key|token|\.env|password|secret|credential|POST |upload|
send to|base64|eval|exec|subprocess|os\.system|rm -rf|Remove-Item|chmod|sudo|
pip install|npm install|ignore previous|disregard
```

Case-insensitive, across every file in each package.

## What this review does not cover

- **Future versions.** This audited the versions listed above on the date given. `uvx` resolves
  `excel-mcp-server` at run time, so a later release will not have been reviewed. Pin the
  version if that matters to you.
- **Transitive dependency source.** `mcp`, `fastmcp`, `openpyxl` and `typer` were accepted on
  reputation rather than read line by line.
- **Runtime behaviour.** This is source review. The server was additionally smoke-tested over
  stdio — it exposed 25 tools and read the demo workbook correctly — but was not run under a
  network monitor.
