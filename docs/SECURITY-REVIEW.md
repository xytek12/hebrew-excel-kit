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
