#!/usr/bin/env bash
# Install the Hebrew Excel Kit on macOS / Linux.
#   ./install.sh [claude-code|codex|both]
set -euo pipefail

AGENT="${1:-claude-code}"
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

step() { printf '\n==> %s\n' "$1"; }
ok()   { printf '    OK  %s\n' "$1"; }
warn() { printf '    !   %s\n' "$1"; }

step "Python dependencies"
PY=$(command -v python3 || command -v python)
"$PY" -m pip install --quiet --disable-pip-version-check openpyxl pandas xlsxwriter
ok "openpyxl, pandas, xlsxwriter (pywin32 is Windows-only; --com verification unavailable here)"

step "Installing hebrew-excel-rtl skill"
case "$AGENT" in
  claude-code) TARGETS=("$HOME/.claude/skills") ;;
  codex)       TARGETS=("$HOME/.codex/skills") ;;
  both)        TARGETS=("$HOME/.claude/skills" "$HOME/.codex/skills") ;;
  *)           echo "usage: $0 [claude-code|codex|both]" >&2; exit 2 ;;
esac
for t in "${TARGETS[@]}"; do
  mkdir -p "$t"
  cp -r "$ROOT/skills/hebrew-excel-rtl" "$t/"
  ok "$t/hebrew-excel-rtl"
done

step "Installing the three supporting skills"
# Not vendored in this repo — they carry their own licences. Fetched from source.
if command -v git >/dev/null 2>&1; then
  TMP=$(mktemp -d)
  while IFS='|' read -r NAME REPO SUBPATH; do
    if git clone --depth 1 --quiet "https://github.com/$REPO" "$TMP/$NAME" 2>/dev/null \
       && [[ -d "$TMP/$NAME/$SUBPATH" ]]; then
      for t in "${TARGETS[@]}"; do cp -r "$TMP/$NAME/$SUBPATH" "$t/"; done
      ok "$NAME  (from $REPO)"
    else
      warn "$NAME: could not fetch from $REPO - upstream may have moved it"
    fi
  done <<'SKILLS'
spreadsheet|davila7/claude-code-templates|cli-tool/components/skills/document-processing/spreadsheet
inventory-manager|jmsktm/claude-settings|skills/inventory-manager
audit-xls|anthropics/financial-services|plugins/agent-plugins/model-builder/skills/audit-xls
SKILLS
  rm -rf "$TMP"
else
  warn "git not found - skipping. See INSTALL.md step 3 for manual commands."
fi

step "Excel MCP server"
if ! command -v uvx >/dev/null 2>&1; then
  warn "uvx not found. Install uv:  curl -LsSf https://astral.sh/uv/install.sh | sh"
  warn "Then re-run this script."
else
  echo "    warming the uvx cache (first run downloads the package)..."
  ( uvx excel-mcp-server stdio >/dev/null 2>&1 & echo $! > /tmp/.xlmcp.pid ) || true
  sleep 60
  kill "$(cat /tmp/.xlmcp.pid)" 2>/dev/null || true
  rm -f /tmp/.xlmcp.pid
  ok "cache warm"

  if [[ "$AGENT" == "claude-code" || "$AGENT" == "both" ]]; then
    if command -v claude >/dev/null 2>&1; then
      # Idempotent: re-running the installer must not fail on an existing server.
      if claude mcp list 2>/dev/null | grep -qE '^\s*excel:'; then
        ok "already registered with Claude Code"
      elif claude mcp add excel --scope user -- uvx excel-mcp-server stdio >/dev/null 2>&1; then
        ok "registered with Claude Code"
      else
        warn "could not register - see docs/INSTALL-claude-code.md"
      fi
      echo "        (restart Claude Code for the MCP tools to appear)"
    else
      warn "'claude' CLI not on PATH - see docs/INSTALL-claude-code.md"
    fi
  fi

  if [[ "$AGENT" == "codex" || "$AGENT" == "both" ]]; then
    CFG="$HOME/.codex/config.toml"
    mkdir -p "$(dirname "$CFG")"
    if [[ -f "$CFG" ]] && grep -q 'mcp_servers.excel' "$CFG"; then
      warn "mcp_servers.excel already present in $CFG - left untouched"
    else
      cat >> "$CFG" <<'TOML'

[mcp_servers.excel]
command = "uvx"
args = ["excel-mcp-server", "stdio"]
TOML
      ok "appended to $CFG"
    fi
  fi
fi

step "Verifying"
DEMO="$ROOT/demo/דוח-כספי-לדוגמה.xlsx"
[[ -f "$DEMO" ]] || ( cd "$ROOT/demo" && PYTHONUTF8=1 "$PY" build_demo.py )
PYTHONUTF8=1 "$PY" "$ROOT/skills/hebrew-excel-rtl/scripts/verify_rtl.py" "$DEMO"

printf '\nDone.\n'
printf '  Claude Code : restart it, then the MCP Excel tools appear\n'
printf '  Codex       : see docs/INSTALL-codex.md\n'
printf '  Desktop     : see docs/INSTALL-claude-desktop.md (needs a ZIP upload)\n'
