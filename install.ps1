<#
.SYNOPSIS
    Install the Hebrew Excel Kit on Windows.
.PARAMETER Agent
    claude-code (default) | codex | both
.PARAMETER SkipMcp
    Skip registering the Excel MCP server.
#>
[CmdletBinding()]
param(
    [ValidateSet('claude-code', 'codex', 'both')]
    [string]$Agent = 'claude-code',
    [switch]$SkipMcp
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

function Step($msg) { Write-Host "`n==> $msg" -ForegroundColor Cyan }
function Ok($msg)   { Write-Host "    OK  $msg" -ForegroundColor Green }
function Warn($msg) { Write-Host "    !   $msg" -ForegroundColor Yellow }

Step "Python dependencies"
$py = (Get-Command py -ErrorAction SilentlyContinue)
if (-not $py) { throw "Python launcher 'py' not found. Install Python 3.10+ from python.org." }
& py -m pip install --quiet --disable-pip-version-check openpyxl pandas xlsxwriter pywin32
if ($LASTEXITCODE -ne 0) { throw "pip install failed" }
Ok "openpyxl, pandas, xlsxwriter, pywin32"

Step "Installing hebrew-excel-rtl skill"
$targets = @()
if ($Agent -in 'claude-code', 'both') { $targets += "$env:USERPROFILE\.claude\skills" }
if ($Agent -in 'codex', 'both')       { $targets += "$env:USERPROFILE\.codex\skills" }
foreach ($t in $targets) {
    New-Item -ItemType Directory -Force -Path $t | Out-Null
    Copy-Item -Recurse -Force "$root\skills\hebrew-excel-rtl" $t
    Ok "$t\hebrew-excel-rtl"
}

Step "Installing the supporting skills"
# Not vendored in this repo - they carry their own licences. Fetched from source,
# one clone per repo. Each was security-reviewed before being listed here - see
# docs/SECURITY-REVIEW.md. -c core.longpaths=true matters: anthropics/skills holds
# schema files whose paths overflow the Windows 260-char limit and break the clone.
$sources = @(
    @{ repo = 'anthropics/skills'; skills = @(
        @{ name = 'xlsx'; path = 'skills/xlsx' }) },
    @{ repo = 'BayramAnnakov/excel-hygiene'; skills = @(
        @{ name = 'excel-hygiene'; path = '.' }) },
    @{ repo = 'anthropics/financial-services'; skills = @(
        @{ name = 'audit-xls';        path = 'plugins/vertical-plugins/financial-analysis/skills/audit-xls' },
        @{ name = 'clean-data-xls';   path = 'plugins/vertical-plugins/financial-analysis/skills/clean-data-xls' },
        @{ name = '3-statement-model'; path = 'plugins/vertical-plugins/financial-analysis/skills/3-statement-model' },
        @{ name = 'dcf-model';        path = 'plugins/vertical-plugins/financial-analysis/skills/dcf-model' },
        @{ name = 'comps-analysis';   path = 'plugins/vertical-plugins/financial-analysis/skills/comps-analysis' },
        @{ name = 'xlsx-author';      path = 'plugins/vertical-plugins/financial-analysis/skills/xlsx-author' }) },
    @{ repo = 'davila7/claude-code-templates'; skills = @(
        @{ name = 'spreadsheet'; path = 'cli-tool/components/skills/document-processing/spreadsheet' }) },
    @{ repo = 'jmsktm/claude-settings'; skills = @(
        @{ name = 'inventory-manager'; path = 'skills/inventory-manager' }) }
)
if (Get-Command git -ErrorAction SilentlyContinue) {
    $tmp = Join-Path $env:TEMP "hek-$(Get-Random)"
    foreach ($srcRepo in $sources) {
        $dest = Join-Path $tmp ($srcRepo.repo -replace '/', '_')
        & git -c core.longpaths=true clone --depth 1 --quiet "https://github.com/$($srcRepo.repo)" $dest 2>&1 | Out-Null
        foreach ($s in $srcRepo.skills) {
            $src = if ($s.path -eq '.') { $dest } else { Join-Path $dest ($s.path -replace '/', '\') }
            if (Test-Path $src) {
                foreach ($t in $targets) {
                    $skillDest = Join-Path $t $s.name
                    if (Test-Path $skillDest) { Remove-Item $skillDest -Recurse -Force }
                    Copy-Item -Recurse -Force $src $skillDest
                    Remove-Item (Join-Path $skillDest '.git') -Recurse -Force -ErrorAction SilentlyContinue
                }
                Ok "$($s.name)  (from $($srcRepo.repo))"
            }
            else { Warn "$($s.name): path not found in $($srcRepo.repo) - upstream may have moved it" }
        }
    }
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
else { Warn "git not found - skipping. See INSTALL.md step 3 for manual commands." }

if (-not $SkipMcp) {
    Step "Excel MCP server"
    $uvx = (Get-Command uvx -ErrorAction SilentlyContinue)
    if (-not $uvx) {
        Warn "uvx not found. Install uv:  powershell -c ""irm https://astral.sh/uv/install.ps1 | iex"""
        Warn "Then re-run this script."
    }
    else {
        Write-Host "    warming the uvx cache (first run downloads the package)..."
        $job = Start-Job -ScriptBlock { param($u) & $u excel-mcp-server stdio } -ArgumentList $uvx.Source
        Start-Sleep -Seconds 75
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        Ok "cache warm"

        if ($Agent -in 'claude-code', 'both') {
            if (Get-Command claude -ErrorAction SilentlyContinue) {
                # Idempotent: re-running the installer must not fail on an existing server.
                $existing = (& claude mcp list 2>&1 | Out-String)
                if ($existing -match '(?m)^\s*excel:') {
                    Ok "already registered with Claude Code"
                }
                else {
                    & claude mcp add excel --scope user -- uvx excel-mcp-server stdio 2>&1 | Out-Null
                    if ($LASTEXITCODE -eq 0) { Ok "registered with Claude Code" }
                    else { Warn "could not register - see docs/INSTALL-claude-code.md" }
                }
                Write-Host "        (restart Claude Code for the MCP tools to appear)"
            }
            else { Warn "'claude' CLI not on PATH - see docs/INSTALL-claude-code.md" }
        }
        if ($Agent -in 'codex', 'both') {
            $cfg = "$env:USERPROFILE\.codex\config.toml"
            New-Item -ItemType Directory -Force -Path (Split-Path $cfg) | Out-Null
            $block = @"

[mcp_servers.excel]
command = "$($uvx.Source -replace '\\','\\')"
args = ["excel-mcp-server", "stdio"]
"@
            if ((Test-Path $cfg) -and (Select-String -Path $cfg -Pattern 'mcp_servers.excel' -Quiet)) {
                Warn "mcp_servers.excel already present in $cfg - left untouched"
            }
            else {
                Add-Content -Path $cfg -Value $block -Encoding utf8
                Ok "appended to $cfg"
            }
        }
    }
}

if (-not $SkipMcp) {
    Step "Live-Excel MCP server (excel-live - drives real Excel via COM)"
    # sbroenne/mcp-server-excel: recalculation, PivotTables, Power Query, screenshots.
    # Windows-only and needs Excel installed. Disclosure: the release build carries
    # opt-out anonymous telemetry (tool names + hashed machine id; never file paths,
    # cell values or formulas) - reviewed in docs/SECURITY-REVIEW.md.
    if (-not (Test-Path 'Registry::HKEY_CLASSES_ROOT\Excel.Application')) {
        Warn "Microsoft Excel not detected - skipping excel-live (the file-level MCP above still works)"
    }
    else {
        $toolDir = "$env:USERPROFILE\tools\excel-mcp"
        $exe = Join-Path $toolDir 'mcp-excel.exe'
        if (-not (Test-Path $exe)) {
            New-Item -ItemType Directory -Force -Path $toolDir | Out-Null
            try {
                $rel = Invoke-RestMethod 'https://api.github.com/repos/sbroenne/mcp-server-excel/releases/latest'
                $asset = $rel.assets | Where-Object { $_.name -like 'ExcelMcp-MCP-Server-*-windows.zip' } | Select-Object -First 1
                $zip = Join-Path $toolDir $asset.name
                Invoke-WebRequest $asset.browser_download_url -OutFile $zip
                Expand-Archive -Path $zip -DestinationPath $toolDir -Force
                Ok "downloaded $($rel.tag_name) to $toolDir"
            }
            catch { Warn "download failed: $($_.Exception.Message) - install manually from github.com/sbroenne/mcp-server-excel/releases" }
        }
        else { Ok "already present: $exe" }

        if (Test-Path $exe) {
            if ($Agent -in 'claude-code', 'both') {
                if (Get-Command claude -ErrorAction SilentlyContinue) {
                    $existing = (& claude mcp list 2>&1 | Out-String)
                    if ($existing -match '(?m)^\s*excel-live:') { Ok "already registered with Claude Code" }
                    else {
                        & claude mcp add excel-live --scope user -- $exe 2>&1 | Out-Null
                        if ($LASTEXITCODE -eq 0) { Ok "registered with Claude Code as excel-live" }
                        else { Warn "could not register excel-live" }
                    }
                }
            }
            if ($Agent -in 'codex', 'both') {
                $cfg = "$env:USERPROFILE\.codex\config.toml"
                New-Item -ItemType Directory -Force -Path (Split-Path $cfg) | Out-Null
                if ((Test-Path $cfg) -and (Select-String -Path $cfg -Pattern 'mcp_servers.excel-live' -Quiet)) {
                    Warn "mcp_servers.excel-live already present in $cfg - left untouched"
                }
                else {
                    $block = "`n[mcp_servers.excel-live]`ncommand = `"$($exe -replace '\\','\\')`"`n"
                    Add-Content -Path $cfg -Value $block -Encoding utf8
                    Ok "appended to $cfg"
                }
            }
        }
    }
}

Step "Verifying"
$env:PYTHONUTF8 = 1
# The demo workbook has a Hebrew filename. Glob for it rather than writing it as a
# literal: Windows PowerShell 5.1 reads .ps1 files as ANSI unless they carry a UTF-8
# BOM, and a mangled Hebrew literal can swallow the closing quote and break the parse.
$demo = Get-ChildItem (Join-Path $root 'demo') -Filter *.xlsx |
        Select-Object -First 1 -ExpandProperty FullName
if (-not $demo) {
    Push-Location (Join-Path $root 'demo'); & py build_demo.py; Pop-Location
    $demo = Get-ChildItem (Join-Path $root 'demo') -Filter *.xlsx |
            Select-Object -First 1 -ExpandProperty FullName
}
$hasExcel = Test-Path 'Registry::HKEY_CLASSES_ROOT\Excel.Application'
if ($hasExcel) {
    # Adds the real PivotTables and saves cached formula values - openpyxl can do neither.
    Push-Location (Join-Path $root 'demo')
    try { & py com_finish.py } catch { Warn "com_finish failed: $($_.Exception.Message)" }
    Pop-Location
    & py (Join-Path $root 'skills\hebrew-excel-rtl\scripts\verify_rtl.py') $demo --com
}
else {
    & py (Join-Path $root 'skills\hebrew-excel-rtl\scripts\verify_rtl.py') $demo
}

Write-Host "`nDone." -ForegroundColor Green
Write-Host '  Claude Code : restart it, then the MCP Excel tools appear'
Write-Host '  Codex       : see docs/INSTALL-codex.md'
Write-Host '  Desktop     : see docs/INSTALL-claude-desktop.md (needs a ZIP upload)'
