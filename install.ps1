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
                & claude mcp add excel --scope user -- uvx excel-mcp-server stdio 2>&1 | Out-Null
                Ok "registered with Claude Code (restart Claude Code to load the tools)"
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

Step "Verifying"
$demo = Join-Path $root "demo\דוח-כספי-לדוגמה.xlsx"
if (-not (Test-Path $demo)) {
    Push-Location "$root\demo"; & py build_demo.py; Pop-Location
}
$env:PYTHONUTF8 = 1
& py "$root\skills\hebrew-excel-rtl\scripts\verify_rtl.py" $demo

Write-Host "`nDone." -ForegroundColor Green
Write-Host "  Claude Code : restart it, then MCP Excel tools appear"
Write-Host "  Codex       : see docs\INSTALL-codex.md"
Write-Host "  Desktop     : see docs\INSTALL-claude-desktop.md (needs a ZIP upload)"
