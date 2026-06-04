<#
.SYNOPSIS
    Interactive installer for the Dataroom Prep Claude Code skill.
.DESCRIPTION
    Copies the skill files into ~/.claude/skills/dataroom-prep, installs
    Python dependencies, and walks you through filling in config.yaml.

    Prerequisites:
    - Windows 10/11 (macOS works too — use install.sh)
    - Python 3.9+ on PATH
    - Claude Code installed (https://claude.com/claude-code)

.EXAMPLE
    Right-click install.ps1 -> "Run with PowerShell"
.NOTES
    Uninstall : Remove-Item -Recurse -Force "$env:USERPROFILE\.claude\skills\dataroom-prep"
#>

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Dataroom Prep - Interactive installer" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Repository: https://github.com/Hectelion-SA/claude-dataroom-prep"
Write-Host ""

# ============================================================
# 1) Check Python + dependencies
# ============================================================
Write-Host "[1/5] Checking Python..." -ForegroundColor Yellow
$pyCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pyCmd) {
    Write-Host "  ERROR: Python not found on PATH." -ForegroundColor Red
    Write-Host "  Install Python 3.9+ from https://python.org (check 'Add Python to PATH') and retry."
    exit 1
}
$pyVersion = & python --version 2>&1
Write-Host "  Found: $pyVersion ($($pyCmd.Source))" -ForegroundColor Green

Write-Host "[2/5] Installing Python packages (pypdf, python-docx, openpyxl, pyyaml)..." -ForegroundColor Yellow
& python -m pip install --quiet pypdf python-docx openpyxl pyyaml
if ($LASTEXITCODE -ne 0) {
    Write-Host "  ERROR: pip install failed." -ForegroundColor Red
    Write-Host "  Run manually: python -m pip install pypdf python-docx openpyxl pyyaml"
    exit 1
}
Write-Host "  All packages OK." -ForegroundColor Green

# ============================================================
# 3) Copy skill files into ~/.claude/skills/dataroom-prep/
# ============================================================
Write-Host "[3/5] Installing skill into Claude Code skills folder..." -ForegroundColor Yellow

$skillsRoot = Join-Path $env:USERPROFILE ".claude\skills"
$skillDest = Join-Path $skillsRoot "dataroom-prep"

if (-not (Test-Path $skillsRoot)) {
    New-Item -ItemType Directory -Path $skillsRoot -Force | Out-Null
    Write-Host "  Created: $skillsRoot" -ForegroundColor Gray
}

if (Test-Path $skillDest) {
    $ans = Read-Host "  Folder $skillDest already exists. Overwrite skill files? (y/N)"
    if ($ans -ne "y" -and $ans -ne "Y") {
        Write-Host "  Aborted by user." -ForegroundColor Yellow
        exit 0
    }
    Remove-Item -Recurse -Force $skillDest
}
New-Item -ItemType Directory -Path $skillDest -Force | Out-Null

# Copy skill/* (SKILL.md, scripts/, templates/)
Copy-Item -Path (Join-Path $here "skill\*") -Destination $skillDest -Recurse -Force
Copy-Item -Path (Join-Path $here "config.example.yaml") -Destination $skillDest -Force

Write-Host "  Copied to: $skillDest" -ForegroundColor Green

# ============================================================
# 4) Interactive config.yaml setup
# ============================================================
Write-Host ""
Write-Host "[4/5] Configuration prompts..." -ForegroundColor Yellow
Write-Host ""

# Q1 — Default output folder
Write-Host "  Default save folder for generated datarooms"
$defaultOutput = Join-Path $env:USERPROFILE "Documents\Datarooms"
$outFolder = Read-Host "  Folder (default: $defaultOutput)"
if ([string]::IsNullOrWhiteSpace($outFolder)) { $outFolder = $defaultOutput }

# Q2 — Folder structure language
Write-Host ""
Write-Host "  Default language for folder structure (N1/N2 names)"
Write-Host "  Options: fr (French) / en (English) / de (German) / it (Italian) / es (Spanish)"
$folderLang = Read-Host "  Language code (default: en)"
if ([string]::IsNullOrWhiteSpace($folderLang)) { $folderLang = "en" }
$folderLang = $folderLang.ToLower()

# Q3 — Excel mapping language
$excelLang = Read-Host "  Excel mapping language (default: same as folder = $folderLang)"
if ([string]::IsNullOrWhiteSpace($excelLang)) { $excelLang = $folderLang }
$excelLang = $excelLang.ToLower()

# Q4 — Default location for Phase 2
Write-Host ""
Write-Host "  Default company location for Phase 2 (DD checklist enrichment)"
Write-Host "  Options: France / Suisse / Belgique / Luxembourg / autre"
$defaultLocation = Read-Host "  Location (default: France)"
if ([string]::IsNullOrWhiteSpace($defaultLocation)) { $defaultLocation = "France" }

# Q5 — Branding
Write-Host ""
Write-Host "  Branding for Excel cover + table headers"
Write-Host "  Type 'default' for Hectelion SA styling (navy #182E4E + Cardo),"
Write-Host "  or provide your own: firm name, primary hex color (no #), font family."
Write-Host "  Example: Smith Partners, 1A3258, Calibri"
$brandInput = Read-Host "  Branding"

# ============================================================
# 5) Write config.yaml
# ============================================================
Write-Host ""
Write-Host "[5/5] Writing config.yaml..." -ForegroundColor Yellow
$configPath = Join-Path $skillDest "config.yaml"

# Parse branding
$brandFirm = "Hectelion SA"
$brandPrimary = "182E4E"
$brandFont = "Cardo"
if ($brandInput -and $brandInput.ToLower() -ne "default") {
    $parts = $brandInput.Split(",") | ForEach-Object { $_.Trim() }
    if ($parts.Count -ge 1 -and $parts[0]) { $brandFirm = $parts[0] }
    if ($parts.Count -ge 2 -and $parts[1]) { $brandPrimary = ($parts[1] -replace "^#", "") }
    if ($parts.Count -ge 3 -and $parts[2]) { $brandFont = $parts[2] }
}

$outFolderEscaped = $outFolder -replace '\\', '\\'

$yamlBody = @"
# Auto-generated by install.ps1 on $(Get-Date -Format "yyyy-MM-dd HH:mm")
# See config.example.yaml for the full schema and comments.

output:
  default_folder: "$outFolderEscaped"
  default_folder_language: "$folderLang"
  default_excel_language: "$excelLang"
  make_zip: true

renaming:
  pattern: "{project} - {theme} - {subject} - {date}"

checklist:
  auto_run: false
  default_location: "$defaultLocation"
  default_sector: ""
  data_sources:
    - "data.gouv.fr"

brand:
  firm_name: "$brandFirm"
  firm_tagline: "Independent M&A advisory"
  primary_color: "$brandPrimary"
  secondary_color: "0E2841"
  accent_pale: "DCEAF7"
  accent_green: "6FCF9A"
  accent_orange: "FFC000"
  accent_grey: "F2F2F2"
  accent_input_cyan: "72C7E7"
  alert_red: "C00000"
  border_grey: "D9D9D9"
  font_family: "$brandFont"
  logo_path: ""

extraction:
  max_chars_per_doc: 2500
  accepted_extensions: [".pdf", ".docx", ".xlsx", ".xlsm", ".doc", ".xls", ".pptx", ".ppt", ".msg", ".rtf", ".txt"]
  exclude_substrings:
    - "node_modules"
    - "linux-unpacked"
    - "/temp/"
    - ".git/"
    - "`$recycle.bin"
"@

$yamlBody | Out-File -FilePath $configPath -Encoding utf8
Write-Host "  Saved: $configPath" -ForegroundColor Green

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Installation complete!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "From any Claude Code project, run:"
Write-Host "  /dataroom-prep" -ForegroundColor White
Write-Host ""
Write-Host "Edit config later at:"
Write-Host "  $configPath" -ForegroundColor White
Write-Host ""
Write-Host "Uninstall:"
Write-Host "  Remove-Item -Recurse -Force `"$skillDest`"" -ForegroundColor White
Write-Host ""
