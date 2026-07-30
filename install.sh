#!/usr/bin/env bash
# ============================================================
# Dataroom Prep — interactive installer (macOS / Linux)
#
# Copies the skill files into ~/.claude/skills/dataroom-prep, installs
# Python dependencies (incl. the OCR stack for scanned PDFs), and walks
# you through filling in config.yaml.
#
# Prerequisites:
# - macOS or Linux
# - Python 3.9+ on PATH
# - Claude Code installed (https://claude.com/claude-code)
#
# Uninstall: rm -rf ~/.claude/skills/dataroom-prep
# ============================================================
set -e

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo ""
echo "============================================================"
echo "  Dataroom Prep - Interactive installer"
echo "============================================================"
echo ""
echo "Repository: https://github.com/Hectelion-SA/claude-dataroom-prep"
echo ""

# ============================================================
# 1) Check Python + dependencies
# ============================================================
echo "[1/5] Checking Python..."
PY=""
for cand in python3 python; do
    if command -v "$cand" >/dev/null 2>&1; then PY="$cand"; break; fi
done
if [ -z "$PY" ]; then
    echo "  ERROR: Python not found on PATH."
    echo "  Install Python 3.9+ from https://python.org and retry."
    exit 1
fi
echo "  Found: $($PY --version) ($(command -v "$PY"))"

echo "[2/5] Installing Python packages (pypdf, python-docx, openpyxl, pyyaml, pymupdf, pytesseract, pillow)..."
if ! "$PY" -m pip install --quiet pypdf python-docx openpyxl pyyaml pymupdf pytesseract pillow; then
    echo "  ERROR: pip install failed."
    echo "  Run manually: $PY -m pip install pypdf python-docx openpyxl pyyaml pymupdf pytesseract pillow"
    exit 1
fi
echo "  All packages OK."

# Tesseract OCR binary (scanned PDF fallback) — best-effort, non-fatal.
# The skill also retries this automatically on every run if still missing
# (see ensure_ocr_stack() in scripts/run_pipeline.py), so a failure here is not blocking.
echo "[2b/5] Checking Tesseract OCR (for scanned PDFs)..."
if command -v tesseract >/dev/null 2>&1; then
    echo "  Tesseract already installed."
elif [[ "$OSTYPE" == "darwin"* ]] && command -v brew >/dev/null 2>&1; then
    brew install tesseract && echo "  Tesseract installed via Homebrew." \
        || echo "  Could not auto-install Tesseract now — the skill will retry automatically on first real run."
elif command -v apt-get >/dev/null 2>&1; then
    sudo apt-get install -y tesseract-ocr && echo "  Tesseract installed via apt." \
        || echo "  Could not auto-install Tesseract now — the skill will retry automatically on first real run."
else
    echo "  No known package manager found — the skill will retry automatically on first real run, or install Tesseract manually."
fi

# ============================================================
# 3) Copy skill files into ~/.claude/skills/dataroom-prep/
# ============================================================
echo "[3/5] Installing skill into Claude Code skills folder..."

SKILLS_ROOT="$HOME/.claude/skills"
SKILL_DEST="$SKILLS_ROOT/dataroom-prep"

mkdir -p "$SKILLS_ROOT"

if [ -d "$SKILL_DEST" ]; then
    read -r -p "  Folder $SKILL_DEST already exists. Overwrite skill files? (y/N) " ans
    if [[ "$ans" != "y" && "$ans" != "Y" ]]; then
        echo "  Aborted by user."
        exit 0
    fi
    rm -rf "$SKILL_DEST"
fi
mkdir -p "$SKILL_DEST"

cp -r "$HERE/skill/." "$SKILL_DEST/"
cp "$HERE/config.example.yaml" "$SKILL_DEST/"

echo "  Copied to: $SKILL_DEST"

# ============================================================
# 4) Interactive config.yaml setup
# ============================================================
echo ""
echo "[4/5] Configuration prompts..."
echo ""

echo "  Default save folder for generated datarooms"
DEFAULT_OUTPUT="$HOME/Documents/Datarooms"
read -r -p "  Folder (default: $DEFAULT_OUTPUT): " OUT_FOLDER
OUT_FOLDER="${OUT_FOLDER:-$DEFAULT_OUTPUT}"

echo ""
echo "  Default language for folder structure (N1/N2 names)"
echo "  Options: fr (French) / en (English) / de (German) / it (Italian) / es (Spanish)"
read -r -p "  Language code (default: en): " FOLDER_LANG
FOLDER_LANG="${FOLDER_LANG:-en}"
FOLDER_LANG="$(echo "$FOLDER_LANG" | tr '[:upper:]' '[:lower:]')"

read -r -p "  Excel mapping language (default: same as folder = $FOLDER_LANG): " EXCEL_LANG
EXCEL_LANG="${EXCEL_LANG:-$FOLDER_LANG}"
EXCEL_LANG="$(echo "$EXCEL_LANG" | tr '[:upper:]' '[:lower:]')"

echo ""
echo "  Default company location (used to pick the DD reference list: France/Suisse)"
echo "  Options: France / Suisse / Belgique / Luxembourg / autre"
read -r -p "  Location (default: France): " DEFAULT_LOCATION
DEFAULT_LOCATION="${DEFAULT_LOCATION:-France}"

echo ""
echo "  Branding for Excel cover + table headers"
echo "  Type 'default' for Hectelion SA styling (navy #182E4E + Cardo),"
echo "  or provide your own: firm name, primary hex color (no #), font family."
echo "  Example: Smith Partners, 1A3258, Calibri"
read -r -p "  Branding: " BRAND_INPUT

# ============================================================
# 5) Write config.yaml
# ============================================================
echo ""
echo "[5/5] Writing config.yaml..."
CONFIG_PATH="$SKILL_DEST/config.yaml"

BRAND_FIRM="Hectelion SA"
BRAND_PRIMARY="182E4E"
BRAND_FONT="Cardo"
if [ -n "$BRAND_INPUT" ] && [ "$(echo "$BRAND_INPUT" | tr '[:upper:]' '[:lower:]')" != "default" ]; then
    IFS=',' read -r p1 p2 p3 <<< "$BRAND_INPUT"
    [ -n "$(echo "$p1" | xargs)" ] && BRAND_FIRM="$(echo "$p1" | xargs)"
    [ -n "$(echo "$p2" | xargs)" ] && BRAND_PRIMARY="$(echo "$p2" | xargs | sed 's/^#//')"
    [ -n "$(echo "$p3" | xargs)" ] && BRAND_FONT="$(echo "$p3" | xargs)"
fi

cat > "$CONFIG_PATH" <<YAML
# Auto-generated by install.sh on $(date "+%Y-%m-%d %H:%M")
# See config.example.yaml for the full schema and comments.

output:
  default_folder: "$OUT_FOLDER"
  default_folder_language: "$FOLDER_LANG"
  default_excel_language: "$EXCEL_LANG"
  make_zip: true

renaming:
  # Short, content-derived title (never the theme, never the raw filename) + yyyymmdd.
  pattern: "{project} - {title} - {yyyymmdd}"

checklist:
  auto_run: false
  default_location: "$DEFAULT_LOCATION"
  default_sector: ""
  data_sources:
    - "data.gouv.fr"

brand:
  firm_name: "$BRAND_FIRM"
  firm_tagline: "Independent M&A advisory"
  primary_color: "$BRAND_PRIMARY"
  secondary_color: "0E2841"
  accent_pale: "DCEAF7"
  accent_green: "6FCF9A"
  accent_orange: "FFC000"
  accent_grey: "F2F2F2"
  accent_input_cyan: "72C7E7"
  alert_red: "C00000"
  border_grey: "D9D9D9"
  font_family: "$BRAND_FONT"
  logo_path: ""

extraction:
  max_chars_per_doc: 2500
  accepted_extensions: [".pdf", ".docx", ".xlsx", ".xlsm", ".doc", ".xls", ".pptx", ".ppt", ".msg", ".rtf", ".txt"]
  exclude_substrings:
    - "node_modules"
    - "linux-unpacked"
    - "/temp/"
    - ".git/"
    - "\$recycle.bin"
YAML

echo "  Saved: $CONFIG_PATH"

echo ""
echo "============================================================"
echo "  Installation complete!"
echo "============================================================"
echo ""
echo "From any Claude Code project, run:"
echo "  /dataroom-prep"
echo ""
echo "Edit config later at:"
echo "  $CONFIG_PATH"
echo ""
echo "Uninstall:"
echo "  rm -rf \"$SKILL_DEST\""
echo ""
