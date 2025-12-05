---
name: xl-cli
description: "LLM-friendly Excel operations via the `xl` CLI. Read cells, view ranges, search, evaluate formulas, export (CSV/JSON/PNG/PDF), style cells, modify rows/columns. Use when working with .xlsx files or spreadsheet data."
---

# XL CLI - Excel Operations

## Installation

Check if installed: `which xl || echo "not installed"`

**If not installed**, download native binary (no JDK required):

**macOS/Linux:**
```bash
# Detect platform and install __XL_VERSION__
PLATFORM="$(uname -s)-$(uname -m)"
case "$PLATFORM" in
  Linux-x86_64)  BINARY="xl-__XL_VERSION__-linux-amd64" ;;
  Darwin-x86_64) BINARY="xl-__XL_VERSION__-darwin-amd64" ;;
  Darwin-arm64)  BINARY="xl-__XL_VERSION__-darwin-arm64" ;;
  *) echo "Unsupported platform: $PLATFORM" && exit 1 ;;
esac
mkdir -p ~/.local/bin
curl -sL "https://github.com/TJC-LP/xl/releases/download/v__XL_VERSION__/$BINARY" -o ~/.local/bin/xl
chmod +x ~/.local/bin/xl
```

**Windows (PowerShell):**
```powershell
Invoke-WebRequest -Uri "https://github.com/TJC-LP/xl/releases/download/v__XL_VERSION__/xl-__XL_VERSION__-windows-amd64.exe" -OutFile "$env:LOCALAPPDATA\xl.exe"
# Add to PATH or move to a directory in PATH
```

**Alternative** (requires JDK 17+):
```bash
curl -sL "https://github.com/TJC-LP/xl/releases/download/v__XL_VERSION__/xl-cli-__XL_VERSION__.tar.gz" | tar xz -C /tmp
cd /tmp && ./install.sh
```

---

## Contents

- [Quick Reference](#quick-reference)
- [Sheet Handling](#sheet-handling)
- [Styling](#styling)
- [Row/Column Operations](#rowcolumn-operations)
- [Output Formats](#output-formats)
- [Workflows](#workflows)
- [Command Reference](#command-reference)

---

## Quick Reference

```bash
# Read operations
xl -f <file> sheets                    # List sheets
xl -f <file> bounds                    # Used range
xl -f <file> view <range>              # View as table
xl -f <file> cell <ref>                # Cell details + dependencies
xl -f <file> search <pattern>          # Find cells
xl -f <file> eval <formula>            # Evaluate formula

# Output formats
xl -f <file> view <range> --format json
xl -f <file> view <range> --format csv --show-labels
xl -f <file> view <range> --format png --raster-output out.png

# Write operations (require -o)
xl -f <file> -o <out> put <ref> <value>
xl -f <file> -o <out> putf <ref> <formula>

# Style operations (require -o)
xl -f <file> -o <out> style <range> --bold --bg yellow
xl -f <file> -o <out> style <range> --bg "#FF6600" --fg white

# Row/Column operations (require -o)
xl -f <file> -o <out> row <n> --height 30
xl -f <file> -o <out> col <letter> --width 20 --hide
```

---

## Sheet Handling

Commands default to first sheet. For multi-sheet files, always specify explicitly:

```bash
# Method 1: --sheet flag
xl -f data.xlsx --sheet "P&L" view A1:D10

# Method 2: Qualified A1 syntax
xl -f data.xlsx view "P&L!A1:D10"
xl -f data.xlsx eval "=SUM(Revenue!A1:A10)"
```

**Workflow**: Always start with `xl -f file.xlsx sheets` to discover sheet names.

---

## Styling

Apply formatting with `style <range>` command.

| Option | Description |
|--------|-------------|
| `--bold` / `--italic` / `--underline` | Text style |
| `--bg <color>` | Background |
| `--fg <color>` | Text color |
| `--font-size <pt>` | Font size |
| `--font-name <name>` | Font family |
| `--align <left\|center\|right>` | Horizontal |
| `--valign <top\|middle\|bottom>` | Vertical |
| `--wrap` | Text wrap |
| `--format <general\|number\|currency\|percent\|date\|text>` | Number format |
| `--border <none\|thin\|medium\|thick>` | Border style |
| `--border-color <color>` | Border color |

**Colors**: Named (`red`, `navy`), hex (`#FF6600`), or RGB (`rgb(100,150,200)`).
See [reference/COLORS.md](reference/COLORS.md) for full color list.

```bash
# Header styling
xl -f data.xlsx -o out.xlsx style A1:E1 --bold --bg navy --fg white --align center

# Currency column
xl -f data.xlsx -o out.xlsx style B2:B100 --format currency
```

---

## Row/Column Operations

```bash
# Row height and visibility
xl -f data.xlsx -o out.xlsx row 5 --height 30
xl -f data.xlsx -o out.xlsx row 10 --hide
xl -f data.xlsx -o out.xlsx row 10 --show

# Column width and visibility
xl -f data.xlsx -o out.xlsx col B --width 20
xl -f data.xlsx -o out.xlsx col C --hide
```

| Option | Description |
|--------|-------------|
| `--height <pt>` | Row height (row only) |
| `--width <chars>` | Column width (col only) |
| `--hide` | Hide row/column |
| `--show` | Unhide row/column |

---

## Output Formats

| Format | Flag | Notes |
|--------|------|-------|
| markdown | Default | Text table |
| json | `--format json` | Structured data |
| csv | `--format csv` | Add `--show-labels` for headers |
| html | `--format html` | Inline CSS |
| svg | `--format svg` | Vector |
| png/jpeg/webp/pdf | `--format <fmt> --raster-output <path>` | Requires ImageMagick |

See [reference/OUTPUT-FORMATS.md](reference/OUTPUT-FORMATS.md) for detailed specs.

**Raster options**: `--dpi <n>`, `--quality <n>`, `--show-labels`

```bash
# Visual analysis (Claude vision)
xl -f data.xlsx view A1:F20 --format png --raster-output /tmp/sheet.png --show-labels
```

---

## Workflows

### Explore Unknown Spreadsheet

```bash
xl -f data.xlsx sheets                     # List sheets
xl -f data.xlsx --sheet "Sheet1" bounds    # Get range
xl -f data.xlsx --sheet "Sheet1" view A1:E20
```

### Formula Analysis

```bash
xl -f data.xlsx view --formulas A1:D10     # Show formulas
xl -f data.xlsx cell C5                    # Dependencies
xl -f data.xlsx eval "=SUM(A1:A10)" --with "A1=500"  # What-if
```

See [reference/FORMULAS.md](reference/FORMULAS.md) for supported functions.

### Create Formatted Report

```bash
xl -f template.xlsx -o report.xlsx put A1 "Sales Report"
xl -f report.xlsx -o report.xlsx style A1:E1 --bold --bg navy --fg white
xl -f report.xlsx -o report.xlsx style B2:B100 --format currency
xl -f report.xlsx -o report.xlsx col A --width 25
```

---

## Command Reference

### Global Options

| Option | Alias | Description |
|--------|-------|-------------|
| `--file <path>` | `-f` | Input file (required) |
| `--sheet <name>` | `-s` | Sheet name |
| `--output <path>` | `-o` | Output file (for writes) |

### View Options

| Option | Description |
|--------|-------------|
| `--format <fmt>` | Output format |
| `--formulas` | Show formulas |
| `--limit <n>` | Max rows (default: 50) |
| `--show-labels` | Row/column headers |
| `--raster-output <path>` | Image output path |
| `--dpi <n>` | Resolution (default: 144) |
| `--quality <n>` | JPEG quality (default: 90) |

### Search Options

| Option | Description |
|--------|-------------|
| `--limit <n>` | Max matches |

### Eval Options

| Option | Alias | Description |
|--------|-------|-------------|
| `--with <overrides>` | `-w` | Cell overrides (e.g., `A1=100,B2=200`) |
