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
- [Sheet Management](#sheet-management)
- [Cell Operations](#cell-operations)
- [Batch Operations](#batch-operations)
- [Styling](#styling)
- [Row/Column Operations](#rowcolumn-operations)
- [Output Formats](#output-formats)
- [Workflows](#workflows)
- [Command Reference](#command-reference)

---

## Quick Reference

```bash
# Info commands
xl functions                           # List all 81 supported functions

# Read operations
xl -f <file> sheets                    # List sheets with stats
xl -f <file> names                     # List defined names (named ranges)
xl -f <file> -s <sheet> bounds         # Used range
xl -f <file> -s <sheet> view <range>   # View as table
xl -f <file> -s <sheet> cell <ref>     # Cell details + dependencies
xl -f <file> -s <sheet> search <pattern>  # Find cells
xl -f <file> -s <sheet> stats <range>  # Calculate statistics (count, sum, min, max, mean)
xl -f <file> -s <sheet> eval <formula> # Evaluate formula

# Output formats
xl -f <file> -s <sheet> view <range> --format json
xl -f <file> -s <sheet> view <range> --format csv --show-labels
xl -f <file> -s <sheet> view <range> --format png --raster-output out.png

# Write operations (require -o)
xl -f <file> -s <sheet> -o <out> put <ref> <value>
xl -f <file> -s <sheet> -o <out> putf <ref> <formula>
xl -f <file> -s <sheet> -o <out> import <csv-file> <ref>
xl -f <file> -o <out> import <csv-file> --new-sheet "Data"

# Style operations (require -o)
xl -f <file> -s <sheet> -o <out> style <range> --bold --bg yellow
xl -f <file> -s <sheet> -o <out> style <range> --bg "#FF6600" --fg white

# Row/Column operations (require -o)
xl -f <file> -s <sheet> -o <out> row <n> --height 30
xl -f <file> -s <sheet> -o <out> col <letter> --width 20 --hide
xl -f <file> -s <sheet> -o <out> col <letter> --auto-fit
xl -f <file> -s <sheet> -o <out> col A:F --auto-fit     # Column range
xl -f <file> -s <sheet> -o <out> autofit                # All used columns
xl -f <file> -s <sheet> -o <out> autofit --columns A:Z  # Specific range

# Sheet management (require -o)
xl -f <file> -o <out> add-sheet "NewSheet"
xl -f <file> -o <out> remove-sheet "OldSheet"
xl -f <file> -o <out> rename-sheet "Old" "New"
xl -f <file> -o <out> move-sheet "Sheet1" --to 0
xl -f <file> -o <out> copy-sheet "Template" "Copy"

# Cell operations (require -o and -s)
xl -f <file> -s <sheet> -o <out> merge A1:C1
xl -f <file> -s <sheet> -o <out> unmerge A1:C1
xl -f <file> -s <sheet> -o <out> comment A1 "Review this" --author "John"
xl -f <file> -s <sheet> -o <out> remove-comment A1

# Batch operations (require -o)
xl -f <file> -s <sheet> -o <out> batch <json-file>
xl -f <file> -s <sheet> -o <out> batch -              # Read from stdin

# Formula dragging (putf with range)
xl -f <file> -s <sheet> -o <out> putf <range> <formula>  # Drags formula over range

# Create new workbook
xl new <output>                                          # Default Sheet1
xl new <output> --sheet Data --sheet Summary             # Multiple sheets
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

## Sheet Management

Modify workbook structure with these commands (all require `-o` for output):

### Add / Remove Sheets

```bash
# Add new empty sheet (appends to end)
xl -f data.xlsx -o out.xlsx add-sheet "NewSheet"

# Add sheet at specific position
xl -f data.xlsx -o out.xlsx add-sheet "Summary" --before "Sheet1"
xl -f data.xlsx -o out.xlsx add-sheet "Notes" --after "Data"

# Remove sheet
xl -f data.xlsx -o out.xlsx remove-sheet "Scratch"
```

### Rename / Move / Copy Sheets

```bash
# Rename
xl -f data.xlsx -o out.xlsx rename-sheet "Sheet1" "Summary"

# Move to position (0-based index)
xl -f data.xlsx -o out.xlsx move-sheet "Notes" --to 0

# Move relative to another sheet
xl -f data.xlsx -o out.xlsx move-sheet "Notes" --after "Summary"
xl -f data.xlsx -o out.xlsx move-sheet "Notes" --before "Data"

# Copy (duplicate with new name)
xl -f data.xlsx -o out.xlsx copy-sheet "Template" "Q1 Report"
```

| Command | Description |
|---------|-------------|
| `add-sheet <name>` | Add empty sheet (`--after`/`--before` for position) |
| `remove-sheet <name>` | Remove sheet from workbook |
| `rename-sheet <old> <new>` | Rename a sheet |
| `move-sheet <name>` | Move sheet (`--to`/`--after`/`--before`) |
| `copy-sheet <src> <dest>` | Duplicate sheet with new name |

---

## Cell Operations

### Comments

Add or remove comments from cells (yellow notes in Excel):

```bash
# Add comment
xl -f file.xlsx -s Sheet1 -o out.xlsx comment A1 "Review this value"

# Add comment with author
xl -f file.xlsx -s Sheet1 -o out.xlsx comment A1 "Q1 data needs verification" --author "Finance Team"

# Remove comment
xl -f file.xlsx -s Sheet1 -o out.xlsx remove-comment A1
```

Comments appear as yellow notes in Excel when hovering over cells. The `cell` command shows existing comments.

### Merge / Unmerge Cells

Merge combines cells into one; unmerge separates them.

```bash
# Merge header row
xl -f data.xlsx -s Sheet1 -o out.xlsx merge A1:D1

# Unmerge
xl -f data.xlsx -s Sheet1 -o out.xlsx unmerge A1:D1
```

**Note**: HTML output shows merged cells with `colspan`/`rowspan`. Markdown tables cannot represent merges.

### Statistics

Calculate statistics for numeric values in a range:

```bash
xl -f data.xlsx -s Sheet1 stats B2:B100
# Output: count: 99, sum: 12345.00, min: 10.00, max: 500.00, mean: 124.70
```

### Formula Dragging

When `putf` receives a range, it drags the formula Excel-style:

```bash
# Fill B2:B10 with formulas, shifting references
xl -f data.xlsx -s Sheet1 -o out.xlsx putf B2:B10 "=A2*1.1"

# Result:
# B2: =A2*1.1
# B3: =A3*1.1
# ...

# Use $ to anchor references
xl -f data.xlsx -s Sheet1 -o out.xlsx putf B2:B10 "=SUM(\$A\$1:A2)"

# Result:
# B2: =SUM($A$1:A2)
# B3: =SUM($A$1:A3)
# ...
```

**Anchor modes**:

| Syntax | Behavior |
|--------|----------|
| `$A$1` | Absolute (never shifts) |
| `$A1` | Column absolute, row relative |
| `A$1` | Column relative, row absolute |
| `A1` | Fully relative (shifts both ways) |

### Putting Negative Numbers

For negative numbers, use the `--value` flag because `-` is interpreted as a flag prefix:

```bash
# ❌ Wrong - interpreted as flags
xl -f data.xlsx -s Sheet1 -o out.xlsx put A1 -100

# ✅ Correct - use --value flag
xl -f data.xlsx -s Sheet1 -o out.xlsx put A1 --value "-100"
```

The `--value` flag works for all values, but it's required for negative numbers:

```bash
# Both work for non-negative values
xl -f data.xlsx -o out.xlsx put A1 100
xl -f data.xlsx -o out.xlsx put A1 --value "100"

# Only --value works for negative values
xl -f data.xlsx -o out.xlsx put A1 --value "-100"
```

---

## Batch Put & Fill Patterns

The `put` command supports three modes automatically based on argument count:

### Mode 1: Single Cell
```bash
xl -f file.xlsx -s Sheet1 -o out.xlsx put A1 100
```

### Mode 2: Fill Pattern
Fill all cells in a range with the same value:

```bash
# Fill column
xl -f file.xlsx -s Sheet1 -o out.xlsx put A1:A10 "TBD"

# Fill row
xl -f file.xlsx -s Sheet1 -o out.xlsx put A1:Z1 0

# Fill 2D range
xl -f file.xlsx -s Sheet1 -o out.xlsx put A1:C3 "Empty"
```

### Mode 3: Batch Values
Map different values to each cell (row-major order: left-to-right, top-to-bottom):

```bash
# Set header row
xl -f file.xlsx -s Sheet1 -o out.xlsx put A1:D1 "Name" "Age" "City" "Active"

# Set column
xl -f file.xlsx -s Sheet1 -o out.xlsx put B2:B4 100 200 300

# Fill 2D grid (row-major)
xl -f file.xlsx -s Sheet1 -o out.xlsx put A1:B2 1 2 3 4
# Result: A1=1, B1=2, A2=3, B2=4
```

**Important**: Value count must match cell count:
```bash
xl put A1:A5 1 2 3
# Error: Range A1:A5 has 5 cells but 3 values provided
```

## Batch Formulas

The `putf` command supports three modes:

### Mode 1: Single Formula
```bash
xl -f file.xlsx -s Sheet1 -o out.xlsx putf B1 "=A1*2"
```

### Mode 2: Formula Dragging (with $ anchors)
Single formula with intelligent reference shifting:

```bash
# Relative references shift
xl -f file.xlsx -s Sheet1 -o out.xlsx putf B1:B10 "=A1*2"
# B1=A1*2, B2=A2*2, B3=A3*2, ...

# $ anchors control shifting
xl -f file.xlsx -s Sheet1 -o out.xlsx putf C1:C10 "=SUM(\$A\$1:A1)"
# C1=SUM($A$1:A1), C2=SUM($A$1:A2), ...
```

### Mode 3: Batch Formulas (explicit, no dragging)
Different formulas for each cell:

```bash
xl -f file.xlsx -s Sheet1 -o out.xlsx putf D1:D3 "=A1+B1" "=A2*B2" "=A3-B3"
# Formulas applied as-is, no shifting

# Mix constants and references
xl -f file.xlsx -s Sheet1 -o out.xlsx putf E1:E2 "=100" "=D1*1.1"
```

**Important**: Formula count must match cell count.

---

## CSV Import

Import CSV data into Excel workbooks with automatic type detection.

```bash
# Import to existing sheet at A1
xl -f file.xlsx -s Sheet1 -o out.xlsx import data.csv A1

# Import at specific position
xl -f file.xlsx -s Sheet1 -o out.xlsx import data.csv B5

# Import to new sheet
xl -f file.xlsx -o out.xlsx import data.csv --new-sheet "Imported Data"

# Custom delimiter (semicolon, tab, etc.)
xl -f file.xlsx -s Sheet1 -o out.xlsx import data.csv A1 --delimiter ";"

# Treat first row as data (not headers)
xl -f file.xlsx -s Sheet1 -o out.xlsx import data.csv A1 --no-header

# Force all values to text (disable type inference)
xl -f file.xlsx -s Sheet1 -o out.xlsx import data.csv A1 --no-type-inference

# Custom encoding
xl -f file.xlsx -s Sheet1 -o out.xlsx import data.csv A1 --encoding "ISO-8859-1"
```

**Type Inference**: Automatically detects and converts:
- **Numbers**: `100`, `29.99`, `-5.5` → Number type with Decimal format
- **Booleans**: `true`, `false` (case-insensitive) → Boolean type
- **Dates**: `2024-01-15` (ISO 8601 format) → DateTime type with Date format
- **Text**: Everything else

**Column-based inference**: Samples first 10 rows per column. If all values match a type, assigns that type; otherwise defaults to Text.

**Options**:
- `--delimiter <char>` - Field separator (default: `,`)
- `--no-header` - First row is data, not headers
- `--encoding <enc>` - Input encoding (default: UTF-8)
- `--new-sheet <name>` - Create new sheet for imported data
- `--no-type-inference` - Treat all values as text (useful for ZIP codes, phone numbers)

**Limitations**:
- **Memory**: Entire CSV loaded into memory (not streamed)
- **Recommended size**: <50k rows for optimal performance
- **Large files**: Split CSVs before importing or use batch operations
- **Date formats**: Only ISO 8601 (YYYY-MM-DD) supported
- **Boolean formats**: Only true/false (case-insensitive)

## Batch Operations

Apply multiple cell operations atomically from JSON input:

```bash
# From file
xl -f data.xlsx -s Sheet1 -o out.xlsx batch operations.json

# From stdin
echo '[{"op":"put","ref":"A1","value":"Hello"},{"op":"putf","ref":"B1","value":"=A1&\" World\""}]' \
  | xl -f data.xlsx -s Sheet1 -o out.xlsx batch -
```

**JSON Format**:
```json
[
  {"op": "put", "ref": "A1", "value": "Hello"},
  {"op": "put", "ref": "A2", "value": "123"},
  {"op": "putf", "ref": "B1", "value": "=A1*2"},
  {"op": "put", "ref": "Sheet2!C1", "value": "Cross-sheet"}
]
```

| Field | Description |
|-------|-------------|
| `op` | Operation: `put` (value) or `putf` (formula) |
| `ref` | Cell reference (supports qualified `Sheet!Ref`) |
| `value` | Value or formula string |

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
| `--border <none\|thin\|medium\|thick>` | Border style (all sides) |
| `--border-top <style>` | Top border only |
| `--border-right <style>` | Right border only |
| `--border-bottom <style>` | Bottom border only |
| `--border-left <style>` | Left border only |
| `--border-color <color>` | Border color |
| `--replace` | Replace entire style (default: merge) |

**Style Behavior**: Styles are merged by default (new properties combine with existing).
Use `--replace` to replace the entire style instead.

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

# Auto-fit column width based on content
xl -f data.xlsx -o out.xlsx col A --auto-fit            # Single column
xl -f data.xlsx -o out.xlsx col A:F --auto-fit          # Column range

# Auto-fit all columns (or specific range)
xl -f data.xlsx -o out.xlsx autofit                     # All used columns
xl -f data.xlsx -o out.xlsx autofit --columns A:Z       # Specific range
```

| Option | Description |
|--------|-------------|
| `--height <pt>` | Row height (row only) |
| `--width <chars>` | Column width (col only) |
| `--auto-fit` | Auto-fit column width based on content (col only) |
| `--columns <range>` | Column range for autofit command (e.g., A:Z) |
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
| png/jpeg/pdf | `--format <fmt> --raster-output <path>` | Auto fallback (Batik/cairosvg/rsvg/ImageMagick) |
| webp | `--format webp --raster-output <path>` | Requires ImageMagick |

See [reference/OUTPUT-FORMATS.md](reference/OUTPUT-FORMATS.md) for detailed specs.

**Raster options**: `--dpi <n>`, `--quality <n>`, `--show-labels`, `--rasterizer <name>`

```bash
# Visual analysis (Claude vision)
xl -f data.xlsx view A1:F20 --format png --raster-output /tmp/sheet.png --show-labels
```

---

## Workflows

### Explore Unknown Spreadsheet

```bash
xl -f data.xlsx sheets                     # List sheets with cell counts
xl -f data.xlsx names                      # List defined names (named ranges)
xl -f data.xlsx -s "Sheet1" bounds         # Get used range
xl -f data.xlsx -s "Sheet1" view A1:E20    # View data
xl -f data.xlsx -s "Sheet1" stats B2:B100  # Quick statistics
```

### Formula Analysis

```bash
xl -f data.xlsx -s Sheet1 view --formulas A1:D10     # Show formulas
xl -f data.xlsx -s Sheet1 cell C5                    # Dependencies
xl -f data.xlsx -s Sheet1 eval "=SUM(A1:A10)" --with "A1=500"  # What-if
```

See [reference/FORMULAS.md](reference/FORMULAS.md) for supported functions.

### Create Formatted Report

```bash
xl -f template.xlsx -s Sheet1 -o report.xlsx put A1 "Sales Report"
xl -f report.xlsx -s Sheet1 -o report.xlsx style A1:E1 --bold --bg navy --fg white
xl -f report.xlsx -s Sheet1 -o report.xlsx style B2:B100 --format currency
xl -f report.xlsx -s Sheet1 -o report.xlsx col A --width 25
```

### Multi-Sheet Workbook Setup

```bash
# Create workbook (default sheet: Sheet1)
xl new output.xlsx

# Create with custom sheet name
xl new output.xlsx --sheet-name "Data"

# Create with multiple sheets in one command
xl new output.xlsx --sheet Data --sheet Summary --sheet Notes

# Add more sheets to existing workbook
xl -f output.xlsx -o output.xlsx add-sheet "Archive" --after "Notes"
xl -f output.xlsx -o output.xlsx copy-sheet "Summary" "Q1 Summary"
```

---

## Performance

### Backend Selection

Choose between two XML backends with the `--backend` flag:

| Backend | Speed | Stability | Use Case |
|---------|-------|-----------|----------|
| `scalaxml` | Baseline | Proven | Default, general use |
| `saxstax` | **36-39% faster** | Production-ready | Heavy writes, batch ops |

**Examples**:
```bash
# Use faster backend for batch operations
xl -f data.xlsx -o out.xlsx --backend saxstax batch ops.json

# Create new file with faster backend
xl new report.xlsx --backend saxstax --sheet Data --sheet Summary

# Style large range with faster backend
xl -f data.xlsx -o out.xlsx --backend saxstax style A1:Z1000 --bold
```

**Recommendation**: Use `saxstax` for:
- Large batch operations
- Styling many cells
- Creating files with `new` command
- Repeated write operations in scripts

---

## Command Reference

### Global Options

| Option | Alias | Description |
|--------|-------|-------------|
| `--file <path>` | `-f` | Input file (required) |
| `--sheet <name>` | `-s` | Sheet name |
| `--output <path>` | `-o` | Output file (for writes) |
| `--backend <type>` | | XML backend: scalaxml (default, stable) or saxstax (36-39% faster) |

### View Options

| Option | Description |
|--------|-------------|
| `--format <fmt>` | Output format (markdown, json, csv, html, svg, png, jpeg, webp, pdf) |
| `--formulas` | Show formulas instead of values |
| `--eval` | Evaluate formulas (compute live values) |
| `--limit <n>` | Max rows (default: 50) |
| `--skip-empty` | Skip empty cells (JSON) or empty rows/columns (tabular) |
| `--header-row <n>` | Use row N as JSON keys (1-based) |
| `--show-labels` | Include row/column headers |
| `--raster-output <path>` | Image output path (required for png/jpeg/webp/pdf) |
| `--dpi <n>` | Resolution (default: 144) |
| `--quality <n>` | JPEG quality (default: 90) |
| `--gridlines` | Show cell gridlines in SVG |
| `--print-scale` | Apply print scaling |
| `--rasterizer <name>` | Force specific rasterizer: batik, cairosvg, rsvg-convert, resvg, imagemagick |

### Search Options

| Option | Description |
|--------|-------------|
| `--limit <n>` | Max matches |
| `--sheets <list>` | Comma-separated sheet names to search (default: all) |

### Eval Options

| Option | Alias | Description |
|--------|-------|-------------|
| `--with <overrides>` | `-w` | Cell overrides (e.g., `A1=100,B2=200`) |

### Cell Options

| Option | Description |
|--------|-------------|
| `--no-style` | Omit style info from output |

### Sheet Management Commands

| Command | Options | Description |
|---------|---------|-------------|
| `add-sheet <name>` | `--after <sheet>`, `--before <sheet>` | Add empty sheet |
| `remove-sheet <name>` | | Remove sheet |
| `rename-sheet <old> <new>` | | Rename sheet |
| `move-sheet <name>` | `--to <idx>`, `--after <sheet>`, `--before <sheet>` | Move sheet |
| `copy-sheet <src> <dest>` | | Duplicate sheet |

### Cell Operation Commands

| Command | Description |
|---------|-------------|
| `merge <range>` | Merge cells in range |
| `unmerge <range>` | Unmerge cells in range |
| `comment <ref> <text>` | Add comment to cell (`--author` for attribution) |
| `remove-comment <ref>` | Remove comment from cell |
| `stats <range>` | Calculate statistics for numeric values |
| `batch [<file>]` | Apply JSON operations (`-` for stdin) |
| `import <csv-file> [<ref>]` | Import CSV data with automatic type detection |

### Row/Column Commands

| Command | Options | Description |
|---------|---------|-------------|
| `row <n>` | `--height <pt>`, `--hide`, `--show` | Set row properties |
| `col <letter(s)>` | `--width <chars>`, `--auto-fit`, `--hide`, `--show` | Set column properties (supports ranges like A:F) |
| `autofit` | `--columns <range>` | Auto-fit all used columns (or specific range) |

### Info Commands

| Command | Description |
|---------|-------------|
| `functions` | List all 81 supported Excel functions (no file required) |
