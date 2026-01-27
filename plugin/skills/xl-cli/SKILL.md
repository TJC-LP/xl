---
name: xl-cli
description: "LLM-friendly Excel operations via the `xl` CLI. Read cells, view ranges, search, evaluate formulas, export (CSV/JSON/PNG/PDF), style cells, modify rows/columns. Use when working with .xlsx files or spreadsheet data."
---

# XL CLI - Excel Operations

## Installation

Check if installed: `which xl || echo "not installed"`

**If not installed**, download the latest native binary (no JDK required):

**macOS/Linux (recommended):**
```bash
# Auto-detect platform and install latest release
REPO="TJC-LP/xl"
LATEST=$(curl -s "https://api.github.com/repos/$REPO/releases/latest" | grep '"tag_name"' | cut -d'"' -f4)
VERSION=${LATEST#v}
case "$(uname -s)-$(uname -m)" in
  Linux-x86_64)  BINARY="xl-$VERSION-linux-amd64" ;;
  Linux-aarch64) BINARY="xl-$VERSION-linux-arm64" ;;
  Darwin-x86_64) BINARY="xl-$VERSION-darwin-amd64" ;;
  Darwin-arm64)  BINARY="xl-$VERSION-darwin-arm64" ;;
  *) echo "Unsupported: $(uname -s)-$(uname -m)" && exit 1 ;;
esac
mkdir -p ~/.local/bin
curl -sL "https://github.com/$REPO/releases/download/$LATEST/$BINARY" -o ~/.local/bin/xl
chmod +x ~/.local/bin/xl
echo "Installed xl $VERSION to ~/.local/bin/xl"
```

**Alternative using GitHub CLI:**
```bash
# If gh is installed (simpler, handles auth for private repos)
gh release download --repo TJC-LP/xl --pattern "xl-*-$(uname -s | tr A-Z a-z)-$(uname -m | sed 's/x86_64/amd64/;s/aarch64/arm64/')" -D /tmp
mv /tmp/xl-* ~/.local/bin/xl && chmod +x ~/.local/bin/xl
```

**Windows (PowerShell):**
```powershell
$repo = "TJC-LP/xl"
$latest = (Invoke-RestMethod "https://api.github.com/repos/$repo/releases/latest").tag_name
$version = $latest -replace '^v', ''
$url = "https://github.com/$repo/releases/download/$latest/xl-$version-windows-amd64.exe"
Invoke-WebRequest -Uri $url -OutFile "$env:LOCALAPPDATA\xl.exe"
Write-Host "Installed xl $version"
```

Ensure `~/.local/bin` is in your PATH: `export PATH="$HOME/.local/bin:$PATH"`

---

> **Self-Documenting CLI**: Run `xl <command> --help` for comprehensive usage, options, and examples.
> Commands like `view`, `style`, `put`, `putf`, `import`, `sort`, and `batch` have detailed built-in help.

---

## Quick Reference

### Info Commands (no file required)
```bash
xl functions                           # List all 81 supported functions
xl rasterizers                         # Check SVG-to-raster backends
```

### Read Operations
```bash
xl -f <file> sheets                    # List sheets with stats
xl -f <file> names                     # List defined names (named ranges)
xl -f <file> -s <sheet> bounds         # Used range
xl -f <file> -s <sheet> view <range>   # View as table
xl -f <file> -s <sheet> cell <ref>     # Cell details + dependencies
xl -f <file> -s <sheet> search <pattern>  # Find cells
xl -f <file> -s <sheet> stats <range>  # Calculate statistics
xl -f <file> -s <sheet> eval <formula> # Evaluate formula
```

### Output Formats
```bash
xl -f <file> -s <sheet> view <range> --format json
xl -f <file> -s <sheet> view <range> --format csv --show-labels
xl -f <file> -s <sheet> view <range> --format png --raster-output out.png
xl -f <file> -s <sheet> view <range> --formulas   # Show formulas
xl -f <file> -s <sheet> view <range> --eval       # Computed values
```

### Write Operations (require `-o`)
```bash
xl -f <file> -s <sheet> -o <out> put <ref> <value>
xl -f <file> -s <sheet> -o <out> putf <ref> <formula>
xl -f <file> -s <sheet> -o <out> style <range> --bold --bg yellow
xl -f <file> -o <out> import <csv-file> --new-sheet "Data"
```

### Row/Column Operations (require `-o`)
```bash
xl -f <file> -s <sheet> -o <out> row <n> --height 30
xl -f <file> -s <sheet> -o <out> col <letter> --width 20
xl -f <file> -s <sheet> -o <out> col A:F --auto-fit
xl -f <file> -s <sheet> -o <out> autofit              # All columns
```

### Sheet Management (require `-o`)
```bash
xl -f <file> -o <out> add-sheet "NewSheet"
xl -f <file> -o <out> remove-sheet "OldSheet"
xl -f <file> -o <out> rename-sheet "Old" "New"
xl -f <file> -o <out> copy-sheet "Template" "Copy"
```

### Cell Operations (require `-o` and `-s`)
```bash
xl -f <file> -s <sheet> -o <out> merge A1:C1
xl -f <file> -s <sheet> -o <out> sort A1:D10 --by B --header
xl -f <file> -s <sheet> -o <out> fill A1 A2:A10           # Fill down
xl -f <file> -s <sheet> -o <out> clear A1:D10 --all
xl -f <file> -s <sheet> -o <out> comment A1 "Note" --author "John"
```

### Batch Operations (require `-o`)
```bash
xl -f <file> -s <sheet> -o <out> batch operations.json
echo '[...]' | xl -f <file> -s <sheet> -o <out> batch -   # From stdin
xl batch --help                                           # Full reference
```

### Create New Workbook
```bash
xl new <output>                              # Default Sheet1
xl new <output> --sheet Data --sheet Summary # Multiple sheets
```

---

## Essential Patterns

### Sheet Selection

Commands default to first sheet. For multi-sheet files, always specify:

```bash
# Method 1: --sheet flag
xl -f data.xlsx --sheet "P&L" view A1:D10

# Method 2: Qualified A1 syntax (no -s needed)
xl -f data.xlsx view "P&L!A1:D10"
xl -f data.xlsx eval "=SUM(Revenue!A1:A10)"
```

**Workflow**: Start with `xl -f file.xlsx sheets` to discover sheet names.

### Formula Dragging (putf with range)

Single formula + range = Excel-style dragging with automatic reference shifting:

```bash
xl -f f.xlsx -s S1 -o o.xlsx putf B2:B10 "=A2*1.1"
# Result: B2: =A2*1.1, B3: =A3*1.1, B4: =A4*1.1, ...
```

**Anchor modes** ($ controls shifting):

| Syntax | Behavior |
|--------|----------|
| `$A$1` | Absolute (never shifts) |
| `$A1`  | Column absolute, row relative |
| `A$1`  | Column relative, row absolute |
| `A1`   | Fully relative (shifts both ways) |

**Running totals**:
```bash
xl -f f.xlsx -s S1 -o o.xlsx putf C2:C10 "=SUM(\$A\$1:A2)"
# Result: C2: =SUM($A$1:A2), C3: =SUM($A$1:A3), ...
```

See `xl putf --help` for full documentation.

### Batch Put & Fill

`put` supports three modes based on argument count:

```bash
# Single cell
xl ... put A1 100

# Fill pattern (same value everywhere)
xl ... put A1:A10 "TBD"

# Batch values (row-major order)
xl ... put A1:D1 "Q1" "Q2" "Q3" "Q4"
```

**Negative numbers**: Use `--value` flag (bare `-` is interpreted as flag):
```bash
xl ... put A1 --value "-100"
```

See `xl put --help` for full documentation.

### Batch JSON Operations

Apply multiple operations atomically:

```json
[
  {"op": "put", "ref": "A1", "value": "Revenue Report"},
  {"op": "style", "range": "A1:D1", "bold": true, "bg": "#4472C4", "fg": "#FFFFFF"},
  {"op": "merge", "range": "A1:D1"},
  {"op": "colwidth", "col": "A", "width": 25},
  {"op": "putf", "ref": "C2", "value": "=B2*1.1"}
]
```

**Operations**: put, putf, style, merge, unmerge, colwidth, rowheight

**Native JSON types** (recommended for numeric data):
```json
{"op": "put", "ref": "A1", "value": 99.0}                    // Number
{"op": "put", "ref": "A2", "value": true}                    // Boolean
{"op": "put", "ref": "A3", "value": 99.0, "format": "currency"}  // $99.00
{"op": "put", "ref": "A4", "value": 0.594, "format": "percent"}  // 59%
{"op": "put", "ref": "A5", "value": 3.5, "format": "0.0x"}   // Custom: 3.5x
```

**Smart detection** (auto-formats string values):
```json
{"op": "put", "ref": "A1", "value": "$1,234.56"}    // → Currency
{"op": "put", "ref": "A2", "value": "59.4%"}        // → Percent (stored as 0.594)
{"op": "put", "ref": "A3", "value": "2025-11-10"}   // → Date
```

**Format names**: `general`, `integer`, `decimal`, `currency`, `percent`, `date`, `datetime`, `time`, `text`, or any custom Excel format code (e.g., `"0.0x"`, `"$#,##0;($#,##0)"`).

**Disable detection**: Set `"detect": false` to treat strings as plain text:
```json
{"op": "put", "ref": "A1", "value": "$99.00", "detect": false}  // → Text, not Currency
```

**Formula dragging** (putf with range):
```json
// Drag formula across range (Excel-style $ anchoring)
{"op": "putf", "ref": "B2:B10", "value": "=SUM($A$1:A2)", "from": "B2"}

// Explicit formulas for each cell (no dragging)
{"op": "putf", "ref": "B2:B4", "values": ["=A2*2", "=A3*2", "=A4*2"]}
```

**Style properties**: Use `align` (not `halign`) for horizontal alignment. Unknown properties emit warnings.

See `xl batch --help` for full reference with all style properties.

### Output Format Summary

| Format | Flag | Notes |
|--------|------|-------|
| markdown | Default | Text table |
| json | `--format json` | Structured data |
| csv | `--format csv` | Add `--show-labels` for headers |
| html | `--format html` | Inline CSS |
| svg | `--format svg` | Vector graphics |
| png/jpeg/pdf | `--format <fmt> --raster-output <path>` | Requires rasterizer |
| webp | `--format webp --raster-output <path>` | ImageMagick only |

**Rasterizer discovery**: `xl rasterizers` shows available backends.

See `xl view --help` for all options.

---

## Workflows

### Explore Unknown Spreadsheet

```bash
xl -f data.xlsx sheets                     # List sheets with cell counts
xl -f data.xlsx names                      # List defined names
xl -f data.xlsx -s "Sheet1" bounds         # Get used range
xl -f data.xlsx -s "Sheet1" view A1:E20    # Preview data
xl -f data.xlsx -s "Sheet1" stats B2:B100  # Quick statistics
```

### Formula Analysis & What-If

```bash
xl -f data.xlsx -s Sheet1 view --formulas A1:D10     # Show formulas
xl -f data.xlsx -s Sheet1 cell C5                    # Dependencies
xl -f data.xlsx -s Sheet1 eval "=SUM(A1:A10)" --with "A1=500"  # What-if
```

See [reference/FORMULAS.md](reference/FORMULAS.md) for 81 supported functions.

### Create Formatted Report

```bash
# Set data and styling
xl -f template.xlsx -s Sheet1 -o report.xlsx put A1 "Sales Report"
xl -f report.xlsx -s Sheet1 -o report.xlsx style A1:E1 --bold --bg navy --fg white
xl -f report.xlsx -s Sheet1 -o report.xlsx style B2:B100 --format currency
xl -f report.xlsx -s Sheet1 -o report.xlsx col A --width 25
```

Or use batch for atomicity:
```bash
echo '[
  {"op": "put", "ref": "A1", "value": "Sales Report"},
  {"op": "style", "range": "A1:E1", "bold": true, "bg": "navy", "fg": "white"},
  {"op": "style", "range": "B2:B100", "numFormat": "currency"},
  {"op": "colwidth", "col": "A", "width": 25}
]' | xl -f template.xlsx -s Sheet1 -o report.xlsx batch -
```

### CSV to Styled Table

```bash
# Import CSV to new sheet
xl -f workbook.xlsx -o out.xlsx import data.csv --new-sheet "Imported"

# Style the header row
xl -f out.xlsx -s Imported -o out.xlsx style A1:Z1 --bold --bg navy --fg white

# Auto-fit columns
xl -f out.xlsx -s Imported -o out.xlsx autofit
```

Import options: `xl import --help`

### Multi-Sheet Workbook Setup

```bash
# Create with multiple sheets
xl new output.xlsx --sheet Data --sheet Summary --sheet Notes

# Or add sheets to existing
xl -f output.xlsx -o output.xlsx add-sheet "Archive" --after "Notes"
xl -f output.xlsx -o output.xlsx copy-sheet "Summary" "Q1 Summary"

# Move sheet to front
xl -f output.xlsx -o output.xlsx move-sheet "Summary" --to 0
```

### Visual Analysis (for Claude Vision)

```bash
xl -f data.xlsx -s Sheet1 view A1:F20 --format png --raster-output /tmp/sheet.png --show-labels
```

### Large File Operations (100k+ rows)

For files with 100k+ rows, use streaming mode for O(1) memory:

**Streaming Read:**
```bash
xl -f huge.xlsx --stream search "pattern" --limit 10    # ~10s for 1M rows
xl -f huge.xlsx --stream stats A1:E100000               # Aggregate without loading
xl -f huge.xlsx --stream bounds                          # Get used range
xl -f huge.xlsx --stream view A1:D100 --format csv      # Export range
```

**Streaming Write:**
```bash
xl -f huge.xlsx -o out.xlsx --stream put A1 "Header"           # Put values
xl -f huge.xlsx -o out.xlsx --stream putf A2 "=B2*1.1"         # Put formulas
xl -f huge.xlsx -o out.xlsx --stream style A1:Z1 --bold --bg navy  # Apply styles
```

**Performance:** Styling row 1 of 100k rows: ~0.3s (early-abort optimization)

**In-memory mode** (when streaming not supported):
```bash
xl -f huge.xlsx --max-size 0 sheets       # Disable 100MB limit
xl -f huge.xlsx --max-size 500 cell A1    # Custom 500MB limit
```

**Streaming supports**: search, stats, bounds, view (markdown/csv/json), put, putf, style

**Requires in-memory**: cell (dependencies), eval (formulas), HTML/SVG/PDF (styles), formula dragging

---

## Command Reference

### Global Options

| Option | Alias | Description |
|--------|-------|-------------|
| `--file <path>` | `-f` | Input file (required) |
| `--sheet <name>` | `-s` | Sheet name |
| `--output <path>` | `-o` | Output file (for writes) |
| `--backend <type>` | | scalaxml (default) or saxstax (36-39% faster) |
| `--max-size <MB>` | | Override 100MB security limit (0 = unlimited) |
| `--stream` | | O(1) memory mode for reads + writes (search/stats/bounds/view/put/putf/style) |

### Info Commands

| Command | Description |
|---------|-------------|
| `functions` | List all 81 supported Excel functions |
| `rasterizers` | List SVG-to-raster backends with status |

### Workbook Commands

| Command | Description |
|---------|-------------|
| `sheets` | List all sheets with stats |
| `names` | List defined names (named ranges) |
| `new <output>` | Create blank workbook (`--sheet` for names) |

### Read Commands

| Command | Options |
|---------|---------|
| `bounds` | Used range of sheet |
| `view <range>` | `--format`, `--formulas`, `--eval`, `--raster-output`, etc. |
| `cell <ref>` | `--no-style` |
| `search <pattern>` | `--limit`, `--sheets` |
| `stats <range>` | Calculate count, sum, min, max, mean |
| `eval <formula>` | `--with` for overrides |

Run `xl view --help` for complete options.

### Write Commands

| Command | Key Options |
|---------|-------------|
| `put <ref> <values>` | `--value` for negatives, `--stream` for O(1) memory |
| `putf <ref> <formulas>` | Supports dragging (no dragging with `--stream`) |
| `style <range>` | `--bold`, `--bg`, `--fg`, `--format`, `--border`, `--stream` for O(1) memory |
| `batch <json-file>` | Operations: put, putf, style, merge, unmerge, colwidth, rowheight |
| `import <csv> [ref]` | `--new-sheet`, `--delimiter`, `--no-type-inference` |

Run `xl <command> --help` for complete options.

### Sheet Management Commands

| Command | Options |
|---------|---------|
| `add-sheet <name>` | `--after`, `--before` |
| `remove-sheet <name>` | |
| `rename-sheet <old> <new>` | |
| `move-sheet <name>` | `--to`, `--after`, `--before` |
| `copy-sheet <src> <dest>` | |

### Cell Commands

| Command | Options |
|---------|---------|
| `merge <range>` | |
| `unmerge <range>` | |
| `comment <ref> <text>` | `--author` |
| `remove-comment <ref>` | |
| `clear <range>` | `--all`, `--styles`, `--comments` |
| `fill <source> <target>` | `--right` |
| `sort <range>` | `--by`, `--then-by`, `--desc`, `--numeric`, `--header` |

Run `xl sort --help` for sorting details.

### Row/Column Commands

| Command | Options |
|---------|---------|
| `row <n>` | `--height`, `--hide`, `--show` |
| `col <letter(s)>` | `--width`, `--auto-fit`, `--hide`, `--show` |
| `autofit` | `--columns` (range like A:Z) |

---

## Links

- `xl <command> --help` for detailed usage and examples
- [reference/FORMULAS.md](reference/FORMULAS.md) for 81 supported functions
- [reference/COLORS.md](reference/COLORS.md) for color names
- [reference/OUTPUT-FORMATS.md](reference/OUTPUT-FORMATS.md) for format specs
