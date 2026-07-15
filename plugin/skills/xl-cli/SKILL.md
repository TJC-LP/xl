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
  # linux-arm64 binaries ship from v0.12.6 (xl#354); on older releases the curl below
  # fails loud — fall back to the JAR distribution (xl-$VERSION-cli.tar.gz, needs a JRE).
  Linux-aarch64) BINARY="xl-$VERSION-linux-arm64" ;;
  Darwin-x86_64) BINARY="xl-$VERSION-darwin-amd64" ;;
  Darwin-arm64)  BINARY="xl-$VERSION-darwin-arm64" ;;
  *) echo "Unsupported: $(uname -s)-$(uname -m)" && exit 1 ;;
esac
mkdir -p ~/.local/bin
curl -fsSL "https://github.com/$REPO/releases/download/$LATEST/$BINARY" -o ~/.local/bin/xl || {
  echo "Error: no $BINARY published for $LATEST — use the JAR distribution xl-cli-$VERSION.tar.gz instead" >&2
  exit 1
}
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
xl functions                           # List all 107 supported functions
xl rasterizers                         # Check SVG-to-raster backends
```

### Read Operations
```bash
xl -f <file> sheets                    # List sheets with visibility state
xl -f <file> names                     # List defined names (named ranges)
xl -f <file> -s <sheet> bounds         # Used range
xl -f <file> -s <sheet> view <range>   # View as table
xl -f <file> -s <sheet> cell <ref>     # Cell details + dependencies
xl -f <file> -s <sheet> search <pattern>  # Find cells
xl -f <file> -s <sheet> stats <range>  # Calculate statistics
xl -f <file> -s <sheet> eval <formula> # Evaluate formula
xl -f <file> -s <sheet> evala <formula>          # Array formula result grid
xl -f <file> -s <sheet> evala <formula> --at <ref>  # Spill to target cell
```

### Output Formats
```bash
xl -f <file> -s <sheet> view <range> --format json
xl -f <file> -s <sheet> view <range> --format csv --show-labels
xl -f <file> -s <sheet> view <range> --format png --raster-output out.png
xl -f <file> -s <sheet> view <range> --formulas   # Show formulas
xl -f <file> -s <sheet> view <range> --eval       # Computed values
```

**Note**: since 0.12.5, `--eval` computes deep multi-hop cross-sheet chains in one pass (global dependency order with memoization). Since 0.12.6, formulas containing **external-workbook references** (`[2]Book!A1`) evaluate from their Excel-written cached values ([#353](https://github.com/TJC-LP/xl/issues/353)); on ≤0.12.5 they fail with `UnexpectedChar([` — read those cells with a plain `view` there.

### Write Operations (require `-o` or `-i`)
```bash
xl -f <file> -s <sheet> -o <out> put <ref> <value>
xl -f <file> -s <sheet> -i put <ref> <value>         # In-place edit (no -o needed)
xl -f <file> -s <sheet> -o <out> putf <ref> <formula>
xl -f <file> -s <sheet> -o <out> style <range> --bold --bg yellow
xl -f <file> -s <sheet> -o <out> copy <source> <target>  # Range copy with formula shift
xl -f <file> -s <sheet> -o <out> freeze <ref>         # Freeze panes
xl -f <file> -s <sheet> -o <out> unfreeze             # Remove freeze panes
xl -f <file> -o <out> import <csv-file> --new-sheet "Data"
xl -f <file> -o <out> import-md <table.md> --start A1   # GFM markdown table import (0.11.3+; '-' = stdin)
```

### Compare & Query (read-only, 0.11.3+)
```bash
xl -f a.xlsx diff -g b.xlsx --format markdown          # Workbook diff (exit 0 identical, 1 differs)
xl -f a.xlsx diff -g b.xlsx --format json              # Stable JSON schema for tooling
xl -f <file> -s <sheet> filter --where "B > 100 AND D = TRUE" --header --format csv
xl -f <file> -s <sheet> filter --where "Name LIKE 'Acme%'" --columns A,C:E --limit 20
```

### Images & Charts (require `-o`, 0.12.0+)
```bash
xl -f <file> -s <sheet> -o <out> add-image logo.png --at B2          # Embed picture (png/jpeg/gif/bmp...)
xl -f <file> -s <sheet> -o <out> chart add --type bar --data B2:D10 --categories A2:A10 --title "Revenue" --at F2:K15
```

### Row/Column Operations (require `-o`)
```bash
xl -f <file> -s <sheet> -o <out> row <n> --height 30
xl -f <file> -s <sheet> -o <out> col <letter> --width 20
xl -f <file> -s <sheet> -o <out> col A:F --auto-fit
xl -f <file> -s <sheet> -o <out> autofit              # All columns
```

### Structural Editing (require `-o`)

Insert/delete rows and columns Excel-style: cells, merges, row/column properties, and freeze
panes shift, and every affected formula (including cross-sheet references) is rewritten.
References to deleted cells become `#REF!`.

```bash
xl -f <file> -s <sheet> -o <out> insert-rows 5        # Insert 1 row before row 5
xl -f <file> -s <sheet> -o <out> insert-rows 5 3      # Insert 3 rows before row 5
xl -f <file> -s <sheet> -o <out> delete-rows 5 2      # Delete rows 5-6
xl -f <file> -s <sheet> -o <out> insert-cols C 2      # Insert 2 columns at C
xl -f <file> -s <sheet> -o <out> delete-cols C        # Delete column C
xl -f <file> -s <sheet> -o <out> delete-cols C:E      # Delete columns C through E
```

### Sheet Management (require `-o`)
```bash
xl -f <file> -o <out> add-sheet "NewSheet"
xl -f <file> -o <out> remove-sheet "OldSheet"
xl -f <file> -o <out> rename-sheet "Old" "New"
xl -f <file> -o <out> copy-sheet "Template" "Copy"
xl -f <file> -o <out> sheets hide "Archive"        # Hide from tabs
xl -f <file> -o <out> sheets hide "Internal" --very # Very hidden (VBA only)
xl -f <file> -o <out> sheets show "Archive"        # Unhide
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

**put vs putf**: `putf` always interprets input as a formula. Using `putf` for text like "Total Revenue" will cause a parse error. Use `put` for text labels, `putf` for formulas.

### Cross-Sheet References in Formulas

**CRITICAL**: Cross-sheet references use Excel's `!` operator (NOT `.` or other separators):

```bash
# Single cell from another sheet
xl -f f.xlsx -s Summary -o o.xlsx putf A1 "=Data!B5"

# Range from another sheet
xl -f f.xlsx -s Summary -o o.xlsx putf A1 "=SUM(Data!A1:A100)"

# SUMIFS with cross-sheet references (common pattern)
xl -f f.xlsx -s Summary -o o.xlsx putf H2 "=SUMIFS(Data!D:D,Data!A:A,A2,Data!C:C,E2)"

# Sheet names with spaces require single quotes AROUND the sheet name
xl -f f.xlsx -s Summary -o o.xlsx putf A1 "=SUM('Q1 Sales'!A1:A100)"
```

**Shell escaping**: The `!` character has special meaning in bash. Use single quotes around the formula:
```bash
# ✓ Correct - single quotes protect !
xl -f f.xlsx -s S1 putf A1 '=Sheet2!B1'

# ✗ Wrong - double quotes allow ! expansion in bash
xl -f f.xlsx -s S1 putf A1 "=Sheet2!B1"  # May fail with "event not found"
```

### Shell Quoting for Sheet Names with Spaces

The parser fully supports `='Sheet Name'!A1` syntax. Use double quotes around the CLI argument so the shell passes the string intact:

```bash
xl -f f.xlsx -s Summary -o o.xlsx putf B4 "='Income Statement'!G8"
xl -f f.xlsx -s Summary -o o.xlsx putf A1 "=SUM('Q1 Sales'!A1:A100)"
```

For complex cases, batch JSON avoids shell quoting entirely:
```bash
echo '[{"op":"putf","ref":"B4","value":"='"'"'Income Statement'"'"'!G8"}]' | xl -f f.xlsx -s Summary -o o.xlsx batch -
```

Alternatively, rename sheets to avoid spaces when CLI manipulation is planned.

### Batch Put & Fill

`put` supports three modes based on argument count:

```bash
# Single cell
xl ... put A1 100

# Fill pattern (same value everywhere)
xl ... put A1:A10 "TBD"

# Batch values (row-major order)
xl ... put A1:D1 "Q1" "Q2" "Q3" "Q4"

# CSV split (opt-in: requires --csv flag)
xl ... put A1:D1 "Q1,Q2,Q3,Q4" --csv
```

**`--csv`**: Opt-in flag that splits a single comma-separated value across the target range. Required because comma-containing values are common in real data (`"Smith, John"`); without `--csv`, the value is written as literal text. The split count must match the range size exactly, otherwise the command errors. Smart type detection applies to each split value.

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

**Operations**: put, putf, style, merge, unmerge, colwidth, rowheight, and more — all 21 listed under "All batch operations" below

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

**Format names**: `general`, `integer`, `decimal`, `currency`, `percent`, `date`, `datetime`, `time`, `text`, or any custom Excel format code (e.g., `"0.0x"`, `"$#,##0;($#,##0)"`, `"#,##0.0_);(#,##0.0)"` for accounting).

**Disable detection**: Set `"detect": false` to treat strings as plain text:
```json
{"op": "put", "ref": "A1", "value": "$99.00", "detect": false}  // → Text, not Currency
```

**Batch values** (put/putf with range):
```json
// Put multiple values in row-major order (supports smart detection)
{"op": "put", "ref": "A1:E1", "values": ["Date", "Company", "Revenue", "Growth", "Status"]}
{"op": "put", "ref": "B2:B4", "values": [1234.56, 5678.90, 9012.34]}
{"op": "put", "ref": "C2:C4", "values": ["$1,234", "$5,678", "$9,012"]}

// Single formula (use "value" or "formula" — both accepted)
{"op": "putf", "ref": "D14", "value": "=SUM(D5:D12)"}
{"op": "putf", "ref": "D14", "formula": "=SUM(D5:D12)"}

// Drag formula across range (Excel-style $ anchoring)
{"op": "putf", "ref": "B2:B10", "value": "=SUM($A$1:A2)", "from": "B2"}

// Explicit formulas for each cell (no dragging)
{"op": "putf", "ref": "B2:B4", "values": ["=A2*2", "=A3*2", "=A4*2"]}
```

**Style properties** (batch JSON property names):

| CLI Flag | JSON Property | Type |
|----------|--------------|------|
| `--bold` | `bold` | boolean |
| `--italic` | `italic` | boolean |
| `--underline` | `underline` | boolean |
| `--bg` | `bg` | string (color name or #hex) |
| `--fg` | `fg` | string (color name or #hex) |
| `--font-size` | `fontSize` | number |
| `--font-name` | `fontName` | string |
| `--format` | `numFormat` | string (format name or code) |
| `--align` | `align` | string (left/center/right) |
| `--valign` | `valign` | string (top/center/bottom) |
| `--wrap` | `wrap` | boolean |
| `--border` | `border` | string (thin/medium/thick) |
| `--border-color` | `borderColor` | string (color) |
| `--replace` | `replace` | boolean (default: merge) |

Use `align` (not `halign`) for horizontal alignment. The JSON property for number format is `numFormat` (camelCase), not `format`. Unknown properties emit warnings.

See `xl batch --help` for full reference.

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

**Note**: `--format html` does not apply cell styles or number formats. Use `--format png` (via rasterizer) for styled output.

**Rasterizer discovery**: `xl rasterizers` shows available backends.

**Installing a rasterizer** (needed for PNG/JPEG/PDF/WebP export):
```bash
# macOS
brew install librsvg

# Linux (Debian/Ubuntu)
apt install librsvg2-bin

# Python alternative
pip install cairosvg
```

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
xl -f data.xlsx -s Sheet1 eval "=SUM(A1:A5)" --with "A1=0,A5=0"  # Multiple overrides (comma-separated)
```

See [reference/FORMULAS.md](reference/FORMULAS.md) for 107 supported functions.

### Create Formatted Report

```bash
# Set data and styling
xl -f template.xlsx -s Sheet1 -o report.xlsx put A1 "Sales Report"
xl -f report.xlsx -s Sheet1 -o report.xlsx style A1:E1 --bold --bg navy --fg white
xl -f report.xlsx -s Sheet1 -o report.xlsx style B2:B100 --format currency
xl -f report.xlsx -s Sheet1 -o report.xlsx style C2:C100 --format "#,##0.00"   # Custom decimal
xl -f report.xlsx -s Sheet1 -o report.xlsx style D2:D100 --format "0.0%"       # Custom percent
xl -f report.xlsx -s Sheet1 -o report.xlsx style E2:E100 --format "yyyy-mm-dd" # Custom date
xl -f report.xlsx -s Sheet1 -o report.xlsx style F2:F100 --format "0.0x"       # Multiples
xl -f report.xlsx -s Sheet1 -o report.xlsx col A --width 25
```

Or use batch for atomicity (preferred for multi-step operations):
```bash
echo '[
  {"op": "put", "ref": "A1", "value": "Sales Report"},
  {"op": "style", "range": "A1:E1", "bold": true, "bg": "navy", "fg": "white"},
  {"op": "style", "range": "B2:B100", "numFormat": "currency"},
  {"op": "colwidth", "col": "A", "width": 25},
  {"op": "comment", "ref": "A1", "text": "Generated report", "author": "Agent"},
  {"op": "autofit", "columns": "A:E"},
  {"op": "row-hide", "row": 2},
  {"op": "add-sheet", "name": "Summary", "after": "Sheet1"}
]' | xl -f template.xlsx -s Sheet1 -o report.xlsx batch -
```

**Dry-run validation** (validate JSON and show summary without writing):
```bash
# Standalone (no --file or --output needed)
echo '[{"op":"putf","ref":"A1","formula":"=SUM(B1:B10)"},{"op":"style","range":"A1","bold":true}]' | xl batch --dry-run -

# Also works on existing batch invocations — skips read/write, just validates
echo '[{"op":"put","ref":"A1","value":"test"}]' | xl -f in.xlsx -o out.xlsx batch --dry-run -
```

**All batch operations** (21): `put`, `putf`, `style`, `merge`, `unmerge`, `colwidth`, `rowheight`, `comment`, `remove-comment`, `hyperlink`, `clear`, `col-hide`, `col-show`, `row-hide`, `row-show`, `autofit`, `add-sheet`, `rename-sheet`, `freeze`, `unfreeze`, `copy`

**Freeze/unfreeze/copy/hyperlink in batch:**
```json
{"op": "freeze", "ref": "B2"}
{"op": "unfreeze"}
{"op": "copy", "source": "A1:D10", "target": "F1"}
{"op": "copy", "source": "A1:D10", "target": "F1", "valuesOnly": true}
{"op": "hyperlink", "ref": "A1", "target": "https://example.com"}
{"op": "hyperlink", "ref": "A1"}
```

`hyperlink` sets a cell's hyperlink target; omit `target` to clear an existing hyperlink.

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

# Move sheet to front (may affect cross-sheet formula references; verify formulas after reordering)
xl -f output.xlsx -o output.xlsx move-sheet "Summary" --to 0

# Hide internal sheets from users
xl -f output.xlsx -o output.xlsx sheets hide "Notes"
xl -f output.xlsx -o output.xlsx sheets hide "Config" --very  # VBA-only
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

**Note**: Streaming CSV shows formula expressions without the `=` prefix (streaming mode reads raw cell content).

**Streaming supports**: search, stats, bounds, view (markdown/csv/json), put, putf, style

**Requires in-memory**: cell (dependencies), eval (formulas), HTML/SVG/PDF (styles), formula dragging, `put --csv` (CSV auto-split)

---

## Field Gotchas (verified in production)

Hard-won rules from fleet use on real deal workbooks. Items marked **fixed in 0.12.6** still apply when the installed binary is older — check `xl --version`.

**Reading**
- **`view` and `search` clip at `--limit` (default 50)** ([#351](https://github.com/TJC-LP/xl/issues/351)). Since 0.12.6 the clip is visible (`… showing N of M rows` trailer; stderr notice for csv; `truncated`/`totalRows` in json; `--limit 0` = unlimited). **On ≤0.12.5 the clip is SILENT** — a `view A1:A259` returning ~50 rows is NOT the whole range; pass `--limit <n>` explicitly and verify the last expected row is present.
- **Use `--show-labels` whenever row numbers matter** in CSV output — hidden rows shift positional counting silently.
- **`--formulas --format json` replaces `value` with the formula text** and emits no separate `formula` key ([#357](https://github.com/TJC-LP/xl/issues/357)). To scan formulas programmatically, use `--format csv --formulas --show-labels` and parse with a CSV reader.
- **Percent literals fail to evaluate** (`=A1*10%`, [#355](https://github.com/TJC-LP/xl/issues/355)) — write `/100` instead. External-workbook refs are **fixed in 0.12.6** (evaluate from Excel-written caches; uncached external cells report a clear per-cell error); on ≤0.12.5 any formula whose precedents include one dies with `UnexpectedChar([` — use plain `view` there.

**Writing**
- **`batch` recalculates at the end since 0.12.6** ([#352](https://github.com/TJC-LP/xl/issues/352)): `putf` cells carry cached values and formula errors appear in the batch summary; `xl recalc` (also 0.12.6+) refreshes any workbook. **On ≤0.12.5 batch writes formulas with NO cached values** — `data_only` readers, pandas, and previewers show blanks for every `putf` cell; prefer single `putf` commands there, verify with `--eval`, or recalculate via the scripting API before shipping.
- **Batch `putf` has no `format` field** ([#356](https://github.com/TJC-LP/xl/issues/356)) — number formats on formula cells need a follow-up `style` op / `xl style RANGE --format '…'`.
- **Batch JSON style keys are camelCase**: `numFormat`, `borderTop`, `borderBottom`, `borderLeft`, `borderRight`, `borderColor`. Kebab-case (`border-top`) is treated as an unknown property — the warning goes to stderr only, so it is easy to miss.
- **Build cross-sheet models in dependency order** (put referenced cells before the formulas that use them) — on ≤0.12.5 batches evaluate nothing and single-command recalcs are dependents-only; on 0.12.6+ the batch-end recalculation makes order matter less, but dependency order remains the safe habit.
- **`-i` (in-place) round-trips preserve hand-patched XML**: sheetPr/tabColor, pageSetup, headerFooter, sheetView gridlines/zoom, sheet-local `_xlnm.Print_Area`, and hand-inserted cached `<v>` values all survive later xl writes.
- **Sheet appearance/print setup (gridlines off, zoom, tab color, landscape, footers) has no CLI yet** ([#358](https://github.com/TJC-LP/xl/issues/358)) — until then it requires zip round-trip XML patching.

**Rendering & environment**
- **PNG/PDF on the native binary needs an external rasterizer** (Batik is JAR-only; [#359](https://github.com/TJC-LP/xl/issues/359)): install `cairosvg` (`pip install cairosvg`) or `rsvg-convert`, and check availability with `xl rasterizers`. Add `--eval` when rendering or formula cells show as text.
- **Real-world workbooks that fail to read**: since 0.12.6, benign `<!DOCTYPE>` prologs are tolerated and parse errors name the construct with line/column ([#350](https://github.com/TJC-LP/xl/issues/350)); the native binary reports real Xerces messages instead of the opaque `Could not load any resource bundle …XMLMessages` ([#349](https://github.com/TJC-LP/xl/issues/349)). If a file still refuses to read (on any version): read with openpyxl, rebuild clean, then xl works on the rebuilt file — and report the 0.12.6+ error message upstream.

---

## Command Reference

### Global Options

| Option | Alias | Description |
|--------|-------|-------------|
| `--file <path>` | `-f` | Input file (required) |
| `--sheet <name>` | `-s` | Sheet name |
| `--output <path>` | `-o` | Output file (for writes) |
| `--in-place` | `-i` | Edit file in-place (mutually exclusive with `-o`) |
| `--backend <type>` | | Write backend: scalaxml (default) or saxstax (36-39% faster). Reads always use StAX. |
| `--max-size <MB>` | | Override 100MB security limit (0 = unlimited) |
| `--stream` | | O(1) memory mode for reads + writes (search/stats/bounds/view/put/putf/style) |
| `--dry-run` | | Validate batch JSON and show summary without writing (batch only) |

### Info Commands

| Command | Description |
|---------|-------------|
| `functions` | List all 107 supported Excel functions |
| `rasterizers` | List SVG-to-raster backends with status |

### Workbook Commands

| Command | Description |
|---------|-------------|
| `sheets` | List sheets with visibility state |
| `sheets list` | Explicit list (`--stats` for cell counts) |
| `sheets hide <name>` | Hide sheet (`--very` for VBA-only access) |
| `sheets show <name>` | Unhide sheet |
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
| `eval <formula>` | `--with` for overrides (comma-separated: `--with "A1=0,A5=0"`) |
| `evala <formula>` | `--at` to spill result starting at ref |

Run `xl view --help` for complete options.

### Write Commands

| Command | Key Options |
|---------|-------------|
| `put <ref> <values>` | `--value` for negatives, `--stream` for O(1) memory, `--csv` to split comma-separated value |
| `putf <ref> <formulas>` | Supports dragging (no dragging with `--stream`) |
| `style <range>` | `--bold`, `--bg`, `--fg`, `--format`, `--border`, `--stream` for O(1) memory |
| `copy <source> <target>` | `--values-only` (no formula adjustment) |
| `freeze <ref>` | Freeze panes (rows above + columns left of ref) |
| `unfreeze` | Remove freeze panes |
| `batch <json-file>` | 21 operations (see below) |
| `import <csv> [ref]` | `--new-sheet`, `--delimiter`, `--no-type-inference` |

Run `xl <command> --help` for complete options.

### Sheet Management Commands

| Command | Options |
|---------|---------|
| `sheets hide <name>` | `--very` (VBA-only access) |
| `sheets show <name>` | |
| `add-sheet <name>` | `--after`, `--before` |
| `remove-sheet <name>` | |
| `rename-sheet <old> <new>` | |
| `move-sheet <name>` | `--to`, `--after`, `--before` (may affect cross-sheet refs) |
| `copy-sheet <src> <dest>` | |

### Cell Commands

| Command | Options |
|---------|---------|
| `merge <range>` | |
| `unmerge <range>` | |
| `copy <source> <target>` | `--values-only` |
| `freeze <ref>` | Rows above + columns left of ref are locked |
| `unfreeze` | |
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

### Structural Editing Commands

| Command | Arguments |
|---------|-----------|
| `insert-rows <at-row> [count]` | Insert `count` rows (default 1) before 1-based row `at-row` |
| `delete-rows <at-row> [count]` | Delete `count` rows (default 1) starting at row `at-row` |
| `insert-cols <at-col> [count]` | Insert `count` columns (default 1) at column letter `at-col` |
| `delete-cols <at-col> [count]` | Delete `count` columns (default 1) starting at `at-col` |

Column commands also accept an inclusive range (`delete-cols C:E`), which derives the count and
overrides the positional `count` argument. All four shift cells, merges, row/column properties,
and freeze panes, then rewrite affected formulas across all sheets; formulas referencing deleted
cells get `#REF!` and ranges shrink correctly.

---

## Links

- `xl <command> --help` for detailed usage and examples
- [reference/FORMULAS.md](reference/FORMULAS.md) for 107 supported functions
- [reference/COLORS.md](reference/COLORS.md) for color names
- [reference/OUTPUT-FORMATS.md](reference/OUTPUT-FORMATS.md) for format specs
