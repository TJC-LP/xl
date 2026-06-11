# XL CLI Reference

The `xl` CLI provides a command-line interface for Excel operations, designed for LLM agents and automation.

**Design Philosophy**:
- **Stateless** — Each command is self-contained
- **Explicit cell refs** — Always use `A1`, `B5:D10` notation
- **Global flags** — `-f` for file, `-s` for sheet, `-o` for output
- **LLM-optimized output** — Markdown tables, token-efficient

---

## Installation

```bash
git clone https://github.com/TJC-LP/xl.git
cd xl
make install
```

This builds a native binary (no JDK required, instant startup) and installs `xl` to `~/.local/bin/`. Ensure it's in your PATH:

```bash
export PATH="$HOME/.local/bin:$PATH"
```

**JAR install** (requires JDK 17+): `make install-jar`

**Uninstall**: `make uninstall`

**Update**: After `git pull`, run `make install` (or `make install-jar`) again.

---

## Quick Reference

```bash
# Global flags (used with all commands)
-f, --file <path>     # Input file (required for most commands)
-s, --sheet <name>    # Sheet to operate on (required for unqualified ranges)
-o, --output <path>   # Output file for mutations
-i, --in-place        # Edit file in place (same as -o matching -f)
--stream              # O(1) memory streaming for large files (search/stats/bounds/view + writes)
--max-size <MB>       # Max uncompressed size for in-memory load (default 100, 0 = unlimited)
--backend <name>      # XML backend: scalaxml (default) or saxstax (faster)

# Read-only operations
xl -f model.xlsx sheets                    # List all sheets
xl -f model.xlsx names                     # List defined names (named ranges)
xl -f model.xlsx -s "P&L" bounds           # Show used range
xl -f model.xlsx -s "P&L" view A1:D20      # View range as markdown
xl -f model.xlsx cell B5                   # Get single cell details (sheet auto-detected if unambiguous)
xl -f model.xlsx search "Revenue"          # Find cells by content (all sheets)
xl -f model.xlsx -s "P&L" stats B1:B100    # Numeric statistics for a range
xl -f model.xlsx -s "P&L" eval "=SUM(B1:B10)"   # Evaluate formula (what-if)
xl -f model.xlsx -s "P&L" evala "=TRANSPOSE(A1:C2)"  # Evaluate array formula (spill grid)

# Mutations (require -o or -i)
xl -f model.xlsx -s S1 -o output.xlsx put B5 1000000       # Write value
xl -f model.xlsx -s S1 -o output.xlsx putf C5 "=B5*1.1"    # Write formula

# What-if analysis with overrides
xl -f model.xlsx -s S1 eval "=B1*1.1" --with "B1=100"      # Evaluate with temporary values

# No file needed
xl new report.xlsx --sheet Data --sheet Summary   # Create a blank workbook
xl functions                                       # List all 104 supported functions
xl rasterizers                                     # List available PNG/PDF backends
```

---

## Commands

### Operation Categories

| Category | Commands | Purpose |
|----------|----------|---------|
| **Info** (no `-f`) | `functions`, `rasterizers`, `new` | Capability listing, blank workbook |
| **Navigate** | `sheets`, `bounds`, `names` | Find your way around |
| **Explore** | `view`, `cell`, `search`, `stats` | Read data incrementally |
| **Analyze** | `eval`, `evala` | What-if formula evaluation (scalar + array) |
| **Mutate cells** | `put`, `putf`, `style`, `fill`, `clear`, `copy`, `sort`, `merge`, `unmerge`, `comment`, `remove-comment`, `batch`, `import` | Make changes (require `-o` or `-i`) |
| **Rows/columns** | `row`, `col`, `autofit`, `insert-rows`, `delete-rows`, `insert-cols`, `delete-cols` | Sizing, visibility, structural editing |
| **Sheets & view** | `add-sheet`, `remove-sheet`, `rename-sheet`, `move-sheet`, `copy-sheet`, `sheets hide/show`, `freeze`, `unfreeze`, `name` | Workbook structure |

### Command Summary

| Command | Arguments | Description |
|---------|-----------|-------------|
| `functions` | | List all 104 supported Excel functions (no `-f` needed) |
| `rasterizers` | | List available SVG-to-raster backends (no `-f` needed) |
| `new` | `<output> [--sheet <name>]...` | Create a blank xlsx file (no `-f` needed) |
| `sheets` | `[list\|hide <name> [--very]\|show <name>]` | List sheets (default) or hide/show one |
| `names` | | List defined names (named ranges) |
| `name` | `add <name> <refers-to>` \| `rm <name>` | Manage named ranges (requires `-o`) |
| `bounds` | `[--scan]` | Show used range of current sheet |
| `view` | `<range> [flags]` | Render range (markdown/json/csv/html/svg/png/jpeg/webp/pdf) |
| `cell` | `<ref> [--no-style]` | Get single cell details |
| `search` | `<pattern> [--limit n] [--sheets a,b]` | Find cells matching pattern (regex, all sheets by default) |
| `stats` | `<range>` | Statistics for numeric values in range |
| `eval` | `<formula> [--with overrides]` | Evaluate formula without modifying |
| `evala` | `<formula> [--at <ref>] [--with overrides]` | Evaluate array formula; display or spill result grid |
| `put` | `<ref\|range> <value...> [--csv]` | Write value(s) to cell or range (requires `-o`) |
| `putf` | `<ref\|range> <formula...>` | Write formula(s); single formula + range drags with `$` anchors (requires `-o`) |
| `style` | `<range> [options]` | Apply styling (requires `-o`) |
| `row` | `<n> [--height pt] [--hide\|--show]` | Set row properties (requires `-o`) |
| `col` | `<letter\|A:F> [--width n] [--auto-fit] [--hide\|--show]` | Set column properties (requires `-o`) |
| `autofit` | `[--columns A:F]` | Auto-fit column widths from content (requires `-o`) |
| `fill` | `<source> <target> [--right]` | Fill cells with source value/formula (requires `-o`) |
| `clear` | `<range> [--all\|--styles\|--comments]` | Clear cell contents/styles/comments (requires `-o`) |
| `copy` | `<source> <target> [--values-only]` | Copy range with formula adjustment (requires `-o`) |
| `sort` | `<range> --by <col> [options]` | Sort rows by one or more columns (requires `-o`) |
| `merge` | `<range>` | Merge cells (requires `-o`) |
| `unmerge` | `<range>` | Unmerge cells (requires `-o`) |
| `comment` | `<ref> <text> [--author name]` | Add cell comment (requires `-o`) |
| `remove-comment` | `<ref>` | Remove cell comment (requires `-o`) |
| `freeze` | `<ref>` | Freeze panes at cell (requires `-o`) |
| `unfreeze` | | Remove freeze panes (requires `-o`) |
| `import` | `<csv-file> [start-ref] [options]` | Import CSV with type detection (requires `-o`) |
| `import-md` | `<md-file\|-> [--start ref] [options]` | Import GFM markdown table with type detection (requires `-o`) |
| `add-sheet` | `<name> [--after s] [--before s]` | Add new empty sheet (requires `-o`) |
| `remove-sheet` | `<name>` | Remove sheet (requires `-o`) |
| `rename-sheet` | `<name> <new-name>` | Rename sheet (requires `-o`) |
| `move-sheet` | `<name> [--to idx] [--after s] [--before s]` | Move sheet to new position (requires `-o`) |
| `copy-sheet` | `<name> <new-name>` | Copy sheet to new name (requires `-o`) |
| `insert-rows` | `<at-row> [count]` | Insert rows; shifts cells, rewrites formulas (requires `-o`) |
| `delete-rows` | `<at-row> [count]` | Delete rows; `#REF!` on lost references (requires `-o`) |
| `insert-cols` | `<at-col> [count]` | Insert columns; shifts cells, rewrites formulas (requires `-o`) |
| `delete-cols` | `<at-col> [count]` | Delete columns; `#REF!` on lost references (requires `-o`) |
| `batch` | `<file\|-> [--dry-run]` | Apply multiple operations from JSON (requires `-o`; `--dry-run` validates without a file) |
| `diff` | `-g <file2> [--format markdown\|json]` | Compare two workbooks; exit 0 identical, 1 differs, 2 error |

---

## Command Details

### `xl sheets [list|hide|show]`

Sheet operations. With no subcommand, defaults to `list`.

**Subcommands**:
| Subcommand | Arguments | Description |
|------------|-----------|-------------|
| `list` | `[--stats]` | List all sheets (`--stats` adds cell/formula counts; slower) |
| `hide` | `<sheet-name> [--very]` | Hide a sheet (`--very` = very hidden, VBA-only; requires `-o`) |
| `show` | `<sheet-name>` | Show a hidden sheet (requires `-o`) |

**Output** (`list --stats`):
```markdown
| # | Name        | Range    | Cells | Formulas |
|---|-------------|----------|-------|----------|
| 1 | Assumptions | A1:F50   | 234   | 12       |
| 2 | Revenue     | A1:M100  | 892   | 156      |
| 3 | P&L         | A1:N120  | 978   | 76       |
```

---

### `xl names` / `xl name add|rm`

List and manage defined names (named ranges).

```bash
xl -f model.xlsx names                                       # List all defined names
xl -f model.xlsx -o out.xlsx name add Tax 'Sheet1!$A$1'      # Add or replace
xl -f model.xlsx -o out.xlsx name rm Tax                     # Remove
```

---

### `xl bounds [--scan]`

Show the used range (bounding box of non-empty cells) for the sheet selected with `-s`. Instant by default (reads the worksheet's dimension element); `--scan` forces a full streaming scan for accurate bounds.

**Output**:
```
Sheet: Revenue
Used range: A1:M100
Rows: 1-100 (100 total)
Columns: A-M (13 total)
Non-empty: 892 cells
```

---

### `xl view <range>`

View a rectangular range — markdown table by default, or JSON/CSV/HTML/SVG/PNG/JPEG/WebP/PDF.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `range` | string | Yes | — | Cell range (e.g., "A1:D20") |
| `--format` | string | No | markdown | Output format: markdown, json, csv, html, svg, png, jpeg, webp, pdf |
| `--formulas` | flag | No | false | Show formulas instead of values |
| `--eval` | flag | No | false | Evaluate formulas (compute live values) |
| `--strict` | flag | No | false | Fail on formula evaluation errors (with `--eval`) |
| `--limit` | int | No | 50 | Max rows to display |
| `--skip-empty` | flag | No | false | Skip empty cells (JSON) or empty rows/columns (tabular) |
| `--show-labels` | flag | No | false | Include column letters and row numbers |
| `--header-row` | int | No | — | Use values from this row as keys in JSON output (1-based) |
| `--raster-output` | path | For raster | — | Output file (required for png/jpeg/webp/pdf) |
| `--dpi` | int | No | 144 | Resolution for raster output |
| `--quality` | int | No | 90 | JPEG quality 1-100 |
| `--rasterizer` | string | No | auto | Force backend: batik, cairosvg, rsvg-convert, resvg, imagemagick |
| `--gridlines` | flag | No | false | Show cell gridlines in SVG output |
| `--print-scale` | flag | No | false | Apply print scaling (for PDF-like output) |

**Output**:
```markdown
|   | A           | B       | C          | D       |
|---|-------------|---------|------------|---------|
| 1 | Revenue     |         | $1,000,000 |         |
| 2 | COGS        |         | $400,000   |         |
| 4 | Gross Profit|         | =C1-C2     |         |
```

---

### `xl cell <ref>`

Get complete information about a single cell.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `ref` | string | Yes | Cell reference (e.g., "A1", "B5") |

**Output** (formula cell):
```
Cell: C4
Type: Formula
Formula: =C1-C2
Cached Value: 600000
Formatted: $600,000
```

**Output** (value cell):
```
Cell: A1
Type: Text
Value: Revenue
```

---

### `xl search <pattern>`

Find cells containing text matching pattern. Searches all sheets by default (no `-s` needed).

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `pattern` | string | Yes | — | Search pattern (supports regex) |
| `--sheets` | string | No | all | Comma-separated list of sheets to search |
| `--limit` | int | No | 50 | Max results |

**Output**:
```markdown
Found 5 matches for "Revenue":

| Ref | Value           | Context (row)              |
|-----|-----------------|----------------------------|
| A1  | Revenue         | Revenue | | $1,000,000     |
| A10 | Revenue Growth  | Revenue Growth | | 5%      |
```

---

### `xl eval <formula> [--with overrides]`

Evaluate a formula without modifying the file (what-if analysis). `-f` is optional for constant formulas (`xl eval "=PI()*2"`).

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `formula` | string | Yes | Formula to evaluate |
| `--with`, `-w` | string | No | Temporary cell overrides (e.g., "B1=100,B2=200"; repeatable) |

**Examples**:
```bash
xl -f model.xlsx -s Sheet1 eval "=SUM(B1:B10)"
xl -f model.xlsx -s Sheet1 eval "=B1*1.1" --with "B1=100"
```

---

### `xl evala <formula> [--at <ref>]`

Evaluate an **array formula** and display the result grid, or spill it into the sheet. Requires `-f` (array formulas need sheet context).

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `formula` | string | Yes | Array formula to evaluate |
| `--at` | string | No | Target cell for array spill (default: display only) |
| `--with`, `-w` | string | No | Temporary cell overrides (repeatable) |

**Examples**:
```bash
xl -f data.xlsx -s Sheet1 evala "=TRANSPOSE(A1:C2)"          # Display result grid
xl -f data.xlsx -s Sheet1 evala "=SEQUENCE(5)" --at E1       # Spill starting at E1
xl -f data.xlsx -s Sheet1 evala "=A1:B2*10"                  # Array arithmetic with broadcasting
```

---

### `xl stats <range>`

Calculate statistics (count, sum, min, max, average, ...) for numeric values in a range. Supports `--stream` for large files.

```bash
xl -f data.xlsx -s Sheet1 stats B2:B10000
xl -f huge.xlsx --stream stats A1:E100000
```

---

### `xl put <ref|range> <value...>`

Write value(s) to a cell or range.

**Modes**:
| Mode | Example | Behavior |
|------|---------|----------|
| Single | `put A1 100` | Write 100 to A1 |
| Fill | `put A1:A10 "TBD"` | Fill range with the same value |
| Batch | `put A1:C1 "X" "Y" "Z"` | One value per cell (row-major) |
| CSV split | `put A1:C1 "X,Y,Z" --csv` | Split one comma-separated value across the range |

**Type Inference**:
- Numbers: `1000`, `1,234.56`, `$100`, `50%`
- Dates: `2024-01-15`, `01/15/2024`
- Booleans: `true`, `false`
- Text: Everything else

**Negative numbers**: use the `--value` flag (a leading `-` is parsed as a flag):
```bash
xl -f input.xlsx -s S1 -o output.xlsx put A1 --value "-500"
```

**Example**:
```bash
xl -f input.xlsx -s S1 -o output.xlsx put B5 1000000
```

---

### `xl putf <ref|range> <formula...>`

Write formula(s) to a cell or range with Excel-style dragging.

**Modes**:
| Mode | Example | Behavior |
|------|---------|----------|
| Single | `putf C1 "=A1+B1"` | One formula, one cell |
| Drag | `putf B2:B10 "=A2*1.1"` | Single formula + range: references shift per cell (`$` anchors pin) |
| Batch | `putf D1:D3 "=A1+B1" "=A2*B2" "=A3-B3"` | One formula per cell, applied as-is (no dragging) |

**Anchor modes** (`$` controls shifting when dragging): `$A$1` absolute, `$A1` column-absolute, `A$1` row-absolute, `A1` fully relative.

**Examples**:
```bash
xl -f input.xlsx -s S1 -o output.xlsx putf C5 "=B5*1.1"
xl -f input.xlsx -s S1 -o output.xlsx putf C2:C10 "=SUM(\$B\$2:B2)"   # Running total
```

(The batch JSON `putf` op additionally accepts a `"from"` field to drag from an explicit source cell.)

---

### `xl style <range> [options]`

Apply styling to cells.

Styles **merge** with existing formatting by default; `--replace` overwrites.

**Arguments**:
| Arg | Type | Description |
|-----|------|-------------|
| `range` | string | Cell/range reference |
| `--bold` | flag | Bold text |
| `--italic` | flag | Italic text |
| `--underline` | flag | Underlined text |
| `--font-size` | double | Font size in points |
| `--font-name` | string | Font family (e.g., "Arial") |
| `--bg` | string | Background color (name, #hex, or rgb(r,g,b)) |
| `--fg` | string | Text color |
| `--align` | string | Horizontal alignment: left, center, right |
| `--valign` | string | Vertical alignment: top, middle, bottom |
| `--wrap` | flag | Enable text wrapping |
| `--format` | string | Number format: general, number, currency, percent, date, text |
| `--border` | string | Border style for all sides: none, thin, medium, thick |
| `--border-top` / `--border-right` / `--border-bottom` / `--border-left` | string | Per-side border style |
| `--border-color` | string | Border color (applies to all specified borders) |
| `--replace` | flag | Replace entire style instead of merging |

**Example**:
```bash
xl -f input.xlsx -s S1 -o output.xlsx style A1:D1 --bold --bg yellow --align center
```

---

### `xl row <n> [options]`

Set row properties (height, hide/show).

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `n` | int | Yes | — | Row number (1-based) |
| `--height` | double | No | — | Row height in points |
| `--hide` | flag | No | false | Hide the row |
| `--show` | flag | No | false | Show (unhide) the row |

**Examples**:
```bash
xl -f input.xlsx -o output.xlsx row 5 --height 30
xl -f input.xlsx -o output.xlsx row 10 --hide
xl -f input.xlsx -o output.xlsx row 10 --show
```

---

### `xl col <letter> [options]`

Set column properties (width, hide/show, auto-fit).

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `letter` | string | Yes | — | Column letter or range (e.g., "A", "AA", "A:F") |
| `--width` | double | No | — | Column width in character units (~8.43 default) |
| `--auto-fit` | flag | No | false | Auto-fit width based on cell content |
| `--hide` | flag | No | false | Hide the column |
| `--show` | flag | No | false | Show (unhide) the column |

**Behavior**:
- `--auto-fit` calculates optimal width based on longest cell content in the column
- Adds 2 characters of padding to the calculated width
- Minimum width is 8.43 (Excel default)
- If both `--width` and `--auto-fit` are specified, `--auto-fit` takes precedence

**Examples**:
```bash
# Set explicit width
xl -f input.xlsx -o output.xlsx col B --width 20

# Auto-fit column width based on content
xl -f input.xlsx -o output.xlsx col A --auto-fit

# Hide/show columns
xl -f input.xlsx -o output.xlsx col C --hide
xl -f input.xlsx -o output.xlsx col C --show
```

---

### `xl clear <range>`

Clear cell contents, styles, or comments from a range.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `range` | string | Yes | — | Cell or range reference (e.g., "A1", "A1:D10") |
| `--all` | flag | No | false | Clear contents, styles, and comments |
| `--styles` | flag | No | false | Clear styles only (reset to default) |
| `--comments` | flag | No | false | Clear comments only |

**Behavior**:
- **Default** (no flags): Clears cell contents only
- **`--all`**: Clears everything (contents, styles, comments)
- **`--styles`**: Clears formatting but keeps contents and comments
- **`--comments`**: Clears comments but keeps contents and styles
- Flags can be combined: `--styles --comments` clears both
- Merged regions overlapping the cleared range are automatically unmerged

**Examples**:
```bash
# Clear contents (default)
xl -f input.xlsx -o output.xlsx clear A1:D10

# Clear everything
xl -f input.xlsx -o output.xlsx clear A1:D10 --all

# Clear styles only (keep data)
xl -f input.xlsx -o output.xlsx clear A1:D10 --styles

# Clear comments only
xl -f input.xlsx -o output.xlsx clear B5 --comments

# Clear styles and comments, keep contents
xl -f input.xlsx -o output.xlsx clear A1:D10 --styles --comments
```

---

### `xl fill <source> <target> [--right]`

Fill cells with source value/formula (Excel Ctrl+D/Ctrl+R equivalent).

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `source` | string | Yes | — | Source cell or range (e.g., "A1", "A1:C1") |
| `target` | string | Yes | — | Target range to fill (e.g., "A1:A10", "A1:C10") |
| `--right` | flag | No | false | Fill rightward instead of downward |

**Behavior**:
- **Fill Down** (default): Source row(s) are repeated down through target range
  - Columns must match between source and target
  - Example: `fill A1 A1:A10` copies A1 to A2:A10
  - Example: `fill A1:C1 A1:C10` copies row 1 to rows 2-10
- **Fill Right** (`--right`): Source column(s) are repeated right through target range
  - Rows must match between source and target
  - Example: `fill A1 A1:E1 --right` copies A1 to B1:E1
  - Example: `fill A1:A5 A1:E5 --right` copies column A to columns B-E
- Formulas are shifted using Excel anchor rules (`$` anchors are preserved)

**Examples**:
```bash
# Fill value down a column (Ctrl+D equivalent)
xl -f input.xlsx -o output.xlsx fill A1 A1:A100

# Fill multiple columns down together
xl -f input.xlsx -o output.xlsx fill A1:E1 A1:E100

# Fill value right across a row (Ctrl+R equivalent)
xl -f input.xlsx -o output.xlsx fill A1 A1:J1 --right

# Fill multiple rows right together
xl -f input.xlsx -o output.xlsx fill A1:A5 A1:J5 --right

# Formula shifting example: =A1*2 in B1 fills to =A2*2, =A3*2, etc.
xl -f input.xlsx -o output.xlsx fill B1 B1:B10
```

---

### `xl autofit [--columns A:F]`

Auto-fit column widths based on content (defaults to all used columns).

```bash
xl -f input.xlsx -s S1 -o output.xlsx autofit
xl -f input.xlsx -s S1 -o output.xlsx autofit --columns A:F
```

---

### `xl copy <source> <target> [--values-only]`

Copy a range to another location with Excel-style formula adjustment (`$` anchors preserved). `--values-only` copies values without adjusting formulas.

```bash
xl -f input.xlsx -s S1 -o output.xlsx copy A1:C10 E1
xl -f input.xlsx -s S1 -o output.xlsx copy A1:C10 E1 --values-only
```

---

### `xl sort <range> --by <col> [options]`

Sort rows in a range by one or more columns.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `range` | string | Yes | Range to sort |
| `--by`, `-b` | string | Yes | Primary sort column |
| `--then-by` | string | No | Secondary sort column(s) (repeatable) |
| `--desc` | flag | No | Sort descending (default: ascending) |
| `--numeric` | flag | No | Force numeric comparison ("10" > "9") |
| `--header` | flag | No | First row is header (exclude from sort) |

**Behavior**: empty cells sort last; formulas sort by cached value; rows move together.

```bash
xl -f f.xlsx -s S1 -o o.xlsx sort A1:D100 --by B --desc --numeric
xl -f f.xlsx -s S1 -o o.xlsx sort A1:D100 --by B --then-by C --header
```

---

### `xl merge <range>` / `xl unmerge <range>`

Merge or unmerge cells in a range.

```bash
xl -f input.xlsx -s S1 -o output.xlsx merge A1:D1
xl -f input.xlsx -s S1 -o output.xlsx unmerge A1:D1
```

---

### `xl comment <ref> <text> [--author name]` / `xl remove-comment <ref>`

Add or remove a cell comment.

```bash
xl -f input.xlsx -s S1 -o output.xlsx comment A1 "Verify this figure" --author "Analyst"
xl -f input.xlsx -s S1 -o output.xlsx remove-comment A1
```

---

### `xl freeze <ref>` / `xl unfreeze`

Freeze panes at a cell reference (rows above and columns to the left are locked), or remove them.

```bash
xl -f input.xlsx -s S1 -o output.xlsx freeze B2    # Freeze row 1 + column A
xl -f input.xlsx -s S1 -o output.xlsx unfreeze
```

---

### `xl import <csv-file> [start-ref] [options]`

Import CSV data with automatic type detection (numbers, booleans, ISO dates).

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `csv-file` | string | Yes | — | CSV file path |
| `start-ref` | string | No | A1 | Cell where import starts |
| `--delimiter` | char | No | `,` | Field separator |
| `--encoding` | string | No | UTF-8 | Input encoding |
| `--skip-header` | flag | No | false | Skip first row (do not import) |
| `--new-sheet` | string | No | — | Create new sheet for imported data |
| `--no-type-inference` | flag | No | false | Treat all values as text |

```bash
xl -f f.xlsx -s S1 -o o.xlsx import data.csv A1 --delimiter ";" --skip-header
xl -f f.xlsx -o o.xlsx import data.csv --new-sheet "Imported"
```

**Limitations**: entire CSV is loaded into memory (recommended <50k rows); dates must be ISO 8601 (`YYYY-MM-DD`).

---

### `xl import-md <md-file|-> [--start ref] [options]`

Import a GFM (GitHub Flavored Markdown) pipe table with smart type detection. Use `-` to read from stdin — handy for LLM agents that generate tables inline.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `md-file` | string | Yes | — | Markdown file path, or `-` for stdin |
| `--start` | string | No | A1 | Top-left cell for the imported table |
| `--skip-header` | flag | No | false | Skip the table's header row (do not import) |
| `--new-sheet` | string | No | — | Create new sheet for imported data |
| `--no-type-inference` | flag | No | false | Treat all values as text |

```bash
xl -f f.xlsx -s S1 -o o.xlsx import-md table.md --start B2
xl -f f.xlsx -o o.xlsx import-md table.md --new-sheet "Data"
printf '| A | B |\n|---|---|\n| 1 | 2 |\n' | xl -f f.xlsx -s S1 -o o.xlsx import-md -
```

**Table format** (GFM): header row, delimiter row (`|---|---|`), body rows. The first table found in the input is imported (preamble prose is skipped); the table ends at the first blank line. Outer pipes are optional, `\|` inside a cell is a literal pipe, and cell whitespace is trimmed. Body rows are padded/truncated to the delimiter row's column count.

**Type detection** (per cell, same smart detection as batch `put`): currency `$1,234.56` → Number + Currency format, percent `45.5%` → `0.455` + Percent format, ISO dates `2025-01-15` → date-formatted cell, plain numbers and `true`/`false` → typed values, everything else → text. Opt out with `--no-type-inference`.

**Alignment**: GFM markers map to cell horizontal alignment — `:---` left, `:---:` center, `---:` right (no marker leaves alignment unset).

**Limitations**: input is read as UTF-8 and parsed in memory; one table per import (first wins).

---

### Sheet management: `add-sheet`, `remove-sheet`, `rename-sheet`, `move-sheet`, `copy-sheet`

```bash
xl -f f.xlsx -o o.xlsx add-sheet Summary --after Sheet1      # or --before <name>
xl -f f.xlsx -o o.xlsx remove-sheet Scratch
xl -f f.xlsx -o o.xlsx rename-sheet "Old Name" "New Name"
xl -f f.xlsx -o o.xlsx move-sheet Summary --to 0             # or --after/--before <name>
xl -f f.xlsx -o o.xlsx copy-sheet Template "Q2 Report"
```

---

### Structural editing: `insert-rows`, `delete-rows`, `insert-cols`, `delete-cols`

Insert or delete rows/columns with full formula-reference rewriting: references at or past the cut shift, straddling ranges shrink, and references to deleted cells become `#REF!`. Cross-sheet references to the edited sheet are rewritten too.

**Arguments**: position (`<at-row>` 1-based, or `<at-col>` letter) and optional `<count>` (default 1). Column commands also accept an inclusive range (`C:E`), which overrides the count.

```bash
xl -f f.xlsx -s S1 -o o.xlsx insert-rows 5 2     # Insert 2 rows at row 5
xl -f f.xlsx -s S1 -o o.xlsx delete-rows 5       # Delete row 5
xl -f f.xlsx -s S1 -o o.xlsx insert-cols B 3     # Insert 3 columns at B
xl -f f.xlsx -s S1 -o o.xlsx delete-cols C:E     # Delete columns C through E
```

Example formula rewriting: deleting row 2 turns `=A1+A3` into `=A1+A2`; `=SUM(A1:A4)` shrinks to `=SUM(A1:A3)`; a direct reference to a deleted cell becomes `#REF!`.

---

### `xl batch <file|->`

Apply multiple operations atomically from JSON input.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `file` | string | No (default `-`) | JSON file path or `-` for stdin |
| `--dry-run` | flag | No | Validate JSON and show a summary without writing (works without `-f`/`-o`) |

**JSON Schema**:

```json
[
  {"op": "put", "ref": "A1", "value": "Hello"},
  {"op": "putf", "ref": "B1", "value": "=A1*2"},
  {"op": "style", "range": "A1:B1", "bold": true},
  {"op": "merge", "range": "A1:D1"},
  {"op": "colwidth", "col": "A", "width": 15.5},
  {"op": "rowheight", "row": 1, "height": 30},
  {"op": "comment", "ref": "A1", "text": "Revenue figure", "author": "Analyst"},
  {"op": "autofit", "columns": "A:D"},
  {"op": "add-sheet", "name": "Summary", "after": "Sheet1"}
]
```

**Supported Operations**:

| Operation | Required Fields | Optional Fields | Description |
|-----------|-----------------|-----------------|-------------|
| `put` | `ref`, `value` | `format`, `values`, `detect` | Write value to cell |
| `putf` | `ref`, `value` | `from`, `values` | Write formula(s) to cell(s) |
| `style` | `range` | styling options | Apply cell styling |
| `merge` | `range` | | Merge cells |
| `unmerge` | `range` | | Unmerge cells |
| `colwidth` | `col`, `width` | | Set column width |
| `rowheight` | `row`, `height` | | Set row height |
| `comment` | `ref`, `text` | `author` | Add cell comment |
| `remove-comment` | `ref` | | Remove cell comment |
| `hyperlink` | `ref` | `target` | Set cell hyperlink (omit `target` to clear) |
| `clear` | `range` | `all`, `styles`, `comments` | Clear cell contents/styles/comments |
| `col-hide` | `col` | | Hide column |
| `col-show` | `col` | | Show column |
| `row-hide` | `row` | | Hide row |
| `row-show` | `row` | | Show row |
| `autofit` | | `columns` | Auto-fit column widths |
| `add-sheet` | `name` | `after` | Add new sheet |
| `rename-sheet` | `from`, `to` | | Rename sheet |
| `freeze` | `ref` | | Freeze panes at cell |
| `unfreeze` | | | Remove freeze panes |
| `copy` | `source`, `target` | `valuesOnly` | Copy range with formula adjustment |

**Native JSON Types** (recommended):

```json
// Numbers are stored as numeric values (not text)
{"op": "put", "ref": "A1", "value": 99.0}

// Booleans
{"op": "put", "ref": "A2", "value": true}

// With explicit format
{"op": "put", "ref": "A3", "value": 99.0, "format": "currency"}
{"op": "put", "ref": "A4", "value": 0.594, "format": "percent"}
```

**Format Options**:

| Format Name | Description | Example Output |
|-------------|-------------|----------------|
| `general` | Default format | `1234.5` |
| `integer` | Whole numbers | `1235` |
| `decimal` | Two decimal places | `1234.50` |
| `currency` | Currency with symbol | `$1,234.50` |
| `percent` | Percentage | `59%` |
| `percent_decimal` | Percentage with decimals | `59.4%` |
| `date` | Date format | `11/10/25` |
| `datetime` | Date and time | `11/10/25 14:30` |
| `time` | Time only | `14:30:00` |
| `text` | Text format | `1234.5` |
| *custom* | Any Excel format code | See below |

**Custom Format Codes**:

```json
// MOIC/Multiple format (3.5x)
{"op": "put", "ref": "A1", "value": 3.5, "format": "0.0x"}

// Accounting format with negatives in parentheses
{"op": "put", "ref": "A2", "value": -1234, "format": "$#,##0;($#,##0)"}

// Basis points
{"op": "put", "ref": "A3", "value": 50, "format": "0 \"bps\""}

// Custom date format
{"op": "put", "ref": "A4", "value": "2025-11-10", "format": "yyyy-mm-dd"}
```

**Smart String Detection** (enabled by default):

Strings are automatically detected and formatted:

```json
// Currency detected from $ prefix
{"op": "put", "ref": "A1", "value": "$99.00"}  // → Number(99.0), Currency

// Percent detected from % suffix
{"op": "put", "ref": "A2", "value": "59.4%"}   // → Number(0.594), Percent

// ISO date detected
{"op": "put", "ref": "A3", "value": "2025-11-10"}  // → DateTime, Date format

// Plain text (no detection pattern)
{"op": "put", "ref": "A4", "value": "Hello"}       // → Text
```

**Disable Detection**: Set `"detect": false` to treat strings as plain text:

```json
{"op": "put", "ref": "A1", "value": "$99.00", "detect": false}  // → Text "$99.00"
{"op": "put", "ref": "A2", "value": "59.4%", "detect": false}   // → Text "59.4%"
```

**Formula Dragging** (putf with range):

```json
// Single formula dragged across range (uses Excel $ anchoring)
{"op": "putf", "ref": "B2:B10", "value": "=SUM($A$1:A2)", "from": "B2"}

// Explicit formulas for each cell (no dragging)
{"op": "putf", "ref": "B2:B4", "values": ["=A2*2", "=A3*2", "=A4*2"]}
```

**Style Options**:

| Option | Type | Description |
|--------|------|-------------|
| `bold` | boolean | Bold text |
| `italic` | boolean | Italic text |
| `underline` | boolean | Underlined text |
| `bg` | string | Background color (hex: `#FF0000`) |
| `fg` | string | Font color (hex: `#0000FF`) |
| `fontSize` | number | Font size in points |
| `fontName` | string | Font family name |
| `align` | string | Horizontal alignment: `left`, `center`, `right`, `justify` |
| `valign` | string | Vertical alignment: `top`, `middle`, `bottom` |
| `wrap` | boolean | Enable text wrapping |
| `numFormat` | string | Number format (see Format Options above) |
| `border` | string | All borders: `none`, `thin`, `medium`, `thick` |
| `borderTop` | string | Top border style |
| `borderRight` | string | Right border style |
| `borderBottom` | string | Bottom border style |
| `borderLeft` | string | Left border style |
| `borderColor` | string | Border color (hex) |
| `replace` | boolean | Replace style instead of merge (default: false) |

**Note**: Use `align` for horizontal alignment, not `halign`. Unknown properties are ignored with a warning.

**Examples**:

```bash
# From file
xl -f input.xlsx -o output.xlsx batch operations.json

# From stdin (pipe)
echo '[{"op": "put", "ref": "A1", "value": 100, "format": "currency"}]' | \
  xl -f input.xlsx -s Sheet1 -o output.xlsx batch -

# Complex workflow
cat <<'EOF' | xl -f input.xlsx -s Sheet1 -o output.xlsx batch -
[
  {"op": "put", "ref": "A1", "value": "Revenue", "format": "text"},
  {"op": "put", "ref": "B1", "value": 1000000, "format": "currency"},
  {"op": "style", "range": "A1:B1", "bold": true, "bg": "#FFFF00"},
  {"op": "merge", "range": "A1:B1"},
  {"op": "colwidth", "col": "A", "width": 20}
]
EOF
```

#### Common Gotchas

| Scenario | Behavior | Workaround |
|----------|----------|------------|
| Leading zeros: `"00123"` | Smart detection converts to number `123` | Use `"detect": false` to preserve as text |
| Mixed patterns: `"50 (50%)"` | First pattern wins (treated as text) | Use explicit `"format"` field |
| `values` array length mismatch | Error raised if array length ≠ range cell count | Ensure exact match |
| Percent as decimal | `"59.4%"` stored as `0.594` | Excel displays correctly with percent format |
| Invalid custom formats | Accepted but may render incorrectly in Excel | Test format codes in Excel first |
| `--stream` mode | Supports formula dragging but not formula evaluation | Use non-streaming for `--eval` |

---

### `xl diff -g <file2> [--format markdown|json]`

Compare two workbooks and report differences. The first file comes from the global `-f`, the second from `-g/--file2`. Optional global `-s/--sheet` restricts the comparison to one sheet.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `-g, --file2` | path | Yes | — | Second file to compare against |
| `--format` | string | No | markdown | `markdown` (human) or `json` (stable schema) |

**Exit codes** (diff-tool convention): `0` identical, `1` differences found, `2` error.

**What is compared** (per sheet, refs in A1, row-major order):
- **Changed cells** — value, formula text, and resolved style (`styleChanged` boolean). Formula cells compare by formula text; cached values are derived and ignored. Styles compare resolved formatting (style id lookup), so equal formatting under different ids is not a difference.
- **Added / removed cells** — a cell with Empty value, default style, and no hyperlink counts as absent.
- **Sheets added / removed** (by name).
- **Merged ranges, comments, hyperlinks** — separate added/removed/changed deltas per sheet.

```bash
xl -f old.xlsx diff -g new.xlsx                      # Markdown report
xl -f old.xlsx -s Sheet1 diff -g new.xlsx            # One sheet only
xl -f old.xlsx diff -g new.xlsx --format json        # Machine-readable
xl -f old.xlsx diff -g new.xlsx && echo "unchanged"  # Exit-code driven
```

**JSON schema** (stable; `sheets` lists only sheets with differences):

```json
{
  "identical": false,
  "sheetsAdded": [], "sheetsRemoved": [],
  "sheets": [{
    "name": "Sheet1",
    "added":   [{"ref": "D5", "value": "New", "formula": null}],
    "removed": [{"ref": "E5", "value": "Old", "formula": null}],
    "changed": [{"ref": "A5",
                 "before": {"value": "Revenue", "formula": null},
                 "after":  {"value": "Total Revenue", "formula": null},
                 "styleChanged": false}],
    "mergesAdded": [], "mergesRemoved": [],
    "commentsAdded": [], "commentsRemoved": [], "commentsChanged": [],
    "hyperlinksAdded": [], "hyperlinksRemoved": [], "hyperlinksChanged": []
  }]
}
```

**Limitations**: both workbooks load in memory (`--max-size` applies to each); no range-level filter yet.

---

## Output Format

### Cell Reference Convention

Always include explicit references in output:

```markdown
# Good - References visible
|   | A        | B       |
|---|----------|---------|
| 1 | Revenue  | $1M     |
| 2 | COGS     | $400K   |
```

### Error Format

```
Error: <ErrorType>
Location: <Context>
Details: <Human-readable explanation>
Suggestion: <How to fix>
```

Example:
```
Error: CircularReference
Location: B10
Details: Formula =A10+B10 creates cycle: B10 → A10 → B10
Suggestion: Use a different cell reference to break the cycle
```

---

## See Also

- [Quick Start Guide](../QUICK-START.md) — Library usage
- [Scripting Guide](scripting.md) — When a task outgrows the CLI (loops, typed extraction, multi-file pipelines)
- [Performance Guide](performance-guide.md) — Streaming for large files
- [GitHub Issues](https://github.com/TJC-LP/xl/issues) — Feature requests and bug reports
