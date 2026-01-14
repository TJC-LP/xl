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

This builds a fat JAR and installs `xl` to `~/.local/bin/`. Ensure it's in your PATH:

```bash
export PATH="$HOME/.local/bin:$PATH"
```

**Uninstall**: `make uninstall`

**Update**: After `git pull`, run `make build` — the wrapper auto-picks up the new JAR.

---

## Quick Reference

```bash
# Global flags (used with all commands)
-f, --file <path>     # Input file (required)
-s, --sheet <name>    # Sheet to operate on (optional, defaults to first)
-o, --output <path>   # Output file for mutations (required for put/putf)

# Read-only operations
xl -f model.xlsx sheets                    # List all sheets
xl -f model.xlsx bounds                    # Show used range
xl -f model.xlsx -s "P&L" view A1:D20      # View range as markdown
xl -f model.xlsx cell B5                   # Get single cell details
xl -f model.xlsx search "Revenue"          # Find cells by content
xl -f model.xlsx eval "=SUM(B1:B10)"       # Evaluate formula (what-if)

# Mutations (require -o)
xl -f model.xlsx -o output.xlsx put B5 1000000       # Write value
xl -f model.xlsx -o output.xlsx putf C5 "=B5*1.1"    # Write formula

# What-if analysis with overrides
xl -f model.xlsx eval "=B1*1.1" --with "B1=100"      # Evaluate with temporary values
```

---

## Commands

### Operation Categories

| Category | Commands | Purpose |
|----------|----------|---------|
| **Navigate** | `sheets`, `bounds` | Find your way around |
| **Explore** | `view`, `cell`, `search` | Read data incrementally |
| **Analyze** | `eval` | What-if formula evaluation |
| **Mutate** | `put`, `putf`, `style`, `row`, `col`, `fill`, `clear` | Make changes (requires `-o`) |

### Command Summary

| Command | Arguments | Description |
|---------|-----------|-------------|
| `sheets` | | List all sheets with stats |
| `bounds` | | Show used range of current sheet |
| `view` | `<range>` | Render range as markdown |
| `cell` | `<ref>` | Get single cell details |
| `search` | `<pattern>` | Find cells matching pattern |
| `eval` | `<formula> [--with overrides]` | Evaluate formula without modifying |
| `put` | `<ref> <value>` | Write value to cell (requires `-o`) |
| `putf` | `<ref> <formula>` | Write formula to cell (requires `-o`) |
| `style` | `<range> [options]` | Apply styling (requires `-o`) |
| `row` | `<n> [options]` | Set row height, hide/show (requires `-o`) |
| `col` | `<letter> [options]` | Set column width, hide/show, auto-fit (requires `-o`) |
| `fill` | `<source> <target> [--right]` | Fill cells with source value/formula (requires `-o`) |
| `clear` | `<range> [--all\|--styles\|--comments]` | Clear cell contents/styles/comments (requires `-o`) |

---

## Command Details

### `xl sheets`

List all sheets as a markdown table.

**Output**:
```markdown
| # | Name        | Range    | Cells | Formulas |
|---|-------------|----------|-------|----------|
| 1 | Assumptions | A1:F50   | 234   | 12       |
| 2 | Revenue     | A1:M100  | 892   | 156      |
| 3 | P&L         | A1:N120  | 978   | 76       |
```

---

### `xl bounds [sheet]`

Show the used range (bounding box of non-empty cells).

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `sheet` | string | No | active | Sheet name |

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

View a rectangular range as a markdown table with row/column headers.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `range` | string | Yes | — | Cell range (e.g., "A1:D20") |
| `--formulas` | flag | No | false | Show formulas instead of values |
| `--eval` | flag | No | false | Evaluate formulas (compute live values) |
| `--limit` | int | No | 50 | Max rows to display |

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

Find cells containing text matching pattern.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `pattern` | string | Yes | — | Search pattern (supports regex) |
| `--in` | string | No | entire sheet | Limit search to range |
| `--limit` | int | No | 20 | Max results |

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

Evaluate a formula without modifying the file (what-if analysis).

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `formula` | string | Yes | Formula to evaluate |
| `--with` | string | No | Temporary cell overrides (e.g., "B1=100") |

**Examples**:
```bash
xl -f model.xlsx eval "=SUM(B1:B10)"
xl -f model.xlsx eval "=B1*1.1" --with "B1=100"
```

---

### `xl put <ref> <value>`

Write a value to a cell.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `ref` | string | Yes | Cell reference |
| `value` | string | Yes | Value to write |

**Type Inference**:
- Numbers: `1000`, `1,234.56`, `$100`, `50%`
- Dates: `2024-01-15`, `01/15/2024`
- Booleans: `true`, `false`
- Text: Everything else

**Example**:
```bash
xl -f input.xlsx -o output.xlsx put B5 1000000
```

---

### `xl putf <ref> <formula>`

Write a formula to a cell.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `ref` | string | Yes | Cell reference or range |
| `formula` | string | Yes | Formula (with or without `=`) |
| `--from` | string | No | Source cell for relative reference adjustment |

**Examples**:
```bash
xl -f input.xlsx -o output.xlsx putf C5 "=B5*1.1"
xl -f input.xlsx -o output.xlsx putf B2:B10 "=SUM(\$A\$1:A2)" --from B2
```

---

### `xl style <range> [options]`

Apply styling to cells.

**Arguments**:
| Arg | Type | Description |
|-----|------|-------------|
| `range` | string | Cell/range reference |
| `--bold` | flag | Bold text |
| `--italic` | flag | Italic text |
| `--bg` | string | Background color (name or hex) |
| `--fg` | string | Text color |
| `--align` | string | Alignment: left, center, right |
| `--format` | string | Number format: currency, percent, date |
| `--border` | string | Border style: none, thin, medium, thick |
| `--replace` | flag | Replace entire style instead of merging |

**Example**:
```bash
xl -f input.xlsx -o output.xlsx style A1:D1 --bold --bg yellow --align center
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
| `letter` | string | Yes | — | Column letter (e.g., "A", "B", "AA") |
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
- [Performance Guide](performance-guide.md) — Streaming for large files
- [GitHub Issues](https://github.com/TJC-LP/xl/issues) — Feature requests and bug reports
