# XL CLI — LLM-Friendly Excel Operations

**Status**: Implemented (Phase 1 - Stateless)
**Target**: `xl-cli` module
**Priority**: High — Core ergonomics for AI/LLM use cases

---

## Executive Summary

The `xl` CLI provides a command-line interface for Excel operations, designed specifically for LLM agents. It follows Claude Code's incremental exploration pattern: don't dump everything at once, explore on demand.

**Design Philosophy**:
- **Stateless by default** — Each command is self-contained, no session state
- **Explicit cell refs** — Always use `A1`, `B5:D10` over "smart" inference
- **Global flags** — `-f` for file, `-s` for sheet, `-o` for output
- **LLM-optimized output** — Markdown tables, token-efficient, always include refs

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

## 1. Tool Taxonomy

### 1.1 Operation Categories (Stateless Mode)

| Category | Commands | Purpose |
|----------|----------|---------|
| **Navigate** | `sheets`, `bounds` | Find your way around |
| **Explore** | `view`, `cell`, `search` | Read data incrementally |
| **Analyze** | `eval` | What-if formula evaluation |
| **Mutate** | `put`, `putf` | Make changes (requires `-o`) |

### 1.2 Command Summary (Currently Implemented)

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

### 1.3 Future Commands (REPL Mode - Phase 2)

| Command | Arguments | Description |
|---------|-----------|-------------|
| `open` | `<path>` | Load Excel file into session |
| `create` | `[--sheets names]` | Create new empty workbook |
| `close` | `[--discard]` | Close current workbook |
| `select` | `<name>` | Set active sheet |
| `save` | | Save to original path |
| `saveas` | `<path>` | Save to new path |
| `describe` | `[sheet]` | Detailed sheet description |
| `labels` | `[range]` | Detect label-value pairs |
| `regions` | `[range]` | Detect data regions |
| `flow` | `[ref]` | Show formula dependencies |
| `batch` | `<ref=value>...` | Write multiple cells |
| `style` | `<range> [options]` | Apply styling |
| `delete` | `<range>` | Clear cell contents |
| `export` | `<range> --format <fmt>` | Export range |

---

## 2. Command Specifications

### 2.1 Open/Initialize Commands

#### `xl open <path>`

Load an Excel file into the session.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `path` | string | Yes | Path to .xlsx file |
| `--readonly` | flag | No | Open in read-only mode |

**Output**:
```
Opened: /path/to/model.xlsx
Sheets: Assumptions, Revenue, P&L, Returns (4 total)
Active: Assumptions
```

**Errors**:
- `FileNotFound` — Path does not exist
- `InvalidFormat` — Not a valid .xlsx file
- `AlreadyOpen` — A workbook is already open (use `close` first)

---

#### `xl create [--sheets <names>]`

Create a new empty workbook in memory.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `--sheets` | string | No | "Sheet1" | Comma-separated sheet names |

**Output**:
```
Created new workbook
Sheets: Sheet1 (1 total)
Active: Sheet1
```

---

#### `xl close [--discard]`

Close the current workbook.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `--discard` | flag | No | Discard unsaved changes without warning |

**Output**:
```
Closed: model.xlsx
```

**Errors**:
- `UnsavedChanges` — Has unsaved changes (use `--discard` to force)

---

### 2.2 Navigate Commands

#### `xl info`

Show high-level workbook summary.

**Output**:
```
Workbook: financial_model.xlsx
Path: /Users/demo/models/financial_model.xlsx
Sheets: 5
Total Cells: 2,847 non-empty
Formulas: 342

Sheet Summary:
  1. Assumptions    | A1:F50   | 234 cells | 12 formulas
  2. Revenue        | A1:M100  | 892 cells | 156 formulas
  3. Costs          | A1:M80   | 543 cells | 98 formulas
  4. P&L            | A1:N120  | 978 cells | 76 formulas
  5. Dashboard      | A1:J30   | 200 cells | 0 formulas
```

---

#### `xl sheets`

List all sheets as a markdown table.

**Output**:
```markdown
| # | Name        | Range    | Cells | Formulas |
|---|-------------|----------|-------|----------|
| 1 | Assumptions | A1:F50   | 234   | 12       |
| 2 | Revenue     | A1:M100  | 892   | 156      |
| 3 | Costs       | A1:M80   | 543   | 98       |
| 4 | P&L         | A1:N120  | 978   | 76       |
| 5 | Dashboard   | A1:J30   | 200   | 0        |
```

---

#### `xl select <name>`

Set the active sheet for subsequent commands.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `name` | string | Yes | Sheet name or 1-based index |

**Output**:
```
Selected: Revenue
Used range: A1:M100
Cells: 892 non-empty, 156 formulas
```

**Errors**:
- `SheetNotFound` — No sheet with that name/index

---

#### `xl bounds [sheet]`

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

### 2.3 Explore Commands

#### `xl view <range>`

View a rectangular range as a markdown table with row/column headers.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `range` | string | Yes | — | Cell range (e.g., "A1:D20") |
| `--formulas` | flag | No | false | Show formulas instead of values |
| `--limit` | int | No | 50 | Max rows to display |

**Output**:
```markdown
|   | A           | B       | C          | D       |
|---|-------------|---------|------------|---------|
| 1 | Revenue     |         | $1,000,000 |         |
| 2 | COGS        |         | $400,000   |         |
| 3 |             |         |            |         |
| 4 | Gross Profit|         | =C1-C2     |         |
| 5 | GP Margin   |         | =C4/C1     |         |
```

**Design Notes**:
- Always include row numbers (1, 2, 3...) and column letters (A, B, C...)
- Empty cells shown as blank (not "null" or "-")
- Numbers formatted per cell style (currency, percent, etc.)
- Formulas shown with `=` prefix when `--formulas` flag is set

---

#### `xl cell <ref>`

Get complete information about a single cell.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `ref` | string | Yes | Cell reference (e.g., "A1", "B5") |
| `--eval` | flag | No | Evaluate formula if present |

**Output** (for a formula cell):
```
Cell: C4
Type: Formula
Formula: =C1-C2
Cached Value: 600000
Formatted: $600,000
Style: Currency, Bold
Dependencies: C1, C2
Dependents: C5, D4
```

**Output** (for a value cell):
```
Cell: A1
Type: Text
Value: Revenue
Formatted: Revenue
Style: Bold
Dependencies: (none)
Dependents: (none)
```

---

#### `xl search <pattern>`

Find cells containing text matching pattern.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `pattern` | string | Yes | — | Search pattern (supports regex) |
| `--in` | string | No | entire sheet | Limit search to range |
| `--case` | flag | No | false | Case-sensitive search |
| `--limit` | int | No | 20 | Max results |

**Output**:
```markdown
Found 5 matches for "Revenue":

| Ref | Value           | Context (row)              |
|-----|-----------------|----------------------------|
| A1  | Revenue         | Revenue \| \| $1,000,000   |
| A10 | Revenue Growth  | Revenue Growth \| \| 5%    |
| B15 | Total Revenue   | \| Total Revenue \| $5M    |
| E3  | =Revenue!C1     | \| \| \| \| =Revenue!C1    |
| A20 | Net Revenue     | Net Revenue \| \| $950,000 |
```

---

#### `xl describe [sheet]`

Get detailed spatial description of a sheet (designed for LLM context injection).

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `sheet` | string | No | active | Sheet name |

**Output**:
```
Sheet: Model
Used Range: A1:Z200
Cells: 847 non-empty | 142 formulas

Regions (4 detected):
  1. A1:D15    | 42 cells  | 70% dense  | Labels: Entry Date, Entry EV
  2. A20:F35   | 84 cells  | 88% dense  | Labels: Revenue Y1, Revenue Y2
  3. A40:C45   | 12 cells  | 67% dense  | Labels: Exit Multiple, IRR
  4. H1:J10    | 24 cells  | 80% dense  | Labels: Tax Rate, Discount Rate

Top Labels: Entry EV, Exit Multiple, Revenue Y1, IRR, Discount Rate
Formula Inputs: 23 cells (assumptions)
Formula Outputs: 8 cells (results, no dependents)
```

---

### 2.4 Analyze Commands

#### `xl labels [range]`

Detect label-value pairs (text cell adjacent to numeric/formula cell).

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `range` | string | No | used range | Range to scan |

**Output**:
```markdown
Found 12 label-value pairs:

| Label          | Ref | Value       | Value Ref | Position |
|----------------|-----|-------------|-----------|----------|
| Entry EV       | A5  | $500,000,000| B5        | LeftOf   |
| Entry Multiple | A6  | 8.0x        | B6        | LeftOf   |
| Exit Multiple  | A7  | 10.0x       | B7        | LeftOf   |
| IRR            | A10 | 22.5%       | B10       | LeftOf   |
| Revenue Y1     | B3  | $100,000,000| B4        | Above    |
```

---

#### `xl regions [range]`

Detect contiguous data regions separated by empty rows/columns.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `range` | string | No | used range | Range to scan |

**Output**:
```markdown
Found 3 regions:

| # | Bounds   | Cells | Density | Has Formulas |
|---|----------|-------|---------|--------------|
| 1 | A1:D15   | 42    | 70%     | Yes          |
| 2 | A20:F35  | 84    | 88%     | Yes          |
| 3 | A40:C45  | 12    | 67%     | No           |
```

---

#### `xl flow [ref]`

Show formula dependency graph.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `ref` | string | No | entire sheet | Focus on specific cell |
| `--depth` | int | No | 3 | Max traversal depth |

**Output** (for specific cell):
```
Formula Flow for B10 (IRR):

Inputs (assumptions):
  B5: Entry EV = $500,000,000
  B6: Entry Multiple = 8.0x
  B7: Exit Multiple = 10.0x
  B8: Hold Period = 5

Calculation Chain:
  B5 → C10 (Entry Price)
  B6, B7 → C11 (Multiple Expansion)
  C10, C11, B8 → B10 (IRR)

Formula: =IRR(C20:C25)
Result: 22.5%

Dependents (cells that use B10):
  D15: Returns Summary
  E20: Dashboard IRR
```

**Output** (for entire sheet):
```
Formula Flow Summary:

Inputs (23 assumption cells):
  B5: Entry EV | B6: Entry Multiple | B7: Exit Multiple
  B8: Hold Period | B9: Revenue Growth | B10: Margin Expansion
  ... (17 more)

Outputs (8 result cells - no dependents):
  B15: IRR | B16: MOIC | B17: NPV
  C30: Exit Value | D30: Total Return
  ... (3 more)

Circular References: None detected
```

---

### 2.5 Mutate Commands

#### `xl put <ref> <value>`

Write a value to a cell.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `ref` | string | Yes | Cell reference |
| `value` | string | Yes | Value to write |
| `--type` | string | No | Force type: text, number, date, bool |

**Output**:
```
Put: A1 = "Revenue"
Type: Text
```

**Type Inference**:
- Numbers: `1000`, `1,234.56`, `$100`, `50%`
- Dates: `2024-01-15`, `01/15/2024`
- Booleans: `true`, `false`, `TRUE`, `FALSE`
- Text: Everything else (or use `--type text` to force)

---

#### `xl putf <ref> <formula>`

Write a formula to a cell.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `ref` | string | Yes | Cell reference |
| `formula` | string | Yes | Formula (with or without `=`) |

**Output**:
```
Put formula: B10 = =IRR(C20:C25)
Dependencies: C20, C21, C22, C23, C24, C25
```

**Errors**:
- `InvalidFormula` — Cannot parse formula
- `CircularReference` — Would create cycle

---

#### `xl batch <assignments...>`

Write multiple cells in one operation.

**Arguments**:
```
xl batch A1="Revenue" B1=1000 C1="=B1*1.05"
```

**Output**:
```
Batch update: 3 cells
  A1 = "Revenue" (text)
  B1 = 1000 (number)
  C1 = =B1*1.05 (formula)
```

---

#### `xl style <range> [options]`

Apply styling to cells.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `range` | string | Yes | Cell/range reference |
| `--bold` | flag | No | Bold text |
| `--italic` | flag | No | Italic text |
| `--bg` | string | No | Background color (name or hex) |
| `--fg` | string | No | Text color |
| `--align` | string | No | Alignment: left, center, right |
| `--format` | string | No | Number format: currency, percent, date |

**Output**:
```
Styled: A1:D1
Applied: bold, bg=yellow, align=center
```

---

#### `xl delete <range>`

Clear cell contents (keeps formatting).

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `range` | string | Yes | Cell/range reference |
| `--all` | flag | No | Also clear formatting |

**Output**:
```
Deleted: A5:D10 (24 cells)
```

---

### 2.6 Persist Commands

#### `xl save`

Save workbook to original path.

**Output**:
```
Saved: /path/to/model.xlsx
Sheets: 5
Cells: 2,847
```

**Errors**:
- `NoWorkbook` — No workbook is open
- `ReadOnly` — Opened in read-only mode
- `PermissionDenied` — Cannot write to path

---

#### `xl saveas <path>`

Save workbook to a new path.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `path` | string | Yes | New file path |
| `--overwrite` | flag | No | Overwrite if exists |

**Output**:
```
Saved: /path/to/new_model.xlsx
Sheets: 5
Cells: 2,847
```

---

#### `xl export <range> --format <format>`

Export a range to different formats.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `range` | string | Yes | — | Cell range |
| `--format` | string | No | markdown | Output format |

**Formats**:
- `markdown` — Markdown table (default)
- `html` — HTML table with inline styles
- `csv` — Comma-separated values
- `json` — JSON array of rows
- `text` — Plain text grid

---

## 3. Session State Model

### 3.1 State Structure

```scala
final case class Session(
  workbook: Option[Workbook],
  path: Option[Path],
  activeSheet: Option[SheetName],
  isDirty: Boolean,
  isReadOnly: Boolean
)
```

### 3.2 State Rules

1. **Single Workbook**: Only one workbook can be open at a time
   - `xl open` fails if already open (use `xl close` first)

2. **Active Sheet Context**: Most commands default to active sheet
   - Changed via `xl select`
   - Override per-command: `xl view Revenue!A1:D20`

3. **Dirty Tracking**: Mutations set `isDirty = true`
   - `xl save` clears dirty flag
   - `xl close` warns if dirty (use `--discard` to force)

4. **Sheet References**: Support both name and cross-sheet syntax
   - Same sheet: `A1:D20`
   - Other sheet: `Revenue!A1:D20` or `"Sheet Name"!A1:D20`

---

## 4. Output Format Guidelines

### 4.1 Token Efficiency

| Output Type | Format | Example |
|-------------|--------|---------|
| Success | Single line | `Put: A1 = "Revenue"` |
| Lists | Markdown table | `\| # \| Name \| ... \|` |
| Ranges | Markdown table | `\| \| A \| B \| C \|` |
| Errors | Prefixed | `Error: FileNotFound \| path.xlsx` |

### 4.2 Cell Reference Convention

**Always include explicit references**:

```markdown
# Good - References visible
|   | A        | B       |
|---|----------|---------|
| 1 | Revenue  | $1M     |
| 2 | COGS     | $400K   |

# Bad - No references
| Label   | Value   |
|---------|---------|
| Revenue | $1M     |
```

### 4.3 Error Format

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

## 5. Example Session

Complete workflow exploring and modifying an LBO model:

```bash
# 1. Open and discover
$ xl open lbo_model.xlsx
Opened: lbo_model.xlsx
Sheets: Assumptions, Revenue, Costs, P&L, Returns (5 total)
Active: Assumptions

$ xl sheets
| # | Name        | Range    | Cells | Formulas |
|---|-------------|----------|-------|----------|
| 1 | Assumptions | A1:D25   | 78    | 0        |
| 2 | Revenue     | A1:M50   | 432   | 120      |
| 3 | Costs       | A1:M40   | 312   | 80       |
| 4 | P&L         | A1:M60   | 520   | 180      |
| 5 | Returns     | A1:F20   | 89    | 45       |

# 2. Explore assumptions
$ xl view A1:D10
|   | A              | B        | C    | D    |
|---|----------------|----------|------|------|
| 1 | Entry EV       | $500M    |      |      |
| 2 | Entry Multiple | 8.0x     |      |      |
| 3 | Exit Multiple  | 10.0x    |      |      |
| 4 | Hold Period    | 5        |      |      |
| 5 | Revenue Growth | 10%      |      |      |
| 6 | EBITDA Margin  | 25%      |      |      |
| 7 |                |          |      |      |
| 8 | Debt/EBITDA    | 5.0x     |      |      |
| 9 | Interest Rate  | 8%       |      |      |
| 10| Tax Rate       | 25%      |      |      |

$ xl labels
Found 9 label-value pairs:

| Label          | Value  | Ref |
|----------------|--------|-----|
| Entry EV       | $500M  | B1  |
| Entry Multiple | 8.0x   | B2  |
| Exit Multiple  | 10.0x  | B3  |
| Hold Period    | 5      | B4  |
| Revenue Growth | 10%    | B5  |
...

# 3. Check returns calculation
$ xl select Returns
Selected: Returns
Used range: A1:F20

$ xl cell B10
Cell: B10
Type: Formula
Formula: =IRR(B5:B9)
Cached Value: 0.225
Formatted: 22.5%
Style: Percent
Dependencies: B5, B6, B7, B8, B9

$ xl flow B10
Formula Flow for B10 (IRR):

Inputs (from Assumptions):
  Assumptions!B1: Entry EV = $500M
  Assumptions!B3: Exit Multiple = 10.0x
  Assumptions!B4: Hold Period = 5

Calculation Chain:
  Assumptions!B1 → B5 (Initial Investment)
  P&L!M60 → B9 (Exit Value)

Formula: =IRR(B5:B9)
Result: 22.5%

# 4. Modify assumption and check impact
$ xl select Assumptions
Selected: Assumptions

$ xl put B3 12.0x
Put: B3 = 12.0x (number)

$ xl cell Returns!B10
Cell: Returns!B10
Type: Formula
Formula: =IRR(B5:B9)
Cached Value: 0.283
Formatted: 28.3%

# 5. Save changes
$ xl save
Saved: lbo_model.xlsx (5 sheets, 1431 cells)
```

---

## 6. Implementation Notes

### 6.1 Module Structure

```
xl-cli/
├── src/com/tjclp/xl/cli/
│   ├── Main.scala           # Entry point
│   ├── Session.scala        # State management
│   ├── commands/
│   │   ├── OpenCmd.scala
│   │   ├── NavigateCmd.scala
│   │   ├── ExploreCmd.scala
│   │   ├── AnalyzeCmd.scala
│   │   ├── MutateCmd.scala
│   │   └── PersistCmd.scala
│   └── output/
│       ├── Markdown.scala
│       └── Format.scala
└── test/
```

### 6.2 Dependencies

- `xl-cats-effect` — IO operations (read, write)
- `xl-evaluator` — Formula parsing, dependency analysis
- `decline-effect` — CLI argument parsing

### 6.3 Integration with llm-exploration-apis.md

This CLI exposes the APIs defined in `llm-exploration-apis.md`:

| Scala API | CLI Command |
|-----------|-------------|
| `RangeView.toMarkdown` | `xl view` |
| `sheet.describe` | `xl describe` |
| `view.labels` | `xl labels` |
| `view.regions` | `xl regions` |
| `sheet.formulaFlow` | `xl flow` |

The CLI is a command-line interface to the same underlying functionality.
