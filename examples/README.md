# XL Examples

Standalone demo scripts showcasing XL library features.

## Running Examples

These examples use [scala-cli](https://scala-cli.virtuslab.org/) for easy standalone execution.

### Prerequisites

1. **Install scala-cli** (if not already installed):
   ```bash
   brew install Virtuslab/scala-cli/scala-cli  # macOS
   # or: curl -sSLf https://scala-cli.virtuslab.org/get | sh
   ```

2. **Publish XL modules locally**:
   ```bash
   ./mill __.publishLocal
   ```

### Running Examples

```bash
scala-cli run examples/<example-name>.sc
```

---

## Example Catalog

### Start Here: Scripting Prelude

| Example | Description |
|---------|-------------|
| `scripting_tour.sc` | **Canonical tour of the one-import scripting prelude** (`com.tjclp.xl.scripting.{*, given}`) — source of truth for the xl-scripting skill snippets |

### Core API

| Example | Description |
|---------|-------------|
| `demo.sc` | Core API showcase: addressing, formatting, DSL, rich text |
| `easy_mode_demo.sc` | Easy Mode API: string refs, inline styling, simplified IO |

### Formula System

| Example | Description |
|---------|-------------|
| `quick_start.sc` | Formula system in 5 minutes: parsing, evaluation, round-trip |
| `dependency_analysis.sc` | Formula dependency graph analysis and visualization |
| `text_functions_demo.sc` | Text functions (TRIM, MID, FIND, SUBSTITUTE, VALUE, TEXT) with expected-vs-actual output |

### Real-World Applications

| Example | Description |
|---------|-------------|
| `financial_model.sc` | Complete 3-year income statement model |
| `sales_pipeline.sc` | CRM analytics with conversion funnels |
| `data_validation.sc` | Data quality validation patterns |
| `table_demo.sc` | Excel tables with AutoFilter and styling |
| `resize_demo.sc` | Row/column sizing and HTML/SVG export |

### Streaming & Performance

| Example | Description |
|---------|-------------|
| `big_demo.sc` | Streams 1,000,000 rows with a 256MB heap — O(1) memory write + read |
| `verify_big.sc` | Verifies the million-row output via streaming reads (sample rows + total count) |
| `streaming_benchmark.sc` | Streaming write benchmark: single-pass vs two-pass |
| `streaming_pressure_test.sc` | O(1) memory pressure test for streaming reads over large datasets |
| `calamine_comparison.sc` | Read-only streaming benchmark vs Rust's calamine (needs a local 1M-row NYC 311 sample) |
| `nyc311_roundtrip.sc` | Round-trip benchmark: stream-read 1M rows, stream-write, verify (same NYC 311 sample) |

### Test Utilities

| Example | Description |
|---------|-------------|
| `readme_test.sc` | Validates README.md code examples work correctly |
| `display_test.sc` | Tests display formatting utilities |
| `dependency_test.sc` | Generates test files for CLI dependency tracking |

---

## Featured Examples

### scripting_tour.sc

```bash
scala-cli run examples/scripting_tour.sc
```

The **canonical prelude tour** — one import (`com.tjclp.xl.scripting.{*, given}`) gives a script
the core API, DSL operators, compile-time literals, formula evaluation, sync `Excel` IO, streaming
`ExcelIO`, smart value detection, and the `.unsafe` boundary. Compile-verified by
`scripts/test-examples.sh` and the source of truth for snippets in the xl-scripting skill.

### demo.sc

```bash
scala-cli run examples/demo.sc
```

Showcases:
- Compile-time validated literals (`ref"A1"`, `money"$1,234.56"`)
- Creating workbooks and sheets
- Addressing operations (shift, intersects, contains)
- Sheet-qualified references (`ref"Sales!A1"`)
- Rich text formatting

### easy_mode_demo.sc

```bash
scala-cli run examples/easy_mode_demo.sc
```

Showcases **Easy Mode API** (LLM-friendly ergonomics):
- String-based cell references (`.put("A1", value)`)
- Template-first styling (`.style("A1:B1", style)`)
- Safe lookups (`.cell("A1")` returns `Option`)
- Rich text formatting (`"Error: ".red.bold + "Fix!"`)
- Simplified IO (`Excel.read/write/modify`)

### quick_start.sc

```bash
scala-cli run examples/quick_start.sc
```

Showcases **Formula System**:
- Parse Excel formula strings to typed AST (`FormulaParser`)
- Build formulas programmatically with GADT type safety (`TExpr[A]`)
- Round-trip verification (parse . print = id)
- Scientific notation support
- Error handling with position-aware diagnostics

### table_demo.sc

```bash
scala-cli run examples/table_demo.sc
```

Showcases **Excel Table Support**:
- Create structured tables with `TableSpec.fromColumnNames`
- Apply table styles (Light, Medium, Dark variants)
- Enable AutoFilter for filterable data
- Multiple tables on one sheet
- Full round-trip (write -> read -> verify)

**Output**: Creates a real Excel file with formatted tables, AutoFilter, and styling.

---

## Adding New Examples

Create a new `.sc` script file with shebang and project reference:

```scala
#!/usr/bin/env -S scala-cli shebang
//> using file project.scala

import com.tjclp.xl.{*, given}

println("=== My Example ===")
// Your code here - top-level statements execute directly
```

Make executable and run:
```bash
chmod +x examples/my_example.sc
./examples/my_example.sc
```

The `project.scala` file centralizes dependencies (`com.tjclp::xl:0.11.1`) for all examples.

## Scripting Prelude & Skill

`scripting_tour.sc` is the canonical tour of the one-import scripting prelude
(`import com.tjclp.xl.scripting.{*, given}`). Self-contained, Maven-pinned recipe scripts
(bulk transforms, typed extraction, streaming filters, diffs, CSV ingest) live in the
xl-scripting Claude Code skill: [`plugin/skills/xl-scripting/reference/RECIPES.md`](../plugin/skills/xl-scripting/reference/RECIPES.md).

All examples here are compile-verified against the local build by `scripts/test-examples.sh`
(CI: the `examples` job).
