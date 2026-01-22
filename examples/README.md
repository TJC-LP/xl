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

### Real-World Applications

| Example | Description |
|---------|-------------|
| `financial_model.sc` | Complete 3-year income statement model |
| `sales_pipeline.sc` | CRM analytics with conversion funnels |
| `data_validation.sc` | Data quality validation patterns |
| `table_demo.sc` | Excel tables with AutoFilter and styling |
| `resize_demo.sc` | Row/column sizing and HTML/SVG export |

### Test Utilities

| Example | Description |
|---------|-------------|
| `readme_test.sc` | Validates README.md code examples work correctly |
| `display_test.sc` | Tests display formatting utilities |
| `dependency_test.sc` | Generates test files for CLI dependency tracking |

---

## Featured Examples

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

Create a new `.sc` script file with scala-cli directives:

```scala
//> using scala 3.7.4
//> using dep com.tjclp::xl:0.6.1

import com.tjclp.xl.{*, given}

@main def myExample(): Unit =
  // Your code here
```

Run with: `scala-cli run examples/my_example.sc`
