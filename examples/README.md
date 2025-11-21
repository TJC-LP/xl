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
   ./mill xl-core.publishLocal
   ./mill xl-evaluator.publishLocal  # Required for formula-demo.sc
   ```

### Running demo.sc

```bash
scala-cli run examples/demo.sc
```

The script showcases:
- Compile-time validated literals (`ref"A1"`, `money"$1,234.56"`)
- Creating workbooks and sheets
- Addressing operations (shift, intersects, contains)
- Sheet-qualified references (`ref"Sales!A1"`)
- Formatted literals (money, percent, date)
- Rich text formatting

### Running easy-mode-demo.sc

```bash
scala-cli run examples/easy-mode-demo.sc
```

The script showcases **Easy Mode API** (LLM-friendly ergonomics):
- String-based cell references (`.put("A1", value)`)
- Template-first styling (`.style("A1:B1", style)`)
- Inline styling (`.put("A1", value, style)`)
- Safe lookups (`.cell("A1")` returns `Option`)
- Rich text formatting (`"Error: ".red.bold + "Fix!"`)
- Simplified IO (`Excel.read/write/modify`)
- Structured exception handling (`XLException` wraps `XLError`)

### Running formula-demo.sc

```bash
# Publish required modules
./mill xl-core.publishLocal && ./mill xl-evaluator.publishLocal

# Run formula demo
scala-cli run examples/formula-demo.sc
```

The script showcases **Formula Parser** (WI-07):
- Parse Excel formula strings to typed AST (`FormulaParser`)
- Build formulas programmatically with GADT type safety (`TExpr[A]`)
- Round-trip verification (parse ∘ print = id)
- Scientific notation support (1.5E10, 3.14E-5, etc.)
- Error handling with position-aware diagnostics (`ParseError`)
- Complex nested formulas (nested IF, AND/OR combinations)
- Type safety demo (GADT prevents type mixing at compile time)

**Sections**:
1. Basic parsing (literals, operators, functions)
2. Programmatic construction (build TExpr manually)
3. Round-trip verification (8 test formulas)
4. Scientific notation (7 edge cases)
5. Error handling (6 invalid formulas with diagnostics)
6. Complex nested formulas (4 real-world examples)
7. Type safety demo (GADT compile-time guarantees)

**Note**: Formula evaluation (WI-08) is not yet implemented. This demo focuses on parsing and AST manipulation.

### Running patch-dsl-demo.sc

```bash
scala-cli run examples/patch-dsl-demo.sc
```

Demonstrates declarative sheet updates using the Patch DSL.

## Adding New Examples

Create a new `.sc` script file with scala-cli directives:

```scala
//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using repository ivy2Local

import com.tjclp.xl.*

@main def myExample(): Unit =
  // Your code here
```

Run with: `scala-cli run examples/my-example.sc`
