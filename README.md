# XL — Pure Functional Excel Library for Scala 3

[![CI](https://github.com/TJC-LP/xl/workflows/CI/badge.svg)](https://github.com/TJC-LP/xl/actions)
[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)

**The best Excel library for Scala.** Type-safe, purely functional, blazing fast.

```scala 3 raw
//> using scala 3.7.4
//> using dep com.tjclp::xl:0.4.2

import com.tjclp.xl.{*, given}

@main def demo(): Unit =
  // Create a financial report
  val report = Sheet("Q1 Report")
    .put(
      "A1" -> "Revenue",     "B1" -> money"$$1,250,000",
      "A2" -> "Expenses",    "B2" -> money"$$875,000",
      "A3" -> "Net Income",  "B3" -> fx"=B1-B2",
      "A4" -> "Margin",      "B4" -> fx"=B3/$$B$$1"
    )
    .style(
      "A1:A4" -> CellStyle.default.bold,
      "B4"    -> CellStyle.default.percent
    )

  Excel.write(Workbook(report), "/tmp/q1-report.xlsx")
  println(s"Created /tmp/q1-report.xlsx with ${report.cellCount} cells")
```

## Why XL?

| | XL | Apache POI |
|-|-----|-----------|
| **Memory** | 10MB constant (streaming) | 800MB for 100k rows |
| **Speed** | 35% faster reads | Baseline |
| **Type Safety** | Compile-time validated | Runtime errors |
| **Purity** | No side effects, immutable | Mutable everywhere |
| **Output** | Byte-identical, diffable | Non-deterministic |

## Quick Start

### Dependencies

```scala 3 ignore
// build.mill — single dependency for everything
def ivyDeps = Agg(ivy"com.tjclp::xl:0.4.2")

// Or individual modules for minimal footprint:
// ivy"com.tjclp::xl-core:0.4.2"        — Pure domain model only
// ivy"com.tjclp::xl-ooxml:0.4.2"       — Add OOXML read/write
// ivy"com.tjclp::xl-cats-effect:0.4.2" — Add IO streaming
// ivy"com.tjclp::xl-evaluator:0.4.2"   — Add formula evaluation
```

### Basic Usage

```scala 3 ignore
import com.tjclp.xl.{*, given}

// Read
val workbook = Excel.read("data.xlsx")

// Modify — string literals are validated at compile time
val updated = workbook.update("Sheet1", sheet =>
  sheet
    .put("A1" -> "Updated!", "B1" -> 42)
    .style("A1:B1" -> CellStyle.default.bold)
)

// Write
Excel.write(updated, "output.xlsx")
```

### Patch DSL (Declarative)

```scala 3 ignore
import com.tjclp.xl.{*, given}

val sheet = Sheet("Sales")
  .put(
    "A1" -> "Product",  "B1" -> "Price",  "C1" -> "Qty",   "D1" -> "Total",
    "A2" -> "Widget",   "B2" -> 19.99,    "C2" -> 100,     "D2" -> fx"=B2*C2",
    "A3" -> "Gadget",   "B3" -> 29.99,    "C3" -> 50,      "D3" -> fx"=B3*C3"
  )
  .style("A1:D1" -> CellStyle.default.bold)
```

## Modern DSL

### Compile-Time Validated References

```scala 3 ignore
val cell: ARef = ref"A1"              // Single cell
val range: CellRange = ref"A1:B10"    // Range
val qualified = ref"Sheet1!A1:C100"   // With sheet name

// Runtime interpolation (returns Either)
val col = "A"; val row = "1"
val dynamic = ref"$col$row"           // Either[XLError, RefType]
```

### Runtime References (Pure FP)

When cell references come from user input or configuration, XL returns `Either` to ensure errors are handled:

```scala 3 ignore
import com.tjclp.xl.{*, given}

val userInput = "B5"  // From config, CLI, etc.

// Pure approach: pattern match on result
sheet.put(userInput -> 100) match
  case Right(updated) => Excel.write(Workbook(updated), "out.xlsx")
  case Left(error)    => println(s"Invalid ref: ${error.message}")

// Escape hatch: .unsafe throws on invalid refs (for scripts/demos)
import com.tjclp.xl.unsafe.*
val updated = sheet.put(userInput -> 100).unsafe  // Throws if invalid
```

This design catches errors like `"ZZZ999999"` at the appropriate boundary — compile time for literals, explicit handling for runtime values.

### Formatted Literals

```scala 3 ignore
val price = money"$$1,234.56"         // Currency format
val growth = percent"12.5%"           // Percent format
val date = date"2025-11-24"           // ISO date format
val loss = accounting"($$500.00)"     // Accounting (negatives in parens)
```

> **Note**: Use `$$` to escape `$` in string interpolators (Scala syntax requirement).

### Fluent Style DSL

```scala 3 ignore
val header = CellStyle.default
  .bold
  .size(14.0)
  .bgBlue
  .white
  .center
  .bordered

val currency = CellStyle.default.currency
val percent = CellStyle.default.percent
```

### Rich Text (Multi-Format in One Cell)

```scala 3 ignore
val text = "Error: ".bold.red + "Fix this!".underline
sheet.put("A1" -> text)  // Literal ref → no .unsafe needed
```

### Patch Composition

```scala 3 ignore
val headerStyle = CellStyle.default.bold.size(14.0)

// Patches use := syntax for complex composition
val patch =
  (ref"A1" := "Title") ++
  ref"A1".styled(headerStyle) ++
  ref"A1:C1".merge

sheet.put(patch)  // Patch application is infallible
```

## Formula System

**30 built-in functions** with type-safe parsing, evaluation, and dependency analysis.

```scala 3 ignore
import com.tjclp.xl.formula.*

// Parse formulas (supports $ anchoring)
FormulaParser.parse("=SUM($A$1:B10)")      // Right(TExpr.FoldRange(...))
FormulaParser.parse("=IF(A1>0, B1, C1)")   // Right(TExpr.If(...))

// Evaluate with cycle detection
import com.tjclp.xl.formula.SheetEvaluator.*
sheet.evaluateWithDependencyCheck() match
  case Right(results) => results.foreach(println)
  case Left(error)    => println(s"Cycle: ${error.message}")

// Analyze dependencies
val graph = DependencyGraph.fromSheet(sheet)
val precedents = DependencyGraph.precedents(graph, ref"B1")
```

**Functions**: SUM, COUNT, AVERAGE, MIN, MAX, IF, AND, OR, NOT, CONCATENATE, LEFT, RIGHT, LEN, UPPER, LOWER, TODAY, NOW, DATE, YEAR, MONTH, DAY, NPV, IRR, VLOOKUP, XLOOKUP, SUMIF, COUNTIF, SUMIFS, COUNTIFS, SUMPRODUCT

## Streaming (Large Files)

```scala 3 ignore
import com.tjclp.xl.io.ExcelIO
import cats.effect.IO

// Constant O(1) memory for any file size
val excel = ExcelIO.instance[IO]

// Read 1M+ rows without OOM
excel.readStream(path).evalMap(row => process(row))

// Write with SAX backend (fastest)
excel.writeFast(workbook, path)
```

**Performance**: 100k rows in ~1.8s read, ~1.1s write, ~10MB memory

## Examples

Run examples with `scala-cli` (requires local publish first):

```bash
# Publish locally (one-time setup)
./mill __.publishLocal

# Run examples
scala-cli run examples/quick-start.sc
scala-cli run examples/financial-model.sc

# Run the README itself!
scala-cli README.md
```

| Example | Description |
|---------|-------------|
| `quick-start.sc` | Formula system in 5 minutes |
| `financial-model.sc` | 3-year income statement |
| `table-demo.sc` | Excel tables with AutoFilter |
| `easy-mode-demo.sc` | Easy Mode showcase |
| `dependency-analysis.sc` | Formula dependency graph |

## CLI

Install the `xl` command-line tool for LLM-friendly Excel operations:

```bash
make install   # Installs to ~/.local/bin/xl
```

Stateless by design — each command is self-contained:

| Command | Description |
|---------|-------------|
| `xl -f data.xlsx sheets` | List all sheets with cell counts |
| `xl -f data.xlsx view A1:D20` | View range as markdown table |
| `xl -f data.xlsx cell A1` | Get cell details + dependencies |
| `xl -f data.xlsx eval "=SUM(A1:A10)"` | Evaluate formula (what-if) |
| `xl -f data.xlsx stats A1:A100` | Calculate min/max/sum/mean/count |
| `xl -f data.xlsx search "pattern"` | Regex search across cells |
| `xl -f in.xlsx -o out.xlsx put B5 1000` | Write value to cell |
| `xl -f in.xlsx -o out.xlsx putf C5 "=B5*1.1"` | Write formula to cell |

```bash
# Example: what-if analysis with temporary overrides
xl -f model.xlsx eval "=A1*1.1" -w "A1=100"

# Example: formula dragging with $ anchoring
xl -f model.xlsx -o out.xlsx putf B2:B10 "=SUM(\$A\$1:A2)" --from B2
```

See [docs/plan/xl-cli.md](docs/plan/xl-cli.md) for full command reference.

## Project Status

| Feature | Status | Tests |
|---------|--------|-------|
| Core Domain | ✅ Production | 500+ |
| OOXML Read/Write | ✅ Production | 145+ |
| Streaming I/O | ✅ Production | 30+ |
| Formula System | ✅ Production | 169+ |
| Excel Tables | ✅ Production | 45+ |
| **Total** | **731+ tests** | |

See [STATUS.md](docs/STATUS.md) for detailed coverage and [roadmap.md](docs/plan/roadmap.md) for future work.

## Development

```bash
./mill __.compile          # Compile all
./mill __.test             # Run tests
./mill __.reformat         # Format code
./mill xl-benchmarks.run   # Run JMH benchmarks
```

## Documentation

| Doc | Purpose |
|-----|---------|
| [CLAUDE.md](CLAUDE.md) | Complete API reference |
| [STATUS.md](docs/STATUS.md) | Current capabilities |
| [roadmap.md](docs/plan/roadmap.md) | Future work |
| [examples/](examples/) | Runnable examples |
| [LIMITATIONS.md](docs/LIMITATIONS.md) | Known limitations |

## Contributing

1. Fork and create feature branch
2. Ensure `./mill __.test` passes
3. Ensure `./mill __.checkFormat` passes
4. Submit PR with clear description

See [CONTRIBUTING.md](docs/CONTRIBUTING.md) for details.

## License

Apache License 2.0 - see [LICENSE](LICENSE)

---

**XL** is maintained by [TJC-LP](https://github.com/TJC-LP). Built with Scala 3.7, Mill, Cats Effect, and fs2.
