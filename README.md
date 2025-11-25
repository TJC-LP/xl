# XL — Pure Functional Excel Library for Scala 3

[![CI](https://github.com/TJC-LP/xl/workflows/CI/badge.svg)](https://github.com/TJC-LP/xl/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

**The best Excel library for Scala.** Type-safe, purely functional, blazing fast.

```scala
import com.tjclp.xl.{*, given}
import com.tjclp.xl.unsafe.*

// Create a financial report
val report = Sheet("Q1 Report")
  .put("A1", "Revenue")      .put("B1", 1250000, CellStyle.default.currency)
  .put("A2", "Expenses")     .put("B2", 875000, CellStyle.default.currency)
  .put("A3", "Net Income")   .put("B3", fx"=B1-B2")
  .put("A4", "Margin")       .put("B4", fx"=B3/B1")
  .style("A1:A4", CellStyle.default.bold)
  .style("B4", CellStyle.default.percent)
  .unsafe

Excel.write(Workbook.empty.put(report).unsafe, "/tmp/q1-report.xlsx")
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

```scala
// build.mill
def ivyDeps = Agg(
  ivy"com.tjclp::xl-core:0.1.0",
  ivy"com.tjclp::xl-ooxml:0.1.0",
  ivy"com.tjclp::xl-cats-effect:0.1.0"  // For IO
)
```

### Basic Usage

```scala
import com.tjclp.xl.{*, given}
import com.tjclp.xl.unsafe.*

// Read
val workbook = Excel.read("data.xlsx")

// Modify
val updated = workbook.update("Sheet1", sheet =>
  sheet
    .put("A1", "Updated!")
    .put("B1", 42)
    .style("A1:B1", CellStyle.default.bold)
    .unsafe
)

// Write
Excel.write(updated, "output.xlsx")
```

### Patch DSL (Declarative)

```scala
import com.tjclp.xl.{*, given}
import com.tjclp.xl.unsafe.*

val sheet = Sheet("Sales")
  .put(
    (ref"A1" := "Product")   ++ (ref"B1" := "Price")   ++ (ref"C1" := "Qty") ++
    (ref"A2" := "Widget")    ++ (ref"B2" := 19.99)     ++ (ref"C2" := 100) ++
    (ref"A3" := "Gadget")    ++ (ref"B3" := 29.99)     ++ (ref"C3" := 50) ++
    (ref"D1" := "Total")     ++ (ref"D2" := fx"=B2*C2") ++ (ref"D3" := fx"=B3*C3") ++
    ref"A1:D1".styled(CellStyle.default.bold)
  )
  .unsafe
```

## Modern DSL

### Compile-Time Validated References

```scala
val cell: ARef = ref"A1"              // Single cell
val range: CellRange = ref"A1:B10"    // Range
val qualified = ref"Sheet1!A1:C100"   // With sheet name

// Runtime interpolation (returns Either)
val col = "A"; val row = "1"
val dynamic = ref"$col$row"           // Either[XLError, RefType]
```

### Formatted Literals

```scala
val price = money"$$1,234.56"         // Currency format
val growth = percent"12.5%"           // Percent format
val date = date"2025-11-24"           // ISO date format
val loss = accounting"($$500.00)"     // Accounting (negatives in parens)
```

> **Note**: Use `$$` to escape `$` in string interpolators (Scala syntax requirement).

### Fluent Style DSL

```scala
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

```scala
val text = "Error: ".bold.red + "Fix this!".underline
sheet.put("A1", text).unsafe
```

### Patch Composition

```scala
val headerStyle = CellStyle.default.bold.size(14.0)

val patch =
  (ref"A1" := "Title") ++
  ref"A1".styled(headerStyle) ++
  ref"A1:C1".merge

sheet.put(patch).unsafe
```

## Formula System

**24 built-in functions** with type-safe parsing, evaluation, and dependency analysis.

```scala
import com.tjclp.xl.formula.*

// Parse formulas
FormulaParser.parse("=SUM(A1:B10)")        // Right(TExpr.FoldRange(...))
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

**Functions**: SUM, COUNT, AVERAGE, MIN, MAX, IF, AND, OR, NOT, CONCATENATE, LEFT, RIGHT, LEN, UPPER, LOWER, TODAY, NOW, DATE, YEAR, MONTH, DAY, NPV, IRR, VLOOKUP

## Streaming (Large Files)

```scala
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

Run with `scala-cli`:

```bash
scala-cli run examples/quick-start.sc
scala-cli run examples/financial-model.sc
scala-cli run examples/table-demo.sc
```

| Example | Description |
|---------|-------------|
| `quick-start.sc` | Formula system in 5 minutes |
| `financial-model.sc` | 3-year income statement |
| `table-demo.sc` | Excel tables with AutoFilter |
| `easy-mode-demo.sc` | Easy Mode showcase |
| `dependency-analysis.sc` | Formula dependency graph |

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

MIT License - see [LICENSE](LICENSE)

---

**XL** is maintained by [TJC-LP](https://github.com/TJC-LP). Built with Scala 3.7, Mill, Cats Effect, and fs2.
