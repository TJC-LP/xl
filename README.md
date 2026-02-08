# XL — Pure Functional Excel Library for Scala 3

[![CI](https://github.com/TJC-LP/xl/workflows/CI/badge.svg)](https://github.com/TJC-LP/xl/actions)
[![Maven Central](https://img.shields.io/maven-central/v/com.tjclp/xl_3.svg)](https://central.sonatype.com/artifact/com.tjclp/xl_3)
[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)

**The best Excel library for Scala.** Type-safe, purely functional, blazing fast.

```scala 3 raw
//> using scala 3.7.4
//> using dep com.tjclp::xl:0.9.5

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

## Table of Contents

- [Why XL?](#why-xl)
- [Quick Start](#quick-start)
- [Key Features](#key-features)
- [CLI](#cli)
- [Claude Code Plugin](#claude-code-plugin)
- [Documentation](#documentation)
- [Development](#development)
- [Contributing](#contributing)

## Why XL?

- **Type-safe**: Compile-time validated cell references (`ref"A1"`) catch errors before runtime
- **Pure functional**: Immutable data structures, no side effects, deterministic output
- **Streaming**: O(1) memory for reads and writes — handle 1M+ rows without OOM
- **Diffable output**: Canonical XML ordering produces byte-identical files for version control
- **Modern Scala 3**: Opaque types, inline macros, Cats Effect integration

## Quick Start

### Installation

```scala 3 ignore
// build.mill — single dependency for everything
def ivyDeps = Agg(ivy"com.tjclp::xl:0.9.5")

// Or individual modules for minimal footprint:
// ivy"com.tjclp::xl-core:0.9.5"        — Pure domain model only
// ivy"com.tjclp::xl-ooxml:0.9.5"       — Add OOXML read/write
// ivy"com.tjclp::xl-cats-effect:0.9.5" — Add IO streaming
// ivy"com.tjclp::xl-evaluator:0.9.5"   — Add formula evaluation
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

## Key Features

### Compile-Time Validated References

```scala 3 ignore
val cell: ARef = ref"A1"              // Single cell
val range: CellRange = ref"A1:B10"    // Range
val qualified = ref"Sheet1!A1:C100"   // With sheet name

// Runtime interpolation (returns Either)
val col = "A"; val row = "1"
val dynamic = ref"$col$row"           // Either[XLError, RefType]
```

### Formatted Literals

```scala 3 ignore
val price = money"$$1,234.56"         // Currency format
val growth = percent"12.5%"           // Percent format
val date = date"2025-11-24"           // ISO date format
```

> **Note**: Use `$$` to escape `$` in string interpolators (Scala syntax requirement).

### Rich Text

```scala 3 ignore
val text = "Error: ".bold.red + "Fix this!".underline
sheet.put("A1" -> text)  // Multiple formats in one cell
```

### Formula System

Type-safe evaluation and dependency analysis.

```scala 3 ignore
// Add formulas to cells
val sheet = Sheet("Model")
  .put(
    "A1" -> 100,
    "A2" -> 200,
    "A3" -> fx"=SUM(A1:A2)",      // Formula literal
    "A4" -> fx"=A3 * 1.1"
  )

// Evaluate all formulas (with cycle detection)
sheet.evaluateWithDependencyCheck() match
  case Right(results) => // Map[ARef, CellValue] of computed values
  case Left(error)    => // Circular reference detected
```

**Functions**: SUM, AVERAGE, MIN, MAX, COUNT, IF, AND, OR, VLOOKUP, XLOOKUP, SUMIF, COUNTIF, NPV, IRR, TODAY, DATE, and [more](docs/STATUS.md).

### Streaming (Large Files)

```scala 3 ignore
import com.tjclp.xl.io.{ExcelIO, RowData}
import cats.effect.IO
import fs2.Stream

val excel = ExcelIO.instance[IO]

// Read 1M+ rows with O(1) memory
excel.readStream(path)
  .filter(_.rowIndex > 1)  // Skip header
  .evalMap(row => process(row))

// Write 1M+ rows with O(1) memory
Stream.range(1, 1_000_001)
  .map(i => RowData(i, Map(0 -> CellValue.Number(i))))
  .through(excel.writeStream(path, "Data"))
```

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
| `xl -f data.xlsx eval "=SUM(A1:A10)"` | Evaluate formula |
| `xl -f data.xlsx search "pattern"` | Regex search across cells |
| `xl -f in.xlsx -o out.xlsx put B5 1000` | Write value to cell |

See [docs/reference/cli.md](docs/reference/cli.md) for full command reference.

## Claude Code Plugin

Install the xl-cli skill in [Claude Code](https://claude.ai/code) for AI-assisted Excel operations:

```
/plugin marketplace add TJC-LP/xl
/plugin install xl-marketplace@xl
```

The skill provides Claude with comprehensive knowledge of xl CLI commands, enabling natural language Excel manipulation:

- "Read the data from sales.xlsx and show me the first 10 rows"
- "Add a SUM formula to cell B10 that totals B2:B9"
- "Style the header row with bold text and a blue background"
- "Search for all cells containing 'Revenue'"

## Documentation

| Doc | Purpose |
|-----|---------|
| **[docs/INDEX.md](docs/INDEX.md)** | Documentation navigation hub |
| [docs/QUICK-START.md](docs/QUICK-START.md) | Get started in 5 minutes |
| [docs/STATUS.md](docs/STATUS.md) | Current capabilities |
| [examples/](examples/) | Runnable scala-cli examples |
| [CLAUDE.md](CLAUDE.md) | Complete API reference |

```bash
# Run examples with scala-cli
./mill __.publishLocal
scala-cli run examples/quick_start.sc
scala-cli run examples/financial_model.sc
```

## Development

```bash
./mill __.compile          # Compile all
./mill __.test             # Run tests
./mill __.reformat         # Format code
```

## Contributing

1. Fork and create feature branch
2. Ensure `./mill __.test` passes
3. Ensure `./mill __.checkFormat` passes
4. Submit PR with clear description

See [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md) for details.

## License

Apache License 2.0 - see [LICENSE](LICENSE)

---

**XL** is maintained by [TJC-LP](https://github.com/TJC-LP). Built with Scala 3.7, Mill, Cats Effect, and fs2.
