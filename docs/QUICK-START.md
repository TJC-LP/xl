# Quick Start Guide

Get started with XL in 5 minutes.

## Installation

XL is published as JVM libraries to Maven Central. Use the aggregate `xl` artifact for easy setup.

### With Mill (build.mill)
```scala
import mill._, scalalib._

object myproject extends ScalaModule {
  def scalaVersion = "3.7.4"
  def ivyDeps = Agg(ivy"com.tjclp::xl:0.9.2")
}
```

### With sbt (build.sbt)
```scala
scalaVersion := "3.7.4"
libraryDependencies += "com.tjclp" %% "xl" % "0.9.2"
```

### With Scala CLI
```scala
//> using dep com.tjclp::xl:0.9.2
```

### Individual Modules (Optional)

For minimal dependencies, use individual modules:

| Module | Description |
|--------|-------------|
| `xl-core` | Pure domain model, macros, DSL |
| `xl-ooxml` | OOXML read/write |
| `xl-cats-effect` | IO streaming with Cats Effect |
| `xl-evaluator` | Formula parser and evaluator |

## Your First Spreadsheet (30 seconds)

```scala
import com.tjclp.xl.*            // Domain model + macros + DSL
import com.tjclp.xl.unsafe.*     // Explicit opt-in for .unsafe
import com.tjclp.xl.io.ExcelIO
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import java.nio.file.Path

// Build a simple sheet (Sales) with two rows
val sheet =
  Sheet("Sales").unsafe
    .put(
      ref"A1" -> "Product",
      ref"B1" -> "Revenue",
      ref"A2" -> "Widget",
      ref"B2" -> 1000
    )
    .unsafe

// Create a workbook and add the sheet
val workbook =
  Workbook.empty
    .flatMap(_.put(sheet))
    .unsafe

// Write to file
ExcelIO.instance[IO].write(workbook, Path.of("sales.xlsx")).unsafeRunSync()
```

**Result**: `sales.xlsx` with a 2×2 grid.

## Reading a File (30 seconds)

```scala
import com.tjclp.xl.io.ExcelIO
import cats.effect.unsafe.implicits.global
import cats.effect.IO
import com.tjclp.xl.*

val excel = ExcelIO.instance[IO]
val path  = Path.of("sales.xlsx")

val program: IO[Unit] =
  excel.read(path).flatMap { workbook =>
    // Access first sheet
    val sheet = workbook.sheets.head

    // Read cell value at B2 (returns a Cell; empty cells have CellValue.Empty)
    val cell = sheet(ref"B2")
    IO.println(s"Revenue cell value: ${cell.value}")
  }

program.unsafeRunSync()
```

## Type-Safe Reading (1 minute)

```scala
import com.tjclp.xl.*
import com.tjclp.xl.codec.syntax.*

val sheet = workbook.sheets.head

// Type-safe reading with Either
sheet.readTyped[Int](ref"B2") match
  case Right(Some(revenue)) => println(s"Revenue: $$${revenue}")
  case Right(None) => println("Empty cell")
  case Left(error) => println(s"Type error: ${error}")
```

## Adding Styles (2 minutes)

```scala
import com.tjclp.xl.*

// Define a header style (bold, centered, gray background)
val headerStyle = CellStyle.default
  .bold.size(14.0).fontFamily("Arial")
  .bgGray.bordered
  .center.middle

// Apply style to A1 and B1 using the sheet extension
val styledSheet =
  sheet
    .withCellStyle(ref"A1", headerStyle)
    .withCellStyle(ref"B1", headerStyle)
```

**Shortcut with Patch DSL**:
```scala
import com.tjclp.xl.*
import com.tjclp.xl.dsl.*
import com.tjclp.xl.unsafe.*

val patch =
  (ref"A1" := "Product") ++
  (ref"B1" := "Revenue") ++
  ref"A1".styled(headerStyle) ++
  ref"B1".styled(headerStyle)

val styledSheet2 = sheet.put(patch).unsafe
```

## Performance Modes

### Small Files (<10k rows)
```scala
// Use in-memory API (simple, full features)
import com.tjclp.xl.*
import com.tjclp.xl.unsafe.*
import com.tjclp.xl.io.ExcelIO
import cats.effect.IO
import cats.effect.unsafe.implicits.global

val sheet =
  Sheet("Data").unsafe
    .put(
      ref"A1" -> "Hello",
      ref"B1" -> 42
    )
    .unsafe

val workbook =
  Workbook.empty
    .flatMap(_.put(sheet))
    .unsafe

ExcelIO.instance[IO].write(workbook, path).unsafeRunSync()
```

**Memory**: ~10MB

---

### Large Files (100k+ rows)
```scala
// Use streaming write (constant memory)
import com.tjclp.xl.*
import com.tjclp.xl.io.{Excel, RowData}
import fs2.Stream

Stream.range(1, 1_000_001)
  .map(i => RowData(i, Map(
    0 -> CellValue.Text(s"Row $i"),
    1 -> CellValue.Number(BigDecimal(i))
  )))
  .through(Excel.forIO.writeStream(path, "Data"))
  .compile.drain
  .unsafeRunSync()
```

**Memory**: ~10MB constant (even for 10M rows!)

**Limitations**: Streaming writers use inline strings and minimal styles (no rich formatting or merges).

---

## Common Tasks

### Create Multi-Sheet Workbook
```scala
import com.tjclp.xl.*
import com.tjclp.xl.unsafe.*

val sheet1 = Sheet("Sales").unsafe.put(ref"A1" -> "Sales Data").unsafe
val sheet2 = Sheet("Inventory").unsafe.put(ref"A1" -> "Inventory").unsafe

val workbook = Workbook(Vector(sheet1, sheet2))
```

### Fill a Range
```scala
import com.tjclp.xl.*

sheet.fillBy(ref"A1:Z10") { (col, row) =>
  CellValue.Text(s"${col.toLetter}${row.index1}")
}
```

### Rich Text (Multiple Formats in One Cell)
```scala
import com.tjclp.xl.richtext.RichText.*

val text = "Error: ".red.bold + "File not found"
sheet.put(ref"A1" -> text)
```

### Export to HTML
```scala
val html = sheet.toHtml(range"A1:B10")
println(html)  // <table>...</table> with inline CSS
```

### Formula Evaluation (Production Ready)
```scala
import com.tjclp.xl.*
import com.tjclp.xl.formula.{FormulaParser, Evaluator}
import com.tjclp.xl.formula.SheetEvaluator.*
import com.tjclp.xl.unsafe.*

// Build a sheet with data and formulas
val sheet = Sheet("Finance").unsafe
  .put(
    ref"A1" -> "Revenue",
    ref"A2" -> BigDecimal("1000000"),
    ref"A3" -> BigDecimal("1500000"),
    ref"B1" -> "Total",
    ref"B2" -> CellValue.Formula("=SUM(A2:A3)"),  // Formula cell
    ref"C1" -> "Average",
    ref"C2" -> CellValue.Formula("=AVERAGE(A2:A3)")
  )
  .unsafe

// Evaluate individual formula
val totalResult = sheet.evaluateFormula("=SUM(A2:A3)")
// Right(CellValue.Number(2500000))

// Evaluate cell with formula
val cellResult = sheet.evaluateCell(ref"B2")
// Right(CellValue.Number(2500000))

// Evaluate all formulas in sheet (with dependency checking)
val allResults = sheet.evaluateWithDependencyCheck()
// Right(Map(
//   ARef(B2) -> CellValue.Number(2500000),
//   ARef(C2) -> CellValue.Number(1250000)
// ))

// Handle circular references safely
val cyclicSheet = Sheet("Cyclic").unsafe
  .put(
    ref"A1" -> CellValue.Formula("=B1"),
    ref"B1" -> CellValue.Formula("=A1")  // Circular reference!
  )
  .unsafe

cyclicSheet.evaluateWithDependencyCheck() match
  case Left(error) =>
    println(s"Cycle detected: ${error.message}")
    // "Circular reference detected: A1 → B1 → A1"
  case Right(_) => // Won't happen
```

**Available Functions** (81 total):
- **Aggregate**: SUM, COUNT, COUNTA, COUNTBLANK, AVERAGE, MEDIAN, MIN, MAX, STDEV, VAR
- **Conditional**: SUMIF, COUNTIF, SUMIFS, COUNTIFS, AVERAGEIF, AVERAGEIFS, SUMPRODUCT
- **Logical**: IF, AND, OR, NOT, ISNUMBER, ISTEXT, ISBLANK, ISERR, ISERROR
- **Text**: CONCATENATE, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, SUBSTITUTE, TEXT, VALUE
- **Date**: TODAY, NOW, DATE, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, EOMONTH, EDATE
- **Math**: ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, POWER, SQRT, LOG, LN, EXP, PI
- **Financial**: NPV, IRR, XNPV, XIRR, PMT, FV, PV, RATE, NPER
- **Lookup**: VLOOKUP, XLOOKUP, INDEX, MATCH

See CLAUDE.md for the complete list.

---

## Next Steps

1. **Read the README**: [README.md](../README.md) for comprehensive documentation
2. **Check examples**: [docs/reference/examples.md](reference/examples.md) for common patterns
3. **Performance guide**: [docs/reference/performance-guide.md](reference/performance-guide.md) for optimization
4. **Migrating from POI**: [docs/reference/migration-from-poi.md](reference/migration-from-poi.md) if coming from Java

## Quick Reference

### Imports You'll Need
```scala
import com.tjclp.xl.*                   // Core types, macros, DSL
import com.tjclp.xl.codec.syntax.*      // Type-safe codecs (readTyped, etc.)
import com.tjclp.xl.dsl.*               // := operator, ++ combinator
import com.tjclp.xl.io.{ExcelIO, Excel} // File and streaming I/O
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import java.nio.file.Path
```

### Key Types
- `Sheet` - Single worksheet
- `Workbook` - Multi-sheet workbook
- `Cell` - Single cell (ref + value + style)
- `ARef` - Cell reference (A1, B2, etc.)
- `CellRange` - Range of cells (A1:B10)
- `CellValue` - Cell content (Text, Number, Bool, DateTime, Formula, etc.)
- `CellStyle` - Formatting (font, fill, border, numFmt, align)

### Key Operations
- `sheet.put(ref, value)` - Set cell value
- `sheet.put(cells)` - Batch set cells
- `sheet.put(updates*)` - Type-safe batch with auto-formatting
- `sheet(ref)` - Read cell (returns a `Cell`, empty cells have `CellValue.Empty`)
- `sheet.readTyped[A](ref)` - Type-safe read (Either[CodecError, Option[A]])
- `sheet.put(patch)` - Apply deferred updates

### Common Patterns
```scala
// Single cell
sheet.put(ref"A1", "Hello")

// Batch cells
sheet.put(
  ref"A1" -> "Name",
  ref"B1" -> 42,
  ref"C1" -> java.time.LocalDate.now()
)

// Apply style
sheet.withCellStyle(ref"A1", headerStyle)

// Range operations
sheet.fillBy(ref"A1:A10") { (_col, row) =>
  CellValue.Number(BigDecimal(row.index1))
}

// Read with type safety
sheet.readTyped[BigDecimal](ref"C1")
```

## Troubleshooting

### "Type mismatch: found IO[Unit], required Unit"
**Solution**: Call `.unsafeRunSync()` to execute IO:
```scala
ExcelIO.instance[IO].write(wb, path).unsafeRunSync()
```

### "Value readTyped is not a member of Sheet"
**Solution**: Import codec givens:
```scala
import com.tjclp.xl.codec.syntax.*
```

### "Macro expansion error: Invalid cell reference"
**Solution**: Check cell reference syntax (must be valid A1 notation):
```scala
ref"A1"      // ✅ Valid
ref"AA100"   // ✅ Valid
ref"1A"      // ❌ Invalid (number first)
ref"XFE1"    // ❌ Invalid (column out of range)
```

### "Sheet name contains invalid character"
**Solution**: Use SheetName.apply for validation:
```scala
import com.tjclp.xl.addressing.SheetName

SheetName("My Sheet!") match
  case Right(name) => Workbook(name)
  case Left(err) => println(s"Invalid: $err")

// Or use unsafe (throws on invalid):
Workbook(SheetName.unsafe("ValidName"))
```

---

## Performance Tips

**Small files (<10k rows)**: Use the in‑memory API and batch with `sheet.put(ref -> value, ...)`.

**Large files (100k+ rows)**: Use streaming write:
```scala
import com.tjclp.xl.*
import com.tjclp.xl.io.{Excel, RowData}
import fs2.Stream
import cats.effect.unsafe.implicits.global

Stream.range(1, 1_000_001)
  .map(i => RowData(i, Map(0 -> CellValue.Text(s"Row $i"))))
  .through(Excel.forIO.writeStream(path, "Data"))
  .compile.drain
  .unsafeRunSync()
```

**See**: [docs/reference/performance-guide.md](reference/performance-guide.md) for mode selection guide

---

## Help & Support

- **Documentation**: [README.md](../README.md)
- **Examples**: [docs/reference/examples.md](reference/examples.md)
- **Issues**: [GitHub Issues](https://github.com/TJC-LP/xl/issues)
- **Contributing**: [CONTRIBUTING.md](CONTRIBUTING.md)

---

Happy spreadsheeting!
