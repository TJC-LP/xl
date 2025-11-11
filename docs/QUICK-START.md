# Quick Start Guide

Get started with XL in 5 minutes.

## Installation

### With Mill (build.sc)
```scala
import mill._, scalalib._

object myproject extends ScalaModule {
  def scalaVersion = "3.7.3"

  def ivyDeps = Agg(
    ivy"com.tjclp::xl-core:0.1.0",
    ivy"com.tjclp::xl-ooxml:0.1.0",
    ivy"com.tjclp::xl-cats-effect:0.1.0"
  )
}
```

### With sbt (build.sbt)
```scala
scalaVersion := "3.7.3"

libraryDependencies ++= Seq(
  "com.tjclp" %% "xl-core" % "0.1.0",
  "com.tjclp" %% "xl-ooxml" % "0.1.0",
  "com.tjclp" %% "xl-cats-effect" % "0.1.0"
)
```

## Your First Spreadsheet (30 seconds)

```scala
import com.tjclp.xl.*
import com.tjclp.xl.macros.*
import com.tjclp.xl.codec.{*, given}
import com.tjclp.xl.io.ExcelIO
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import java.nio.file.Path

// Create a simple sheet
val sheet = Sheet("Sales").get.putMixed(
  cell"A1" -> "Product",
  cell"B1" -> "Revenue",
  cell"A2" -> "Widget",
  cell"B2" -> 1000
)

// Write to file
val workbook = Workbook(Vector(sheet))
ExcelIO.instance.write[IO](workbook, Path.of("sales.xlsx")).unsafeRunSync()
```

**Result**: `sales.xlsx` with 2x2 grid

## Reading a File (30 seconds)

```scala
import com.tjclp.xl.io.ExcelIO
import cats.effect.IO
import cats.effect.unsafe.implicits.global

ExcelIO.instance.read[IO](Path.of("sales.xlsx")).map {
  case Right(workbook) =>
    // Access first sheet
    val sheet = workbook.sheets.head

    // Read cell value
    sheet.get(cell"B2") match
      case Some(cell) => println(s"Revenue: ${cell.value}")
      case None => println("Cell empty")

  case Left(error) =>
    println(s"Error: ${error.message}")
}.unsafeRunSync()
```

## Type-Safe Reading (1 minute)

```scala
import com.tjclp.xl.codec.{*, given}

val sheet = workbook.sheets.head

// Type-safe reading with Either
sheet.readTyped[Int](cell"B2") match
  case Right(Some(revenue)) => println(s"Revenue: $$${revenue}")
  case Right(None) => println("Empty cell")
  case Left(error) => println(s"Type error: ${error}")
```

## Adding Styles (2 minutes)

```scala
import com.tjclp.xl.style.*

// Define a header style
val headerStyle = CellStyle.default
  .withFont(Font("Arial", 14.0, bold = true))
  .withFill(Fill.Solid(Color.fromHex("#CCCCCC")))
  .withAlign(Align(HAlign.Center, VAlign.Middle))

// Register and apply style
val (registry, styleId) = sheet.styleRegistry.register(headerStyle)
val styled = sheet
  .copy(styleRegistry = registry)
  .withCellStyle(cell"A1", styleId)
  .withCellStyle(cell"B1", styleId)
```

**Shortcut with Patch DSL**:
```scala
import com.tjclp.xl.dsl.*

val patch = (cell"A1" := "Product") ++
            (cell"B1" := "Revenue") ++
            range"A1:B1".styled(headerStyle)

sheet.applyPatch(patch).get
```

## Performance Modes

### Small Files (<10k rows)
```scala
// Use in-memory API (simple, full features)
val sheet = Sheet("Data").get.putMixed(
  // ... cells
)
ExcelIO.instance.write[IO](Workbook(Vector(sheet)), path).unsafeRunSync()
```

**Memory**: ~10MB

---

### Large Files (100k+ rows)
```scala
// Use streaming write (constant memory)
import com.tjclp.xl.io.{Excel, RowData}
import fs2.Stream

Stream.range(1, 1_000_001)
  .map(i => RowData(i, Map(
    0 -> CellValue.Text(s"Row $i"),
    1 -> CellValue.Number(BigDecimal(i))
  )))
  .through(Excel.forIO.writeStreamTrue(path, "Data"))
  .compile.drain
  .unsafeRunSync()
```

**Memory**: ~10MB constant (even for 10M rows!)

**Limitation**: No SST or styles in streaming write yet (see roadmap)

---

## Common Tasks

### Create Multi-Sheet Workbook
```scala
val sheet1 = Sheet("Sales").get.putMixed(cell"A1" -> "Sales Data")
val sheet2 = Sheet("Inventory").get.putMixed(cell"A1" -> "Inventory")

val workbook = Workbook(Vector(sheet1, sheet2))
```

### Fill a Range
```scala
sheet.fillBy(range"A1:Z10") { (col, row) =>
  CellValue.Text(s"${col.toLetter}${row.index1}")
}
```

### Rich Text (Multiple Formats in One Cell)
```scala
import com.tjclp.xl.RichText.*

val text = "Error: ".red.bold + "File not found"
sheet.putMixed(cell"A1" -> text)
```

### Export to HTML
```scala
val html = sheet.toHtml(range"A1:B10")
println(html)  // <table>...</table> with inline CSS
```

---

## Next Steps

1. **Read the README**: [README.md](../README.md) for comprehensive documentation
2. **Check examples**: [docs/reference/examples.md](reference/examples.md) for common patterns
3. **Performance guide**: [docs/reference/performance-guide.md](reference/performance-guide.md) for optimization
4. **Migrating from POI**: [docs/reference/migration-from-poi.md](reference/migration-from-poi.md) if coming from Java

## Quick Reference

### Imports You'll Need
```scala
import com.tjclp.xl.*                  // Core types
import com.tjclp.xl.macros.*           // cell"A1", range"A1:B10"
import com.tjclp.xl.codec.{*, given}   // Type-safe codecs (putMixed, readTyped)
import com.tjclp.xl.style.*            // CellStyle, Font, Fill, etc.
import com.tjclp.xl.dsl.*              // := operator, ++ combinator
import com.tjclp.xl.io.ExcelIO         // File I/O
import cats.effect.IO                  // Effect system
import cats.effect.unsafe.implicits.global  // unsafeRunSync
import java.nio.file.Path              // File paths
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
- `sheet.putAll(cells)` - Batch set cells
- `sheet.putMixed(updates*)` - Type-safe batch with auto-formatting
- `sheet.get(ref)` - Read cell (Option[Cell])
- `sheet.readTyped[A](ref)` - Type-safe read (Either[Error, Option[A]])
- `sheet.applyPatch(patch)` - Apply deferred updates

### Common Patterns
```scala
// Single cell
sheet.put(cell"A1", "Hello")

// Batch cells
sheet.putMixed(
  cell"A1" -> "Name",
  cell"B1" -> 42,
  cell"C1" -> LocalDate.now()
)

// Apply style
sheet.withCellStyle(cell"A1", styleId)

// Range operations
sheet.fillBy(range"A1:A10")((_col, row) => CellValue.Number(row.index1))

// Read with type safety
sheet.readTyped[BigDecimal](cell"C1")
```

## Troubleshooting

### "Type mismatch: found IO[Unit], required Unit"
**Solution**: Call `.unsafeRunSync()` to execute IO:
```scala
ExcelIO.instance.write[IO](wb, path).unsafeRunSync()
```

### "Value readTyped is not a member of Sheet"
**Solution**: Import codec givens:
```scala
import com.tjclp.xl.codec.{*, given}
```

### "Macro expansion error: Invalid cell reference"
**Solution**: Check cell reference syntax (must be valid A1 notation):
```scala
cell"A1"    // ✅ Valid
cell"AA100" // ✅ Valid
cell"1A"    // ❌ Invalid (number first)
cell"XFE1"  // ❌ Invalid (column out of range)
```

### "Sheet name contains invalid character"
**Solution**: Use SheetName.apply for validation:
```scala
SheetName("My Sheet!") match
  case Right(name) => Workbook(name)
  case Left(err) => println(s"Invalid: $err")

// Or use unsafe (throws on invalid):
Workbook(SheetName.unsafe("ValidName"))
```

---

## Performance Tips

**Small files (<10k rows)**: Use in-memory API with putMixed for batching

**Large files (100k+ rows)**: Use streaming write:
```scala
Stream.range(1, 1_000_001)
  .map(i => RowData(i, Map(0 -> CellValue.Text(s"Row $i"))))
  .through(Excel.forIO.writeStreamTrue(path, "Data"))
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
