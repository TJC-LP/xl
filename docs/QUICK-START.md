# Quick Start Guide

Get started with XL in 5 minutes.

## Installation

XL is published as a set of JVM libraries. A typical setup depends on:

- `xl-core` for the pure domain model and macros
- `xl-ooxml` for in-memory OOXML read/write
- `xl-cats-effect` for streaming and effectful I/O

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
  "com.tjclp" %% "xl-core"       % "0.1.0",
  "com.tjclp" %% "xl-ooxml"      % "0.1.0",
  "com.tjclp" %% "xl-cats-effect"% "0.1.0"
)
```

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
  .through(Excel.forIO.writeStreamTrue(path, "Data"))
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
