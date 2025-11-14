# XL — Pure Functional Excel Library for Scala 3

[![CI](https://github.com/TJC-LP/xl/workflows/CI/badge.svg)](https://github.com/TJC-LP/xl/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

A purely functional, mathematically rigorous Excel (OOXML) library for Scala 3.7 with **zero-overhead opaque types**, **law-governed semantics**, and **deterministic output**.

## Features

- **Type-Safe Addressing**: Zero-cost opaque types for columns, rows, and cell references with compile-time validated literals
- **Pure Domain Model**: 100% pure core with no side effects, nullable values, or partial functions
- **Immutable Workbooks**: Persistent data structures with efficient structural sharing
- **Monoid Composition**: Declarative updates via lawful Patch and StylePatch monoids
- **Ergonomic DSL**: Clean operators (`:=`, `++`, `.merge`) without type ascription friction
- **Comprehensive Styling**: Colors, fonts, fills, borders, number formats, and alignment
- **Cell-Level Codecs**: Type-safe encoding/decoding with auto-inferred formatting for 9 primitive types (including RichText)
- **Rich Text DSL**: Composable intra-cell formatting with String extensions (`.bold.red + " text"`)
- **HTML Export**: Convert sheet ranges to HTML tables with inline CSS
- **Pure Error Channels**: ExcelR[F] for explicit error handling without exceptions
- **Optics & Focus**: Lens and Optional for composable, law-governed transformations
- **Range Combinators**: fillBy, tabulate, putRow, putCol for declarative sheet building
- **Configurable Writer**: Control compression (DEFLATED/STORED), SST policy, and XML formatting
- **Deterministic Output**: Canonical XML ordering for stable git diffs and reproducible builds
- **Property-Based Testing**: ScalaCheck generators and law tests for all core algebras

## Status

**Current**: ~89% complete with 263/263 tests passing ✅

**Implemented** (P0-P6):
- ✅ Core domain types (Cell, Sheet, Workbook)
- ✅ Addressing system with compile-time macros (`ref"A1"`, `ref"A1:B10"`)
- ✅ Patch system with Monoid composition for updates
- ✅ Complete style system (fonts, colors, fills, borders, alignment)
- ✅ **OOXML I/O**: Read/write XLSX files with SST and styles.xml
- ✅ **Streaming**: True constant-memory streaming with fs2-data-xml (100k+ rows)
- ✅ **Cell Codecs**: Type-safe encoding/decoding with auto-inferred formatting
- ✅ Elegant syntax: given conversions, batch put macro, formatted literals

**Roadmap**: Formulas, codecs with case class derivation, charts, drawings, tables.

See [docs/plan/18-roadmap.md](docs/plan/18-roadmap.md) for full implementation plan.

## Quick Start

### Prerequisites

- **JDK**: 25 LTS (Temurin recommended)
- **Scala**: 3.7.3 (managed by Mill)
- **Build Tool**: Mill 0.12.x (included via `./mill` wrapper)

### Installation

```bash
# Clone the repository
git clone https://github.com/TJC-LP/xl.git
cd xl

# Verify setup
./mill __.compile

# Run tests
./mill __.test
```

### Basic Usage

```scala
import com.tjclp.xl.*
// Unified import: all domain types, syntax, macros (ref", money", date", etc.)

// Create workbook and update sheet (for-comprehension handles errors!)
val result = for
  workbook <- Workbook.empty
  updatedSheet <- Sheet("Sheet1").flatMap(_.put(
    ref"A1" -> "Product",
    ref"B1" -> "Price",
    ref"A2" -> "Widget",
    ref"B2" -> money"$$19.99",    // Formatted literal preserves Currency!
    ref"C2" -> 42
  ))
  final <- workbook.put(updatedSheet)
yield final

// Or build sheet first, then add to workbook
val salesSheet = Sheet("Sales").flatMap(_.put(
  ref"A1" -> "Revenue",
  ref"B1" -> money"$$10,000"
))

val wb = for
  workbook <- Workbook.empty
  sheet <- salesSheet
  final <- workbook.put(sheet)      // Add-or-replace by sheet name
yield final

// Batch add multiple sheets
workbook.put(Sheet("Sales"), Sheet("Marketing"), Sheet("Finance"))

// Patch DSL for conditional updates (pure Either handling)
val patch = (ref"A1" := "Title") ++ range"A1:C1".merge
val updatedSheet = sheet.put(patch) match
  case Right(s) => s
  case Left(err) => handleError(err)

// For demos/REPLs only - requires explicit opt-in:
// import com.tjclp.xl.unsafe.*
// sheet.put(patch).unsafe
```

### Styling Example

```scala
import com.tjclp.xl.*

// CellStyle DSL with fluent method chaining
val headerStyle = CellStyle.default
  .bold.size(14.0).fontFamily("Arial")
  .bgGray.bordered
  .center.middle

// Custom colors with RGB
val brandStyle = CellStyle.default
  .bold.rgb(68, 114, 196)        // Custom text color
  .bgRgb(240, 240, 240)          // Custom background

// Or use hex codes (e.g., brand guidelines)
val tjcStyle = CellStyle.default
  .white.hex("#003366")          // TJC blue text
  .bgHex("#F5F5F5")              // Light gray background
  .bold.center

// Prebuilt constants for common styles
val header = Style.header          // Bold, 14pt, blue bg, white text, centered
val currency = Style.currencyCell  // Currency format, right-aligned

// Apply styles with unified put (returns XLResult[Sheet])
val result = sheet.put(
  (ref"A1" := "Revenue Report") ++
  ref"A1".styled(headerStyle) ++
  range"A1:C1".merge
)  // Right(updatedSheet) or Left(error)
```

### Type-Safe Cell Operations with Codecs

XL provides cell-level codecs for type-safe reading and writing with auto-inferred formatting, ideal for unstructured Excel sheets like financial models:

```scala
import com.tjclp.xl.codec.syntax.*

// Batch updates with mixed types (auto-infers formats) - pure Either handling
val result = for
  initialSheet <- Sheet("Q1 Forecast")
  withData <- initialSheet.put(
    ref"A1" -> "Revenue",                        // String: no format
    ref"B1" -> LocalDate.of(2025, 11, 10),      // Auto: date format
    ref"C1" -> BigDecimal("1000000.50"),        // Auto: decimal format
    ref"D1" -> 42                                // Auto: general number format
  )
  final <- withData.put(ref"E1", CellValue.from(fx"C1*D1"))
yield final

// Type-safe reading (result is XLResult[Sheet], so flatMap for chaining)
result.flatMap { sheet =>
  sheet.readTyped[LocalDate](ref"B1").left.map(_.toXLError(ref"B1")) match
    case Right(Some(date)) => Right(s"Date: $date")
    case Right(None) => Right("Cell empty")
    case Left(error) => Left(error)
}

// Or simpler - direct pattern match on result
result match
  case Right(sheet) =>
    val revenue = sheet.readTyped[BigDecimal](ref"C1")  // Either[CodecError, Option[BigDecimal]]
    val count = sheet.readTyped[Int](ref"D1")           // Either[CodecError, Option[Int]]
    val name = sheet.readTyped[String](ref"A1")         // Either[CodecError, Option[String]]
    // ... use the values
  case Left(err) => handleError(err)
```

**Supported types**: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime

**Why codecs?** Financial models are cell-oriented and irregular (not tabular). Codecs provide type safety and automatic format inference without requiring schema definitions, making them perfect for one-off imports/exports and manual cell manipulation.

### Rich Text Formatting (Intra-Cell Styling)

Apply multiple formats within a single cell using the composable DSL:

```scala
import com.tjclp.xl.richtext.RichText.*

// String extension DSL
val text = "Bold".bold.red + " normal " + "Italic".italic.blue

sheet.put(ref"A1", CellValue.RichText(text))

// Or batch put with multiple rich text values
val result = sheet.put(
  ref"A1" -> ("Error: ".red.bold + "File not found"),
  ref"A2" -> ("Revenue: ".bold + "+12.5%".green),
  ref"A3" -> ("Q1 ".size(18.0).bold + "Report".size(18.0).italic)
)  // Returns XLResult[Sheet]
```

**Supported formatting**: bold, italic, underline, colors (red/green/blue/black/white/custom), font size, font family

**Why rich text?** Excel supports multiple formatting runs within a cell (OOXML `<is><r>` structure). This is essential for financial reports where values need color-coding (green for positive, red for negative) alongside descriptive text.

### HTML Export

Export sheet ranges to HTML tables for web display:

```scala
// Export with inline CSS styling
val html = sheet.toHtml(ref"A1:B10")

// Export without styles (plain table)
val plainHtml = sheet.toHtml(ref"A1:B10", includeStyles = false)
```

Rich text formatting and cell styles are preserved as HTML tags (`<b>`, `<i>`, `<span style="">`) and inline CSS.

### Performance & I/O Mode Selection

XL provides two I/O modes optimized for different use cases. **Choose the right mode** based on your dataset size and styling needs:

| Dataset Size | Styling Needs | Recommended Mode | Memory Usage | Features |
|--------------|---------------|------------------|--------------|----------|
| **< 10k rows** | Any | In-memory write/read | ~10MB | Full SST, styles, formatting |
| **10k-100k rows** | Full styling | In-memory write/read | ~50-100MB | Full SST, styles, formatting |
| **100k+ rows** | Minimal | Streaming write/read | ~10MB (constant) | No SST, default styles only |
| **100k+ rows** | Full styling | Not yet supported | N/A | See roadmap (P7.5) |

**Current Limitations**:
- ⚠️ **Streaming write** doesn't support SST or rich styles. Use for data-heavy, minimal-styling workloads only.

**Recommendation**: Use **in-memory API** for most use cases. Use streaming for very large files (read or write) with minimal styling.

See [docs/reference/performance-guide.md](docs/reference/performance-guide.md) for detailed guidance.

### Streaming API (For Large Files)

XL provides **streaming write** with constant memory for files with 100k+ rows using fs2 and fs2-data-xml:

#### Write Large Files

```scala
import com.tjclp.xl.io.{Excel, RowData}
import com.tjclp.xl.CellValue
import cats.effect.IO
import fs2.Stream

val excel = Excel.forIO

// Generate 1 million rows with constant memory
Stream.range(1, 1_000_001)
  .map(i => RowData(i, Map(
    0 -> CellValue.Text(s"Row $i"),
    1 -> CellValue.Number(BigDecimal(i * 100)),
    2 -> CellValue.Bool(i % 2 == 0)
  )))
  .through(excel.writeStreamTrue(path, "BigData"))
  .compile.drain
  .unsafeRunSync()

// Memory usage: ~10MB constant (not 10GB!)
// Throughput: ~88k rows/second
// Note: Inline strings only (no SST), default styles only
```

#### Read Large Files

```scala
// Process 100k rows with constant memory
excel.readStream(path)
  .filter(_.rowIndex > 1)  // Skip header
  .map { row =>
    // Transform row data
    row.cells.get(0) match
      case Some(CellValue.Number(n)) => n * 2
      case _ => BigDecimal(0)
  }
  .compile.toList
  .unsafeRunSync()
```

#### Performance Characteristics

**Streaming Write** (✅ Working):
- **Memory**: O(1) constant (~10MB for any file size) ✅
- **Write Throughput**: ~88,000 rows/second ✅
- **Scalability**: Handles 1M+ rows (vs POI OOM at ~500k) ✅
- **Limitations**: No SST (inline strings only), default styles only ⚠️

**Streaming Read** (✅ Constant-Memory):
- **Memory**: O(1) constant (~10MB for any file size) ✅
- **Read Throughput**: ~55,000 rows/second ✅
- **Implementation**: Uses `fs2.io.readInputStream` with 4KB chunks
- **Scalability**: Handles 500k+ rows without OOM ✅

**Comparison to Apache POI**:
- **Write**: 4.5x faster, 80x less memory (10MB vs 800MB) ✅
- **Read**: 4.4x faster, 100x less memory (10MB vs 1GB) ✅

**API Methods**:
- `writeStreamTrue`: True constant-memory streaming (recommended for large writes)
- `writeStream`: Hybrid approach (materializes rows, then writes)
- `readStream`: True constant-memory streaming ✅

**When to Use**:
- ✅ Large data generation (1M+ rows, minimal styling)
- ✅ ETL pipelines (database → Excel with simple formatting)
- ✅ Large file reading (500k+ rows with constant memory) ✅
- ❌ Rich styling at scale (use in-memory for < 100k rows, wait for P7.5 for larger)

### Pure Error Handling with ExcelR

For pure functional programming without exceptions in the effect type, use `ExcelR[F]`:

```scala
import com.tjclp.xl.io.{ExcelIO, ExcelR}
import com.tjclp.xl.error.{XLResult, XLError}
import cats.effect.IO

val excel: ExcelR[IO] = ExcelIO.instance

// Explicit error channel - no exceptions
excel.readR(path).flatMap {
  case Right(workbook) =>
    // Successfully loaded
    processWorkbook(workbook)
  case Left(error: XLError) =>
    // Handle error explicitly
    IO.println(s"Failed to read: ${error.message}")
}

// Streaming with explicit errors
excel.readStreamR(path)
  .evalMap {
    case Right(rowData) => processRow(rowData)
    case Left(error) => handleError(error)
  }
  .compile.drain
```

**When to use**:
- `Excel[F]`: Simpler API, errors raised as F[_] failures (default choice)
- `ExcelR[F]`: Pure FP, explicit error types, no exceptions in F[_]

Both traits coexist - choose based on your error handling style.

### Configurable Writer

Control compression, SST usage, and XML formatting for advanced use cases:

```scala
import com.tjclp.xl.ooxml.{XlsxWriter, WriterConfig, SstPolicy, Compression}

// Default (compressed, compact XML)
XlsxWriter.write(workbook, path)  // Uses WriterConfig.default

// Debug mode (uncompressed, readable XML)
XlsxWriter.writeWith(
  workbook,
  path,
  WriterConfig.debug  // prettyPrint=true, compression=Stored
)

// Custom compression settings
XlsxWriter.writeWith(
  workbook,
  path,
  WriterConfig(
    compression = Compression.Stored,  // No compression (faster, 5-10x larger)
    prettyPrint = true,                // Readable XML (for debugging)
    sstPolicy = SstPolicy.Always       // Force shared strings table
  )
)
```

**Options**:
- `compression`: `Compression.Deflated` (default, 5-10x smaller files), `Compression.Stored` (no compression, faster)
- `sstPolicy`: `SstPolicy.Auto` (default heuristics), `Always` (force SST), `Never` (inline only)
- `prettyPrint`: `false` (default, compact XML), `true` (readable XML for debugging)

### Range Combinators & Utilities

Fill ranges declaratively with functional combinators:

```scala
// Fill using Excel coordinates
sheet.fillBy(ref"A1:C10") { (col, row) =>
  CellValue.Number(BigDecimal(col.index0 + row.index0))
}

// Fill using 0-based indices
sheet.tabulate(ref"A1:C10") { (colIdx, rowIdx) =>
  CellValue.Number(BigDecimal(colIdx * rowIdx))
}

// Populate row and column (returns XLResult for strict bounds checking)
val withRow = sheet.putRow(Row.from1(1), Column.from1(1), Vector(
  CellValue.Text("Q1"), CellValue.Text("Q2"), CellValue.Text("Q3")
)).fold(err => throw new RuntimeException(err.message), identity)

val withColumn = withRow.putCol(Column.from1(1), Row.from1(2), Vector(
  CellValue.Text("North"), CellValue.Text("South")
)).fold(err => throw new RuntimeException(err.message), identity)

// Strict range population (value count must match range size)
val filled = sheet.putRange(ref"A2:B3", Vector(
  CellValue.Text("North"),
  CellValue.Number(BigDecimal(42)),
  CellValue.Text("South"),
  CellValue.Number(BigDecimal(55))
)).fold(err => throw new RuntimeException(err.message), identity)

// Deterministic iteration (sorted)
sheet.cellsSorted.foreach(c => println(s"${c.ref.toA1}: ${c.value}"))
```

### Optics & Focus DSL

Composable, law-governed updates with Lens and Optional:

```scala
import Optics.*

// Modify cell value functionally
sheet.modifyValue(ref"A1") {
  case CellValue.Text(s) => CellValue.Text(s.toUpperCase)
  case other => other
}

// Conditional transformations
sheet.modifyCell(ref"B5") { cell =>
  cell.value match
    case CellValue.Number(n) if n < 0 => cell.withStyle(negativeStyleId)
    case _ => cell
}

// Focus on specific cells
val updated = sheet.focus(ref"C1").modify(_.withValue(CellValue.Number(42)))(sheet)
```

**Benefits**: Composable, total functions, zero allocation for simple paths.

## Development

### Build Commands

```bash
# Compile all modules
./mill __.compile

# Compile specific module
./mill xl-core.compile

# Run all tests
./mill __.test

# Run specific module tests
./mill xl-core.test

# Format code (Scalafmt)
./mill __.reformat

# Check formatting (CI mode)
./mill __.checkFormat

# Clean build artifacts
./mill clean
```

### Running Tests

```bash
# All tests (263 tests: 221 core + 24 OOXML + 18 streaming)
./mill __.test

# Individual test suites
./mill xl-core.test.testOnly com.tjclp.xl.AddressingSpec
./mill xl-core.test.testOnly com.tjclp.xl.PatchSpec
./mill xl-core.test.testOnly com.tjclp.xl.StyleSpec
./mill xl-core.test.testOnly com.tjclp.xl.codec.CellCodecSpec
./mill xl-core.test.testOnly com.tjclp.xl.codec.BatchUpdateSpec
./mill xl-ooxml.test.testOnly com.tjclp.xl.ooxml.OoxmlRoundTripSpec
./mill xl-cats-effect.test.testOnly com.tjclp.xl.io.ExcelIOSpec
```

### Code Formatting

This project uses **Scalafmt 3.10.1**. All code must be formatted before commit:

```bash
# Format everything
./mill __.reformat

# Check if formatting is needed
./mill __.checkFormat
```

See [docs/plan/29-linting.md](docs/plan/29-linting.md) for details.

## Project Structure

```
xl/
├── xl-core/          # Pure domain model (Cell, Sheet, Workbook, Patch, Style)
│   ├── src/          # Core types, addressing, patches, styles
│   └── test/         # Property-based tests with ScalaCheck
├── xl-ooxml/         # Pure XML serialization (no IO)
├── xl-cats-effect/   # IO interpreters for streaming (future)
├── xl-evaluator/     # Formula evaluator (future)
└── xl-testkit/       # Test laws and generators (future)
```

### Core Modules

**xl-core**: Pure, lawful domain model
- `addressing.scala` → Column, Row, ARef (packed 64-bit), CellRange
- `cell.scala` → CellValue enum, Cell case class
- `sheet.scala` → Sheet, Workbook with immutable operations
- `patch.scala` → Patch Monoid for composing updates
- `style.scala` → CellStyle, Font, Fill, Border, Color, NumFmt
- `codec/` → CellCodec instances, batch update extensions
- `error.scala` → XLError ADT for total error handling

**xl-ooxml**: OOXML (Office Open XML) serialization
- `xml.scala` → XmlWritable/XmlReadable traits, utilities
- `ContentTypes.scala` → [Content_Types].xml
- `Relationships.scala` → .rels files
- `Workbook.scala` → xl/workbook.xml
- `Worksheet.scala` → xl/worksheets/sheet#.xml

## Design Principles

### 1. Purity & Totality

No `null`, no exceptions, no partial functions. All errors as values:

```scala
// ✅ Total function
def parseCell(s: String): Either[XLError, ARef] =
  ARef.parse(s).left.map(err => XLError.InvalidCellRef(s, err))

// ❌ Never do this
def parseCell(s: String): ARef =
  ARef.parse(s).getOrElse(throw new Exception("Invalid!"))
```

### 2. Zero-Overhead Abstractions

Opaque types compile to primitives with no runtime overhead:

```scala
opaque type Column = Int      // Zero-cost wrapping
opaque type ARef = Long       // 64-bit packed: (row << 32) | col
```

### 3. Law-Governed Semantics

All core types have property-based tests proving algebraic laws:

- **Monoid laws**: Patch and StylePatch composition (associativity, identity)
- **Round-trip laws**: Parsing and printing are inverses
- **Idempotence**: Repeated identical operations yield same result

### 4. Deterministic Output

Same workbook always produces **byte-identical** XML output:

- Sorted attributes and elements
- Structural hashing for style deduplication
- Canonical formatting for stable git diffs

## Key Concepts

### Addressing

```scala
import com.tjclp.xl.*

// Compile-time validated (errors at compile time!)
val cellRef = ref"A1"           // Type: ARef
val rangeRef = ref"A1:B10"      // Type: CellRange

// Runtime parsing (returns Either)
val parsed = ARef.parse("A1")      // Either[String, ARef]
val range = CellRange.parse("A1:B10")  // Either[String, CellRange]

// Construction from indices
val col = Column.from0(0)    // Column A (0-based)
val row = Row.from1(1)       // Row 1 (1-based)
val ref2 = ARef(col, row)    // Cell A1
```

### Patches (Monoid Composition)

```scala
import com.tjclp.xl.*
import cats.syntax.all.*

// Build declarative updates
val updates =
  (Patch.Put(ref"A1", CellValue.Text("Title")): Patch) |+|
  (Patch.Put(ref"B1", CellValue.Number(100)): Patch) |+|
  (Patch.Merge(ref"A1:B1"): Patch)

// Apply to sheet
val result: Either[XLError, Sheet] = sheet.put(updates)
```

**Note**: Type ascription `(patch: Patch)` is required for `|+|` operator due to Scala enum limitations.

### Ergonomic Patch DSL (Simplified Syntax)

For cleaner syntax without type ascription, use the DSL operators:

```scala
import com.tjclp.xl.*

// Build patches with := operator and ++ composition
val patch =
  (ref"A1" := "Title") ++
  (ref"B1" := 100) ++
  ref"A1:B1".merge

// Apply to sheet
val result: Either[XLError, Sheet] = sheet.put(patch)
```

**Operators**:
- `:=` - Assign values (auto-converts primitives to CellValue)
- `++` - Compose patches (alternative to `|+|`, no type ascription needed)
- `.merge` / `.unmerge` / `.remove` on CellRange (returns Patch)
- `.styled(style)` / `.styleId(id)` / `.clearStyle` on ARef

**Benefits**: Cleaner syntax, no type ascription, same semantics as Monoid composition.

### Styles

```scala
import com.tjclp.xl.*

// Create a styled cell
val style = CellStyle.default
  .withFont(Font("Arial", 14.0, bold = true, color = Some(Color.fromRgb(255, 0, 0))))
  .withFill(Fill.Solid(Color.fromRgb(240, 240, 240)))
  .withBorder(Border.all(BorderStyle.Thin))
  .withNumFmt(NumFmt.Currency)
  .withAlign(Align(HAlign.Center, VAlign.Middle, wrapText = true))

// Unit conversions
val fontSizePx = Pt(12.0).toPx   // Points to pixels
val emu = Px(96.0).toEmu          // Pixels to EMU (OOXML unit)
```

## Testing

### Running Tests

```bash
# All tests (263 total)
./mill __.test

# Specific test suite
./mill xl-core.test.testOnly com.tjclp.xl.AddressingSpec
./mill xl-core.test.testOnly com.tjclp.xl.PatchSpec
./mill xl-core.test.testOnly com.tjclp.xl.StyleSpec
```

### Test Coverage

- **Addressing**: 17 property tests for Column, Row, ARef, CellRange
- **Patches**: 21 tests for Monoid laws and application semantics
- **Styles**: 60 tests for style system, unit conversions, canonicalization
- **Codecs**: 58 tests for encoding/decoding, identity laws, batch updates

All core algebras are tested with:
- Property-based testing (ScalaCheck)
- Monoid laws (associativity, identity)
- Idempotence properties
- Round-trip invariants
- Identity laws for codecs

## Documentation

Comprehensive design documentation in `docs/plan/`:

- **[00-executive-summary.md](docs/plan/00-executive-summary.md)** → Vision and non-negotiables
- **[01-purity-charter.md](docs/plan/01-purity-charter.md)** → Effect isolation and totality
- **[02-domain-model.md](docs/plan/02-domain-model.md)** → Full type algebra
- **[18-roadmap.md](docs/plan/18-roadmap.md)** → Implementation phases (P0-P11)
- **[25-style-guide.md](docs/plan/25-style-guide.md)** → Coding conventions

29 planning documents covering domain model, OOXML mapping, performance, security, testing, and examples.

## Contributing

### Development Workflow

1. **Format code**: `./mill __.reformat`
2. **Check formatting**: `./mill __.checkFormat`
3. **Run tests**: `./mill __.test`
4. **Commit**: Follow conventional commit style

### Code Style

- Use **opaque types** for domain quantities
- Use **enums** with `derives CanEqual` for closed sets
- Return `Either[XLError, A]` for fallible operations (no exceptions)
- Write **property-based tests** for all core algebras
- Ensure **deterministic output** (sorted attributes, canonical ordering)

See [docs/plan/25-style-guide.md](docs/plan/25-style-guide.md) for details.

### CI Requirements

All PRs must pass:
1. Format check (`./mill __.checkFormat`)
2. Compilation (`./mill __.compile`)
3. All tests (`./mill __.test`)

## Architecture Highlights

### Opaque Types (Zero-Overhead)

```scala
opaque type Column = Int       // No wrapper allocation
opaque type Row = Int          // Compiles to primitive
opaque type ARef = Long        // Packed: (row << 32) | col
```

### Patch Monoid (Declarative Updates)

Patches form a Monoid, enabling composition:

```scala
val p1 = Patch.Put(ref, value)
val p2 = Patch.SetStyle(ref, styleId)
val combined = (p1: Patch) |+| (p2: Patch)  // Batch(Vector(p1, p2))
```

Laws enforced via property tests:
- Associativity: `(p1 |+| p2) |+| p3 == p1 |+| (p2 |+| p3)`
- Identity: `Patch.empty |+| p == p`
- Idempotence: `apply(apply(s, p), p) == apply(s, p)`

### ARef Bit-Packing

Efficient 64-bit cell reference representation:

```
┌─────────────────────────┬─────────────────────────┐
│  Row (32 bits)          │  Column (32 bits)       │
└─────────────────────────┴─────────────────────────┘
```

Supports Excel max: 16,384 columns × 1,048,576 rows.

### Deterministic Canonicalization

Styles deduplicate via structural hashing:

```scala
val key1 = CellStyle.canonicalKey(style1)
val key2 = CellStyle.canonicalKey(style2)
// Equal keys → same index in styles.xml
```

XML output is byte-identical on re-serialization (sorted attributes/elements).

## License

MIT License - See [LICENSE](LICENSE) for details.

## Credits

Built with Scala 3.7, Mill, Cats, MUnit, and ScalaCheck.

## Links

- **Documentation**: [docs/plan/](docs/plan/)
- **Roadmap**: [docs/plan/18-roadmap.md](docs/plan/18-roadmap.md)
- **Examples**: [docs/plan/19-examples.md](docs/plan/19-examples.md)
- **Issue Tracker**: GitHub Issues
