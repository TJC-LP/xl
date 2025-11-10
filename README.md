# XL — Pure Functional Excel Library for Scala 3

[![CI](https://github.com/TJC-LP/xl/workflows/CI/badge.svg)](https://github.com/TJC-LP/xl/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

A purely functional, mathematically rigorous Excel (OOXML) library for Scala 3.7 with **zero-overhead opaque types**, **law-governed semantics**, and **deterministic output**.

## Features

- **Type-Safe Addressing**: Zero-cost opaque types for columns, rows, and cell references with compile-time validated literals
- **Pure Domain Model**: 100% pure core with no side effects, nullable values, or partial functions
- **Immutable Workbooks**: Persistent data structures with efficient structural sharing
- **Monoid Composition**: Declarative updates via lawful Patch and StylePatch monoids
- **Comprehensive Styling**: Colors, fonts, fills, borders, number formats, and alignment
- **Deterministic Output**: Canonical XML ordering for stable git diffs and reproducible builds
- **Property-Based Testing**: ScalaCheck generators and law tests for all core algebras

## Status

**Current**: ~75% complete with 112/112 tests passing ✅

**Implemented** (P0-P5):
- ✅ Core domain types (Cell, Sheet, Workbook)
- ✅ Addressing system with compile-time macros (`cell"A1"`, `range"A1:B10"`)
- ✅ Patch system with Monoid composition for updates
- ✅ Complete style system (fonts, colors, fills, borders, alignment)
- ✅ **OOXML I/O**: Read/write XLSX files with SST and styles.xml
- ✅ **Streaming Write**: True constant-memory streaming with fs2-data-xml (100k+ rows)
- ✅ Elegant syntax: given conversions, batch put macro, formatted literals

**In Progress** (P5):
- 🚧 Streaming Read: Constant-memory reading (currently hybrid)

**Roadmap**: Complete streaming read, codecs with derivation, formulas, charts, drawings, tables.

See [docs/plan/18-roadmap.md](docs/plan/18-roadmap.md) for full implementation plan.

## Quick Start

### Prerequisites

- **JDK**: 17 or 21 (Temurin recommended)
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
import com.tjclp.xl.macros.{cell, range}

// Create a workbook
val wb = Workbook("MySheet").map { workbook =>
  val sheet = workbook.sheets(0)

  // Use compile-time validated cell references
  val ref = cell"A1"

  // Build updates with Monoid composition
  val updates =
    (Patch.Put(cell"A1", CellValue.Text("Hello")): Patch) |+|
    (Patch.Put(cell"B1", CellValue.Number(42)): Patch) |+|
    (Patch.SetStyle(cell"A1", 1): Patch)

  // Apply patches to get updated sheet
  sheet.applyPatch(updates).map { updated =>
    workbook.updateSheet(0, updated)
  }
}
```

### Styling Example

```scala
import com.tjclp.xl.*

// Define a style with the builder API
val headerStyle = CellStyle.default
  .withFont(Font("Arial", 14.0, bold = true))
  .withFill(Fill.Solid(Color.fromRgb(200, 200, 200)))
  .withBorder(Border.all(BorderStyle.Thin))
  .withAlign(Align(HAlign.Center, VAlign.Middle))

// Apply style patches
val styleUpdates =
  (StylePatch.SetFont(Font("Calibri", 12.0)): StylePatch) |+|
  (StylePatch.SetFill(Fill.Solid(Color.Rgb(0xFFFFFFFF))): StylePatch)

val newStyle = CellStyle.default.applyPatch(styleUpdates)
```

### Streaming API (For Large Files)

XL provides constant-memory streaming for files with 100k+ rows using fs2 and fs2-data-xml:

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

// Memory usage: ~50MB constant (not 10GB!)
// Throughput: ~88k rows/second
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

- **Memory**: O(1) constant (~50MB for any file size)
- **Write Throughput**: ~88,000 rows/second
- **Scalability**: Handles 1M+ rows (vs POI OOM at ~500k)
- **Use Case**: Files >10k rows, ETL pipelines, data generation

**API Methods**:
- `writeStreamTrue`: True constant-memory streaming with fs2-data-xml events
- `writeStream`: Hybrid approach (materializes rows, then writes)
- `readStream`: Currently hybrid (materializes workbook, then streams) - true streaming coming soon

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
# All tests (112 tests: 95 core + 9 OOXML + 8 streaming)
./mill __.test

# Individual test suites
./mill xl-core.test.testOnly com.tjclp.xl.AddressingSpec
./mill xl-core.test.testOnly com.tjclp.xl.PatchSpec
./mill xl-core.test.testOnly com.tjclp.xl.StyleSpec
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
├── xl-macros/        # Compile-time validated literals (cell"", range"")
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
- `error.scala` → XLError ADT for total error handling

**xl-macros**: Compile-time DSLs
- `macros.scala` → `cell""` and `range""` string interpolators with compile-time validation

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
import com.tjclp.xl.macros.{cell, range}

// Compile-time validated (errors at compile time!)
val ref = cell"A1"           // Type: ARef
val rng = range"A1:B10"      // Type: CellRange

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
  (Patch.Put(cell"A1", CellValue.Text("Title")): Patch) |+|
  (Patch.Put(cell"B1", CellValue.Number(100)): Patch) |+|
  (Patch.Merge(range"A1:B1"): Patch)

// Apply to sheet
val result: Either[XLError, Sheet] = sheet.applyPatch(updates)
```

**Note**: Type ascription `(patch: Patch)` is required for `|+|` operator due to Scala enum limitations.

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
# All tests (77 total)
./mill __.test

# Specific test suite
./mill xl-core.test.testOnly com.tjclp.xl.AddressingSpec
./mill xl-core.test.testOnly com.tjclp.xl.PatchSpec
./mill xl-core.test.testOnly com.tjclp.xl.StyleSpec
```

### Test Coverage

- **Addressing**: 17 property tests for Column, Row, ARef, CellRange
- **Patches**: 19 tests for Monoid laws and application semantics
- **Styles**: 41 tests for style system, unit conversions, canonicalization

All core algebras are tested with:
- Property-based testing (ScalaCheck)
- Monoid laws (associativity, identity)
- Idempotence properties
- Round-trip invariants

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
