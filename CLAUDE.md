# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**XL** is a purely functional, mathematically rigorous Excel (OOXML) library for Scala 3.7. The design prioritizes **purity, totality, determinism, and law-governed semantics** with zero-overhead opaque types and compile-time DSLs.

**Package**: `com.tjclp.xl`
**Build**: Mill 0.12.x
**Scala**: 3.7.3

## Core Philosophy (Non-Negotiables)

1. **Purity & Totality**: No `null`, no partial functions, no thrown exceptions, no hidden effects
2. **Strong Typing**: Opaque types for Column, Row, ARef, SheetName; enums with `derives CanEqual`
3. **Deterministic Output**: Canonical ordering for XML/styles; byte-identical output on re-serialization
4. **Law-Governed**: Monoid laws for Patch/StylePatch; round-trip laws for parsers/printers
5. **Effect Isolation**: Core (`xl-core`, `xl-ooxml`) is 100% pure; only `xl-cats-effect` has IO

## Module Architecture

```
xl-core/         → Pure domain model (Cell, Sheet, Workbook, Patch, Style)
xl-macros/       → Compile-time validated literals (cell"A1", range"A1:B10")
xl-ooxml/        → Pure XML serialization/deserialization (no IO)
xl-cats-effect/  → IO interpreters (ZIP, streaming, file system) [future]
xl-evaluator/    → Optional formula evaluator [future]
xl-testkit/      → Test laws, generators, golden files [future]
```

### Key Type Relationships

**Addressing** (`xl-core/src/com/tjclp/xl/addressing.scala`):
- `Column`, `Row` → opaque types wrapping Int (zero-overhead)
- `ARef` → opaque Long (64-bit packed: high 32 bits = row, low 32 bits = col)
- `CellRange(start: ARef, end: ARef)` → normalized, inclusive range

**Domain Model** (`xl-core/src/com/tjclp/xl/`):
- `Cell(ref: ARef, value: CellValue, styleId: Option[Int], ...)` → single cell
- `Sheet(name: SheetName, cells: Map[ARef, Cell], ...)` → immutable worksheet
- `Workbook(sheets: Vector[Sheet], ...)` → multi-sheet workbook

**Patches** (`xl-core/src/com/tjclp/xl/patch.scala`):
- `Patch` enum → Monoid for composing Sheet updates (Put, SetStyle, Merge, Remove, Batch)
- `StylePatch` enum → Monoid for composing CellStyle updates (SetFont, SetFill, etc.)
- Law-governed: associativity, identity, idempotence

**Styles** (`xl-core/src/com/tjclp/xl/style.scala`):
- `CellStyle(font, fill, border, numFmt, align)` → complete cell formatting
- `CellStyle.canonicalKey` → deterministic deduplication for styles.xml
- Units: `Pt`, `Px`, `Emu` opaque types with bidirectional conversions

**Codecs** (`xl-core/src/com/tjclp/xl/codec/`):
- `CellCodec[A]` → Bidirectional type-safe encoding/decoding for 8 primitive types
- `putMixed(updates)` → Batch updates with auto-inferred formatting
- `readTyped[A](ref)` → Type-safe reading with Either[CodecError, Option[A]]

**OOXML Layer** (`xl-ooxml/src/com/tjclp/xl/ooxml/`):
- Pure XML serialization with `XmlWritable`/`XmlReadable` traits
- `ContentTypes`, `Relationships`, `OoxmlWorkbook`, `OoxmlWorksheet` → OOXML parts
- Deterministic attribute/element ordering for stable diffs

## Build Commands

```bash
# Compile everything
./mill __.compile

# Run all tests (229 tests: 187 core + 24 OOXML + 18 streaming)
./mill __.test

# Test specific module
./mill xl-core.test

# Format code (Scalafmt 3.10.1)
./mill __.reformat

# Check formatting (CI mode)
./mill __.checkFormat

# Clean build artifacts
./mill clean
```

## Development Workflow

### Running Tests

```bash
# All tests
./mill __.test

# Single module
./mill xl-core.test
./mill xl-ooxml.test

# Single test class (run from module test directory)
./mill xl-core.test.testOnly com.tjclp.xl.AddressingSpec
```

### Code Style

All code must pass `./mill __.checkFormat` before commit. Format with `./mill __.reformat`.

**Style rules** (from `docs/plan/25-style-guide.md`):
- Prefer **opaque types** for domain quantities
- Use **enums** for closed sets; `derives CanEqual` everywhere
- Keep public functions **total**; return Either/Option for errors
- Prefer **extension methods** over implicit classes
- Macros must emit **clear diagnostics**

### Error Handling

Always use `XLResult[A] = Either[XLError, A]`:

```scala
// ✅ Good: Total function with explicit error
def parseCell(s: String): XLResult[ARef] =
  ARef.parse(s).left.map(err => XLError.InvalidCellRef(s, err))

// ❌ Bad: Partial function or null
def parseCell(s: String): ARef = ???  // throws!
```

## Architecture Patterns

### 1. Opaque Types for Zero-Overhead Safety

All domain types use opaque types to prevent mixing units:

```scala
opaque type Column = Int  // Cannot accidentally use Int as Column
opaque type ARef = Long   // Packed representation: (row << 32) | col
```

### 2. Monoid Composition for Updates

Both `Patch` and `StylePatch` are Monoids, enabling declarative composition:

```scala
val updates =
  (Patch.Put(cell"A1", CellValue.Text("Hello")): Patch) |+|
  (Patch.SetStyle(cell"A1", 1): Patch) |+|
  (Patch.Merge(range"A1:B2"): Patch)

val result = sheet.applyPatch(updates)  // Either[XLError, Sheet]
```

**Important**: Enum cases need type ascription for Monoid syntax (`|+|`).

### 3. Compile-Time Validated Literals

Macros in `xl-macros` provide zero-cost validated literals:

```scala
import com.tjclp.xl.macros.{cell, range}

val ref = cell"A1"        // Validated at compile time, emits packed Long
val rng = range"A1:B10"   // Parsed at compile time
```

**Implementation**: Zero-allocation parsers emit constants directly as `Expr[ARef]`.

### 4. Deterministic Canonicalization

For stable diffs and deduplication:

```scala
// Styles with same canonical key → same index in styles.xml
val key = CellStyle.canonicalKey(style)

// XML attributes/elements in sorted order
XmlUtil.elem("sheet", "name" -> "Sheet1", "sheetId" -> "1")(...)
```

### 5. Law-Based Testing

All core algebras have property-based law tests:

```scala
// Monoid laws
property("associativity") { forAll { (p1, p2, p3) =>
  ((p1 |+| p2) |+| p3) == (p1 |+| (p2 |+| p3))
}}

// Round-trip laws
property("parse . print = id") { forAll { (ref: ARef) =>
  ARef.parse(ref.toA1) == Right(ref)
}}
```

## Testing Infrastructure

**Framework**: MUnit + ScalaCheck (munit-scalacheck)
**Generators**: `xl-core/test/src/com/tjclp/xl/Generators.scala`

Key test files:
- `AddressingSpec.scala` → Column, Row, ARef, CellRange property tests
- `PatchSpec.scala` → Patch Monoid laws and semantics
- `StyleSpec.scala` → Style system, unit conversions, color parsing
- `StylePatchSpec.scala` → StylePatch Monoid laws

### Adding New Tests

1. Add generators to `Generators.scala` with `given Arbitrary[T]`
2. Extend `ScalaCheckSuite` for property tests
3. Test laws: associativity, identity, idempotence, round-trips
4. Add golden file tests for XML serialization (future)

## Implementation Status (Roadmap)

**✅ Complete (P0-P3, ~55%)**:
- P0: Bootstrap, Mill build, module structure
- P1: Addressing & Literals (opaque types, macros, 17 tests)
- P2: Core + Patches (Monoid composition, 19 tests)
- P3: Styles & Themes (full formatting system, 41 tests)
- P4: OOXML MVP (50% - XML foundation done, need SST/Styles/ZIP)

**🚧 In Progress (P4)**:
- Shared Strings Table (SST) with deduplication
- Styles.xml mapping
- ZIP reader/writer
- Round-trip tests

**📋 Not Started (P5-P11)**:
- P5: Streaming (fs2, constant memory)
- P6: Codecs & Named Tuples (derivation, header binding)
- P7-P11: Advanced features (drawings, charts, tables, safety)

## Important Constraints

### ARef Packing

`ARef` uses bit-packing for efficient storage:

```scala
// Pack: (row << 32) | (col & 0xFFFFFFFF)
def apply(col: Column, row: Row): ARef =
  (row.index0.toLong << 32) | (col.index0.toLong & 0xFFFFFFFFL)

// Unpack:
def col: Column = Column.from0((ref & 0xFFFFFFFFL).toInt)
def row: Row = Row.from0((ref >> 32).toInt)
```

Valid ranges: Column 0-16383 (A-XFD), Row 0-1048575 (1-1048576).

### CellRange Normalization

`CellRange(start, end)` is always normalized (start ≤ end). Use the constructor to auto-normalize:

```scala
CellRange(ref1, ref2)  // Automatically orders min/max
```

### Style Deduplication

Styles are deduplicated by `CellStyle.canonicalKey`:

```scala
val key = CellStyle.canonicalKey(style)  // Structural hash
// Equal keys → same style index in styles.xml
```

**Critical**: When implementing styles.xml writer, build style index first, then emit cells.

### XML Determinism

All XML output must be deterministic:

1. **Attribute order**: Always sort by name (`XmlUtil.elem` does this)
2. **Element order**: Sort by natural key (style index, cell ref, etc.)
3. **Whitespace**: Use `XmlUtil.prettyPrint` for consistent formatting

## Common Patterns

### Working with Either

```scala
// Chain operations
for
  ref <- ARef.parse("A1")
  sheet <- workbook("Sheet1")
  updated <- sheet.applyPatch(Patch.Put(ref, value))
yield updated

// Convert to XLError
ARef.parse(s).left.map(err => XLError.InvalidCellRef(s, err))
```

### Extension Methods

```scala
// Define in companion object
extension (sheet: Sheet)
  @annotation.targetName("applyPatchExt")  // Avoid erasure conflicts
  def applyPatch(patch: Patch): XLResult[Sheet] = ...
```

### Macro Patterns

Macros in `xl-macros/src/com/tjclp/xl/macros.scala`:

- Use `quotes.reflect.report.errorAndAbort` for compile errors
- Emit constants directly: `'{ (${Expr(value)}: T).asInstanceOf[OpaqueType] }`
- Zero-allocation parsers (no regex, manual char iteration)

### Codec Patterns

Cell-level codecs (`xl-core/src/com/tjclp/xl/codec/`) provide type-safe encoding/decoding with auto-inferred formatting.

#### When to Use Codecs

- ✅ Financial models (cell-oriented, irregular layouts)
- ✅ Unstructured sheets (no schema)
- ✅ Type-safe reading from user input
- ✅ Auto-inferred date/number formats

#### Batch Updates

Always prefer `putMixed` for multi-cell updates:

```scala
import com.tjclp.xl.codec.{*, given}

sheet.putMixed(
  cell"A1" -> "Revenue",
  cell"B1" -> LocalDate.of(2025, 11, 10),  // Auto: date format
  cell"C1" -> BigDecimal("1000000.50"),    // Auto: decimal format
  cell"D1" -> 42
)
```

**Benefits**: Cleaner syntax, auto-inferred styles, builds on existing `putAll` (no performance overhead).

#### Type-Safe Reading

```scala
// Reading returns Either[CodecError, Option[A]]
sheet.readTyped[BigDecimal](cell"C1") match
  case Right(Some(value)) => // Success: cell has value
  case Right(None) => // Success: cell is empty
  case Left(error) => // Error: type mismatch or parse error

// Convert to XLError if needed
val xlResult: Either[XLError, Option[BigDecimal]] =
  sheet.readTyped[BigDecimal](ref)
    .left.map(_.toXLError(ref))
```

#### Supported Types

8 primitive types with inline given instances (zero-overhead):
- `String`, `Int`, `Long`, `Double`, `BigDecimal`, `Boolean`, `LocalDate`, `LocalDateTime`

#### Auto-Inferred Formats

- `LocalDate` → `NumFmt.Date`
- `LocalDateTime` → `NumFmt.DateTime`
- `BigDecimal` → `NumFmt.Decimal`
- `Int`, `Long`, `Double` → `NumFmt.General`
- `String`, `Boolean` → No format (plain)

#### Identity Laws

All codecs satisfy: `codec.read(Cell(ref, codec.write(value)._1)) == Right(Some(value))` (up to formatting precision).

## Documentation

**Primary source**: `docs/plan/` directory (29 markdown files)

Key documents:
- `00-executive-summary.md` → Vision and non-negotiables
- `01-purity-charter.md` → Effect isolation, totality, laws
- `02-domain-model.md` → Full type algebra
- `10-patches-and-optics.md` → Patch Monoid design
- `18-roadmap.md` → Implementation phases (P0-P11)
- `25-style-guide.md` → Coding conventions
- `29-linting.md` → Scalafmt setup

## CI/CD

**GitHub Actions**: `.github/workflows/ci.yml`

Pipeline:
1. Check formatting (`./mill __.checkFormat`)
2. Compile all modules (`./mill __.compile`)
3. Run tests (`./mill __.test`)

**Caching**: Coursier + Mill artifacts cached for 2-5x speedup.

## Common Tasks

### Adding a New Domain Type

1. Define as opaque type or enum in `xl-core/src/com/tjclp/xl/`
2. Add `derives CanEqual` for enums
3. Provide smart constructors returning `Either[String, T]`
4. Add ScalaCheck generator in `Generators.scala`
5. Write property tests for laws (round-trips, invariants)

### Adding a New Patch Type

1. Add case to `Patch` enum in `patch.scala`
2. Implement in `applyPatch` match
3. Add tests in `PatchSpec.scala` (application, idempotence)
4. Verify Monoid laws still hold

### Adding OOXML Serialization

1. Create model in `xl-ooxml/src/com/tjclp/xl/ooxml/`
2. Extend `XmlWritable` with `toXml: Elem`
3. Extend `XmlReadable[A]` with `fromXml(elem: Elem): Either[String, A]`
4. Use `XmlUtil.elem` for deterministic attribute ordering
5. Add round-trip tests

## Known Issues & Gotchas

### Monoid Syntax Requires Type Ascription

Enum cases don't automatically widen to enum type:

```scala
// ❌ Fails: Patch.Put doesn't have |+| method
val p = Patch.Put(ref, value) |+| Patch.SetStyle(ref, 1)

// ✅ Works: Explicitly type as Patch
val p = (Patch.Put(ref, value): Patch) |+| (Patch.SetStyle(ref, 1): Patch)
```

### @targetName for Extension Methods

Prevent erasure conflicts when adding extensions to types that already have methods:

```scala
extension (sheet: Sheet)
  @annotation.targetName("applyPatchExt")  // Different from Patch.applyPatch
  def applyPatch(patch: Patch): XLResult[Sheet] = ...
```

### Import Organization

Currently manual (Scalafix deferred). Canonical order:

```scala
// 1. Java/javax
import java.time.LocalDateTime

// 2. Scala stdlib
import scala.xml.*

// 3. Cats
import cats.Monoid
import cats.syntax.all.*

// 4. Project internal
import com.tjclp.xl.*

// 5. Test frameworks (in tests)
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
```

### Unit Conversion Precision

Emu (Long) conversions have rounding:

```scala
property("Px to Emu round-trip") {
  forAll { (px: Px) =>
    val diff = math.abs(px.toEmu.toPx.value - px.value)
    assert(diff < 0.001)  // Use 0.001, not 0.0001 due to Long truncation
  }
}
```

## Reference Documentation

- **Spec directory**: `docs/plan/` contains 29 detailed design documents
- **Roadmap**: `docs/plan/18-roadmap.md` for implementation phases
- **Style guide**: `docs/plan/25-style-guide.md` for coding conventions
- **OOXML mapping**: `docs/plan/11-ooxml-mapping.md` for XML structure
- **Security**: `docs/plan/23-security.md` for ZIP/XXE/formula injection guards

## Test Coverage Goal

229/229 tests passing (as of P6 completion):
- 17 addressing tests (Column, Row, ARef, CellRange laws)
- 21 patch tests (Monoid laws, application semantics)
- 60 style tests (units, colors, builders, canonicalization, StylePatch, StyleRegistry)
- 8 datetime tests (Excel serial number conversions)
- 42 codec tests (CellCodec identity laws, type safety, auto-inference)
- 16 batch update tests (putMixed, readTyped, style deduplication)
- 18 elegant syntax tests (given conversions, batch put, formatted literals)
- 24 OOXML tests (round-trip, serialization, styles)
- 18 streaming tests (fs2-data-xml, constant-memory I/O)

Target: 90%+ coverage with property-based tests for all algebras.
