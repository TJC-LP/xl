# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**XL** is a purely functional, mathematically rigorous Excel (OOXML) library for Scala 3.7. The design prioritizes **purity, totality, determinism, and law-governed semantics** with zero-overhead opaque types and compile-time DSLs.

> **Guiding Principle**: You are working on **the best Excel library in the world**. Before making any decision, ask yourself: **"What would the best Excel library in the world do?"**

**Package**: `com.tjclp.xl` | **Build**: Mill 0.12.x | **Scala**: 3.7.3

## Core Philosophy (Non-Negotiables)

1. **Purity & Totality**: No `null`, no partial functions, no thrown exceptions, no hidden effects
2. **Strong Typing**: Opaque types for Column, Row, ARef, SheetName; enums with `derives CanEqual`
3. **Deterministic Output**: Canonical ordering for XML/styles; byte-identical output on re-serialization
4. **Law-Governed**: Monoid laws for Patch/StylePatch; round-trip laws for parsers/printers
5. **Effect Isolation**: Core (`xl-core`, `xl-ooxml`) is 100% pure; only `xl-cats-effect` has IO

## Performance Summary (JMH Validated)

**XL vs Apache POI** (Apple Silicon, JDK 25):
- **Streaming reads**: XL 35% faster at 1k rows, competitive at 10k rows (SAX parser, O(1) memory)
- **In-memory reads**: XL 26% faster at 1k rows, competitive at 10k rows
- **Writes**: POI 49% faster (optimization planned for Phase 3)

*Use `ExcelIO.readStream()` for production (<5k rows fastest, constant memory), `ExcelIO.read()` for random access.*

## Module Architecture

```
xl-core/         → Pure domain model (Cell, Sheet, Workbook, Patch, Style), macros, DSL
xl-ooxml/        → Pure OOXML mapping (XlsxReader, XlsxWriter, SharedStrings, Styles)
xl-cats-effect/  → IO interpreters and streaming (Excel[F], ExcelIO, SAX-based streaming)
xl-benchmarks/   → JMH performance benchmarks
xl-evaluator/    → Formula parser/evaluator (TExpr GADT, 21 functions, dependency graphs)
xl-testkit/      → Test laws, generators, helpers [future]
```

## Import Patterns

```scala
import com.tjclp.xl.{*, given}     // Everything: core + formula + display + type class instances
import com.tjclp.xl.unsafe.*       // .unsafe boundary (explicit opt-in)

// Core API
val sheet = Sheet("Demo").put(ref"A1", 100).unsafe
Excel.write(Workbook(sheet), "output.xlsx")

// Formula evaluation (when xl-evaluator is a dependency)
sheet.evaluateFormula("=A1*2")     // SheetEvaluator extension method
FormulaParser.parse("=SUM(A1:A10)") // Parser at package level

// Display formatting
given Sheet = sheet
println(excel"Value: ${ref"A1"}")  // excel interpolator
sheet.displayCell(ref"A1")         // explicit display method
```

Note: `{*, given}` is required because Scala 3's `*` doesn't include given instances by default.

For production code with Cats Effect, use `ExcelIO` instead of `Excel`:
```scala
import com.tjclp.xl.io.ExcelIO
val excel = ExcelIO.instance[IO]
excel.read(path).flatMap(wb => excel.write(wb, outPath))
```

## Key Types

**Addressing** (`xl-core/src/com/tjclp/xl/addressing/`):
- `Column`, `Row` → opaque Int (zero-overhead)
- `ARef` → opaque Long (64-bit packed: `(row << 32) | col`)
- `CellRange(start, end)` → normalized, inclusive range

**Domain** (`xl-core/src/com/tjclp/xl/`):
- `Cell(ref, value, styleId, ...)` | `Sheet(name, cells, ...)` | `Workbook(sheets, ...)`
- `Patch` enum → Monoid for Sheet updates (Put, SetStyle, Merge, Remove, Batch)
- `StylePatch` enum → Monoid for CellStyle updates
- `CellStyle(font, fill, border, numFmt, align)` → complete cell formatting

**Codecs** (`xl-core/src/com/tjclp/xl/codec/`):
- `CellCodec[A]` → Bidirectional for 9 types: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText
- Auto-inferred formats: LocalDate→Date, LocalDateTime→DateTime, BigDecimal→Decimal

**Formula** (`xl-evaluator/src/com/tjclp/xl/formula/`):
- `TExpr[A]` GADT → typed formula AST
- `FormulaParser.parse()` / `FormulaPrinter.print()` → round-trip verified
- `DependencyGraph` → cycle detection via Tarjan's SCC, topological sort via Kahn's

## Build Commands

```bash
./mill __.compile          # Compile all
./mill __.test             # Run all tests (731+)
./mill xl-core.test        # Test specific module
./mill __.reformat         # Format (Scalafmt 3.10.1)
./mill __.checkFormat      # CI check
./mill clean               # Clean artifacts
```

## Code Style

All code must pass `./mill __.checkFormat`. See `docs/design/style-guide.md`.

**Rules**: opaque types for domain quantities | enums with `derives CanEqual` | `final case class` for data | total functions returning Either/Option | extension methods over implicit classes

**WartRemover** (compile-time enforcement):
- ❌ Error: `null`, `.head/.tail`, `.get` on Try/Either
- ⚠️ Warning: `.get` on Option, `var`/`while`/`return` (acceptable in tests/macros)

**Error Handling**: Always use `XLResult[A] = Either[XLError, A]`

## Architecture Patterns

### 1. Opaque Types
```scala
opaque type Column = Int   // Cannot mix Int as Column
opaque type ARef = Long    // Packed: (row << 32) | col
```

### 2. Monoid Composition
```scala
import com.tjclp.xl.dsl.*
val patch = (ref"A1" := "Hello") ++ ref"A1".styled(boldStyle) ++ ref"A1:B2".merge
```
Note: Using Cats `|+|` requires type ascription on enum cases.

### 3. Compile-Time Macros
```scala
val cellRef: ARef = ref"A1"           // Validated at compile time
val formula = fx"=SUM(A1:B10)"        // CellValue.Formula
val price = money"$$1,234.56"         // Formatted(value, NumFmt.Currency)
```

### 4. Deterministic XML
- Attributes sorted by name (`XmlUtil.elem`)
- Elements sorted by natural key
- `XmlUtil.compact` for production, `XmlUtil.prettyPrint` for debug

### 5. Law-Based Testing
```scala
property("parse . print = id") { forAll { (ref: ARef) => ARef.parse(ref.toA1) == Right(ref) }}
```

## Essential APIs

### Sheet Operations
```scala
// Batch put with type inference
sheet.put(ref"A1" -> "Revenue", ref"B1" -> LocalDate.now, ref"C1" -> BigDecimal("1000.50"))

// Type-safe reading
sheet.readTyped[BigDecimal](ref"C1") // Either[CodecError, Option[A]]

// Styling
sheet.style(ref"B19:D21", CellStyle.default.withNumFmt(NumFmt.Percent))
```

### Formula Evaluation
```scala
// SheetEvaluator extension methods available from com.tjclp.xl.{*, given}
sheet.evaluateFormula("=SUM(A1:A10)")      // XLResult[CellValue]
sheet.evaluateWithDependencyCheck()         // Safe eval with cycle detection
```

**21 Functions**: SUM, COUNT, AVERAGE, MIN, MAX, IF, AND, OR, NOT, CONCATENATE, LEFT, RIGHT, LEN, UPPER, LOWER, TODAY, NOW, DATE, YEAR, MONTH, DAY

### Rich Text
```scala
val text = "Bold".bold.red + " normal " + "Italic".italic.blue
sheet.put(ref"A1" -> text)
```

### Comments & HTML Export
```scala
sheet.comment(ref"A1", Comment.plainText("Note", Some("Author")))
sheet.toHtml(ref"A1:B10")  // HTML table with inline CSS
```

## Important Constraints

### ARef Packing
```scala
// Pack: (row << 32) | (col & 0xFFFFFFFF)
// Valid: Column 0-16383 (A-XFD), Row 0-1048575 (1-1048576)
```

### Style Deduplication
Styles deduplicated by `CellStyle.canonicalKey`. Build style index before emitting cells.

### CellRange Normalization
`CellRange(start, end)` auto-normalizes (start ≤ end).

## Testing

**Framework**: MUnit + ScalaCheck | **Generators**: `xl-core/test/src/com/tjclp/xl/Generators.scala`

**731+ tests**: addressing (17), patch (21), style (60), datetime (8), codec (42), batch (16), syntax (18), optics (34), OOXML (24), streaming (18), RichText (5), formula (51+)

## Documentation

- **Roadmap**: `docs/plan/roadmap.md` (single source of truth for work scheduling)
- **Status**: `docs/STATUS.md` (current capabilities, 731+ tests)
- **Design**: `docs/design/*.md` (architecture, purity charter, domain model)
- **Reference**: `docs/reference/*.md` (examples, scaffolds, performance guide)

## AI Agent Workflow

1. Check `docs/plan/roadmap.md` TL;DR for available work (🔵 blue nodes)
2. Run `gtr list` to verify no conflicting worktrees
3. Read corresponding plan doc for implementation details
4. After PR merge: update roadmap.md, STATUS.md, LIMITATIONS.md

**Module Conflict Matrix**:
- High risk: `xl-core/Sheet.scala` (serialize work)
- Medium risk: `xl-ooxml/Worksheet.scala` (coordinate)
- Low risk: `xl-evaluator/` (parallelize freely)

## Known Gotchas

**Monoid syntax needs type ascription**:
```scala
val p = (Patch.Put(ref, value): Patch) |+| (Patch.SetStyle(ref, 1): Patch)
```

**Extension methods need @targetName**:
```scala
extension (sheet: Sheet)
  @annotation.targetName("applyPatchExt")
  def applyPatch(patch: Patch): XLResult[Sheet] = ...
```

**Import order**: Java/javax → Scala stdlib → Cats → Project → Tests

## CI/CD

GitHub Actions: `./mill __.checkFormat` → `./mill __.compile` → `./mill __.test`

Coursier + Mill artifacts cached for 2-5x speedup.
