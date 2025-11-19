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
xl-core/         → Pure domain model (Cell, Sheet, Workbook, Patch, Style), macros, DSL
xl-ooxml/        → Pure OOXML mapping (XlsxReader, XlsxWriter, SharedStrings, Styles)
xl-cats-effect/  → IO interpreters and streaming (Excel[F], ExcelIO, fs2-based streaming)
xl-evaluator/    → Optional formula evaluator [future]
xl-testkit/      → Test laws, generators, helpers [future]
```

**Import Pattern (single-import ergonomics)**:
```scala
import com.tjclp.xl.*  // Domain model, macros, DSL, rich text
```

Macros (`ref`, `fx`, money/percent/date/accounting) are bundled in `xl-core` and surfaced through this unified import.

### Key Type Relationships

**Addressing** (`xl-core/src/com/tjclp/xl/addressing/`):
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
- `CellCodec[A]` → Bidirectional type-safe encoding/decoding for 9 primitive types (String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText)
- Batch `Sheet.put(ref -> value, ...)` → auto-inferred formatting (this is the former `putMixed` behavior, now folded into `Sheet.put`)
- `readTyped[A](ref)` → Type-safe reading with `Either[CodecError, Option[A]]`

**Rich Text** (`xl-core/src/com/tjclp/xl/richtext.scala`):
- `TextRun` → Single formatted text segment (OOXML `<r>` element)
- `RichText` → Multiple runs with different formatting within one cell
- String extensions: `.bold`, `.italic`, `.red`, `.green`, `.size()`, etc.

**HTML Export** (`xl-core/src/com/tjclp/xl/html/`):
- `sheet.toHtml(range)` → Convert cell range to HTML table with inline CSS

**OOXML Layer** (`xl-ooxml/src/com/tjclp/xl/ooxml/`):
- Pure XML serialization with `XmlWritable`/`XmlReadable` traits
- `ContentTypes`, `Relationships`, `OoxmlWorkbook`, `OoxmlWorksheet` → OOXML parts
- Deterministic attribute/element ordering for stable diffs

## Build Commands

```bash
# Compile everything
./mill __.compile

# Run all tests (see docs/STATUS.md for current counts)
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

**Style rules** (see also `docs/design/style-guide.md`):
- Prefer **opaque types** for domain quantities
- Use **enums** for closed sets; `derives CanEqual` everywhere
- Keep public functions **total**; return Either/Option for errors
- Prefer **extension methods** over implicit classes
- Macros must emit **clear diagnostics**

### WartRemover

XL uses [WartRemover](https://www.wartremover.org/) to enforce purity and totality at compile time.

**Policy**: See `docs/design/wartremover-policy.md` for complete wart list and rationale.

**Key enforcements**:
- ❌ No `null` (Tier 1 - error)
- ❌ No `.head/.tail` on collections (Tier 1 - error)
- ❌ No `.get` on Try/Either projections (Tier 1 - error)
- ⚠️ Warn on `.get` on Option (Tier 2 - warning, acceptable in tests)
- ⚠️ Warn on `var`/`while`/`return` (Tier 2 - warning, acceptable in macros)

**Suppressing false positives**:
```scala
// Inline comment for clarity (preferred)
parts(0)  // Safe: length == 1 verified above

// @SuppressWarnings for entire methods
@SuppressWarnings(Array("org.wartremover.warts.Var"))
def performanceOptimizedCode(): Unit = {
  var accumulator = 0  // Intentional for performance
  // ...
}
```

**Running**: WartRemover runs automatically during `./mill __.compile` and in pre-commit hooks.

### Error Handling

Always use `XLResult[A] = Either[XLError, A]`:

```scala
// ✅ Good: Total function with explicit errors
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
import com.tjclp.xl.*
import com.tjclp.xl.dsl.*

val patch: Patch =
  (ref"A1" := "Hello") ++
  ref"A1".styled(CellStyle.default.bold) ++
  ref"A1:B2".merge

val result: XLResult[Sheet] = sheet.put(patch)
```

**Important**: If you use Cats’ `|+|` instead of the DSL `++`, enum cases still need type ascription.

### 3. Compile-Time Validated Literals

Macros in `xl-core/src/com/tjclp/xl/macros/` provide zero-cost validated literals:

```scala
import com.tjclp.xl.*  // Macros included in unified import

val cellRef: ARef = ref"A1"          // Validated at compile time
val range: CellRange = ref"A1:B10"   // Parsed at compile time

val formula = fx"=SUM(A1:B10)"       // Validated CellValue.Formula
val price   = money"$$1,234.56"      // Formatted(value, NumFmt.Currency)
```

**Implementation**: Zero-allocation parsers emit constants directly as `Expr[ARef]`, `Expr[CellRange]`, or `Expr[CellValue]`. Runtime paths use `Either` for rich error reporting.

### 4. Deterministic Canonicalization

For stable diffs and deduplication:

```scala
// Styles with same canonical key → same index in styles.xml
val key = CellStyle.canonicalKey(style)

// XML attributes/elements in sorted order
XmlUtil.elem("sheets", "name" -> "Sheet1", "sheetId" -> "1")(...)
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

> **For detailed phase completion status, see [docs/plan/roadmap.md](docs/plan/roadmap.md)**

Implementation phases and current status (including test counts and performance results) are tracked in:

- `docs/plan/roadmap.md` – phase definitions and completion status
- `docs/STATUS.md` – current capabilities, coverage, and performance

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
// Equal keys → same styles index in styles.xml
```

**Critical**: When implementing styles.xml writer, build style index first, then emit cells.

### XML Determinism

All XML output must be deterministic:

1. **Attribute order**: Always sort by name (`XmlUtil.elem` does this)
2. **Element order**: Sort by natural key (style index, cell ref, etc.)
3. **Whitespace**: Prefer `XmlUtil.compact` for production (stable single-line elements) and `XmlUtil.prettyPrint` for debug builds (via `WriterConfig.debug`)

## Common Patterns

### Working with Either

```scala
// Chain operations
for
  ref <- ARef.parse("A1")
  sheet <- workbook("Sheet1")
  updated <- sheet.put(Patch.Put(ref, value))
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

Macros in `xl-core/src/com/tjclp/xl/macros/`:

- Use `quotes.reflect.report.errorAndAbort` for compile errors
- Emit constants directly: `'{ (${Expr(value)}: T).asInstanceOf[OpaqueType] }`
- Zero-allocation parsers (no regex, manual char iteration)
- All macros are `transparent inline` (zero runtime cost)

**Hybrid Compile-Time/Runtime Validation Pattern**:

XL uses a smart pattern for methods that accept both literals and dynamic strings (e.g., `.hex()` in Style DSL):

```scala
// Macro checks if parameter is compile-time constant
transparent inline def hex(code: String): CellStyle = ${ hexMacro('code, 'style) }

def hexMacro(code: Expr[String], style: Expr[CellStyle])(using Quotes): Expr[CellStyle] =
  code.value match
    case Some(literal) =>
      // Compile-time validation for string literals
      Color.fromHex(literal) match
        case Right(c) => emitConstant(c, style)
        case Left(err) => quotes.reflect.report.errorAndAbort(s"Invalid hex: $err")
    case None =>
      // Runtime validation for dynamic strings (pure, silent fail)
      '{ Color.fromHex($code).fold(_ => $style, c => applyColor($style, c)) }
```

**Benefits**:
- String literals → compile errors if invalid (`style.hex("#GGGGGG")` fails build)
- Dynamic strings → runtime validation (`style.hex(userInput)` compiles, fails gracefully)
- Single API (no separate methods needed)
- Zero overhead (literals emit constants, runtime uses efficient Either)

**String Interpolation Pattern** (P7-P8):

XL macros support runtime string interpolation with compile-time optimization:

```scala
// Phase 1: Runtime interpolation for all 4 macro families
val sheet = "Sheet1"
val col = "B"
val row = "10"

// Runtime validation (returns Either[XLError, RefType])
val cellRef = ref"$sheet!$col$row"  // Runtime: parses "Sheet1!B10"
val money = money"$$$amount"        // Runtime: parses with dollar sign
val pct = percent"${value}%"        // Runtime: parses with percent

// Phase 2: Compile-time optimization when all parts are literals
val optimized = ref"${"A"}${"1"}"  // Optimized: emits ARef constant at compile time
val fallback = ref"$sheet!A1"      // Fallback: uses runtime path (sheets is variable)
```

**Implementation** (`MacroUtil.scala`):
- `MacroUtil.reconstructString`: Rebuilds interpolated string from parts
- Hybrid path: Detects if all interpolated values are compile-time constants
- Compile-time optimization: Emits opaque type constant directly (zero runtime cost)
- Runtime fallback: Uses pure Either-based validation for dynamic strings

**Benefits**:
- Dynamic references from user input (`ref"${userSheet}!${userCell}"`)
- Type-safe: Returns `Either[XLError, RefType]` for error handling
- Zero overhead when all parts are literals (Phase 2 optimization)
- Single unified API (no separate runtime vs compile-time methods)

### DSL Patterns

**Patch DSL** (`xl-core/src/com/tjclp/xl/dsl.scala`):
```scala
import com.tjclp.xl.dsl.*

// Ergonomic operators (no type ascription needed)
val patch = (cell"A1" := "Title") ++ (cell"A1".styled(headerStyle)) ++ range"A1:B1".merge

// Compared to Monoid syntax (requires type ascription)
val patch = (Patch.Put(cell"A1", value): Patch) |+| (Patch.SetStyle(cell"A1", id): Patch)
```

**Optics** (`xl-core/src/com/tjclp/xl/optics.scala`):
```scala
import Optics.*

// Modify cell value functionally
sheet.modifyValue(cell"A1") {
  case CellValue.Text(s) => CellValue.Text(s.toUpperCase)
  case other => other
}

// Focus on specific cell
sheet.focus(cell"B1").modify(_.withValue(CellValue.Number(42)))(sheet)
```

**ExcelR** (Pure error handling):
```scala
import com.tjclp.xl.io.ExcelR

val excel: ExcelR[IO] = ExcelIO.instance
excel.readR(path).flatMap {
  case Right(wb) => processWorkbook(wb)
  case Left(err) => handleError(err)
}
```

### Codec Patterns

Cell-level codecs (`xl-core/src/com/tjclp/xl/codec/`) provide type-safe encoding/decoding with auto-inferred formatting.

#### When to Use Codecs

- ✅ Financial models (cell-oriented, irregular layouts)
- ✅ Unstructured sheets (no schema)
- ✅ Type-safe reading from user input
- ✅ Auto-inferred date/number formats

#### Batch Updates

Use batch `put` for multi-cell updates with automatic type inference:

```scala
sheet.put(
  ref"A1" -> "Revenue",
  ref"B1" -> LocalDate.of(2025, 11, 10),  // Auto: date format
  ref"C1" -> BigDecimal("1000000.50"),    // Auto: decimal format
  ref"D1" -> 42,
  ref"E1" -> money"$$1,234.56"            // Preserves Currency format!
)
```

**Benefits**: Cleaner syntax, auto-inferred styles, preserves Formatted literal metadata.

#### Type-Safe Reading

```scala
// Reading returns Either[CodecError, Option[A]]
sheet.readTyped[BigDecimal](cell"C1") match
  case Right(Some(value)) => // Success: cell has value
  case Right(None) => // Success: cell is empty
  case Left(error) => // Error: type mismatch or parse errors

// Convert to XLError if needed
val xlResult: Either[XLError, Option[BigDecimal]] =
  sheet.readTyped[BigDecimal](ref)
    .left.map(_.toXLError(ref))
```

#### Supported Types

9 primitive types with inline given instances (zero-overhead):
- `String`, `Int`, `Long`, `Double`, `BigDecimal`, `Boolean`, `LocalDate`, `LocalDateTime`, `RichText`

#### Auto-Inferred Formats

- `LocalDate` → `NumFmt.Date`
- `LocalDateTime` → `NumFmt.DateTime`
- `BigDecimal` → `NumFmt.Decimal`
- `Int`, `Long`, `Double` → `NumFmt.General`
- `String`, `Boolean`, `RichText` → No cell-level format (RichText has run-level formatting)

#### Identity Laws

All codecs satisfy: `codec.read(Cell(ref, codec.write(value)._1)) == Right(Some(value))` (up to formatting precision).

#### Rich Text DSL

For intra-cell formatting (multiple formats within one cell):

```scala
import com.tjclp.xl.richtext.RichText.*

// Composable with + operator
val text = "Bold".bold.red + " normal " + "Italic".italic.blue

// Use with putMixed
sheet.put(
  cell"A1" -> ("Error: ".red.bold + "Fix this!"),
  cell"A2" -> ("Q1 ".size(18.0).bold + "Report".italic)
)

// Or use directly
sheet.put(cell"A1", CellValue.RichText(text))
```

**Formatting methods**: `.bold`, `.italic`, `.underline`, `.red`, `.green`, `.blue`, `.black`, `.white`, `.size(pt)`, `.fontFamily(name)`, `.withColor(color)`

**OOXML mapping**: Each `TextRun` → `<r>` element with `<rPr>` (run properties). Uses `xml:space="preserve"` to preserve whitespace.

#### HTML Export

Export sheet ranges to HTML for web display:

```scala
val html = sheet.toHtml(range"A1:B10")  // With inline CSS
val plain = sheet.toHtml(range"A1:B10", includeStyles = false)  // No CSS
```

Rich text and cell styles preserved as HTML tags and inline CSS. Useful for dashboards, reporting, email generation.

## Documentation

Documentation is organized by purpose (no numbering for easier maintenance):

### Active Plans (`docs/plan/`) - 9 files
**Future work only** (P7-P11). Completed phases archived.
- `roadmap.md` → Master status tracker (P0-P11)
- `error-model-and-safety.md` → P11 security hardening
- `formula-system.md` → P7+ evaluator
- `drawings.md`, `charts.md`, `tables-and-pivots.md` → P8-P10 features
- `benchmarks.md` → Performance testing plan
- `security.md` → Additional P11 features

### Design Docs (`docs/design/`) - 6 files
**Architectural decisions** (timeless)
- `executive-summary.md` → Vision and non-negotiables
- `purity-charter.md` → Effect isolation, totality, laws
- `domain-model.md` → Full type algebra (Cell, Sheet, Workbook, RichText, StyleRegistry)
- `decisions.md` → ADR log (10 decisions documented)
- `style-guide.md` → Coding conventions
- `query-api.md` → Unimplemented design spec

### Reference (`docs/reference/`) - 5 files
**Quick reference material**
- `testing-guide.md` → Test coverage breakdown (263 tests)
- `examples.md` → Code samples
- `glossary.md` → Terminology
- `ooxml-cheatsheet.md` → Quick reference
- `ooxml-research.md` → OOXML schemas (ChatGPT research, 484 lines)

### Archived Plans (`docs/archive/plan/`) - 11 files
**Completed implementation plans** (P0-P6, P31), organized by phase:
- `p0-bootstrap/` → Build system, linting
- `p1-addressing/` → Opaque types, macros
- `p2-patches/` → Patch Monoid
- `p3-styles/` → Style system
- `p4-ooxml/` → XML serialization
- `p5-streaming/` → fs2-data-xml streaming
- `p6-codecs/` → CellCodec primitives
- `p31-refactor/` → Optics, RichText, HTML export

### Root Docs
- `docs/STATUS.md` → Detailed current state (263 tests, performance, limitations)
- `docs/LIMITATIONS.md` → Current roadmap and missing features
- `docs/CONTRIBUTING.md` → Contribution guidelines
- `docs/FAQ.md` → Frequently asked questions

## CI/CD

**GitHub Actions**: `.github/workflows/ci.yml`

Pipeline:
1. Check formatting (`./mill __.checkFormat`)
2. Compile all modules (`./mill __.compile`)
3. Run tests (`./mill __.test`)

**Caching**: Coursier + Mill artifacts cached for 2-5x speedup.

## Reviewing PR Feedback

### Working with Claude Code and Codex Reviews

This repository uses both **Claude Code** (AI-assisted development) and **Codex** (automated PR review) to maintain code quality.

#### Viewing PR Feedback

```bash
# View PR details and comments
gh pr view <number>

# View PR with all comments inline
gh pr view <number> --comments

# View PR review status
gh pr view <number> --json reviews --jq '.reviews[] | "\(.author.login): \(.state)\n\(.body)"'
```

#### Common PR Review Patterns

**1. Codex Automated Reviews**
- Triggered automatically on PR creation or when marked ready for review
- Provides detailed code analysis with specific line-by-line suggestions
- Comments include: spec compliance issues, performance concerns, security risks, test coverage gaps
- Reviews are comprehensive (6-10 detailed issues typical)

**2. Review Severity Levels**

Reviews categorize issues by priority:
- **High-Priority (Before Merge)**: Spec violations, data loss bugs, security issues
- **Medium Priority**: Performance concerns, code quality, test coverage
- **Future Improvements (Not Blockers)**: Optimizations, refactoring opportunities

**3. Addressing Feedback Workflow**

```bash
# 1. Plan mode: Review feedback and create fix plan
# Use Task agent to analyze current state vs feedback
# Present plan with ExitPlanMode tool

# 2. Execute fixes systematically
# Fix high-priority issues first
# Add regression tests for each fix
# Verify existing tests still pass

# 3. Verify and commit
./mill __.compile  # Zero warnings
./mill __.test     # All tests passing
./mill __.reformat # Format code
git add -A && git commit -m "Fix: Address PR feedback - <issue>"
git push

# 4. Update PR with response
gh pr comment <number> --body "Addressed Issue X: <details>"
```

#### Real Example: PR #4 Feedback Response

**Review Identified**:
- Issue 1: RichText missing double-space check (High-Priority)
- Issue 5: VAlign.Middle mapping concern (turned out to be already correct)

**Response Process**:
1. Used Task agent to analyze both issues against current code
2. Found Issue 1 needed fix, Issue 5 was false positive (already implemented correctly)
3. Fixed Worksheet.scala and StreamingXmlWriter.scala (added `|| run.text.contains("  ")`)
4. Added regression test for RichText with internal double spaces
5. Verified all tests passing, formatted code
6. Committed with reference to PR #4 and specific issue number
7. Posted PR comment explaining both fixes/verifications

**Key Patterns**:
- Always verify feedback against current code (use agents to analyze)
- Some "issues" may already be fixed (verify before changing)
- Add regression tests for each fix
- Reference PR number in commit messages
- Post summary comment on PR explaining what was addressed

#### Common Feedback Categories

**Spec Compliance**:
- OOXML/ECMA-376 spec violations
- XML namespace handling
- Attribute ordering and determinism

**Data Integrity**:
- Round-trip fidelity issues
- Data loss scenarios
- Whitespace/formatting preservation

**Code Quality**:
- Performance (O(n²) → O(1) optimizations)
- Code duplication (extract to utilities)
- Proper error handling

**Security**:
- XXE prevention
- ZIP bomb protection
- Path traversal guards
- Input validation

**Testing**:
- Missing edge case tests
- Property-based law verification
- Round-trip tests

#### Best Practices

**DO**:
- ✅ Address high-priority issues immediately before merge
- ✅ Add regression tests for each fix
- ✅ Verify "issues" against actual code (may already be fixed)
- ✅ Document why certain suggestions are deferred (with issue tracking)
- ✅ Reference PR numbers in commit messages
- ✅ Post summary comments explaining what was addressed

**DON'T**:
- ❌ Blindly apply suggestions without verification
- ❌ Skip tests for "trivial" fixes
- ❌ Ignore medium/future suggestions (track them for later)
- ❌ Fix issues without understanding root cause

#### Tracking Future Improvements

Create issues for non-blocking suggestions:

```bash
# Example: Create issue for performance optimization
gh issue create --title "Perf: Optimize style indexOf to O(1)" \
  --body "From PR #4 review: Vector.indexOf is O(n), becomes O(n²) for many styles.

  Recommendation: Use zipWithIndex.toMap for O(1) lookups.

  References: PR #4 Issue 3"
```

Or add to `docs/plan/README.md` under "Future Work" with priority labels.

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
import com.tjclp.xl.api.*

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

263/263 tests passing (as of P6 + P31 completion):
- 17 addressing tests (Column, Row, ARef, CellRange laws)
- 21 patch tests (Monoid laws, application semantics)
- 60 style tests (units, colors, builders, canonicalization, StylePatch, StyleRegistry)
- 8 datetime tests (Excel serial number conversions)
- 42 codec tests (CellCodec identity laws, type safety, auto-inference)
- 16 batch update tests (putMixed, readTyped, style deduplication)
- 18 elegant syntax tests (given conversions, batch put, formatted literals)
- 34 optics tests (Lens/Optional laws, focus DSL, real-world use cases)
- 24 OOXML tests (round-trip, serialization, styles)
- 18 streaming tests (fs2-data-xml, constant-memory I/O)
- 5 RichText tests (composition, formatting, DSL)

Target: 90%+ coverage with property-based tests for all algebras.
