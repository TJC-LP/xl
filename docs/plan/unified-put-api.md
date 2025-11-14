# Unified `put` API + Style DSL

**Status**: Design phase
**Target**: v0.2.0 (breaking changes acceptable pre-1.0)
**Created**: 2025-11-14

## Executive Summary

Consolidate Sheet's fragmented `put*` methods (`putMixed`, `putAll`, `putTyped`, etc.) into a single, intelligently-overloaded `put` method. Combine this with a fluent CellStyle builder DSL to create the most ergonomic Excel library API possible.

**Core principle**: Make simple things simple (safe chaining), complex things explicit (error handling).

---

## Problem Statement

###Current Fragmentation

The Sheet API has **8+ different `put` methods**, each solving a narrow use case:

```scala
def put(cell: Cell): Sheet
def put(ref: ARef, value: CellValue): Sheet
def putAll(cells: IterableOnce[Cell]): Sheet
def putMixed(updates: (ARef, Any)*): Sheet
def putTyped[A: CellCodec](ref: ARef, value: A): Sheet
def putFormatted(ref: ARef, formatted: Formatted): Sheet
def applyPatch(patch: Patch): XLResult[Sheet]
def applyPatches(patches: Patch*): XLResult[Sheet]
```

**Pain points**:
1. **Discoverability**: Users don't know which method to use
2. **Inconsistency**: Some return `Sheet`, others `XLResult[Sheet]`
3. **Verbosity**: Multiple concepts for the same operation
4. **Mental overhead**: Learn 8 methods instead of 1

### Missing Style Composition

CellStyle construction is verbose:

```scala
// Current: verbose constructors
val style = CellStyle(
  font = Font("Arial", 14.0, bold = true, italic = false, underline = false, color = None),
  fill = Fill.Solid(Color.fromRgb(68, 114, 196)),
  align = Align(HAlign.Center, VAlign.Middle, wrapText = false, indent = 0)
)

// Desired: fluent builder
val style = CellStyle.bold.size(14.0).bgBlue.center.middle
```

---

## Solution: Unified `put` with Smart Overloading

### Design Philosophy

**Three pillars**:
1. **Type-directed returns**: Compiler knows which signature you called
2. **Infallible by default**: Simple operations return `Sheet` directly
3. **Explicit failure**: Complex operations return `XLResult[Sheet]`

### Core API

```scala
extension (sheet: Sheet)
  // Simple operations (infallible) → Sheet
  def put(cell: Cell): Sheet
  def put(ref: ARef, value: CellValue): Sheet

  // Batch operations (infallible, macro-based) → Sheet
  transparent inline def put(pairs: (ARef, Any)*): Sheet

  // Patch operations (fallible) → XLResult[Sheet]
  def put(patch: Patch): XLResult[Sheet]
```

**Key insight**: The compiler distinguishes these via argument types. No ambiguity, no `@targetName` needed!

---

## Usage Examples

### Simple Chaining (Just Works™)

```scala
val sheet = Sheet("Revenue")
  .put(ref"A1", "Product")          // Returns Sheet
  .put(ref"B1", "Price")             // Returns Sheet
  .put(                               // Returns Sheet (macro)
    ref"A2" -> "Widget",
    ref"B2" -> money"$19.99",        // Preserves Currency format
    ref"C2" -> date"2025-11-10"      // Preserves Date format
  )
```

**Why it works**:
- All operations are infallible (simple cell puts can't fail)
- Chaining is natural and ergonomic
- No error handling needed for straightforward operations

### Complex Operations (Explicit Errors)

```scala
val result: XLResult[Sheet] = sheet.put(
  (ref"A1" := "Title") ++           // Build patch
  range"A1:C1".merge                 // Can fail (overlap!)
)

result match
  case Right(s) => println("Success!")
  case Left(XLError.MergeOverlap(msg)) => println(s"Overlap: $msg")
  case Left(err) => println(s"Error: ${err.message}")
```

**Why explicit**:
- Merges can overlap with existing merges
- Range operations can be invalid
- Users must handle failure explicitly

### Mixed: Use `.unsafe` for Known-Safe Operations

```scala
val sheet = Sheet("Data")
  .put(ref"A1", "Title")
  .put(range"A1:C1".merge).unsafe    // I know it's safe, unwrap it
  .put(ref"A2", "More data")         // Continue chaining
```

**Extension method**:
```scala
extension (result: XLResult[Sheet])
  def unsafe: Sheet =
    result.getOrElse(throw new IllegalStateException("Patch failed"))
```

---

## Implementation Strategy

### Phase 1: Consolidate Core Methods

**Remove**:
```scala
def putAll(cells: IterableOnce[Cell]): Sheet
def putMixed(updates: (ARef, Any)*): Sheet
def putTyped[A: CellCodec](ref: ARef, value: A): Sheet
def applyPatch(patch: Patch): XLResult[Sheet]
def applyPatches(patches: Patch*): XLResult[Sheet]
```

**Keep** (as unified `put`):
```scala
def put(cell: Cell): Sheet =
  copy(cells = cells.updated(cell.ref, cell))

def put(ref: ARef, value: CellValue): Sheet =
  put(Cell(ref, value))

transparent inline def put(pairs: (ARef, Any)*): Sheet =
  ${ putBatchMacro('pairs) }

def put(patch: Patch): XLResult[Sheet] =
  patch match
    case Patch.Put(ref, value) =>
      Right(put(ref, value))  // Delegate to infallible
    case Patch.Merge(range) =>
      validateNoOverlap(range).map(_ => applyMerge(range))
    case Patch.Batch(patches) =>
      patches.foldLeft(Right(this): XLResult[Sheet]) { (acc, p) =>
        acc.flatMap(_.put(p))
      }
```

### Phase 2: Enhance Batch Put Macro

**Fix Formatted bug**:
```scala
// Current: Formatted loses NumFmt info
sheet.put(ref"A1" -> money"$123.56")  // Auto-converts to CellValue, loses Currency!

// Fixed: Preserve Formatted metadata
inline case formatted: Formatted =>
  val style = CellStyle.default.withNumFmt(formatted.numFmt)
  Cell(ref, formatted.value).withInferredStyle(style)
```

**Support all value types**:
- Primitives (String, Int, BigDecimal, LocalDate) → CellCodec
- `Formatted` literals → Preserve NumFmt
- `RichText` → Preserve intra-cell formatting
- `Styled` (future) → Preserve cell-level styling

### Phase 3: Internal Unification via Patches

Make all operations build patches internally:

```scala
private def applyPatchUnsafe(patch: Patch): Sheet =
  put(patch) match
    case Right(sheet) => sheet
    case Left(err) =>
      throw new IllegalStateException(s"Infallible patch failed: ${err.message}")

def put(ref: ARef, value: CellValue): Sheet =
  applyPatchUnsafe(Patch.Put(ref, value))

transparent inline def put(pairs: (ARef, Any)*): Sheet =
  val patches = compiletime.summonAll[patches for pairs]
  applyPatchUnsafe(Patch.Batch(patches))
```

**Benefits**:
- One unified execution path (all operations are patches)
- Consistent semantics
- Easy to reason about performance
- Debuggable (inspect patch before application)

---

## CellStyle Builder DSL

### Design

Fluent shortcuts on `CellStyle` (NOT on values):

```scala
// xl-core/src/com/tjclp/xl/style/dsl.scala
extension (style: CellStyle)
  // Font
  inline def bold: CellStyle =
    style.withFont(style.font.withBold(true))
  inline def italic: CellStyle =
    style.withFont(style.font.withItalic(true))
  inline def size(pt: Double): CellStyle =
    style.withFont(style.font.withSize(pt))

  // Colors (font)
  inline def red: CellStyle =
    style.withFont(style.font.withColor(Color.fromRgb(255, 0, 0)))
  inline def blue: CellStyle =
    style.withFont(style.font.withColor(Color.fromRgb(0, 0, 255)))
  inline def white: CellStyle =
    style.withFont(style.font.withColor(Color.fromRgb(255, 255, 255)))

  // Background
  inline def bgBlue: CellStyle =
    style.withFill(Fill.Solid(Color.fromRgb(68, 114, 196)))
  inline def bgRed: CellStyle =
    style.withFill(Fill.Solid(Color.fromRgb(255, 0, 0)))
  inline def bgYellow: CellStyle =
    style.withFill(Fill.Solid(Color.fromRgb(255, 255, 0)))

  // Alignment
  inline def center: CellStyle =
    style.withAlign(style.align.withHAlign(HAlign.Center))
  inline def middle: CellStyle =
    style.withAlign(style.align.withVAlign(VAlign.Middle))
  inline def right: CellStyle =
    style.withAlign(style.align.withHAlign(HAlign.Right))
  inline def wrap: CellStyle =
    style.withAlign(style.align.withWrap(true))

  // Borders
  inline def bordered: CellStyle =
    style.withBorder(Border.all(BorderStyle.Thin))

  // Number formats
  inline def currency: CellStyle =
    style.withNumFmt(NumFmt.Currency)
  inline def percent: CellStyle =
    style.withNumFmt(NumFmt.Percent)
  inline def date: CellStyle =
    style.withNumFmt(NumFmt.Date)

// Prebuilt constants
object Style:
  val header = CellStyle.default.bold.size(14.0).center.bgBlue.white
  val currency = CellStyle.default.withNumFmt(NumFmt.Currency).right
  val dateFormat = CellStyle.default.withNumFmt(NumFmt.Date).center
```

### Usage

```scala
// Build style
val headerStyle = CellStyle.bold.size(16.0).white.bgBlue.center.middle

// Apply via patch
sheet.put(
  (ref"A1" := "Title") ++
  ref"A1".styled(headerStyle) ++
  range"A1:C1".merge
)

// Or use prebuilt
sheet.put(ref"A1".styled(Style.header))
```

**Why NOT on values**:
```scala
// ❌ Confusing: cell-level vs intra-cell
"Revenue".bold  // RichText (intra-cell) or CellStyle (cell-level)?

// ✅ Clear: styles are properties of cells (refs), not values
ref"A1".styled(CellStyle.bold)
```

---

## StylePatch Composition Operator

Remove type ascription tax:

```scala
// xl-core/src/com/tjclp/xl/style/patch/StylePatch.scala
extension (p1: StylePatch)
  infix def ++(p2: StylePatch): StylePatch = StylePatch.combine(p1, p2)
```

**Before**:
```scala
val patch =
  (StylePatch.SetFont(font): StylePatch) |+|
  (StylePatch.SetFill(fill): StylePatch)
```

**After**:
```scala
val patch =
  StylePatch.SetFont(font) ++
  StylePatch.SetFill(fill)
```

---

## Migration Guide

### Breaking Changes

1. **Removed methods**:
   - `putAll` → Use `put(cells: _*)`
   - `putMixed` → Use `put(pairs: _*)`
   - `putTyped` → Use `put(pairs: _*)` with typed values
   - `applyPatch` → Use `put(patch)`
   - `applyPatches` → Use `put(Patch.Batch(patches))`

2. **Return type changes**:
   - Patch operations now return `XLResult[Sheet]` instead of `Sheet`
   - Use `.unsafe` to unwrap when you know it's safe

### Migration Patterns

```scala
// Before
sheet.putMixed(ref"A1" -> "x", ref"B1" -> 42)
// After
sheet.put(ref"A1" -> "x", ref"B1" -> 42)

// Before
sheet.applyPatch(patch)
// After
sheet.put(patch)

// Before
sheet.putAll(cells)
// After
sheet.put(cells.toSeq: _*)

// Before (error handling)
sheet.applyPatch(patch) match
  case Right(s) => s
  case Left(err) => handleError(err)

// After (same)
sheet.put(patch) match
  case Right(s) => s
  case Left(err) => handleError(err)
```

### IntelliJ Structural Search/Replace

```regex
# Pattern: putMixed(...)
Search: \.putMixed\((.*)\)
Replace: .put($1)

# Pattern: applyPatch(...)
Search: \.applyPatch\((.*)\)
Replace: .put($1)

# Pattern: putAll(cells)
Search: \.putAll\((.*)\)
Replace: .put($1.toSeq: _*)
```

---

## Complete Example (Final Vision)

```scala
import com.tjclp.xl.*
import com.tjclp.xl.io.ExcelIO
import cats.effect.IO

// 1. Create sheet with mixed value types
val sheet = Sheet("Financial Report")
  .put(
    ref"A1" -> "Product",
    ref"B1" -> "Revenue",
    ref"C1" -> "Margin",
    ref"D1" -> "Date",
    ref"A2" -> "Widgets",
    ref"B2" -> money"$1,234.56",       // Formatted: preserves Currency
    ref"C2" -> percent"15.5%",         // Formatted: preserves Percent
    ref"D2" -> date"2025-11-10"        // Formatted: preserves Date
  )
  .put(
    ref"A3" -> "Gadgets",
    ref"B3" -> money"$2,345.67",
    ref"C3" -> percent"22.3%",
    ref"D3" -> date"2025-11-11"
  )

// 2. Build styles with fluent DSL
val headerStyle = CellStyle.bold.size(14.0).white.bgBlue.center.middle
val currencyStyle = CellStyle.currency.right.bold
val dateStyle = CellStyle.date.center

// 3. Apply styles + layout via patches
val styled = sheet.put(
  (range"A1:D1".styled(headerStyle)) ++   // Header row
  (range"B2:B3".styled(currencyStyle)) ++ // Currency column
  (range"D2:D3".styled(dateStyle)) ++      // Date column
  range"A1:D1".merge                       // Merge header (returns XLResult!)
)

// 4. Handle potential merge error
val final = styled match
  case Right(s) => s
  case Left(err) =>
    println(s"Warning: ${err.message}")
    sheet  // Fallback to unstyled

// 5. Write to Excel
ExcelIO.instance[IO]
  .write(Workbook(Vector(final)), Paths.get("report.xlsx"))
  .unsafeRunSync()
```

**Output**: Professional Excel file with:
- Formatted numbers (currency, percent, dates)
- Styled header (bold, blue background, centered, merged)
- Properly aligned columns
- Zero error handling boilerplate for simple operations
- Explicit error handling for complex operations

---

## Implementation Timeline

### Phase 1: Core Unification (1-2 days)
- [ ] Consolidate `put` methods in Sheet.scala
- [ ] Update batch put macro to handle Formatted
- [ ] Add `.unsafe` extension for XLResult
- [ ] Remove deprecated methods

### Phase 2: Style DSL (1 day)
- [ ] Create `xl-core/src/com/tjclp/xl/style/dsl.scala`
- [ ] Add 30+ fluent shortcuts
- [ ] Add prebuilt Style constants
- [ ] Export from main syntax

### Phase 3: StylePatch Operator (30 minutes)
- [ ] Add `++` operator to StylePatch
- [ ] Update documentation

### Phase 4: Testing & Migration (1-2 days)
- [ ] Update all test files to use `put`
- [ ] Add ~100 new tests for style DSL
- [ ] Update documentation
- [ ] Update examples/demo

### Phase 5: Documentation (ongoing)
- [ ] Update README with new API
- [ ] Add migration guide
- [ ] Update ScalaDoc
- [ ] Add cookbook entries

**Total estimate**: ~1 week for complete implementation + testing

---

## Success Criteria

- [ ] All `put*` methods consolidated into single `put`
- [ ] CellStyle has 30+ fluent shortcuts
- [ ] StylePatch `++` works without type ascription
- [ ] All 263+ existing tests pass
- [ ] ~100 new tests for unified API and style DSL
- [ ] Demo uses clean unified syntax
- [ ] Documentation updated
- [ ] No ambiguous overload errors
- [ ] Compile times unchanged
- [ ] Zero runtime overhead (all inline)

---

## Future Considerations

### Value-Level Styling (Deferred)

```scala
case class Styled(value: CellValue, style: CellStyle)

extension (s: String)
  def withStyle(style: CellStyle): Styled = Styled(CellValue.Text(s), style)

// Usage
sheet.put(
  ref"A1" -> "Title".withStyle(headerStyle)  // One operation!
)
```

**Why deferred**: Styles are conceptually properties of cells (refs), not values. This needs more design work to avoid confusion with RichText (intra-cell formatting).

### Range Styling Methods

```scala
def putWithStyle(range: CellRange, values: Iterable[CellValue], style: CellStyle): XLResult[Sheet]
```

Could simplify bulk styling operations.

---

## References

- **Current Patch DSL**: `xl-core/src/com/tjclp/xl/dsl/syntax.scala`
- **RichText DSL**: `xl-core/src/com/tjclp/xl/richtext/RichText.scala`
- **Formatted**: `xl-core/src/com/tjclp/xl/formatted/Formatted.scala`
- **StylePatch**: `xl-core/src/com/tjclp/xl/style/patch/StylePatch.scala`
- **Related**: Phase P6 (Codecs), Phase P31 (RichText, Optics)

---

**Next steps**: Implement Phase 1 on new branch after PR #9 merges.
