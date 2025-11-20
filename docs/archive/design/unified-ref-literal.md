# Unified `ref` Literal Design

> **Status**: ✅ Completed – Unified `ref"..."` literal implemented, replacing `cell"..."`/`range"..."`.
> This document is retained as a historical design record.
> **Archived**: 2025-11-20

**Status:** Implemented ✅
**Date:** 2025-11-12
**Context:** Discovered during package structure refactoring

## Problem Statement

Naming conflict between `cell` package and `cell` macro literal:

```scala
// Package
import com.tjclp.xl.cell.{Cell, CellValue, CellError}

// Macro (creates ARef, NOT Cell!)
import com.tjclp.xl.macros.cell
val address: ARef = cell"A1"  // Confusing: creates ARef, not Cell
```

**Issues:**
1. `cell"A1"` creates an **ARef** (address reference), not a Cell object
2. Semantic mismatch: "cell" suggests Cell type, but actually produces addressing
3. Cannot use `import com.tjclp.xl.api.*` + `import com.tjclp.xl.macros.*` together
4. Package name `cell` conflicts with macro name `cell`

## Implemented Solution: Unified `ref` Literal

### Design Overview

Single `ref` literal with **format-dependent return type** using `transparent inline`:

- **Simple refs** (`A1`, `A1:B10`) → **Unwrapped** `ARef` or `CellRange` for backwards compatibility
- **Sheet-qualified refs** (`Sales!A1`, `'Q1 Sales'!A1:B10`) → **Wrapped** `RefType` enum to preserve sheet info

```scala
// Location: xl-macros/src/com/tjclp/xl/macros/RefLiteral.scala
extension (inline sc: StringContext)
  transparent inline def ref(): ARef | CellRange | RefType = ${ refImpl0('sc) }
```

### RefType Enum

New enum for sheet-qualified references:

```scala
// Location: xl-core/src/com/tjclp/xl/addressing/RefType.scala
enum RefType derives CanEqual:
  case Cell(ref: ARef)
  case Range(range: CellRange)
  case QualifiedCell(sheet: SheetName, ref: ARef)
  case QualifiedRange(sheet: SheetName, range: CellRange)
```

### Usage Examples

```scala
import com.tjclp.xl.macros.ref
import com.tjclp.xl.addressing.RefType

// Simple refs (unwrapped for backwards compatibility)
val cell: ARef = ref"A1"                      // Returns ARef directly
val range: CellRange = ref"A1:B10"            // Returns CellRange directly

// Sheet-qualified refs (wrapped in RefType)
val qcell = ref"Sales!A1"                     // RefType.QualifiedCell(SheetName("Sales"), ARef(...))
val qrange = ref"'Q1 Sales'!A1:B10"           // RefType.QualifiedRange(SheetName("Q1 Sales"), CellRange(...))

// Workbook access with qualified refs
workbook(ref"Sales!A1") match
  case Right(cell: Cell) => println(s"Found cell: ${cell.value}")
  case Left(err) => println(s"Error: ${err.message}")

// Sheet access (ignores sheet qualifier if present)
sheet(ref"A1")              // Returns Cell
sheet(ref"Sales!A1")        // Returns Cell (ignores "Sales!")
sheet(ref"A1:B10")          // Returns Iterable[Cell]
```

### Macro Implementation

Zero-allocation parser with compile-time validation:

```scala
private def refImpl0(sc: Expr[StringContext])(using Quotes): Expr[ARef | CellRange | RefType] =
  import quotes.reflect.report
  val s = literal(sc)

  try
    val bangIdx = findUnquotedBang(s)  // Find '!' outside quotes
    if bangIdx < 0 then
      // No sheet qualifier → return unwrapped type
      if s.contains(':') then
        val ((cs, rs), (ce, re)) = parseRangeLit(s)
        constCellRange(cs, rs, ce, re)  // Expr[CellRange]
      else
        val (c0, r0) = parseCellLit(s)
        constARef(c0, r0)  // Expr[ARef]
    else
      // Has sheet qualifier → return RefType wrapper
      val sheetPart = s.substring(0, bangIdx)
      val refPart = s.substring(bangIdx + 1)

      // Parse sheet name (handle 'quoted names')
      val sheetName = if sheetPart.startsWith("'") then
        validateSheetName(unquote(sheetPart))
      else
        validateSheetName(sheetPart)

      // Parse ref part
      if refPart.contains(':') then
        constQualifiedRange(sheetName, ...)  // Expr[RefType.QualifiedRange]
      else
        constQualifiedCell(sheetName, ...)   // Expr[RefType.QualifiedCell]
  catch
    case e: IllegalArgumentException =>
      report.errorAndAbort(s"Invalid ref literal '$s': ${e.getMessage}")
```

### Sheet Name Support

Excel-compatible parsing:
- Simple names: `Sales!A1`
- Quoted names with spaces: `'Q1 Sales'!A1`
- Escaped quotes: `'It''s Q1'!A1` ('' → ') ✅ Implemented
- Multiple escaped quotes: `'It''s ''Q1'''!A1` → "It's 'Q1'"

Validation rules (Excel spec):
- Max 31 characters
- No `: \ / ? * [ ]`

**Escaping rules (Excel convention):**
- Single quote inside sheet name must be doubled: `'` → `''`
- Examples:
  - Sheet "It's Q1" → `'It''s Q1'!A1`
  - Sheet "'Test'" → `'''Test'''!A1`

### Integration Points

**Workbook:**
```scala
def apply(ref: RefType): XLResult[Cell | Iterable[Cell]] =
  ref match
    case RefType.Cell(_) | RefType.Range(_) =>
      Left(XLError.InvalidReference("Workbook access requires sheet-qualified ref"))
    case RefType.QualifiedCell(sheet, cellRef) =>
      apply(sheet).map(s => s(cellRef))
    case RefType.QualifiedRange(sheet, range) =>
      apply(sheet).map(s => s.getRange(range))
```

**Sheet:**
```scala
def apply(ref: RefType): Cell | Iterable[Cell] =
  ref match
    case RefType.Cell(cellRef) => apply(cellRef)
    case RefType.Range(range) => getRange(range)
    case RefType.QualifiedCell(_, cellRef) => apply(cellRef)  // Ignores sheet
    case RefType.QualifiedRange(_, range) => getRange(range)  // Ignores sheet
```

## Benefits

1. **Resolves Naming Conflict**
   - `ref` doesn't conflict with `cell` package
   - Can now use `import com.tjclp.xl.api.*` + `import com.tjclp.xl.macros.*` together

2. **Semantic Clarity**
   - `ref` accurately describes what's created (address reference)
   - No confusion with Cell domain type

3. **Unified Concept**
   - Excel has one concept: "references" (cell, range, or qualified)
   - Syntax disambiguates format automatically

4. **Backwards Compatible**
   - Simple refs return unwrapped `ARef`/`CellRange` (drop-in replacement for `cell`/`range`)
   - Existing code works without changes (except import statements)

5. **Type Safety**
   - `transparent inline` provides exact type to compiler
   - Union type only when needed for qualified refs
   - No runtime overhead

6. **Sheet-Qualified Access**
   - New capability: direct workbook access with `ref"Sales!A1"`
   - Enables cross-sheet operations (future)

## Migration Path

### Phase 1: Transition Period (0.2.0 - 0.9.x)

Deprecated macros still work with warnings:

```scala
// OLD (deprecated, shows warnings)
import com.tjclp.xl.macros.CellRangeLiterals.{cell, range}
val ref = cell"A1"
val rng = range"A1:B10"

// NEW (recommended)
import com.tjclp.xl.macros.ref
val ref = ref"A1"
val rng = ref"A1:B10"
```

**Note:** `cell` and `range` are NOT exported at top-level (`import com.tjclp.xl.macros.*`) due to naming conflicts. Users must import explicitly from `CellRangeLiterals` or switch to `ref`.

### Phase 2: Removal (1.0.0)

- Remove `cell` and `range` macros completely
- Only `ref` remains

## Design Principles Alignment

✅ **Purity** - Compile-time validation, zero runtime cost
✅ **Strong Typing** - Transparent inline provides exact types
✅ **Ergonomics** - One concept, intuitive syntax
✅ **Explicit** - Format clearly indicates type (`:` = range, `!` = qualified)
✅ **Modern Scala 3** - Union types, transparent inline, top-level exports

## Implementation Files

- **RefType enum:** `xl-core/src/com/tjclp/xl/addressing/RefType.scala`
- **ref macro:** `xl-macros/src/com/tjclp/xl/macros/RefLiteral.scala`
- **Deprecations:** `xl-macros/src/com/tjclp/xl/macros/CellRangeLiterals.scala`
- **Exports:** `xl-macros/src/com/tjclp/xl/syntax.scala`, `xl-core/src/com/tjclp/xl/api.scala`
- **Integration:** `xl-core/src/com/tjclp/xl/workbook/Workbook.scala`, `xl-core/src/com/tjclp/xl/sheet/Sheet.scala`

## Future Enhancements (Not Yet Implemented)

### String Interpolation

Support dynamic refs with Scala variables:

```scala
val sheetName = "Sales"
val colName = "A"
val rowNum = 1

ref"${sheetName}!${colName}${rowNum}"  // "Sales!A1"
```

**Challenge:** Conflicts with Excel's absolute ref syntax (`$A$1`). Solution: require escaping `$$A$$1`.

### Full Column/Row References

```scala
ref"A:A"    // Full column A
ref"1:1"    // Full row 1
ref"A:C"    // Columns A through C
ref"1:10"   // Rows 1 through 10
```

**Status:** Deferred. Unclear semantics for operations (what does "merge A:A" mean?). Mainly useful for formula evaluation.

### Absolute References

```scala
ref"$$A$$1"  // Absolute ref (requires escaping $)
ref"$$A1"    // Column absolute, row relative
ref"A$$1"    // Row absolute, column relative
```

**Status:** Deferred until string interpolation design finalized.

## Implemented Features

### Escaped Quote Support ✅

Excel's single quote escaping is now fully supported:

```scala
// Sheet name with apostrophe
ref"'It''s Q1'!A1"           // → RefType.QualifiedCell(SheetName("It's Q1"), ...)

// Multiple quotes
ref"'It''s ''Q1'''!A1"       // → Sheet name: "It's 'Q1'"

// Round-trip preserved
val ref = RefType.QualifiedCell(SheetName.unsafe("It's Q1"), ARef.from1(1, 1))
ref.toA1                      // → "'It''s Q1'!A1"
RefType.parse(ref.toA1)       // → Right(ref) ✅
```

**Implementation:**
- Runtime parser (`RefType.parse`): Unescapes `''` → `'` when parsing
- Macro parser (`ref` literal): Unescapes at compile time with validation
- `toA1` method: Escapes `'` → `''` when generating output

## Related Documentation

- **RefType runtime parser:** `xl-core/src/com/tjclp/xl/addressing/RefType.scala` (companion object)
- **Property tests:** `xl-core/test/src/com/tjclp/xl/addressing/RefTypeSpec.scala` ✅ Implemented
- **Integration tests:** Included in RefTypeSpec (workbook qualified access, pattern matching)

## Notes

- The unified `ref` literal leverages Scala 3's `transparent inline` for type-level programming
- Return type `ARef | CellRange | RefType` provides exact type when context allows
- Unwrapped returns for simple refs maintain backwards compatibility with existing code
- This follows the "compress where unified" principle from the design charter
