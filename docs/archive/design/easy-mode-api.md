# Easy Mode API: Ultra-Clean Single-Import Ergonomics

> **Status**: ✅ Completed – Easy Mode API implemented via type-class consolidation (PR #20).
> This document is retained as a historical design record.
> **Archived**: 2025-11-20

**Status**: ~~Design Proposal~~ → ✅ Implemented (PR #20)
**Author**: Claude (AI Assistant)
**Date**: 2025-11-19
**Goal**: Make `import com.tjclp.xl.*` so ergonomic that incorrect code is harder to write than correct code

---

## Executive Summary

XL's current API is mathematically rigorous and type-safe, but requires ceremony for simple tasks. This proposal adds an **"easy mode" layer** via exports and minimal DSL wrapping, achieving:

- ✅ **Single import**: `import com.tjclp.xl.*` (everything you need for core domain)
- ✅ **LLM-friendly**: Predictable, discoverable, forgiving API
- ✅ **Zero compromise**: Pure API still available underneath
- ✅ **Hard to get wrong**: Sensible defaults, clear errors, guided experience
- ✅ **Structured Errors**: Throws `XLException` wrapping `XLError` for programmatic error recovery

**Implementation**: 4-5 files, ~200 LOC, no new dependencies, fully backwards compatible.

---

## Philosophy: Three Tiers of Safety

### Understanding "Unsafe" in Modern Scala 3

Modern Scala 3 libraries (ZIO, Cats Effect, circe, http4s) reserve the term **"unsafe"** for:
1. **IO effects** (reading files, network, databases)
2. **Breaking referential transparency** (mutable state, side effects)
3. **Bypassing the type system** (casts, nulls)

**Not** for validation! Validation can throw deterministically and still be "safe" in a scripting context if it wraps a rigorous error model.

### XL's Three Tiers

**Tier 1: Pure API** (production code, libraries)
```scala
import com.tjclp.xl.*

// Explicit Either-based error handling
val result: Either[XLError, Workbook] = for {
  sheet <- Sheet("Sales")
  updated <- sheet.put(ref"A1" := "Product")
  wb <- Workbook.empty
  final <- wb.put(updated)
} yield final
```

**Tier 2: Easy Mode Extensions** (scripts, REPL, LLM-generated)
```scala
import com.tjclp.xl.*

// No .unsafe visible anywhere! Hidden in extension implementations
// All validation failures throw XLException (wrapping XLError)
val sheet = Sheet("Sales")                    // Throws on invalid name
  .put("A1", "Product")                       // Throws on invalid ref
  .put("B1", "Revenue", headerStyle)          // Throws on invalid ref

// Lookups naturally safe (return Option)
val cell: Option[Cell] = sheet.cell("A1")
val html: String = sheet.toHtml(range"A1:B10")  // Pure, can't fail

val wb = Workbook.empty.addSheet(sheet)
val sales = wb.sheet("Sales")                  // Throws if not found
```

**Tier 3: IO Boundary** (true unsafe operations)
```scala
import com.tjclp.xl.easy.*

// ONLY IO operations are "truly unsafe"
val wb = Excel.read("data.xlsx")      // Can throw IOException / XLException
Excel.write(wb, "output.xlsx")        // Can throw IOException
```

### Key Insight: Validation ≠ Unsafe

| Operation Type | Return Type | Throws? | Truly Unsafe? |
|---------------|-------------|---------|---------------|
| **Validation** (parse "A1") | Either or throw | `XLException` | ❌ No - deterministic |
| **Lookup** (get cell) | `Option[Cell]` | Never | ❌ No - pure |
| **Transformation** (toHtml) | `String` | Never | ❌ No - pure |
| **Constructor** (Sheet("x")) | Either or throw | `XLException` | ❌ No - validation |
| **IO** (read file) | Either or throw | `IOException` | ✅ **YES** - effects |

**Philosophy**:
- `.unsafe` is an **implementation detail** of easy mode extensions
- Users **never type** `.unsafe` in their code
- The word "unsafe" is reserved for **true IO boundaries** or explicit opt-in
- Exceptions preserve the structure of `XLError` via `XLException`

---

## Design Principles

### 1. **Unified Error Handling (XLException)**

Instead of throwing generic `IllegalArgumentException` or `IllegalStateException`, we introduce `XLException` which wraps the underlying `XLError`. This allows scripts to catch errors and inspect the structural reason if needed, maintaining a link to the rigorous core model.

```scala
try {
  sheet.put("InvalidRef", "Value")
} catch {
  case e: XLException => println(e.error) // e.g., XLError.InvalidCellRef(...)
}
```

### 2. **Architecture: Layered Modules**

To avoid circular dependencies between `xl-core` and `xl-cats-effect`:
- `xl-core`: Exports domain, macros, extensions, and `.unsafe` (validation only).
- `xl-cats-effect`: Exports `easy` object which includes `xl-core` syntax **PLUS** the `Excel` IO object.

### 3. **Data vs Style Separation**

XL provides three distinct styling approaches, each optimized for different use cases:

**1. Range Styling (for templates and structure)**
```scala
// Style first (formatting/structure)
sheet
  .style("A1:E1", titleStyle)       // Title row
  .style("A2:E2", headerStyle)      // Header row
  .style("A3:E100", dataStyle)      // Data rows

// Then add data (content-focused)
  .put("A1", "Q1 Sales Report")
  .put("A2:E2", headers)
  .put("A3:E100", data)
```

**2. Inline Styling (for one-offs)**
```scala
// Data + style together
sheet.put("A1", "Title", headerStyle)
```

**3. RichText Styling (for intra-cell formatting)**
```scala
// Multiple formats within one cell
sheet.put("A1", "Error: ".bold.red + "Fix this")
```

**Philosophy**:
- Use `.style()` for **structure** (templates, bulk formatting before data)
- Use `.put(..., style)` for **one-off** inline styles
- Use RichText for **intra-cell** formatting (multiple styles in one cell)

**Why separate `.style()`?**
- Enables template creation before data is available
- Clear distinction between presentation (formatting) and content (data)
- Efficient bulk styling operations
- Natural workflow: format structure → fill data → save

**Rejected**: `value.style(...)` extensions on String/Int (would pollute primitives and conflict with RichText)

---

## Implementation Strategy

### Phase 1: Foundation & Exception Model (3 files)

**Step 1a: Create XLException wrapper**

**File**: `xl-core/src/com/tjclp/xl/error/XLException.scala` (NEW)

```scala
package com.tjclp.xl.error

final class XLException(val error: XLError) extends RuntimeException(error.message)
```

**Step 1b: Migrate unsafe to Scala 3 style**

**File**: `xl-core/src/com/tjclp/xl/unsafe.scala` (NEW - replaces unsafe/package.scala)

```scala
package com.tjclp.xl

import com.tjclp.xl.error.{XLError, XLException, XLResult}

object unsafe:
  extension [A](result: XLResult[A])
    /** Unwrap result, throwing XLException if Left */
    def unsafe: A = result match
      case Right(value) => value
      case Left(err) => throw new XLException(err)

    /** Unwrap with fallback value */
    def getOrElse(default: => A): A = result.toOption.getOrElse(default)

export unsafe.*
```

**Step 1c: Add unsafe to syntax exports**

**File**: `xl-core/src/com/tjclp/xl/syntax.scala` (MODIFY)

```scala
object syntax:
  // ... existing exports

  // ⭐ NEW: Export unsafe extensions
  export com.tjclp.xl.unsafe.*

export syntax.*
```

### Phase 2: Simplified IO (1 file)

**Constraint**: `xl-core` cannot depend on `xl-cats-effect`.
**Solution**: `xl-cats-effect` provides the unified import for IO users.

**File**: `xl-cats-effect/src/com/tjclp/xl/easy.scala` (NEW)

```scala
package com.tjclp.xl

/**
 * Easy mode: Single import for everything including IO.
 *
 * Usage:
 * {{{
 * import com.tjclp.xl.easy.*
 *
 * val wb = Excel.read("data.xlsx")
 * // ... use wb
 * Excel.write(wb, "output.xlsx")
 * }}}
 */
object easy:
  // Re-export core syntax (domain, macros, extensions)
  export com.tjclp.xl.syntax.*
  export com.tjclp.xl.extensions.* // String-based extensions from Phase 3

  // Export simplified Excel IO
  export com.tjclp.xl.io.Excel

export easy.*
```

Note: `com.tjclp.xl.io.Excel` (the simplified IO object) needs to be created in `xl-cats-effect/src/com/tjclp/xl/io/easy.scala` (or similar) wrapping `ExcelIO`.

### Phase 3: String-Based Extensions (2 files)

**Step 3a: Create extensions with style overloads**

**File**: `xl-core/src/com/tjclp/xl/extensions.scala` (NEW)

```scala
package com.tjclp.xl

import com.tjclp.xl.api.{Sheet, Workbook}
import com.tjclp.xl.cell.Cell
import com.tjclp.xl.style.CellStyle
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.codec.CellCodec
import com.tjclp.xl.unsafe.unsafe // Internal use of .unsafe

/**
 * String-based extensions for ergonomic API.
 *
 * These extensions accept string cell references and throw XLException on parse errors.
 * Automatically exported via `import com.tjclp.xl.syntax.*` (or separate import).
 */
object extensions:

  extension (sheet: Sheet)
    // --- Standard Put (data only) ---
    def put(cellRef: String, value: String): Sheet =
      ARef.parse(cellRef)
        .flatMap(ref => sheet.put(ref := value))
        .unsafe

    def put(cellRef: String, value: Int): Sheet = ...
    def put(cellRef: String, value: Double): Sheet = ...
    // ... other primitive overloads

    // --- Styled Put (data + style inline) ---
    def put(cellRef: String, value: String, style: CellStyle): Sheet =
      ARef.parse(cellRef)
        .flatMap { ref =>
             sheet.put(ref := value)
                  .flatMap(_.styleCell(ref, style))
        }
        .unsafe

    def put(cellRef: String, value: Int, style: CellStyle): Sheet = ...
    // ... other primitive overloads with style

    // --- Range Put (auto-detects ":" in ref) ---
    def put[A: CellCodec](rangeRef: String, values: List[A]): Sheet =
      CellRange.parse(rangeRef)
        .flatMap { range =>
          values.zipWithIndex.foldLeft(Right(sheet): XLResult[Sheet]) {
            case (Right(s), (value, idx)) =>
              val cell = range.start.offset(0, idx)
              s.put(cell := value)
            case (left, _) => left
          }
        }
        .unsafe

    def put[A: CellCodec](rangeRef: String, values: List[A], style: CellStyle): Sheet =
      // Put data + style the range
      put(rangeRef, values).style(rangeRef, style)

    // --- Style Method (formatting without data) ---
    def style(ref: String, style: CellStyle): Sheet =
      if ref.contains(":") then
        // Range styling
        CellRange.parse(ref)
          .flatMap(range => sheet.styleRange(range, style))
          .unsafe
      else
        // Single cell styling
        ARef.parse(ref)
          .flatMap(aref => sheet.styleCell(aref, style))
          .unsafe

    // --- Accessors (lookups) ---
    def cell(cellRef: String): Option[Cell] =
      ARef.parse(cellRef).toOption.flatMap(ref => sheet.cells.get(ref))

    def range(rangeRef: String): List[Cell] =
      CellRange.parse(rangeRef)
        .map(range => range.cells.flatMap(sheet.cells.get).toList)
        .unsafe

    def merge(rangeRef: String): Sheet =
      CellRange.parse(rangeRef)
        .flatMap(range => sheet.merge(range))
        .unsafe

export extensions.*
```

**WartRemover**: The implementation of `extensions.scala` will need `@SuppressWarnings(Array("org.wartremover.warts.Throw"))` because it intentionally throws `XLException`.

**Required Sheet API** (check if these exist or need to be added):
- `sheet.styleCell(ref: ARef, style: CellStyle): XLResult[Sheet]` - Apply style to single cell
- `sheet.styleRange(range: CellRange, style: CellStyle): XLResult[Sheet]` - Apply style to range

If these don't exist, they can be implemented as:
```scala
extension (sheet: Sheet)
  def styleCell(ref: ARef, style: CellStyle): XLResult[Sheet] =
    // Apply style to the cell at ref, preserving its value
    sheet.cells.get(ref) match
      case Some(cell) => sheet.put(Patch.Put(ref, cell.value, Some(style)))
      case None => sheet.put(Patch.Put(ref, CellValue.Empty, Some(style)))

  def styleRange(range: CellRange, style: CellStyle): XLResult[Sheet] =
    range.cells.foldLeft(Right(sheet): XLResult[Sheet]) { (acc, ref) =>
      acc.flatMap(_.styleCell(ref, style))
    }
```

**Step 3b: Export from syntax**

**File**: `xl-core/src/com/tjclp/xl/syntax.scala` (MODIFY)

```scala
object syntax:
  // ...
  export com.tjclp.xl.unsafe.*
  export com.tjclp.xl.extensions.* // Make string extensions available by default in core
```

---

## Usage Examples

### Example 1: Template-First Workflow (Separation of Concerns)

```scala
import com.tjclp.xl.easy.*

// Step 1: Create template (formatting/structure)
val template = Sheet("Q1 Sales")
  .style("A1:E1", titleStyle)           // Title row
  .style("A2:E2", headerStyle)          // Header row
  .style("A3:E100", dataStyle)          // Data rows
  .style("F3:F100", formulaStyle)       // Formula column
  .style("A101:F101", totalStyle)       // Total row

// Step 2: Fill with data (content-focused)
val report = template
  .put("A1", "Q1 2025 Sales Report")
  .put("A2:E2", List("Product", "Q1", "Q2", "Q3", "Total"))
  .put("A3:A10", products)
  .put("B3:D10", quarterlyRevenue)       // 2D data (rows of quarters)
  .put("F3:F10", formulas)                // SUM formulas
  .put("A101", "Total")
  .put("B101:D101", totals)

// Step 3: Save
Excel.write(Workbook.empty.addSheet(report), "q1-sales.xlsx")
```

**Benefits**:
- Clear separation: format → fill → save
- Can reuse template with different data
- Easier to reason about structure vs content

### Example 2: Inline Styling (Quick One-Offs)

```scala
import com.tjclp.xl.easy.*

val sheet = Sheet("Quick Report")
  .put("A1", "Sales Report", titleStyle)    // Inline style
  .put("A2", "Revenue", headerStyle)
  .put("A3:A10", products)                   // Data only
  .put("B3:B10", revenues, currencyStyle)   // Data + style

Excel.write(Workbook.empty.addSheet(sheet), "quick.xlsx")
```

### Example 3: RichText (Intra-Cell Formatting)

```scala
import com.tjclp.xl.easy.*

val sheet = Sheet("Alerts")
  .put("A1", "Status: ".bold + "ERROR".red.bold)
  .put("A2", "Warning: ".yellow + "Fix ASAP")
  .put("A3", "Success: ".green + "All tests passed")

Excel.write(Workbook.empty.addSheet(sheet), "alerts.xlsx")
```

### Example 4: Pure Core (No IO)

```scala
import com.tjclp.xl.* // No IO, but includes string extensions

val sheet = Sheet("Sales")
  .style("A1:B1", headerStyle)   // ✓ Works
  .put("A1", "Product")           // ✓ Works
  .put("B1", "Revenue")           // ✓ Works

// Excel.read/write not available (need xl.easy.* for IO)
```

---

## Error Handling Philosophy

### Easy Mode: Fail Fast with XLException

```scala
try {
  sheet.put("InvalidRef", "Value")
} catch {
  case ex: XLException =>
    // Access structural error info
    ex.error match {
      case XLError.InvalidCellRef(ref, reason) => ...
      case _ => ...
    }
}
```

**Why this is good**:
- Immediate feedback for scripts
- Structural data available for tools/IDE integrations
- Maintains link to rigorous domain model

---

## Complete Easy Mode API Surface

### Sheet Extensions

**Data Operations**:
```scala
// Single cell
sheet.put("A1", "Value")                    // Data only
sheet.put("A1", "Value", style)             // Data + inline style
sheet.put("A1", 42)                         // Type-safe primitives
sheet.put("A1", 42, currencyStyle)

// Range (auto-detects ":" in ref)
sheet.put("A1:A10", products)               // List[String]
sheet.put("B1:B10", revenues, currencyStyle) // List + style
```

**Styling Operations**:
```scala
// Style without data (template pattern)
sheet.style("A1", headerStyle)              // Single cell
sheet.style("A1:Z100", dataStyle)           // Range
```

**Lookups** (naturally safe):
```scala
sheet.cell("A1")                            // Option[Cell]
sheet.range("A1:B10")                       // List[Cell]
```

**Transformations** (pure):
```scala
sheet.toHtml(range"A1:B10")                 // String
sheet.merge("A1:B1")                        // Sheet (throws on parse error)
```

### Workbook Extensions

```scala
// Construction
Workbook.empty.addSheet(sheet)              // Add sheet
wb.sheet("Sales")                           // Get by name (throws if not found)
wb.removeSheet("OldData")                   // Remove by name
```

### IO Operations (xl.easy.* only)

```scala
Excel.read("data.xlsx")                     // Workbook (throws IOException)
Excel.write(wb, "output.xlsx")              // Unit (throws IOException)
Excel.modify("file.xlsx") { wb => ... }     // Read → modify → write
```

---

## Implementation Roadmap (Refined)

1.  **Foundation**: Create `XLException.scala` and migrate `unsafe` to object (with throw XLException). Update `syntax.scala` to export it.
2.  **Extensions**: Create `extensions.scala` in `xl-core` with:
    - String overloads for `put` (data only and data+style)
    - `.style()` method for range/cell styling
    - Helper methods: `styleCell`, `styleRange` (if not already in Sheet API)
    - Suppress WartRemover for intentional throws
3.  **IO Layer**: Create `io/easy.scala` (Excel object) and `easy.scala` (unified export) in `xl-cats-effect`.
4.  **Validation**: Ensure no circular dependencies and that `xl-core` remains pure (no IO).
5.  **Testing**: Add tests for string parsing edge cases, style application, and error messages.

**Backwards Compatibility**: 100% preserved. Pure API remains untouched.
