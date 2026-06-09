package com.tjclp.xl.dsl

import com.tjclp.xl.addressing.{ARef, CellRange, RefType}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.units.StyleId

/**
 * DSL extensions for ergonomic patch building and sheet operations.
 *
 * Provides operator syntax for creating and composing patches without requiring Cats Monoid syntax
 * or type ascription on enum cases.
 *
 * Usage:
 * {{{
 *   import com.tjclp.xl.dsl.*
 *
 *   val patch = (ref"A1" := "Hello") ++ (ref"A1".styled(boldStyle)) ++ ref"A1:B1".merge
 *   sheet.applyPatch(patch)
 * }}}
 */

object syntax:
  // ========== ARef Extensions ==========

  /**
   * ARef extensions for creating patches without explicit Patch constructors.
   *
   * The `:=` operator creates Put patches with automatic type conversion, eliminating the need for
   * explicit CellValue constructors or type ascription.
   */
  extension (ref: ARef)
    /** Create a Put patch to set cell value */
    inline def :=(cv: CellValue): Patch = Patch.Put(ref, cv)

    /** Create a Put patch with automatic conversion from String */
    inline def :=(value: String): Patch = Patch.Put(ref, CellValue.Text(value))

    /** Create a Put patch with automatic conversion from Int */
    inline def :=(value: Int): Patch = Patch.Put(ref, CellValue.Number(BigDecimal(value)))

    /** Create a Put patch with automatic conversion from Long */
    inline def :=(value: Long): Patch = Patch.Put(ref, CellValue.Number(BigDecimal(value)))

    /** Create a Put patch with automatic conversion from Double */
    inline def :=(value: Double): Patch = Patch.Put(ref, CellValue.Number(BigDecimal(value)))

    /** Create a Put patch with automatic conversion from BigDecimal */
    inline def :=(value: BigDecimal): Patch = Patch.Put(ref, CellValue.Number(value))

    /** Create a Put patch with automatic conversion from Boolean */
    inline def :=(value: Boolean): Patch = Patch.Put(ref, CellValue.Bool(value))

    /** Create a Put patch with automatic conversion from LocalDateTime */
    inline def :=(value: java.time.LocalDateTime): Patch =
      Patch.Put(ref, CellValue.DateTime(value))

    /**
     * Create a SetCellStyle patch using a CellStyle object.
     *
     * The style will be auto-registered in the sheet's StyleRegistry when applied.
     */
    def styled(style: CellStyle): Patch = Patch.SetCellStyle(ref, style)

    /** Create a SetStyle patch using a StyleId (for pre-registered styles) */
    def styleId(id: StyleId): Patch = Patch.SetStyle(ref, id)

    /** Create a ClearStyle patch to remove styling from a cell */
    def clearStyle: Patch = Patch.ClearStyle(ref)

    /**
     * Navigate down by n rows (default 1). Total; like `shift`, staying in bounds is the caller's
     * responsibility (Excel rows end at 1048576).
     */
    def down(n: Int = 1): ARef = ref.shift(0, n)

    /** Navigate up by n rows (default 1). Total; see `down` for the bounds contract. */
    def up(n: Int = 1): ARef = ref.shift(0, -n)

    /** Navigate right by n columns (default 1). Total; see `down` for the bounds contract. */
    def right(n: Int = 1): ARef = ref.shift(n, 0)

    /** Navigate left by n columns (default 1). Total; see `down` for the bounds contract. */
    def left(n: Int = 1): ARef = ref.shift(-n, 0)

  // ========== CellRange Extensions ==========

  /**
   * CellRange extensions for creating patches.
   *
   * These methods return Patch objects (not Sheet objects), enabling composition with other patches
   * before application.
   */
  extension (range: CellRange)
    /**
     * Fill every cell in the range with the same value (Excel Ctrl+Enter semantics).
     *
     * Desugars to a Batch of Puts — no new Patch case, so monoid laws and exhaustive matches are
     * unaffected. A 1x1 range is equivalent to a single-cell `:=`.
     */
    @annotation.targetName("rangeAssignCellValue")
    def :=(cv: CellValue): Patch =
      Patch.Batch(range.cells.map(r => Patch.Put(r, cv): Patch).toVector)

    /** Fill every cell in the range with a String (Excel Ctrl+Enter semantics). */
    @annotation.targetName("rangeAssignString")
    def :=(value: String): Patch = range := CellValue.Text(value)

    /** Fill every cell in the range with an Int (Excel Ctrl+Enter semantics). */
    @annotation.targetName("rangeAssignInt")
    def :=(value: Int): Patch = range := CellValue.Number(BigDecimal(value))

    /** Fill every cell in the range with a Long (Excel Ctrl+Enter semantics). */
    @annotation.targetName("rangeAssignLong")
    def :=(value: Long): Patch = range := CellValue.Number(BigDecimal(value))

    /** Fill every cell in the range with a Double (Excel Ctrl+Enter semantics). */
    @annotation.targetName("rangeAssignDouble")
    def :=(value: Double): Patch = range := CellValue.Number(BigDecimal(value))

    /** Fill every cell in the range with a BigDecimal (Excel Ctrl+Enter semantics). */
    @annotation.targetName("rangeAssignBigDecimal")
    def :=(value: BigDecimal): Patch = range := CellValue.Number(value)

    /** Fill every cell in the range with a Boolean (Excel Ctrl+Enter semantics). */
    @annotation.targetName("rangeAssignBoolean")
    def :=(value: Boolean): Patch = range := CellValue.Bool(value)

    /** Fill every cell in the range with a LocalDateTime (Excel Ctrl+Enter semantics). */
    @annotation.targetName("rangeAssignLocalDateTime")
    def :=(value: java.time.LocalDateTime): Patch = range := CellValue.DateTime(value)

    /** Create a Merge patch */
    def merge: Patch = Patch.Merge(range)

    /** Create an Unmerge patch */
    def unmerge: Patch = Patch.Unmerge(range)

    /** Create a RemoveRange patch to clear all cells in the range */
    def remove: Patch = Patch.RemoveRange(range)

    /** Create a SetRangeStyle patch to apply style to all cells in range */
    def styled(style: CellStyle): Patch = Patch.SetRangeStyle(range, style)

  // ========== RefType Extensions (Runtime Interpolation Support) ==========

  /**
   * RefType extensions for seamless patch DSL integration with runtime-interpolated refs.
   *
   * These extensions allow `ref"$var"` (which returns `Either[XLError, RefType]`) to work directly
   * with patch DSL operators without manual extraction of ARef/CellRange.
   *
   * `:=` is total across all RefType cases: cells produce a single Put; ranges fill every cell with
   * the value (Excel Ctrl+Enter semantics, a Batch of Puts). Before 0.11.0, ranges silently
   * returned `Patch.empty` — that was a bug, fixed by fill semantics.
   *
   * {{{
   *   val row = "5"
   *   for cellRef <- ref"A$row"        // RefType.Cell
   *   yield (cellRef := "Value")       // single Put
   *
   *   for rangeRef <- ref"A1:B$row"    // RefType.Range
   *   yield (rangeRef := 0)            // fills A1:B5 with 0
   * }}}
   *
   * `styled`, `merge`, and `remove` also work on both cells and ranges (merge of a single cell is
   * meaningless and returns the empty patch).
   */
  extension (refType: RefType)
    /** Create Put patch from RefType (delegates to underlying ARef) */
    @annotation.targetName("refTypeAssignCellValue")
    inline def :=(cv: CellValue): Patch = refType match
      case RefType.Cell(aref) => aref := cv
      case RefType.QualifiedCell(_, aref) => aref := cv
      case RefType.Range(range) => range := cv
      case RefType.QualifiedRange(_, range) => range := cv

    /** Create Put patch with automatic conversion from String */
    @annotation.targetName("refTypeAssignString")
    inline def :=(value: String): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case RefType.Range(range) => range := value
      case RefType.QualifiedRange(_, range) => range := value

    /** Create Put patch with automatic conversion from Int */
    @annotation.targetName("refTypeAssignInt")
    inline def :=(value: Int): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case RefType.Range(range) => range := value
      case RefType.QualifiedRange(_, range) => range := value

    /** Create Put patch with automatic conversion from Long */
    @annotation.targetName("refTypeAssignLong")
    inline def :=(value: Long): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case RefType.Range(range) => range := value
      case RefType.QualifiedRange(_, range) => range := value

    /** Create Put patch with automatic conversion from Double */
    @annotation.targetName("refTypeAssignDouble")
    inline def :=(value: Double): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case RefType.Range(range) => range := value
      case RefType.QualifiedRange(_, range) => range := value

    /** Create Put patch with automatic conversion from BigDecimal */
    @annotation.targetName("refTypeAssignBigDecimal")
    inline def :=(value: BigDecimal): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case RefType.Range(range) => range := value
      case RefType.QualifiedRange(_, range) => range := value

    /** Create Put patch with automatic conversion from Boolean */
    @annotation.targetName("refTypeAssignBoolean")
    inline def :=(value: Boolean): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case RefType.Range(range) => range := value
      case RefType.QualifiedRange(_, range) => range := value

    /** Create Put patch with automatic conversion from LocalDateTime */
    @annotation.targetName("refTypeAssignLocalDateTime")
    inline def :=(value: java.time.LocalDateTime): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case RefType.Range(range) => range := value
      case RefType.QualifiedRange(_, range) => range := value

    /** Apply style to RefType (works for both cells and ranges) */
    @annotation.targetName("refTypeStyled")
    def styled(style: CellStyle): Patch = refType match
      case RefType.Cell(aref) => aref.styled(style)
      case RefType.QualifiedCell(_, aref) => aref.styled(style)
      case RefType.Range(range) => range.styled(style)
      case RefType.QualifiedRange(_, range) => range.styled(style)

    /** Merge RefType (only works for ranges, returns empty for cells) */
    @annotation.targetName("refTypeMerge")
    def merge: Patch = refType match
      case RefType.Range(range) => range.merge
      case RefType.QualifiedRange(_, range) => range.merge
      case _ => Patch.empty // Cells can't be merged alone

    /** Remove RefType (works for both cells and ranges) */
    @annotation.targetName("refTypeRemove")
    def remove: Patch = refType match
      case RefType.Cell(aref) => Patch.Remove(aref)
      case RefType.Range(range) => range.remove
      case RefType.QualifiedCell(_, aref) => Patch.Remove(aref)
      case RefType.QualifiedRange(_, range) => range.remove

  // ========== Patch Composition Extensions ==========

  /**
   * Patch composition operator without requiring Cats Monoid syntax.
   *
   * The `++` operator provides an alternative to Cats' `|+|` that doesn't require type ascription
   * on enum cases.
   *
   * Example:
   * {{{
   *   val patch = (ref"A1" := "Title") ++ (ref"A1".styled(headerStyle))
   * }}}
   */
  extension (p1: Patch)
    /** Compose two patches using Patch.combine */
    infix def ++(p2: Patch): Patch = Patch.combine(p1, p2)

  /**
   * Build a Batch patch from multiple patches.
   *
   * Provides a convenient way to create batched patches without manual Vector construction.
   *
   * Example:
   * {{{
   *   val patch = PatchBatch(
   *     ref"A1" := "Title",
   *     ref"A1".styled(headerStyle),
   *     ref"A1:B1".merge
   *   )
   * }}}
   */
  object PatchBatch:
    def apply(patches: Patch*): Patch = Patch.Batch(patches.toVector)

export syntax.*

// Re-export all of RowColumnDsl (builders, entry points, and extension methods)
export RowColumnDsl.*
