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

  // ========== CellRange Extensions ==========

  /**
   * CellRange extensions for creating patches.
   *
   * These methods return Patch objects (not Sheet objects), enabling composition with other patches
   * before application.
   */
  extension (range: CellRange)
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
   * '''Important Limitation: Assignment operators (`:=`) only work on single cells'''
   *
   * When applied to ranges, assignment operators return `Patch.empty` (silent no-op) rather than
   * throwing errors. This is intentional for composability in fold operations, but users should be
   * aware of this behavior.
   *
   * '''Working examples (single cells):'''
   * {{{
   *   val row = "5"
   *   for {
   *     cellRef <- ref"A$row"      // RefType.Cell
   *   } yield (cellRef := "Value")  // ✓ Works - creates Put patch
   * }}}
   *
   * '''Silent no-op examples (ranges):'''
   * {{{
   *   val rangeRef = ref"A1:B10"    // RefType.Range
   *   rangeRef := "Value"            // ✗ Silent no-op - returns Patch.empty
   *
   *   // To style ranges, use .styled() or apply to individual cells:
   *   rangeRef.styled(style)         // ✓ Works for ranges
   * }}}
   *
   * '''Other operations work on both cells and ranges:'''
   * {{{
   *   cellRef.styled(style)          // ✓ Works
   *   rangeRef.styled(style)         // ✓ Works (cells only, not ranges in current impl)
   *   rangeRef.merge                 // ✓ Works
   *   rangeRef.remove                // ✓ Works
   * }}}
   */
  extension (refType: RefType)
    /** Create Put patch from RefType (delegates to underlying ARef) */
    @annotation.targetName("refTypeAssignCellValue")
    inline def :=(cv: CellValue): Patch = refType match
      case RefType.Cell(aref) => aref := cv
      case RefType.QualifiedCell(_, aref) => aref := cv
      case _ => Patch.empty // Ranges can't be assigned single values

    /** Create Put patch with automatic conversion from String */
    @annotation.targetName("refTypeAssignString")
    inline def :=(value: String): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case _ => Patch.empty

    /** Create Put patch with automatic conversion from Int */
    @annotation.targetName("refTypeAssignInt")
    inline def :=(value: Int): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case _ => Patch.empty

    /** Create Put patch with automatic conversion from Long */
    @annotation.targetName("refTypeAssignLong")
    inline def :=(value: Long): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case _ => Patch.empty

    /** Create Put patch with automatic conversion from Double */
    @annotation.targetName("refTypeAssignDouble")
    inline def :=(value: Double): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case _ => Patch.empty

    /** Create Put patch with automatic conversion from BigDecimal */
    @annotation.targetName("refTypeAssignBigDecimal")
    inline def :=(value: BigDecimal): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case _ => Patch.empty

    /** Create Put patch with automatic conversion from Boolean */
    @annotation.targetName("refTypeAssignBoolean")
    inline def :=(value: Boolean): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case _ => Patch.empty

    /** Create Put patch with automatic conversion from LocalDateTime */
    @annotation.targetName("refTypeAssignLocalDateTime")
    inline def :=(value: java.time.LocalDateTime): Patch = refType match
      case RefType.Cell(aref) => aref := value
      case RefType.QualifiedCell(_, aref) => aref := value
      case _ => Patch.empty

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
