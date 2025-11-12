package com.tjclp.xl.dsl

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.style.CellStyle
import com.tjclp.xl.style.units.StyleId

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
 *   val patch = (cell"A1" := "Hello") ++ (cell"A1".styled(boldStyle)) ++ range"A1:B1".merge
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

  // ========== Patch Composition Extensions ==========

  /**
   * Patch composition operator without requiring Cats Monoid syntax.
   *
   * The `++` operator provides an alternative to Cats' `|+|` that doesn't require type ascription
   * on enum cases.
   *
   * Example:
   * {{{
   *   val patch = (cell"A1" := "Title") ++ (cell"A1".styled(headerStyle))
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
   *     cell"A1" := "Title",
   *     cell"A1".styled(headerStyle),
   *     range"A1:B1".merge
   *   )
   * }}}
   */
  object PatchBatch:
    def apply(patches: Patch*): Patch = Patch.Batch(patches.toVector)

export syntax.*
