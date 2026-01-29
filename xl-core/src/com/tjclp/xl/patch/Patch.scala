package com.tjclp.xl.patch

import cats.Monoid
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.error.XLResult
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties, Sheet}
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.units.StyleId

/**
 * Patch ADT for Sheet updates with monoid semantics.
 *
 * Patches can be composed using the Monoid instance, allowing batch operations to be built
 * declaratively.
 *
 * Laws:
 *   - Associativity: (p1 |+| p2) |+| p3 == p1 |+| (p2 |+| p3)
 *   - Identity: Patch.empty |+| p == p == p |+| Patch.empty
 *   - Idempotence: Applying the same patch twice yields the same result
 */
enum Patch:
  /** Put a cell value at a reference */
  case Put(ref: ARef, value: CellValue)

  /** Set style for a cell */
  case SetStyle(ref: ARef, styleId: StyleId)

  /** Set style for a cell using CellStyle object (auto-registers in styleRegistry) */
  case SetCellStyle(ref: ARef, style: CellStyle)

  /** Set style for a range of cells (auto-registers in styleRegistry) */
  case SetRangeStyle(range: CellRange, style: CellStyle)

  /** Clear style for a cell */
  case ClearStyle(ref: ARef)

  /** Merge cells in a range */
  case Merge(range: CellRange)

  /** Unmerge cells in a range */
  case Unmerge(range: CellRange)

  /** Set column properties */
  case SetColumnProperties(col: Column, props: ColumnProperties)

  /** Set row properties */
  case SetRowProperties(row: Row, props: RowProperties)

  /** Remove cell at reference */
  case Remove(ref: ARef)

  /** Remove all cells in range */
  case RemoveRange(range: CellRange)

  /** Put a grid of values starting at origin (for array formula spill) */
  case PutArray(origin: ARef, values: Vector[Vector[CellValue]])

  /** Batch multiple patches together */
  case Batch(patches: Vector[Patch])

object Patch:
  /** Empty patch (identity element) */
  val empty: Patch = Batch(Vector.empty)

  /**
   * Combine two patches into a batch.
   *
   * Flattens nested batches to maintain a flat structure. Later patches override earlier ones for
   * the same reference.
   */
  def combine(p1: Patch, p2: Patch): Patch = (p1, p2) match
    case (Batch(ps1), Batch(ps2)) => Batch(ps1 ++ ps2)
    case (Batch(ps1), p2) => Batch(ps1 :+ p2)
    case (p1, Batch(ps2)) => Batch(p1 +: ps2)
    case (p1, p2) => Batch(Vector(p1, p2))

  /** Monoid instance for Patch composition */
  given Monoid[Patch] with
    def empty: Patch = Patch.empty
    def combine(x: Patch, y: Patch): Patch = Patch.combine(x, y)

  /**
   * Apply a patch to a sheet, returning the modified sheet.
   *
   * Patches are applied left-to-right. Later patches override earlier ones for conflicting
   * operations (e.g., two Puts to the same reference).
   *
   * This operation is infallible - all patch types are pure transformations on validated data.
   *
   * @param sheet
   *   The sheet to modify
   * @param patch
   *   The patch to apply
   * @return
   *   The modified sheet
   */
  def applyPatch(sheet: Sheet, patch: Patch): Sheet = patch match
    case Put(ref, value) =>
      sheet.put(ref, value)

    case SetStyle(ref, styleId) =>
      val cell = sheet(ref).withStyle(styleId)
      sheet.put(cell)

    case SetCellStyle(ref, style) =>
      // Register style and apply to cell automatically
      sheet.withCellStyle(ref, style)

    case SetRangeStyle(range, style) =>
      // Register style and apply to all cells in range
      sheet.withRangeStyle(range, style)

    case ClearStyle(ref) =>
      val cell = sheet(ref).clearStyle
      sheet.put(cell)

    case Merge(range) =>
      sheet.merge(range)

    case Unmerge(range) =>
      sheet.unmerge(range)

    case SetColumnProperties(col, props) =>
      sheet.setColumnProperties(col, props)

    case SetRowProperties(row, props) =>
      sheet.setRowProperties(row, props)

    case Remove(ref) =>
      sheet.remove(ref)

    case RemoveRange(range) =>
      sheet.removeRange(range)

    case PutArray(origin, values) =>
      val newCells = values.zipWithIndex.foldLeft(sheet.cells) { case (cellsAcc, (row, rowIdx)) =>
        row.zipWithIndex.foldLeft(cellsAcc) { case (cellsAcc2, (value, colIdx)) =>
          val ref = origin.shift(colIdx, rowIdx)
          val cell = cellsAcc2.get(ref) match
            case Some(existing) => existing.withValue(value)
            case None => Cell(ref, value)
          cellsAcc2.updated(ref, cell)
        }
      }
      sheet.copy(cells = newCells)

    case Batch(patches) =>
      patches.foldLeft(sheet) { (acc, p) =>
        applyPatch(acc, p)
      }

  /** Apply multiple patches in sequence */
  def applyPatches(sheet: Sheet, patches: Iterable[Patch]): Sheet =
    applyPatch(sheet, Batch(patches.toVector))
