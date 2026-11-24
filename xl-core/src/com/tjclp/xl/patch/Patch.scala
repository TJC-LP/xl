package com.tjclp.xl.patch

import cats.Monoid
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue, Comment}
import com.tjclp.xl.cf.CfRule
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.error.XLResult
import com.tjclp.xl.ops.{Edit, PatchLowering}
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties, Sheet}
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.border.Border
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
 *   - Idempotence: Applying the same patch twice yields the same result (SetConditionalFormat
 *     excepted — it appends one CF block per application, matching Sheet.conditionalFormat's
 *     append-only authoring semantics)
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

  /**
   * Merge border sides into a cell's existing style (auto-registers in styleRegistry).
   *
   * Unlike [[SetCellStyle]], which replaces the whole style, this preserves the cell's existing
   * font, fill, number format, and alignment, and only overlays the border sides that are set
   * (non-none) in `border` — the CLI's additive border semantics (`--border-top` etc.). Sides left
   * as `BorderSide.none` keep the cell's existing sides.
   */
  case MergeBorder(ref: ARef, border: Border)

  /** Clear style for a cell */
  case ClearStyle(ref: ARef)

  /**
   * Merge cells in a range ([[com.tjclp.xl.sheets.Sheet.merge]], GH-935): a non-empty upper-left
   * cell keeps its value and the others are cleared like Clear Contents. Over an empty upper-left
   * cell the first non-empty cell is kept in place (the kernel cannot rewrite readers of a moved
   * cell); `Edit.Merge` performs Excel's move.
   */
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

  /**
   * Attach a comment to a cell (later comments to the same ref win). Comments live on
   * `Sheet.comments`, so no cell is materialized.
   */
  case SetComment(ref: ARef, comment: Comment)

  /**
   * Append ONE conditional-formatting block applying `rules` to `ranges` (multi-range sqref).
   *
   * Applied via [[com.tjclp.xl.sheets.Sheet.conditionalFormat]], which owns auto-priority stamping:
   * rules left at [[CfRule.AutoPriority]] are assigned above every priority already on the sheet.
   * Append-only like the Sheet method — applying the same patch twice appends two blocks.
   */
  case SetConditionalFormat(ranges: Vector[CellRange], rules: Vector[CfRule])

  /** Batch multiple patches together */
  case Batch(patches: Vector[Patch])

object Patch:

  /**
   * The edits that mean the same as `patch` on `sheet` (ADR-017 §2.12), or `None` when no edit
   * sequence reproduces it EXACTLY (a `ClearStyle` on a cell that does not exist, a `Remove` inside
   * a merged region, column/row properties beyond width/height/visibility). Desugar-coherence law
   * (EditLawsSpec): applying them under the sheet's scope equals `applyPatch(sheet, patch)`.
   */
  def toEdits(patch: Patch, sheet: Sheet): Option[Vector[Edit]] =
    PatchLowering.toEdits(patch, sheet)

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

    case MergeBorder(ref, border) =>
      // Overlay border sides onto the cell's existing style (font/fill/numFmt/align preserved)
      val existing = sheet.getCellStyle(ref).getOrElse(CellStyle.default)
      sheet.withCellStyle(ref, existing.withBorder(existing.border.merge(border)))

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
      // the array's references are distinct, so each reads the sheet as it was; one merge (#891)
      val written = values.iterator.zipWithIndex.flatMap { case (row, rowIdx) =>
        row.iterator.zipWithIndex.map { case (value, colIdx) =>
          val ref = origin.shift(colIdx, rowIdx)
          sheet.cells.get(ref).fold(Cell(ref, value))(_.withValue(value))
        }
      }
      sheet.copy(cells = sheet.cells.putAll(written))

    case SetComment(ref, cmt) =>
      sheet.comment(ref, cmt)

    case SetConditionalFormat(ranges, rules) =>
      // Route through Sheet.conditionalFormat — it owns AutoPriority stamping
      sheet.conditionalFormat(ranges, rules)

    case Batch(patches) =>
      patches.foldLeft(sheet) { (acc, p) =>
        applyPatch(acc, p)
      }

  /** Apply multiple patches in sequence */
  def applyPatches(sheet: Sheet, patches: Iterable[Patch]): Sheet =
    applyPatch(sheet, Batch(patches.toVector))
