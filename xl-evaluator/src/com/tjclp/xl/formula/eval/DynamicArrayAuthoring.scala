package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-714: author an Excel 365 dynamic-array formula — what typing `=SORT(B1:B3)` into a cell of
 * Excel 365 stores.
 *
 * The anchor holds the formula as a dynamic-array record whose `ref` is the spill extent, cached at
 * the array's top-left value; every other cell of the extent holds its element as a plain constant
 * (Excel's own on-disk shape). The extent comes from evaluating the formula as an array at the
 * anchor — a scalar result is a 1x1 record, which is how Excel stores `=SUM(A1:A10*B1:B10)` typed
 * into one cell (no implicit intersection).
 *
 * A spill Excel would block (`#SPILL!`) is refused and nothing is written: xl does not model a
 * blocked spill as a value, and a later recalculation would overwrite a written `#SPILL!` with the
 * array's top-left value. Pure and total.
 */
private[xl] object DynamicArrayAuthoring:

  /**
   * The authored sheet, the record's extent, every cell whose value the authoring changed (the old
   * extent it cleared plus the new one — the seeds of the dependent refresh), and whether the
   * anchor carries a cache (false when the formula could not be evaluated: it is then written as an
   * uncached 1x1 record, like any putf whose evaluation fails).
   */
  final case class Authored(sheet: Sheet, extent: CellRange, touched: Set[ARef], cached: Boolean)

  private val maxListed = 10

  def author(
    wb: Workbook,
    sheet: Sheet,
    anchor: ARef,
    formula: String,
    clock: Clock
  ): XLResult[Authored] =
    val text = CellValue.canonicalFormulaText(formula)
    val full = s"=$text"
    SheetEvaluator.parseFormula(full).flatMap { _ =>
      val (cleared, clearedRefs) = clearOwnSpill(sheet, anchor)
      val single = CellRange(anchor, anchor)
      val placed =
        cleared.put(anchor, CellValue.Formula(text, None, FormulaKind.dynamicArray(single)))
      SheetEvaluator.evaluateArrayValues(
        placed,
        full,
        anchor,
        Evaluator.instance,
        clock,
        Some(wb.put(placed))
      ) match
        // a host failure (not an error value): an uncached 1x1 record, reported by the refresh
        case Left(_) => Right(Authored(placed, single, clearedRefs + anchor, cached = false))
        case Right(grid0) =>
          val grid =
            if grid0.isEmpty || grid0.exists(_.isEmpty) then
              Vector(Vector(CellValue.Error(CellError.Calc)))
            else grid0
          val rows = grid.size
          val cols = grid.map(_.size).maxOption.getOrElse(1)
          spillExtent(anchor, rows, cols).left.map(XLError.FormulaError(full, _)).flatMap {
            extent =>
              blocked(placed, anchor, extent)
                .map(reason => XLError.FormulaError(full, reason)) match
                case Some(err) => Left(err)
                case None =>
                  val kind = FormulaKind.dynamicArray(extent)
                  val written = grid.iterator.zipWithIndex.foldLeft(placed) { case (s, (row, r)) =>
                    row.iterator.zipWithIndex.foldLeft(s) { case (acc, (v, c)) =>
                      if r == 0 && c == 0 then
                        acc.put(anchor, CellValue.Formula(text, Some(v), kind))
                      else acc.put(anchor.shift(c, r), v)
                    }
                  }
                  Right(Authored(written, extent, clearedRefs ++ extent.cells.toSet, cached = true))
          }
    }

  /**
   * Re-authoring a dynamic anchor first clears the cells its previous spill filled (values only;
   * styles stay), so a shrinking spill leaves no stale children. A legacy (CSE) extent is left
   * alone: its cells are the user's to clear.
   */
  private def clearOwnSpill(sheet: Sheet, anchor: ARef): (Sheet, Set[ARef]) =
    sheet.cells.get(anchor).map(_.value) match
      case Some(CellValue.Formula(_, _, arr: FormulaKind.ArrayFormula))
          if arr.isDynamicArray && arr.ref.start == anchor =>
        val children = arr.ref.cells
          .filter(_ != anchor)
          .filter { r =>
            sheet.cells.get(r).exists(_.value != CellValue.Empty)
          }
          .toSet
        val cells = children.foldLeft(sheet.cells) { (acc, r) =>
          acc.get(r) match
            case Some(cell) if cell.styleId.isDefined || cell.hyperlink.isDefined =>
              acc.updated(r, cell.withValue(CellValue.Empty))
            case Some(_) => acc.removed(r)
            case None => acc
        }
        (sheet.copy(cells = cells), children)
      case _ => (sheet, Set.empty)

  private def spillExtent(anchor: ARef, rows: Int, cols: Int): Either[String, CellRange] =
    val lastCol = anchor.col.index0.toLong + cols - 1
    val lastRow = anchor.row.index0.toLong + rows - 1
    if lastCol > Column.MaxIndex0 || lastRow > Row.MaxIndex0 then
      Left(
        s"a ${rows}x$cols spill from ${anchor.toA1} runs past the sheet's last cell " +
          "XFD1048576 (Excel shows #SPILL!); choose another anchor"
      )
    else Right(CellRange(anchor, ARef.from0(lastCol.toInt, lastRow.toInt)))

  /**
   * Why Excel would show `#SPILL!` for `extent`, in Excel's own order of checks; None when clear.
   */
  private def blocked(sheet: Sheet, anchor: ARef, extent: CellRange): Option[String] =
    val span = extent.toA1
    val occupied = extent.cellsRowMajor
      .filter(_ != anchor)
      .filter(r => sheet.cells.get(r).exists(_.value != CellValue.Empty))
      .toVector
    lazy val merges = sheet.mergedRanges.filter(_.intersects(extent)).toVector.sortBy(rowMajor)
    lazy val arrays = sheet.cells.valuesIterator
      .filter(_.ref != anchor)
      .flatMap(cell =>
        cell.value match
          case CellValue.Formula(_, _, arr: FormulaKind.ArrayFormula)
              if arr.ref.intersects(extent) =>
            Some(cell.ref -> arr.ref)
          case _ => None
      )
      .toVector
      .sortBy((ref, _) => (ref.row.index0, ref.col.index0))
    lazy val tables = sheet.tables.values.filter(_.range.intersects(extent)).toVector.sortBy(_.name)
    if occupied.nonEmpty then
      val them = if occupied.size == 1 then "it" else "them"
      Some(
        s"spill range $span is blocked by ${listed(occupied.map(_.toA1))} " +
          s"(Excel shows #SPILL!); clear $them or choose another anchor"
      )
    else if merges.nonEmpty then
      Some(
        s"spill range $span is blocked by merged cells ${listed(merges.map(_.toA1))} " +
          "(Excel shows #SPILL!); unmerge them or choose another anchor"
      )
    else if arrays.nonEmpty then
      val named = arrays.map((ref, r) => s"${ref.toA1} (${r.toA1})")
      Some(
        s"spill range $span is blocked by the array formula at ${listed(named)} " +
          "(Excel shows #SPILL!); choose another anchor"
      )
    else if tables.nonEmpty then
      Some(
        s"spill range $span overlaps table ${listed(tables.map(t => s"${t.name} (${t.range.toA1})"))}; " +
          "a dynamic array cannot spill into a table (Excel shows #SPILL!); choose another anchor"
      )
    else None

  private def rowMajor(r: CellRange): (Int, Int) = (r.start.row.index0, r.start.col.index0)

  /** At most ten names, then `… (+N more)`. */
  private def listed(names: Vector[String]): String =
    val shown = names.take(maxListed).mkString(", ")
    if names.size > maxListed then s"$shown, … (+${names.size - maxListed} more)" else shown
