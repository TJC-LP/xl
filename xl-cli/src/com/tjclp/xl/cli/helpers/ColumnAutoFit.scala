package com.tjclp.xl.cli.helpers

import com.tjclp.xl.addressing.{ARef, Column, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.formula.eval.DependentRecalculation
import com.tjclp.xl.sheets.{Sheet, SheetEdits}
import com.tjclp.xl.workbooks.Workbook

/**
 * Forwarder onto `Sheet.autoFitWidth` (ADR-017 §2.12, W2.1: the GH-156 font-metric width — Calibri
 * 11 max-digit-width convention, char-count heuristic when fonts are unavailable — lives in xl-core
 * now), shared by `col --auto-fit`, `autofit` and batch `autofit` so the three cannot drift.
 */
object ColumnAutoFit:

  /** The width `col` needs for its formatted content, in Excel character units. */
  def calculateWidth(sheet: Sheet, col: Column): Double = sheet.autoFitWidth(col)

  /**
   * GH-613: the sheet to MEASURE — `sheet` with every uncached formula cell in `columns` evaluated
   * against `wb` and given its value as a cache, where evaluation succeeds. The core fits a column
   * to what it displays and an uncached formula displays nothing yet; the CLI has the evaluator, so
   * a fit that follows `putf` in one batch (the recalculation runs after the ops) sizes to the
   * values the formulas will show. The evaluation is ONE dependency-ordered pass over the uncached
   * cells and their cone (`DependentRecalculation.recalculateAfterEdit`, the machinery every write
   * ends with), so a chain of formulas evaluates each precedent once. The result is for measuring
   * only: the caller sets widths on its own sheet, and the end-of-run recalculation writes the real
   * caches.
   */
  def withEvaluatedCaches(sheet: Sheet, wb: Workbook, columns: Iterable[Column]): Sheet =
    val wanted = columns.toSet
    val uncached: Set[ARef] = sheet.cells.iterator.collect {
      case cell if wanted.contains(cell.ref.col) && isUncachedFormula(cell.value) => cell.ref
    }.toSet
    if uncached.isEmpty then sheet
    else
      DependentRecalculation
        .recalculateAfterEdit(wb.put(sheet), sheet.name, uncached, Clock.system)
        .workbook(sheet.name)
        .getOrElse(sheet)

  /** A formula with no cache that evaluation can supply one for (a data-table record cannot). */
  private def isUncachedFormula(value: CellValue): Boolean = value match
    case CellValue.Formula(_, None, _: FormulaKind.DataTable) => false
    case CellValue.Formula(_, None, _) => true
    case _ => false

  /**
   * GH-869: widen, as Excel does on entry, every never-sized column where a written cell's
   * formatted number would not fit and so would show `#####` (`SheetEdits.widenForEntries`). Run on
   * the final, recalculated workbook so a `putf` cell is measured by the value it will display.
   * Returns the workbook and one `(sheet, column, width)` per column widened.
   */
  def widenForWrites(
    wb: Workbook,
    written: Iterable[(SheetName, ARef)]
  ): (Workbook, Vector[(SheetName, Column, Double)]) =
    // grouped by the tab's own name, so `data` and `Data` are measured as the one sheet they are
    val byTab = written.groupMap((name, _) => wb(name).fold(_ => name, _.name))(_._2)
    byTab.toVector.sortBy(_._1.value).foldLeft((wb, Vector.empty)) {
      case ((acc, notes), (name, refs)) =>
        acc(name).toOption match
          case None => (acc, notes)
          case Some(sheet) =>
            val (widened, cols) = SheetEdits.widenForEntries(sheet, refs)
            if cols.isEmpty then (acc, notes)
            else (acc.put(widened), notes ++ cols.map((c, w) => (name, c, w)))
    }

  /** The summary line for the columns [[widenForWrites]] widened; empty when there were none. */
  def widenedNote(widened: Vector[(SheetName, Column, Double)]): String =
    if widened.isEmpty then ""
    else
      val cols = widened.map((_, c, w) => f"${c.toLetter} to $w%.2f").mkString(", ")
      s"\nWidened column(s) $cols so formatted values fit (Excel would show #####)"
