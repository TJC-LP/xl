package com.tjclp.xl.cli.helpers

import com.tjclp.xl.addressing.{ARef, Column}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.formula.eval.DependentRecalculation
import com.tjclp.xl.sheets.Sheet
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
    val uncached: Set[ARef] = sheet.cells.valuesIterator.collect {
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
