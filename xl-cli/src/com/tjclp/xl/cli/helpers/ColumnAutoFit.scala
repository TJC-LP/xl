package com.tjclp.xl.cli.helpers

import com.tjclp.xl.addressing.Column
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.SheetEvaluator
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
   * values the formulas will show. The result is for measuring only: the caller sets widths on its
   * own sheet, and the end-of-run recalculation writes the real caches.
   */
  def withEvaluatedCaches(sheet: Sheet, wb: Workbook, columns: Iterable[Column]): Sheet =
    val wanted = columns.toSet
    val book = wb.put(sheet)
    sheet.cells.valuesIterator.filter(cell => wanted.contains(cell.ref.col)).foldLeft(sheet) {
      (measured, cell) =>
        cell.value match
          case CellValue.Formula(_, None, _: FormulaKind.DataTable) => measured
          case CellValue.Formula(expr, None, kind) =>
            SheetEvaluator
              .evaluateCell(sheet)(cell.ref, workbook = Some(book))
              .fold(
                _ => measured,
                value => measured.put(cell.ref, CellValue.Formula(expr, Some(value), kind))
              )
          case _ => measured
    }
