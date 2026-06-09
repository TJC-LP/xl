package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError
import com.tjclp.xl.workbooks.Workbook

/**
 * A formula cell that failed to evaluate during `Workbook.recalculate`, with its location and
 * structured error. The cell is left uncached — Excel recalculates it on open.
 */
final case class CellEvalError(
  sheet: SheetName,
  ref: ARef,
  error: XLError
) derives CanEqual:
  /** Human-readable one-liner: `Sales!B2: Formula error in '=A1/0': ...` */
  def render: String = s"${sheet.value}!${ref.toA1}: ${error.message}"

/**
 * Result of a total, whole-workbook recalculation (`wb.recalculate()`).
 *
 * Evaluation is per-cell: a failing or cyclic formula is collected into `errors` and left uncached,
 * while every other formula — including those on the same sheet — still evaluates and caches.
 * Cross-sheet references are resolved against the workbook automatically.
 *
 * @param workbook
 *   The workbook with every successfully evaluated formula cached (`Formula(expr, Some(value))`)
 * @param evaluated
 *   Computed values per sheet for inspection without re-reading cells
 * @param errors
 *   Per-cell failures: parse/eval errors, cycle participants, and cells blocked by a cycle
 */
final case class RecalcResult(
  workbook: Workbook,
  evaluated: Map[SheetName, Map[ARef, CellValue]],
  errors: Vector[CellEvalError]
) derives CanEqual:

  /** True when every formula in the workbook evaluated successfully. */
  def isClean: Boolean = errors.isEmpty

  /**
   * Right(workbook) when clean, Left(errors) otherwise — for scripts that must not proceed on
   * partial results.
   */
  def toEither: Either[Vector[CellEvalError], Workbook] =
    if isClean then Right(workbook) else Left(errors)
