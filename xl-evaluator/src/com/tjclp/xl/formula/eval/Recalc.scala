package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError
import com.tjclp.xl.workbooks.{CalcPr, Workbook}

/**
 * A formula cell that failed to evaluate during `Workbook.recalculate`, with its location and
 * structured error. The cell is left uncached — Excel recalculates it on open.
 */
final case class CellEvalError(
  sheet: SheetName,
  ref: ARef,
  error: XLError
) derives CanEqual:
  /**
   * Human-readable one-liner: `Sales!B2: Formula error in '=A1/0': ...` — GH-280: cell-ref-shaped
   * sheet names quote (`'A1'!B2: ...`) so the location reads unambiguously.
   */
  def render: String = s"${SheetName.quoteForFormula(sheet.value)}!${ref.toA1}: ${error.message}"

/**
 * GH-373: opt-in bounded iterative calculation for circular workbooks.
 *
 * Professional models are routinely circular by design (interest on average debt balances) and ship
 * with `<calcPr iterate="1">`. Passing an `IterativeCalc` to `recalculate` fixpoints cycle members
 * with Jacobi iteration (every member reads PREVIOUS-iteration values, matching Excel): members
 * seed to 0, iterate until every member's |Δ| < `maxChange` or `maxIter` rounds, and
 * non-convergence keeps the last values with NO error — Excel's semantics, and the deliberate
 * inversion of the default path's circular-reference errors.
 *
 * Deliberately NOT auto-derived from `wb.metadata.calcPr` — iteration is opt-in so the default
 * `recalculate()` stays byte-identical. Bridge explicitly when honoring a file's settings:
 * {{{
 * wb.metadata.calcPr
 *   .filter(_.iterativeCalculation)
 *   .map(IterativeCalc.fromCalcPr)
 *   .fold(wb.recalculate())(it => wb.recalculate(it))
 * }}}
 *
 * @param maxIter
 *   Iteration cap (Excel UI default 100; values < 1 are treated as 1)
 * @param maxChange
 *   Convergence threshold: iteration stops when every cycle member changes by LESS than this
 *   between rounds (Excel UI default 0.001)
 */
final case class IterativeCalc(maxIter: Int, maxChange: BigDecimal) derives CanEqual

object IterativeCalc:
  /**
   * Lift a workbook's modeled `<calcPr>` into iteration settings, applying Excel's defaults (100,
   * 0.001) for absent attributes. Callers gate on `calcPr.iterativeCalculation` themselves — see
   * the bridge example on [[IterativeCalc]].
   */
  def fromCalcPr(cp: CalcPr): IterativeCalc =
    IterativeCalc(cp.maxIterations.getOrElse(100), cp.maxChange.getOrElse(BigDecimal("0.001")))

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
