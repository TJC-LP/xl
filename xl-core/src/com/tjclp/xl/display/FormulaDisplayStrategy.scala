package com.tjclp.xl.display

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Type class for displaying formula cells.
 *
 * This provides capability-based formula display:
 *   - '''xl-core only''': Shows raw formula text `"=SUM(A1:A10)"`
 *   - '''xl-core + xl-evaluator''': Evaluates and formats `"$1,000,000"`
 *
 * The active strategy is determined by which given instances are in scope.
 *
 * @since 0.2.0
 */
trait FormulaDisplayStrategy:
  /**
   * Format a formula cell for display.
   *
   * @param formula
   *   The formula expression (without leading =)
   * @param sheet
   *   The sheet context for evaluation (if evaluating)
   * @return
   *   Formatted display string
   */
  def format(formula: String, sheet: Sheet): String

  /**
   * Format a formula cell for display, with access to its cached value and number format.
   *
   * Called by the display layer with the full formula-cell context: the cached value (populated by
   * `Workbook.recalculate()` / read from disk) and the cell's number format. The default
   * implementation ignores both and delegates to [[format]]; strategies may override to prefer the
   * cached value (see EvaluatingFormulaDisplay in xl-evaluator, GH-275).
   *
   * @param formula
   *   The formula expression
   * @param cached
   *   The cached computed value carried by the formula cell, if any
   * @param numFmt
   *   The cell's number format (from its style; General when unstyled)
   * @param sheet
   *   The sheet context for evaluation (if evaluating)
   * @return
   *   Formatted display string
   */
  def formatCached(
    formula: String,
    cached: Option[CellValue],
    numFmt: NumFmt,
    sheet: Sheet
  ): String = format(formula, sheet)

/**
 * Low-priority given instances for FormulaDisplayStrategy.
 *
 * This trait provides the default (non-evaluating) strategy at low priority. Objects that extend
 * this trait can provide higher-priority givens that override the default.
 *
 * This pattern (used by Cats, Scalaz) resolves ambiguity via inheritance depth: givens defined in
 * more-specific objects win over givens in parent traits.
 */
trait LowPriorityFormulaDisplay:
  /**
   * Default strategy: Show the cached value when present, raw formula text otherwise.
   *
   * This is the fallback when xl-evaluator is not imported. It never evaluates, but it does prefer
   * the cached value carried by `CellValue.Formula` (populated by `Workbook.recalculate()` or read
   * from disk), formatted via the cell's number format — files Excel or xl already evaluated thus
   * display meaningfully for xl-core-only consumers (GH-282, mirroring the evaluating strategy's
   * cache preference from GH-275). Uncached formulas display as `"=SUM(A1:A10)"`.
   */
  given default: FormulaDisplayStrategy with
    def format(formula: String, sheet: Sheet): String =
      // Formula expression already includes "=" prefix
      if formula.startsWith("=") then formula else s"=$formula"

    override def formatCached(
      formula: String,
      cached: Option[CellValue],
      numFmt: NumFmt,
      sheet: Sheet
    ): String =
      cached match
        case Some(value) => NumFmtFormatter.formatValue(value, numFmt)
        case None => format(formula, sheet)

object FormulaDisplayStrategy extends LowPriorityFormulaDisplay
