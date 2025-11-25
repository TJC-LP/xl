package com.tjclp.xl.formula.display

import com.tjclp.xl.display.{FormulaDisplayStrategy, LowPriorityFormulaDisplay, NumFmtFormatter}
import com.tjclp.xl.formula.{Clock, SheetEvaluator}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Formula display strategy with automatic evaluation.
 *
 * When xl-evaluator is imported via `import com.tjclp.xl.{*, given}`, formula cells are
 * automatically evaluated and formatted for display. The `evaluating` given has higher priority
 * than the `default` strategy due to inheritance depth (LowPriority pattern).
 *
 * {{{
 * import com.tjclp.xl.{*, given}  // Brings in evaluating strategy automatically!
 *
 * given Sheet = mySheet
 * println(excel"Total: \${ref"B1"}")  // Shows "$1,000,000" (evaluated!)
 * }}}
 *
 * Without xl-evaluator, formulas display as raw text: `"=SUM(A1:A10)"`.
 *
 * @since 0.2.0
 */
object EvaluatingFormulaDisplay extends LowPriorityFormulaDisplay:

  /**
   * Evaluating formula display strategy.
   *
   * Parses and evaluates formulas, then formats the result according to cell NumFmt. Falls back to
   * raw formula text if evaluation fails.
   *
   * Uses Clock.system for date/time functions.
   */
  given evaluating: FormulaDisplayStrategy = withClock(Clock.system)

  /**
   * Create an evaluating strategy with a custom clock.
   *
   * Useful for testing date/time functions with deterministic results:
   * {{{
   * val clock = Clock.fixedDate(LocalDate.of(2025, 1, 15))
   * given FormulaDisplayStrategy = EvaluatingFormulaDisplay.withClock(clock)
   * }}}
   *
   * @param clock
   *   The clock to use for date/time functions (TODAY, NOW, etc.)
   * @return
   *   A FormulaDisplayStrategy that evaluates formulas using the given clock
   */
  def withClock(clock: Clock): FormulaDisplayStrategy = new FormulaDisplayStrategy:
    def format(formula: String, sheet: Sheet): String =
      import SheetEvaluator.*

      // Ensure formula has "=" prefix for evaluator
      val formulaWithEquals = if formula.startsWith("=") then formula else s"=$formula"

      // Evaluate the formula using provided clock
      sheet.evaluateFormula(formulaWithEquals, clock) match
        case Right(result) =>
          // Successfully evaluated - format the result
          // Try to infer NumFmt from result type if cell has no explicit format
          val inferredFormat = inferFormatFromValue(result)
          NumFmtFormatter.formatValue(result, inferredFormat)

        case Left(error) =>
          // Evaluation failed - show raw formula as fallback
          formulaWithEquals

  /**
   * Infer appropriate NumFmt from a computed value.
   *
   * This provides reasonable defaults when a formula result doesn't have explicit formatting.
   *
   * @param value
   *   The computed cell value
   * @return
   *   Inferred number format
   */
  private def inferFormatFromValue(value: com.tjclp.xl.cells.CellValue): NumFmt =
    import com.tjclp.xl.cells.CellValue
    value match
      case CellValue.DateTime(_) => NumFmt.DateTime
      case CellValue.Number(n) =>
        // Infer percent format for values between 0 and 1 (heuristic)
        if n >= 0 && n <= 1 && n != n.setScale(0, BigDecimal.RoundingMode.HALF_UP) then
          NumFmt.PercentDecimal
        else NumFmt.General
      case _ => NumFmt.General
