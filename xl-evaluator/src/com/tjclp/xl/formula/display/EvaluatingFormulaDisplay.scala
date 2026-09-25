package com.tjclp.xl.formula.display

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.display.{FormulaDisplayStrategy, LowPriorityFormulaDisplay, NumFmtFormatter}
import com.tjclp.xl.error.XLResult
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.formula.eval.SheetEvaluator
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
      val formulaWithEquals = withEquals(formula)

      // Evaluate the formula using provided clock. No cell position: it evaluates as a new
      // Excel 365 cell would, as a dynamic array (formatAt evaluates a real cell at its position)
      displayEvaluated(sheet.evaluateFormula(formulaWithEquals, clock), formulaWithEquals)

    override def formatCached(
      formula: String,
      cached: Option[CellValue],
      numFmt: NumFmt,
      sheet: Sheet
    ): String =
      // Prefer the cached value (populated by Workbook.recalculate() or read from
      // disk): sheet-local re-evaluation has no workbook context, so cross-sheet
      // formulas would degrade to raw text despite a correct cached value (GH-275).
      // Only uncached formulas fall back to local evaluation.
      cached match
        case Some(value) => displayCached(value, numFmt)
        case None => format(formula, sheet)

    override def formatAt(
      formula: CellValue.Formula,
      numFmt: NumFmt,
      sheet: Sheet,
      at: ARef
    ): String =
      formula.cachedValue match
        case Some(value) => displayCached(value, numFmt)
        case None =>
          import SheetEvaluator.*
          // An uncached cell evaluates as recalculate() would evaluate it: at its position (a
          // plain formula is Excel's legacy formula, its references implicitly intersected) and by
          // its kind (an array-formula record evaluates as an array and shows element (0,0)). A
          // cell the sheet does not hold (a detached Cell given to the interpolator) is placed at
          // its position first.
          val holder =
            if sheet.cells.get(at).exists(_.value == formula) then sheet
            else sheet.put(at, formula)
          displayEvaluated(holder.evaluateCell(at, clock), withEquals(formula.expression))

  /** The formula text with its leading `=`, the display of a formula that fails to evaluate. */
  private def withEquals(formula: String): String =
    if formula.startsWith("=") then formula else s"=$formula"

  /**
   * An evaluated formula formatted by the NumFmt its value suggests; a formula that fails to
   * evaluate shows its raw text.
   */
  private def displayEvaluated(result: XLResult[CellValue], rawText: String): String =
    result match
      case Right(value) => NumFmtFormatter.formatValue(value, inferFormatFromValue(value))
      case Left(_) => rawText

  /** A cached value formatted by the cell's NumFmt, or the one its value suggests under General. */
  private def displayCached(value: CellValue, numFmt: NumFmt): String =
    val effectiveFmt = numFmt match
      case NumFmt.General => inferFormatFromValue(value)
      case explicit => explicit
    NumFmtFormatter.formatValue(value, effectiveFmt)

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
