package com.tjclp.xl.formula

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CodecError
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet
import java.time.{LocalDate, LocalDateTime}

/**
 * Extension methods for evaluating formulas in Excel sheets.
 *
 * Provides high-level API for formula evaluation:
 *   - Parse formula strings and evaluate against sheet
 *   - Evaluate formula cells (CellValue.Formula)
 *   - Bulk evaluation of all formulas in sheet
 *
 * Design principles:
 *   - Pure functional (no mutations, no side effects)
 *   - Total error handling (XLResult[A] = Either[XLError, A])
 *   - Clock parameter for deterministic date/time functions
 *   - Type conversion: TExpr results → CellValue
 *
 * Example:
 * {{{
 * import com.tjclp.xl.formula.{*, given}
 *
 * // Evaluate formula string
 * sheet.evaluateFormula("=SUM(A1:A10)") // XLResult[CellValue]
 *
 * // Evaluate cell with formula
 * sheet.evaluateCell(ref"B1") // XLResult[CellValue]
 *
 * // Evaluate all formulas
 * sheet.evaluateAllFormulas() // XLResult[Map[ARef, CellValue]]
 * }}}
 */
object SheetEvaluator:
  extension (sheet: Sheet)
    /**
     * Evaluate formula string against this sheet.
     *
     * Parses formula, evaluates against sheet, converts result to CellValue.
     *
     * @param formula
     *   Excel formula string (with or without leading =)
     * @param clock
     *   Clock for date/time functions (defaults to system clock)
     * @return
     *   Either XLError or evaluated CellValue
     *
     * Example:
     * {{{
     * sheet.evaluateFormula("=A1+B2") // Right(CellValue.Number(15))
     * sheet.evaluateFormula("=SUM(A1:A10)") // Right(CellValue.Number(55))
     * sheet.evaluateFormula("=TODAY()") // Right(CellValue.DateTime(...))
     * }}}
     */
    def evaluateFormula(formula: String, clock: Clock = Clock.system): XLResult[CellValue] =
      for
        // Parse formula string to TExpr AST
        expr <- FormulaParser
          .parse(formula)
          .left
          .map(parseError =>
            XLError.FormulaError(
              formula,
              s"Parse error: $parseError"
            )
          )

        // Evaluate TExpr against sheet
        result <- Evaluator.instance
          .eval(expr, sheet, clock)
          .left
          .map(evalError => evalErrorToXLError(evalError, Some(formula)))

        // Convert typed result to CellValue
        cellValue = toCellValue(result)
      yield cellValue

    /**
     * Evaluate cell at ref (if it contains formula).
     *
     * If cell contains CellValue.Formula, parses and evaluates it. Otherwise returns cell value
     * unchanged.
     *
     * @param ref
     *   Cell reference to evaluate
     * @param clock
     *   Clock for date/time functions (defaults to system clock)
     * @return
     *   Either XLError or evaluated CellValue
     *
     * Example:
     * {{{
     * // A1 contains "=B1+C1"
     * sheet.evaluateCell(ref"A1") // Right(CellValue.Number(result))
     *
     * // D1 contains plain number
     * sheet.evaluateCell(ref"D1") // Right(CellValue.Number(42)) - unchanged
     * }}}
     */
    def evaluateCell(ref: ARef, clock: Clock = Clock.system): XLResult[CellValue] =
      val cell = sheet(ref)
      cell.value match
        case CellValue.Formula(expr, _) =>
          evaluateFormula(expr, clock)
        case other =>
          scala.util.Right(other)

    /**
     * Evaluate all formula cells in sheet with dependency checking.
     *
     * This method:
     *   1. Builds dependency graph from all formula cells
     *   2. Detects circular references (fails fast if found)
     *   3. Performs topological sort to determine evaluation order
     *   4. Evaluates formulas in dependency order (dependencies before dependents)
     *
     * This is the safe, production-ready evaluation method that prevents infinite loops and ensures
     * correct evaluation order.
     *
     * @param clock
     *   Clock for date/time functions (defaults to system clock)
     * @return
     *   Either CircularRef error or map of ref → evaluated value
     *
     * Example:
     * {{{
     * // Sheet with formulas: A1="=10", B1="=A1*2", C1="=B1+5"
     * sheet.evaluateWithDependencyCheck() // Right(Map(A1 -> 10, B1 -> 20, C1 -> 25))
     *
     * // Sheet with cycle: A1="=B1", B1="=A1"
     * sheet.evaluateWithDependencyCheck() // Left(XLError.FormulaError(..., CircularRef))
     * }}}
     */
    @SuppressWarnings(Array("org.wartremover.warts.Var"))
    def evaluateWithDependencyCheck(clock: Clock = Clock.system): XLResult[Map[ARef, CellValue]] =
      // Suppression rationale: Mutable accumulator for building results map during iteration.
      // Functional fold alternative less clear for this use case.
      // Build dependency graph
      val graph = DependencyGraph.fromSheet(sheet)

      // Detect cycles first (fail fast)
      DependencyGraph.detectCycles(graph) match
        case scala.util.Left(circularRef) =>
          // Convert EvalError.CircularRef to XLError
          scala.util.Left(evalErrorToXLError(circularRef, None))
        case scala.util.Right(_) =>
          // No cycles, get evaluation order
          DependencyGraph.topologicalSort(graph) match
            case scala.util.Left(circularRef) =>
              // Topological sort found cycle (shouldn't happen after detectCycles passed)
              scala.util.Left(evalErrorToXLError(circularRef, None))
            case scala.util.Right(evalOrder) =>
              // Evaluate in dependency order
              // Create a mutable temp sheet to accumulate evaluated values
              var tempSheet = sheet
              var results = Map.empty[ARef, CellValue]

              val evalResult = evalOrder.foldLeft[XLResult[Unit]](scala.util.Right(())) {
                case (scala.util.Right(_), ref) =>
                  // Evaluate this cell against the temp sheet (which has previously evaluated values)
                  tempSheet.evaluateCell(ref, clock) match
                    case scala.util.Right(value) =>
                      // Update temp sheet with evaluated value (for dependent formulas to use)
                      // put(ref, CellValue) returns Sheet directly - it cannot fail
                      tempSheet = tempSheet.put(ref, value)
                      results = results + (ref -> value)
                      scala.util.Right(())
                    case scala.util.Left(error) =>
                      scala.util.Left(error)
                case (left, _) => left
              }

              evalResult.map(_ => results)

    /**
     * Evaluate all formula cells in sheet (unsafe, no cycle detection).
     *
     * Iterates through all cells, evaluates Formula cells, leaves others unchanged. This method
     * does NOT check for circular references and may result in stack overflow or incorrect results
     * if cycles exist.
     *
     * @deprecated
     *   Use evaluateWithDependencyCheck for production code. This method is kept for backwards
     *   compatibility and simple cases where circular references are known not to exist.
     * @param clock
     *   Clock for date/time functions (defaults to system clock)
     * @return
     *   Either first error encountered or map of ref → evaluated value
     *
     * Example:
     * {{{
     * // Sheet has formulas in A1, B1, C1 (no cycles)
     * sheet.evaluateAllFormulas() // Right(Map(A1 -> CellValue.Number(10), B1 -> ...))
     * }}}
     */
    def evaluateAllFormulas(clock: Clock = Clock.system): XLResult[Map[ARef, CellValue]] =
      // For now, delegate to the safe method
      // In the future, we may optimize this to skip dependency checking for known-safe cases
      evaluateWithDependencyCheck(clock)

  // ========== Helper Functions ==========

  /**
   * Convert typed TExpr evaluation result to CellValue.
   *
   * Handles all result types: BigDecimal, String, Boolean, Int, LocalDate, LocalDateTime.
   */
  private def toCellValue(result: Any): CellValue =
    result match
      case bd: BigDecimal => CellValue.Number(bd)
      case s: String => CellValue.Text(s)
      case b: Boolean => CellValue.Bool(b)
      case i: Int => CellValue.Number(BigDecimal(i))
      case ld: LocalDate => CellValue.DateTime(ld.atStartOfDay())
      case ldt: LocalDateTime => CellValue.DateTime(ldt)
      case other =>
        // Fallback for unexpected types (should never happen with well-typed TExpr)
        CellValue.Text(other.toString)

  /**
   * Convert EvalError to XLError for integration.
   *
   * @param error
   *   The evaluation error
   * @param formulaContext
   *   Optional formula string for context
   * @return
   *   XLError with detailed message
   */
  private def evalErrorToXLError(error: EvalError, formulaContext: Option[String]): XLError =
    val contextStr = formulaContext.map(f => s" in formula: $f").getOrElse("")

    error match
      case EvalError.DivByZero(num, denom) =>
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          s"Division by zero: $num / $denom$contextStr"
        )
      case EvalError.CodecFailed(ref, codecErr) =>
        val codecMsg = codecErr match
          case CodecError.TypeMismatch(expected, actual) =>
            s"Expected $expected, got $actual"
          case CodecError.ParseError(value, targetType, detail) =>
            s"Cannot parse '$value' as $targetType: $detail"
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          s"Type mismatch at $ref: $codecMsg$contextStr"
        )
      case EvalError.CircularRef(cycle) =>
        val cyclePath = cycle
          .map { ref =>
            // Format ARef to A1 notation
            // (Simplified - in production would use ARef.toA1 extension)
            s"cell_${ref}"
          }
          .mkString(" → ")
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          s"Circular reference detected: $cyclePath"
        )
      case EvalError.EvalFailed(reason, ctx) =>
        val fullContext = (ctx.toList ++ formulaContext.toList).mkString(", ")
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          s"Evaluation failed: $reason${if fullContext.nonEmpty then s" ($fullContext)" else ""}"
        )
      case other =>
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          s"Evaluation error: $other$contextStr"
        )
