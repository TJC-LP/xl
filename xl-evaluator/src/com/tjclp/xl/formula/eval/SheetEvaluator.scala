package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.Clock

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CodecError
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.workbooks.Workbook
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
     * @param workbook
     *   Pass `Some(wb)` iff the formula references other sheets (`='Other'!A1`); intra-sheet
     *   formulas don't need it. Or use `wb.evaluateFormula(formula, onSheet)` which wires the
     *   context automatically.
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
    def evaluateFormula(
      formula: String,
      clock: Clock = Clock.system,
      workbook: Option[Workbook] = None,
      currentCell: Option[ARef] = None
    ): XLResult[CellValue] =
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
          .eval(expr, sheet, clock, workbook, currentCell)
          .left
          .map(evalError => evalErrorToXLError(evalError, Some(formula)))

        // Convert typed result to CellValue
        cellValue = toCellValue(result)
      yield cellValue

    /**
     * Evaluate cell at ref (if it contains formula).
     *
     * If cell contains CellValue.Formula, parses and evaluates it against the current sheet state.
     * Formula dependencies that are themselves formulas will be referenced as-is (unevaluated).
     *
     * For formulas with dependencies on other formulas, use evaluateWithDependencyCheck() instead,
     * which evaluates all dependencies in correct order.
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
     * // A1 contains "=B1+C1" where B1=10, C1=20
     * sheet.evaluateCell(ref"A1") // Right(CellValue.Number(30))
     *
     * // D1 contains plain number
     * sheet.evaluateCell(ref"D1") // Right(CellValue.Number(42)) - unchanged
     * }}}
     */
    def evaluateCell(
      ref: ARef,
      clock: Clock = Clock.system,
      workbook: Option[Workbook] = None
    ): XLResult[CellValue] =
      val cell = sheet(ref)
      cell.value match
        case CellValue.Formula(expr, _) =>
          // Pass the current cell ref for ROW()/COLUMN() without arguments
          evaluateFormula(expr, clock, workbook, Some(ref))
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
    def evaluateWithDependencyCheck(
      clock: Clock = Clock.system,
      workbook: Option[Workbook] = None
    ): XLResult[Map[ARef, CellValue]] =
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
              // Evaluate in dependency order, threading the partially evaluated sheet so
              // dependent formulas see previously computed values. Fail-fast on first error.
              val evalResult = evalOrder.foldLeft[XLResult[(Sheet, Map[ARef, CellValue])]](
                scala.util.Right((sheet, Map.empty))
              ) {
                case (scala.util.Right((tempSheet, results)), ref) =>
                  tempSheet.evaluateCell(ref, clock, workbook) match
                    case scala.util.Right(value) =>
                      scala.util.Right((tempSheet.put(ref, value), results + (ref -> value)))
                    case scala.util.Left(error) =>
                      scala.util.Left(error)
                case (left, _) => left
              }

              evalResult.map(_._2)

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
    def evaluateAllFormulas(
      clock: Clock = Clock.system,
      workbook: Option[Workbook] = None
    ): XLResult[Map[ARef, CellValue]] =
      // For now, delegate to the safe method
      // In the future, we may optimize this to skip dependency checking for known-safe cases
      evaluateWithDependencyCheck(clock, workbook)

    /**
     * Evaluate formula cells within a specific range, plus their transitive dependencies.
     *
     * This is an optimized evaluation method for viewing a subset of a sheet. Instead of evaluating
     * all formula cells, it only evaluates:
     *   1. Formula cells within the specified range
     *   2. Formula cells that those formulas depend on (transitively)
     *
     * This is much more efficient than evaluateWithDependencyCheck when viewing a small range from
     * a large sheet with many formulas.
     *
     * @param range
     *   The cell range to evaluate
     * @param clock
     *   Clock for date/time functions (defaults to system clock)
     * @param workbook
     *   Optional workbook context for cross-sheet formula references
     * @return
     *   Either XLError or map of ref → evaluated value for cells within the range
     *
     * Example:
     * {{{
     * // Sheet with formulas throughout, but we only need A1:C10
     * sheet.evaluateForRange(CellRange.parse("A1:C10").toOption.get)
     * // Only evaluates formulas in A1:C10 + their dependencies
     * }}}
     */
    def evaluateForRange(
      range: CellRange,
      clock: Clock = Clock.system,
      workbook: Option[Workbook] = None
    ): XLResult[Map[ARef, CellValue]] =
      // 1. Find formula cells within the range (using pattern match, not isInstanceOf)
      val rangeFormulaCells = sheet.cells
        .collect {
          case (ref, cell) if range.contains(ref) =>
            cell.value match
              case _: CellValue.Formula => Some(ref)
              case _ => None
        }
        .flatten
        .toSet

      // If no formulas in range, return empty (nothing to evaluate)
      if rangeFormulaCells.isEmpty then scala.util.Right(Map.empty)
      else
        // 2. Build full dependency graph
        val graph = DependencyGraph.fromSheet(sheet)

        // 3. Find all transitive dependencies (cells that range formulas depend on)
        val transitiveDeps = DependencyGraph.transitiveDependencies(graph, rangeFormulaCells)

        // 4. Target cells = range formulas + (transitive deps that are also formulas)
        val allFormulaCells = graph.dependencies.keySet
        val targetCells = rangeFormulaCells ++ (transitiveDeps & allFormulaCells)

        // 5. Check for cycles in the subgraph (could still have cycles if dependencies have cycles)
        // We use the full graph for cycle detection since dependencies may form cycles outside the range
        DependencyGraph.detectCycles(graph) match
          case scala.util.Left(circularRef) =>
            scala.util.Left(evalErrorToXLError(circularRef, None))
          case scala.util.Right(_) =>
            // 6. Get topological order from full graph, but filter to only our target cells
            DependencyGraph.topologicalSort(graph) match
              case scala.util.Left(circularRef) =>
                scala.util.Left(evalErrorToXLError(circularRef, None))
              case scala.util.Right(fullEvalOrder) =>
                // Filter to only include cells we need to evaluate
                val evalOrder = fullEvalOrder.filter(targetCells.contains)

                // 7. Evaluate in dependency order, threading the partially evaluated sheet.
                // Fail-fast on first error; only cells in the original range are reported.
                val evalResult = evalOrder.foldLeft[XLResult[(Sheet, Map[ARef, CellValue])]](
                  scala.util.Right((sheet, Map.empty))
                ) {
                  case (scala.util.Right((tempSheet, results)), ref) =>
                    tempSheet.evaluateCell(ref, clock, workbook) match
                      case scala.util.Right(value) =>
                        val nextResults =
                          if rangeFormulaCells.contains(ref) then results + (ref -> value)
                          else results
                        scala.util.Right((tempSheet.put(ref, value), nextResults))
                      case scala.util.Left(error) =>
                        scala.util.Left(error)
                  case (left, _) => left
                }

                evalResult.map(_._2)

    /**
     * Evaluate an array formula and spill results into adjacent cells.
     *
     * Array formulas like TRANSPOSE return multiple values that "spill" into adjacent cells
     * starting from the origin cell. This method:
     *   1. Parses and evaluates the formula
     *   2. If result is ArrayResult, applies PutArray patch to spill values
     *   3. If result is single value, puts it at origin only
     *   4. Returns the updated sheet and the range of cells affected
     *
     * @param formula
     *   Excel formula string (with or without leading =)
     * @param originRef
     *   The cell where the array formula is entered (spill starts here)
     * @param clock
     *   Clock for date/time functions (defaults to system clock)
     * @param workbook
     *   Optional workbook context for cross-sheet formula references
     * @return
     *   Either XLError or tuple of (updated sheet, affected range)
     *
     * Example:
     * {{{
     * // TRANSPOSE a 2x3 range into a 3x2 result
     * sheet.evaluateArrayFormula("=TRANSPOSE(A1:C2)", ref"E1")
     * // Returns (updatedSheet, CellRange(E1:F3)) with transposed values
     * }}}
     */
    def evaluateArrayFormula(
      formula: String,
      originRef: ARef,
      clock: Clock = Clock.system,
      workbook: Option[Workbook] = None
    ): XLResult[(Sheet, CellRange)] =
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
        result <- Evaluator.arrayInstance
          .eval(expr, sheet, clock, workbook, Some(originRef))
          .left
          .map(evalError => evalErrorToXLError(evalError, Some(formula)))

        // Handle array vs scalar result
        updated <- result match
          case ar: ArrayResult =>
            val patch = Patch.PutArray(originRef, ar.values)
            val endRef = originRef.shift(ar.cols - 1, ar.rows - 1)
            scala.util.Right((Patch.applyPatch(sheet, patch), CellRange(originRef, endRef)))
          case other =>
            val cv = toCellValue(other)
            scala.util.Right((sheet.put(originRef, cv), CellRange(originRef, originRef)))
      yield updated

    /**
     * Put a formula at ref, inheriting the number format of its referenced cells (GH-184).
     *
     * Excel applies the referenced cells' format automatically when a formula is entered into a
     * General cell (`=B2-B3` over currency cells displays as `$400,000.00`); plain `put` keeps
     * formula cells at General. This opt-in variant parses the formula, infers a format via
     * [[FormulaFormatting.inferFormatFromReferences]], and — matching Excel — applies it only when
     * the target cell's current format is General (an explicit target format is preserved, as are
     * its other style properties).
     *
     * Note: explicit overloads instead of a default workbook parameter — extension methods with
     * default arguments crash the compiler when merged through the formulaExports wildcard export
     * (see the DependentRecalculation note in exports.scala).
     *
     * @param ref
     *   Target cell for the formula
     * @param formula
     *   Excel formula string (with or without leading =)
     * @return
     *   Left on parse failure; Right(updated sheet) otherwise (inference itself never fails —
     *   unresolvable or unformatted references simply contribute nothing)
     */
    def putFormulaInheriting(ref: ARef, formula: String): XLResult[Sheet] =
      putFormulaInheritingImpl(sheet, ref, formula, None)

    /**
     * Workbook-aware [[putFormulaInheriting]]: cross-sheet references (`=Data!B2`) resolve their
     * formats against the given workbook.
     */
    @annotation.targetName("putFormulaInheritingWorkbook")
    def putFormulaInheriting(ref: ARef, formula: String, workbook: Workbook): XLResult[Sheet] =
      putFormulaInheritingImpl(sheet, ref, formula, Some(workbook))

  // ========== Helper Functions ==========

  private def putFormulaInheritingImpl(
    sheet: Sheet,
    ref: ARef,
    formula: String,
    workbook: Option[Workbook]
  ): XLResult[Sheet] =
    FormulaParser
      .parse(formula)
      .left
      .map(parseError => XLError.FormulaError(formula, s"Parse error: $parseError"))
      .map { expr =>
        val normalized = if formula.startsWith("=") then formula else s"=$formula"
        val withFormula = sheet.put(ref, CellValue.Formula(normalized))
        val currentStyle =
          withFormula.cells.get(ref).flatMap(_.styleId).flatMap(withFormula.styleRegistry.get)
        // Excel parity: inherit only into a General-formatted target (formats are inferred
        // against the pre-put sheet so a self-reference can't observe the formula being written)
        FormulaFormatting.inferFormatFromReferences(expr, sheet, workbook) match
          case Some(fmt) if currentStyle.forall(_.numFmt == NumFmt.General) =>
            import com.tjclp.xl.sheets.styleSyntax.withCellStyle
            withFormula.withCellStyle(
              ref,
              currentStyle.getOrElse(CellStyle.default).withNumFmt(fmt)
            )
          case _ => withFormula
      }

  /**
   * Convert typed TExpr evaluation result to CellValue.
   *
   * Handles all result types: BigDecimal, String, Boolean, Int, LocalDate, LocalDateTime.
   */
  private def toCellValue(result: Any): CellValue =
    result match
      case cv: CellValue => cv // IFERROR and similar functions return CellValue directly
      case bd: BigDecimal => CellValue.Number(bd)
      case s: String => CellValue.Text(s)
      case b: Boolean => CellValue.Bool(b)
      case i: Int => CellValue.Number(BigDecimal(i))
      case ld: LocalDate => CellValue.DateTime(ld.atStartOfDay())
      case ldt: LocalDateTime => CellValue.DateTime(ldt)
      case ar: ArrayResult =>
        // For non-array formula contexts, return top-left value (Excel behavior)
        if ar.isEmpty then CellValue.Empty
        else ar(0, 0)
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
