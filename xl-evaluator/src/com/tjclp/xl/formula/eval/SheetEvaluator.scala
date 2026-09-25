package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.{Clock, Rng}

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
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
      evaluateFormulaWith(sheet, formula, Evaluator.instance, clock, workbook, currentCell)

    /**
     * Evaluate a formula with an explicit randomness source (GH-115) — deterministic
     * RAND/RANDBETWEEN via `Rng.seeded(seed)`.
     *
     * Note: explicit overloads instead of a default rng parameter — extension methods with default
     * arguments crash the compiler when merged through the formulaExports wildcard export (see the
     * DependentRecalculation note in exports.scala).
     */
    @annotation.targetName("evaluateFormulaWithRng")
    def evaluateFormula(formula: String, clock: Clock, rng: Rng): XLResult[CellValue] =
      evaluateFormulaWith(sheet, formula, Evaluator.instance(rng), clock, None, None)

    /** Evaluate a formula with an explicit randomness source and workbook context (GH-115). */
    @annotation.targetName("evaluateFormulaWithRngWorkbook")
    def evaluateFormula(
      formula: String,
      clock: Clock,
      rng: Rng,
      workbook: Option[Workbook]
    ): XLResult[CellValue] =
      evaluateFormulaWith(sheet, formula, Evaluator.instance(rng), clock, workbook, None)

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
      evaluateCellWithEvaluator(sheet, ref, Evaluator.instance, clock, workbook)

    /** Evaluate a formula cell with an explicit randomness source (GH-115). */
    @annotation.targetName("evaluateCellWithRng")
    def evaluateCell(ref: ARef, clock: Clock, rng: Rng): XLResult[CellValue] =
      evaluateCell(ref, clock, rng, None)

    /** Evaluate a formula cell with explicit randomness and workbook context (GH-115). */
    @annotation.targetName("evaluateCellWithRngWorkbook")
    def evaluateCell(
      ref: ARef,
      clock: Clock,
      rng: Rng,
      workbook: Option[Workbook]
    ): XLResult[CellValue] =
      evaluateCellWithEvaluator(sheet, ref, Evaluator.instance(rng), clock, workbook)

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
      evaluateWithDependencyCheckImpl(sheet, clock, workbook, None)

    /** Dependency-ordered evaluation with an explicit randomness source (GH-115). */
    @annotation.targetName("evaluateWithDependencyCheckRng")
    def evaluateWithDependencyCheck(clock: Clock, rng: Rng): XLResult[Map[ARef, CellValue]] =
      evaluateWithDependencyCheckImpl(sheet, clock, None, Some(rng))

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
      rangePlan(sheet, Vector(range), workbook) match
        // If no formulas in range, return empty (nothing to evaluate)
        case None => scala.util.Right(Map.empty)
        case Some(RangePlan(rangeFormulaCells, graph, _, targetCells, dynamic)) =>
          // Fail fast on any cycle in the sheet: the full graph is checked, since dependencies may
          // form cycles outside the range (evaluateForRangePerCell narrows this to the closure)
          DependencyGraph.detectCycles(graph) match
            case scala.util.Left(circularRef) =>
              scala.util.Left(evalErrorToXLError(circularRef, None))
            case scala.util.Right(_) =>
              // Topological order from the full graph, filtered to the target cells
              DependencyGraph.topologicalSort(graph) match
                case scala.util.Left(circularRef) =>
                  scala.util.Left(evalErrorToXLError(circularRef, None))
                case scala.util.Right(fullEvalOrder) =>
                  // Filter to only include cells we need to evaluate
                  // GH-274: dynamic cells + their static dependents evaluate last (caches stripped)
                  val (evalOrder, initial) = deferDynamicWithStrip(
                    sheet,
                    graph,
                    fullEvalOrder.filter(targetCells.contains),
                    dynamic
                  )

                  // Evaluate in dependency order, threading the partially evaluated sheet.
                  // Fail-fast on first error; only cells in the original range are reported.
                  val evalResult = evalOrder.foldLeft[XLResult[(Sheet, Map[ARef, CellValue])]](
                    scala.util.Right((initial, Map.empty))
                  ) {
                    case (scala.util.Right((tempSheet, results)), ref) =>
                      val currentWorkbook = workbook.map(_.put(tempSheet))
                      tempSheet.evaluateCell(ref, clock, currentWorkbook) match
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
     * Evaluate the formulas of a range cell by cell — the total counterpart of
     * [[evaluateForRange]], and what `xl view --eval` renders.
     *
     * The same formulas evaluate (the range's formulas plus their transitive formula precedents, or
     * every formula when the sheet has dynamic references), in dependency order, but a formula that
     * cannot evaluate no longer fails the range: it is listed in [[RangeEvalResult.failures]] and
     * the formulas depending on it are listed in [[RangeEvalResult.blocked]], never evaluated (no
     * stale cache is mixed into a live value). A cycle fails its members — and blocks their
     * dependents — only when it lies inside that closure; a cycle elsewhere on the sheet no longer
     * affects the range. Never throws.
     *
     * Note: explicit overloads instead of default parameters — extension methods with default
     * arguments crash the compiler when merged through the formulaExports wildcard export.
     *
     * @param range
     *   the cell range whose formula values to report
     * @param clock
     *   Clock for date/time functions
     * @param workbook
     *   workbook context for cross-sheet references
     */
    def evaluateForRangePerCell(
      range: CellRange,
      clock: Clock,
      workbook: Option[Workbook]
    ): RangeEvalResult =
      evaluateForRangesPerCell(sheet, Vector(range), clock, workbook)

    /** [[evaluateForRangePerCell]] with the system clock and no workbook context. */
    @annotation.targetName("evaluateForRangePerCellDefault")
    def evaluateForRangePerCell(range: CellRange): RangeEvalResult =
      evaluateForRangesPerCell(sheet, Vector(range), Clock.system, None)

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
      evaluateArrayFormulaImpl(sheet, formula, originRef, Evaluator.arrayInstance, clock, workbook)

    /** Array formula evaluation with an explicit randomness source (GH-115). */
    @annotation.targetName("evaluateArrayFormulaWithRng")
    def evaluateArrayFormula(
      formula: String,
      originRef: ARef,
      clock: Clock,
      rng: Rng
    ): XLResult[(Sheet, CellRange)] =
      evaluateArrayFormulaImpl(sheet, formula, originRef, Evaluator.arrayInstance(rng), clock, None)

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

  // ========== Implementation shared by the public overloads ==========

  /**
   * What a range evaluation evaluates: the ranges' formula cells, the sheet's dependency graph, the
   * ranges' static closure (their formulas plus their transitive formula precedents), the target
   * formulas (the closure, or every formula when the sheet has dynamic references) and the sheet's
   * dynamic cells.
   */
  private final case class RangePlan(
    rangeFormulaCells: Set[ARef],
    graph: DependencyGraph,
    closure: Set[ARef],
    targets: Set[ARef],
    dynamic: Set[ARef]
  )

  /** The [[RangePlan]] of `ranges`, None when they hold no formula. */
  private def rangePlan(
    sheet: Sheet,
    ranges: Vector[CellRange],
    workbook: Option[Workbook]
  ): Option[RangePlan] =
    // a small range is a handful of point lookups: many of them (a formula rule's reads at every
    // cell of a render window) never make the scan below quadratic
    val (small, large) = ranges.partition(r => r.width.toLong * r.height <= 256L)
    val points = small.iterator.flatMap(_.cells).toSet
    def covered(ref: ARef): Boolean = points.contains(ref) || large.exists(_.contains(ref))
    val rangeFormulaCells = sheet.cells.iterator.collect {
      case (ref, cell) if isFormula(cell.value) && covered(ref) => ref
    }.toSet
    Option.when(rangeFormulaCells.nonEmpty) {
      val graph = DependencyGraph.fromSheet(sheet)
      val transitiveDeps = DependencyGraph.transitiveDependencies(graph, rangeFormulaCells)
      // GH-274: when the sheet has dynamic references (INDIRECT), widen to EVERY formula cell —
      // dynamic targets are invisible to the static walk, so correctness requires the whole sheet
      // computed in order (correctness over narrowness; results still report only the range).
      val dynamic = dynamicCellsFor(sheet, workbook)
      val allFormulaCells = graph.dependencies.keySet
      val closure = rangeFormulaCells ++ (transitiveDeps & allFormulaCells)
      val targets = if dynamic.isEmpty then closure else allFormulaCells
      RangePlan(rangeFormulaCells, graph, closure, targets, dynamic)
    }

  private def isFormula(value: CellValue): Boolean = value match
    case _: CellValue.Formula => true
    case _ => false

  /**
   * The per-cell evaluation of the formulas in `ranges` behind [[evaluateForRangePerCell]] (one
   * range) and `view --eval`'s pictures ([[LiveRender]]: the window and the cells its conditional
   * formatting reads). Cycle members inside the targets fail as circular and block their
   * dependents; the rest evaluates in topological order against a threaded sheet. A failure blocks
   * its transitive dependents, and every failed or blocked cell loses its cache on the threaded
   * sheet, so a dynamic reader (INDIRECT, invisible to the static graph) re-derives it and fails
   * too instead of reading a stale value. Only failures and blocked cells in the range's static
   * closure are reported: with dynamic references every formula is a target, but one that feeds the
   * range only through a dynamic read makes that reader — in the closure — fail on its own, so an
   * unrelated failure elsewhere stays silent.
   */
  private[eval] def evaluateForRangesPerCell(
    sheet: Sheet,
    ranges: Vector[CellRange],
    clock: Clock,
    workbook: Option[Workbook]
  ): RangeEvalResult =
    rangePlan(sheet, ranges, workbook) match
      case None => RangeEvalResult(Map.empty, Vector.empty, Vector.empty)
      case Some(RangePlan(rangeFormulaCells, graph, closure, targets, dynamic)) =>
        def formulaText(ref: ARef): String = sheet(ref).value match
          case CellValue.Formula(expression, _, _) => expression
          case _ => ref.toA1
        def failure(ref: ARef, reason: String): CellEvalError =
          CellEvalError(sheet.name, ref, XLError.FormulaError(formulaText(ref), reason))
        def dependentsOf(refs: Set[ARef]): Set[ARef] =
          (DependencyGraph.transitiveDependents(graph, refs) & targets) -- refs

        val cyclic = DependencyGraph.cyclicNodes(graph) & targets
        val cycleBlocked = if cyclic.isEmpty then Set.empty[ARef] else dependentsOf(cyclic)
        val live = targets -- cyclic -- cycleBlocked
        // Every formula a live target depends on is live: a dependent of a cycle member is itself
        // cycle-blocked, and the targets are closed under formula precedents (all formulas when
        // dynamic). Kahn's ordering counts only the nodes it is given, so the live subgraph sorts.
        val liveGraph =
          DependencyGraph(graph.dependencies.filter((ref, _) => live(ref)), graph.dependents)

        type State = (Sheet, Map[ARef, CellValue], Vector[CellEvalError], Set[ARef])
        val start: State =
          (
            stripFormulaCaches(sheet, cyclic ++ cycleBlocked),
            Map.empty,
            cyclic.toVector.map(failure(_, "Circular reference")),
            cycleBlocked
          )
        val (_, values, failures, blocked) = DependencyGraph.topologicalSort(liveGraph) match
          // Unreachable by the argument above; kept total rather than trusted.
          case scala.util.Left(circular) =>
            val (_, reported, failed, blockedSoFar) = start
            val unordered = live.toVector.map(failure(_, s"Unresolvable order: $circular"))
            (sheet, reported, failed ++ unordered, blockedSoFar)
          case scala.util.Right(order) =>
            val (tempSheet, reported, failed, blockedSoFar) = start
            // GH-274: dynamic cells and their static dependents evaluate last, caches stripped
            val (evalOrder, initial) = deferDynamicWithStrip(tempSheet, graph, order, dynamic)
            evalOrder.foldLeft[State]((initial, reported, failed, blockedSoFar)) {
              case (state @ (_, _, _, blockedNow), ref) if blockedNow(ref) => state
              case ((current, results, failedNow, blockedNow), ref) =>
                current.evaluateCell(ref, clock, workbook.map(_.put(current))) match
                  case scala.util.Right(value) =>
                    val next =
                      if rangeFormulaCells.contains(ref) then results + (ref -> value) else results
                    (current.put(ref, value), next, failedNow, blockedNow)
                  case scala.util.Left(error) =>
                    val newlyBlocked = dependentsOf(Set(ref)) -- blockedNow
                    (
                      stripFormulaCaches(current, newlyBlocked + ref),
                      results,
                      failedNow :+ CellEvalError(sheet.name, ref, error),
                      blockedNow ++ newlyBlocked
                    )
            }
        def rowMajor(ref: ARef): (Int, Int) = (ref.row.index0, ref.col.index0)
        RangeEvalResult(
          values,
          failures.filter(failed => closure(failed.ref)).sortBy(failed => rowMajor(failed.ref)),
          blocked.filter(closure).toVector.sortBy(rowMajor)
        )

  /**
   * Internal cell-evaluation boundary for WorkbookEvaluator.
   *
   * Supplying an evaluator lets one recalculation generation carry a narrowly scoped aggregate memo
   * through every cell and parallel worker. Public overloads call the same boundary with a normal
   * uncached evaluator, so their semantics remain unchanged.
   */
  private[formula] def evaluateCellWithEvaluator(
    sheet: Sheet,
    ref: ARef,
    evaluator: => Evaluator,
    clock: Clock,
    workbook: Option[Workbook]
  ): XLResult[CellValue] =
    sheet(ref).value match
      case value @ CellValue.Formula(expr, _, kind) =>
        pinnedCache(value) match
          // GH-353/GH-430: pinned-cache semantics — the Excel-written cache IS the value
          case Some(cached) => scala.util.Right(cached)
          case None =>
            // an ArrayFormula record (CSE or dynamic-array anchor) evaluates as an array and the
            // anchor holds element (0,0); a plain formula evaluates as a scalar cell
            val cellEvaluator = kind match
              case _: FormulaKind.ArrayFormula => evaluator.withArrayResults
              case _ => evaluator
            // Pass the current cell ref for ROW()/COLUMN() without arguments
            evaluateFormulaWith(sheet, expr, cellEvaluator, clock, workbook, Some(ref))
      case other => scala.util.Right(other)

  private def evaluateArrayFormulaImpl(
    sheet: Sheet,
    formula: String,
    originRef: ARef,
    evaluator: Evaluator,
    clock: Clock,
    workbook: Option[Workbook]
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

      // Evaluate TExpr against sheet. GH-344: a Left carrying an Excel error VALUE promotes to
      // a 1x1 CellValue.Error spilled at the origin; host failures stay loud Lefts.
      result <- evaluator.eval(expr, sheet, clock, workbook, Some(originRef)) match
        case scala.util.Right(value) => scala.util.Right(value)
        case scala.util.Left(evalError) =>
          EvalError.toErrorValue(evalError) match
            case Some(code) => scala.util.Right(CellValue.Error(code): Any)
            case None => scala.util.Left(evalErrorToXLError(evalError, Some(formula)))

      // Handle array vs scalar result
      updated <- result match
        case ar: ArrayResult =>
          val patch = Patch.PutArray(originRef, ar.values)
          val endRef = originRef.shift(ar.cols - 1, ar.rows - 1)
          scala.util.Right((Patch.applyPatch(sheet, patch), CellRange(originRef, endRef)))
        case other =>
          val cv = EvalResult.toCellValue(other)
          scala.util.Right((sheet.put(originRef, cv), CellRange(originRef, originRef)))
    yield updated

  // ========== Helper Functions ==========

  private def putFormulaInheritingImpl(
    sheet: Sheet,
    ref: ARef,
    formula: String,
    workbook: Option[Workbook]
  ): XLResult[Sheet] =
    // GH-479: the model's one canonical rule (trim, one leading '=' off, trim) is applied first, so
    // a padded `" = B2 - B3 "` parses and is stored exactly as fx / the CLI would store it
    val canonical = CellValue.canonicalFormulaText(formula)
    FormulaParser
      .parse(canonical)
      .left
      .map(parseError => XLError.FormulaError(formula, s"Parse error: $parseError"))
      .map { expr =>
        val withFormula = sheet.put(ref, CellValue.Formula(canonical))
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
   * Shared parse → evaluate → CellValue pipeline, parameterized by evaluator (GH-115: rng).
   *
   * GH-344: THE boundary-promotion site — a Left classified as an Excel error VALUE by
   * [[EvalError.toErrorValue]] becomes `Right(CellValue.Error(code))` here, funneling every
   * `evaluateFormula` overload, `evaluateCell`, `evaluateWithDependencyCheck`,
   * `wb.evaluateFormula`, `recalculate`, and the CLI. Headline contract:
   * `sheet.evaluateFormula("=1/0")` is `Right(CellValue.Error(Div0))`. Host failures (parse,
   * missing workbook/sheet, unknown function, cycles) stay loud Lefts.
   */
  private def evaluateFormulaWith(
    sheet: Sheet,
    formula: String,
    evaluator: Evaluator,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): XLResult[CellValue] =
    // A formula with a cell position evaluates as that plain cell would: Excel's legacy formula,
    // references in value positions implicitly intersected. Without one there is no cell to
    // intersect with, so it evaluates as the same formula typed into a new Excel 365 cell would:
    // as an array, showing its top-left value. `@range` still needs the position.
    val positioned = if currentCell.isDefined then evaluator else evaluator.withArrayResults
    parseFormula(formula).flatMap(expr =>
      evaluateParsedWith(sheet, formula, expr, positioned, clock, workbook, currentCell)
    )

  /** The parse half of [[evaluateFormulaWith]]: a parse failure is the `Parse error:` XLError. */
  private[eval] def parseFormula(formula: String): XLResult[TExpr[?]] =
    FormulaParser
      .parse(formula)
      .left
      .map(parseError =>
        XLError.FormulaError(
          formula,
          s"Parse error: $parseError"
        )
      )

  /**
   * The evaluate half of [[evaluateFormulaWith]] for an ALREADY PARSED formula — GH-537: an
   * iterative fixpoint parses each member once and evaluates its `TExpr` every round. `formulaText`
   * is the source text the diagnostics quote; it must be the text `expr` was parsed from. The
   * GH-344 boundary promotion lives here and nowhere else.
   */
  private[eval] def evaluateParsedWith(
    sheet: Sheet,
    formulaText: String,
    expr: TExpr[?],
    evaluator: Evaluator,
    clock: Clock,
    workbook: Option[Workbook],
    currentCell: Option[ARef]
  ): XLResult[CellValue] =
    evaluator.eval(expr, sheet, clock, workbook, currentCell) match
      case scala.util.Right(value) =>
        // A formula cell is never blank in Excel: a result that is a reference to an empty cell
        // (INDEX, INDIRECT, OFFSET, CHOOSE, a lookup) reads 0, as `=Z1` already does. Only the
        // cell's final value changes — inside a formula the reference stays blank (ISBLANK, COUNTA).
        EvalResult.toCellValue(value) match
          case CellValue.Empty => scala.util.Right(CellValue.Number(BigDecimal(0)))
          case other => scala.util.Right(other)
      case scala.util.Left(evalError) =>
        EvalError.toErrorValue(evalError) match
          case Some(code) => scala.util.Right(CellValue.Error(code))
          case None => scala.util.Left(evalErrorToXLError(evalError, Some(formulaText)))

  /** Shared dependency-ordered evaluation, optionally with an explicit rng (GH-115). */
  private def evaluateWithDependencyCheckImpl(
    sheet: Sheet,
    clock: Clock,
    workbook: Option[Workbook],
    rngOpt: Option[Rng]
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
            // GH-274: dynamic (INDIRECT-bearing) cells and their static dependents evaluate
            // last, against a temp whose bucket caches are stripped, so dynamic reads see
            // computed values. A dynamic depth-cap error fails the call, consistent with
            // this method's fail-fast cycle behavior.
            val (ordered, initial) =
              deferDynamicWithStrip(sheet, graph, evalOrder, dynamicCellsFor(sheet, workbook))
            // Evaluate in dependency order, threading the partially evaluated sheet so
            // dependent formulas see previously computed values. Fail-fast on first error.
            val evalResult = ordered.foldLeft[XLResult[(Sheet, Map[ARef, CellValue])]](
              scala.util.Right((initial, Map.empty))
            ) {
              case (scala.util.Right((tempSheet, results)), ref) =>
                val currentWorkbook = workbook.map(_.put(tempSheet))
                val evaluated = rngOpt match
                  case Some(rng) => tempSheet.evaluateCell(ref, clock, rng, currentWorkbook)
                  case None => tempSheet.evaluateCell(ref, clock, currentWorkbook)
                evaluated match
                  case scala.util.Right(value) =>
                    scala.util.Right((tempSheet.put(ref, value), results + (ref -> value)))
                  case scala.util.Left(error) =>
                    scala.util.Left(error)
              case (left, _) => left
            }

            evalResult.map(_._2)

  /**
   * GH-353: Excel closed-workbook semantics — Some(cached) when this is a formula cell whose
   * expression touches an external workbook ([2]Book1!A1) AND Excel wrote a cached value.
   *
   * Such cells are PINNED: the cached value is the only source of truth (the external workbook is
   * not loaded, so re-evaluation can only fail) and must be preserved verbatim. Callers skip
   * evaluation and use the cache directly; dependents then read it exactly as they read any cached
   * precedent. Uncached external formulas return None and fall through to normal evaluation, which
   * yields a clear per-cell error (Evaluator.externalRefUnsupported).
   *
   * The '[' pre-filter keeps the common recalculation path parse-free (same trick as
   * DependencyGraph.dynamicCells).
   *
   * Parse-failure fallback: some external shapes are beyond xl's parser (e.g. an external range in
   * an XLOOKUP array slot, or external defined names `[2]!name`). Re-evaluating such a cell can
   * only produce a parse error, and clearing its cache would destroy the sole source of truth, so a
   * CACHED formula that fails to parse but lexically carries the external-workbook prefix (`[n]…` /
   * `'[n]…`) is pinned too. A false positive (a `[n]`-looking substring in an unparseable
   * non-external formula) merely keeps that cell's existing cache — the same value every earlier xl
   * version reported for it.
   */
  private[formula] def pinnedExternalCache(value: CellValue): Option[CellValue] =
    value match
      case CellValue.Formula(expr, Some(cached), _) if expr.contains('[') =>
        FormulaParser.parse(expr) match
          case scala.util.Right(ast) if TExpr.containsExternalRef(ast) => Some(cached)
          case scala.util.Right(_) => None
          case scala.util.Left(_) if externalPrefixPattern.matcher(expr).find() => Some(cached)
          case scala.util.Left(_) => None
      case _ => None

  /**
   * GH-430: the generalized GH-353 seam. A data-table record's cache is its only truthful value —
   * xl does not evaluate `TABLE(...)` (the record, not the text, is the formula), so evaluation
   * pins the Excel-written cache exactly like closed-workbook externals; an uncached record pins to
   * Empty rather than parse-failing. ArrayFormula kinds are NOT pinned: their text evaluates as an
   * array and the anchor takes element (0,0) (evaluateCellWithEvaluator).
   */
  private[eval] def pinnedCache(value: CellValue): Option[CellValue] =
    value match
      case CellValue.Formula(_, cached, _: FormulaKind.DataTable) =>
        Some(cached.getOrElse(CellValue.Empty))
      case _ => pinnedExternalCache(value)

  /** GH-353: lexical external-workbook prefix — `[2]Book1!`, `'[3]Sheet Name'!`, `[2]!name`. */
  private val externalPrefixPattern = java.util.regex.Pattern.compile("""'?\[\d+\]""")

  /**
   * GH-274: strip stale formula caches from the given cells.
   *
   * Used on the deferred dynamic bucket before threading an evaluation fold: a dynamic read
   * (INDIRECT) of a not-yet-evaluated bucket cell then recursively evaluates the formula fresh
   * (depth-guarded) instead of trusting a previous generation's cache. Normal and ArrayFormula
   * records are both evaluatable; DataTable records alone retain their pinned cache, which is their
   * only source of truth. Only the threaded temp sheet is affected — final cache write-back
   * overlays computed results on the original sheet.
   */
  private[eval] def stripFormulaCaches(sheet: Sheet, refs: Set[ARef]): Sheet =
    refs.foldLeft(sheet) { (s, r) =>
      s.cells.get(r).map(_.value) match
        // GH-430: a DataTable cache is the sole source of truth. Graph analysis normally keeps
        // these records out of the dynamic bucket; keep this guard as structural safety.
        case Some(CellValue.Formula(_, Some(_), _: FormulaKind.DataTable)) => s
        case Some(f @ CellValue.Formula(_, Some(_), _)) =>
          s.put(r, f.copy(cachedValue = None))
        case _ => s
    }

  /** Resolve name-hidden dynamic references when the caller supplied workbook metadata. */
  private def dynamicCellsFor(sheet: Sheet, workbook: Option[Workbook]): Set[ARef] =
    workbook match
      case None => DependencyGraph.dynamicCells(sheet)
      case Some(wb) =>
        DependencyGraph
          .dynamicCells(wb.put(sheet))
          .iterator
          .filter(_.sheet == sheet.name)
          .map(_.ref)
          .toSet

  /**
   * GH-274: defer dynamic (INDIRECT-bearing) cells and their static dependents to the end of a
   * topological evaluation order, and strip their stale caches from the sheet the fold threads.
   * Identity when the sheet has no dynamic references.
   */
  private[eval] def deferDynamicWithStrip(
    sheet: Sheet,
    graph: DependencyGraph,
    evalOrder: List[ARef],
    dynamic: Set[ARef]
  ): (List[ARef], Sheet) =
    if dynamic.isEmpty then (evalOrder, sheet)
    else
      val bucket = DependencyGraph.dynamicClosure(graph, dynamic)
      (DependencyGraph.deferDynamic(evalOrder, bucket), stripFormulaCaches(sheet, bucket))

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
      // GH-344: unreachable after boundary promotion (toErrorValue promotes every ErrorValue),
      // kept total for direct callers of this table.
      case EvalError.ErrorValue(err, ctx) =>
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          ctx.fold(err.toExcel)(c => s"${err.toExcel} ($c)")
        )
      case other =>
        XLError.FormulaError(
          formulaContext.getOrElse(""),
          s"Evaluation error: $other$contextStr"
        )
