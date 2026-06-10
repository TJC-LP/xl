package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.{Clock, Rng}

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

// Import SheetEvaluator extension methods
import SheetEvaluator.*

/**
 * Extension methods for evaluating and caching formula values in workbooks.
 *
 * Usage:
 * {{{
 * import com.tjclp.xl.{*, given}
 *
 * val wb = Workbook(sheets).withCachedFormulas()
 * Excel.write(wb, "output.xlsx")
 * }}}
 */
object WorkbookEvaluator:

  extension (wb: Workbook)

    /**
     * Evaluate a formula in the context of the named sheet, with cross-sheet references resolved
     * against this workbook.
     *
     * Prefer this over `Sheet.evaluateFormula` in scripts: the workbook context is wired
     * automatically, so formulas like `='Other Sheet'!A1 * 2` just work.
     *
     * Note: explicit overloads instead of a default clock parameter — extension methods with
     * default arguments crash the compiler when merged through the formulaExports wildcard export
     * (see the DependentRecalculation note in exports.scala).
     */
    def evaluateFormula(
      formula: String,
      onSheet: SheetName,
      clock: Clock
    ): XLResult[CellValue] =
      wb(onSheet).flatMap(s => SheetEvaluator.evaluateFormula(s)(formula, clock, Some(wb)))

    /** Evaluate a formula on the named sheet with the system clock. */
    @annotation.targetName("evaluateFormulaOnSheetDefaultClock")
    def evaluateFormula(formula: String, onSheet: SheetName): XLResult[CellValue] =
      evaluateFormula(formula, onSheet, Clock.system)

    /**
     * Evaluate a formula on the named sheet (string variant). The sheet name is resolved at
     * runtime.
     */
    @annotation.targetName("evaluateFormulaOnSheetString")
    def evaluateFormula(
      formula: String,
      onSheet: String,
      clock: Clock
    ): XLResult[CellValue] =
      wb(onSheet).flatMap(s => SheetEvaluator.evaluateFormula(s)(formula, clock, Some(wb)))

    /** Evaluate a formula on the named sheet (string variant, system clock). */
    @annotation.targetName("evaluateFormulaOnSheetStringDefaultClock")
    def evaluateFormula(formula: String, onSheet: String): XLResult[CellValue] =
      evaluateFormula(formula, onSheet, Clock.system)

    /**
     * Evaluate a formula on the named sheet with an explicit randomness source (GH-115) —
     * deterministic RAND/RANDBETWEEN via `Rng.seeded(seed)`.
     */
    @annotation.targetName("evaluateFormulaOnSheetWithRng")
    def evaluateFormula(
      formula: String,
      onSheet: SheetName,
      clock: Clock,
      rng: Rng
    ): XLResult[CellValue] =
      wb(onSheet).flatMap(s => SheetEvaluator.evaluateFormula(s)(formula, clock, rng, Some(wb)))

    /** Evaluate a formula on the named sheet with an explicit rng (string variant, GH-115). */
    @annotation.targetName("evaluateFormulaOnSheetStringWithRng")
    def evaluateFormula(
      formula: String,
      onSheet: String,
      clock: Clock,
      rng: Rng
    ): XLResult[CellValue] =
      wb(onSheet).flatMap(s => SheetEvaluator.evaluateFormula(s)(formula, clock, rng, Some(wb)))

    /**
     * Evaluate all formulas and cache their computed values.
     *
     * Useful for:
     *   - CLI tools that need to display computed values
     *   - Files opened by tools other than Excel
     *   - Faster Excel open times (no recalculation needed)
     *
     * Formulas that fail to evaluate (unsupported functions, circular refs) are left without cached
     * values - Excel will recalculate on open.
     *
     * Implemented as `recalculate(clock).workbook`; use `recalculate` directly when you need to
     * know which cells failed.
     *
     * @param clock
     *   Clock for volatile functions (TODAY, NOW). Defaults to system clock.
     * @return
     *   Workbook with formula cells containing cached values
     */
    def withCachedFormulas(clock: Clock = Clock.system): Workbook =
      recalculate(clock).workbook

    /** Cache formula values with an explicit randomness source (GH-115). */
    @annotation.targetName("withCachedFormulasWithRng")
    def withCachedFormulas(clock: Clock, rng: Rng): Workbook =
      recalculate(clock, rng).workbook

    /**
     * Total whole-workbook recalculation with per-cell error reporting.
     *
     * Evaluates every formula in every sheet in dependency order, with cross-sheet references
     * resolved against this workbook. Failures never propagate: a failing cell becomes a
     * [[CellEvalError]] and stays uncached; cycle participants are isolated (reported as circular,
     * their downstream dependents as blocked) while the acyclic remainder still evaluates.
     *
     * {{{
     * val result = wb.recalculate()
     * result.errors.foreach(e => println(e.render))
     * Excel.write(result.workbook, "out.xlsx")  // partial results; failed cells stay uncached
     * result.toEither                           // Left(errors) when not clean
     * }}}
     *
     * Scope notes: cycle isolation is per-sheet (Tarjan over each sheet's graph) — a cycle that
     * spans sheets is caught instead by the evaluator's recursion depth guard and surfaces as a
     * generic per-cell eval error (still total, still collected; see RecalcSpec). Cross-sheet
     * references to upstream formulas evaluate on demand against the original workbook snapshot and
     * are not memoized across sheets — fine for typical workbooks; a workbook-level topological
     * order is future work for deep cross-sheet fan-out.
     */
    def recalculate(clock: Clock = Clock.system): RecalcResult =
      recalculateImpl(wb, clock, None)

    /**
     * Total whole-workbook recalculation with an explicit randomness source (GH-115).
     *
     * Volatile semantics: RAND/RANDBETWEEN draw fresh values on every recalculate; with
     * `Rng.seeded(seed)` the whole recalculation is reproducible.
     */
    @annotation.targetName("recalculateWithRng")
    def recalculate(clock: Clock, rng: Rng): RecalcResult =
      recalculateImpl(wb, clock, Some(rng))

  // ========== recalculate internals ==========

  private def recalculateImpl(wb: Workbook, clock: Clock, rngOpt: Option[Rng]): RecalcResult =
    val perSheet = wb.sheets.map(sheet => recalcSheet(sheet, wb, clock, rngOpt))
    RecalcResult(
      workbook = wb.copy(sheets = perSheet.map(_.sheet)),
      evaluated = perSheet.map(r => r.sheet.name -> r.evaluated).toMap,
      errors = perSheet.flatMap(_.errors).toVector
    )

  private final case class SheetRecalc(
    sheet: Sheet,
    evaluated: Map[ARef, CellValue],
    errors: Vector[CellEvalError]
  )

  private def formulaText(sheet: Sheet, ref: ARef): String =
    sheet.cells.get(ref).map(_.value) match
      case Some(CellValue.Formula(expr, _)) => expr
      case _ => ref.toA1

  private def recalcSheet(
    sheet: Sheet,
    wb: Workbook,
    clock: Clock,
    rngOpt: Option[Rng]
  ): SheetRecalc =
    val graph = DependencyGraph.fromSheet(sheet)
    val cyclicCore = DependencyGraph.cyclicNodes(graph)
    val blocked = DependencyGraph.transitiveDependents(graph, cyclicCore)

    val cycleErrors = cyclicCore.toVector.map { ref =>
      CellEvalError(
        sheet.name,
        ref,
        XLError.FormulaError(formulaText(sheet, ref), "Circular reference")
      )
    }
    // transitiveDependents excludes its seed set, so `blocked` is already disjoint from the core
    val blockedErrors = blocked.toVector.map { ref =>
      CellEvalError(
        sheet.name,
        ref,
        XLError.FormulaError(formulaText(sheet, ref), "Blocked by an upstream circular reference")
      )
    }

    val removed = cyclicCore ++ blocked
    val pruned = DependencyGraph(
      dependencies = (graph.dependencies -- removed).view.mapValues(_ -- removed).toMap,
      dependents = (graph.dependents -- removed).view.mapValues(_ -- removed).toMap
    )

    DependencyGraph.topologicalSort(pruned) match
      case Left(circular) =>
        // Unreachable: cyclicNodes removed every cycle participant. Stay total regardless.
        val residual = pruned.dependencies.keySet.toVector.map { ref =>
          CellEvalError(
            sheet.name,
            ref,
            XLError.FormulaError(formulaText(sheet, ref), s"Unresolvable order: $circular")
          )
        }
        SheetRecalc(sheet, Map.empty, cycleErrors ++ blockedErrors ++ residual)

      case Right(evalOrder) =>
        val (tempSheet, results, evalErrors) = evalOrder.foldLeft(
          (sheet, Map.empty[ARef, CellValue], Vector.empty[CellEvalError])
        ) { case ((temp, acc, errs), ref) =>
          val evaluated = rngOpt match
            case Some(rng) => temp.evaluateCell(ref, clock, rng, Some(wb))
            case None => temp.evaluateCell(ref, clock, Some(wb))
          evaluated match
            case Right(value) => (temp.put(ref, value), acc + (ref -> value), errs)
            case Left(error) => (temp, acc, errs :+ CellEvalError(sheet.name, ref, error))
        }

        // Cache computed values into formula cells; failed cells stay uncached
        val cachedSheet = results.foldLeft(sheet) { case (s, (ref, computed)) =>
          s.cells.get(ref).map(_.value) match
            case Some(CellValue.Formula(expr, _)) =>
              s.put(ref, CellValue.Formula(expr, Some(computed)))
            case _ => s
        }
        SheetRecalc(cachedSheet, results, cycleErrors ++ blockedErrors ++ evalErrors)
