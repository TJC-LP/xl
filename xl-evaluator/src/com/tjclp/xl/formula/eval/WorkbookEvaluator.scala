package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
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
     * Scope notes (GH-346): ordering and cycle isolation are workbook-level — one topological order
     * over the qualified (sheet!cell) graph, so every precedent (same-sheet or cross-sheet) is
     * computed exactly once before its dependents and never re-derived on demand, and a cycle that
     * spans sheets is detected by Tarjan exactly like a same-sheet one (participants reported
     * circular, downstream dependents blocked, the acyclic remainder still evaluating; see
     * RecalcSpec).
     *
     * Dynamic references (GH-274): the static graph sees INDIRECT's *arguments*, not its resolved
     * *targets*. INDIRECT-bearing cells and their static dependents evaluate after all other
     * formulas (stable evaluate-last partition, `DependencyGraph.deferDynamic`) with stale caches
     * stripped, resolving not-yet-evaluated targets on demand under the depth-100 recursion guard —
     * so INDIRECT chains compute fresh values every recalculation. Dynamic cycles (INDIRECT
     * resolving into its own dependents) are not pre-detected by Tarjan; they surface as per-cell
     * recursion-guard errors while the rest of the workbook still evaluates. Because `recalculate`
     * is always full-workbook, Excel's "volatile" marking is moot here; the targeted
     * `recalculateDependents` treats dynamic cells as always dirty instead.
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
    val (deps, dependents) = DependencyGraph.fromWorkbookBounded(wb)

    def formulaText(q: QualifiedRef): String =
      wb(q.sheet).toOption.flatMap(_.cells.get(q.ref)).map(_.value) match
        case Some(CellValue.Formula(expr, _)) => expr
        case _ => q.ref.toA1

    // GH-346: cycle isolation on the WORKBOOK-level graph — same-sheet and cross-sheet cycles
    // alike are detected by Tarjan and removed up front (previously a cross-sheet cycle fell
    // through to the evaluator's recursion depth guard).
    val cyclicCore = DependencyGraph.qualifiedCyclicNodes(deps)
    val blocked = DependencyGraph.qualifiedTransitiveDependents(dependents, cyclicCore)

    val cycleErrors = cyclicCore.toVector.map { q =>
      CellEvalError(q.sheet, q.ref, XLError.FormulaError(formulaText(q), "Circular reference"))
    }
    // qualifiedTransitiveDependents excludes its seed set, so `blocked` is disjoint from the core
    val blockedErrors = blocked.toVector.map { q =>
      CellEvalError(
        q.sheet,
        q.ref,
        XLError.FormulaError(formulaText(q), "Blocked by an upstream circular reference")
      )
    }

    val removed = cyclicCore ++ blocked
    val prunedDeps = (deps -- removed).view.mapValues(_ -- removed).toMap
    val prunedDependents = (dependents -- removed).view.mapValues(_ -- removed).toMap

    DependencyGraph.qualifiedTopologicalSort(prunedDeps, prunedDependents) match
      case Left(circular) =>
        // Unreachable: qualifiedCyclicNodes removed every cycle participant. Stay total anyway.
        val residual = prunedDeps.keySet.toVector.map { q =>
          CellEvalError(
            q.sheet,
            q.ref,
            XLError.FormulaError(formulaText(q), s"Unresolvable order: $circular")
          )
        }
        RecalcResult(
          workbook = wb,
          evaluated = wb.sheets.map(s => s.name -> Map.empty[ARef, CellValue]).toMap,
          errors = cycleErrors ++ blockedErrors ++ residual
        )

      case Right(evalOrder) =>
        // GH-274/GH-301: dynamic-reference cells (INDIRECT/OFFSET) and their transitive static
        // dependents evaluate last (stable partition — preserves topological validity per the
        // deferDynamic lemma), with stale caches stripped so a dynamic read of a
        // not-yet-evaluated bucket cell recursively evaluates fresh (depth-guarded) instead of
        // trusting a previous generation's cache. The closure is workbook-level (GH-346): a
        // cross-sheet static dependent of a dynamic cell defers with it. Dynamic-free workbooks
        // take the identity path.
        val dynamic: Set[QualifiedRef] =
          wb.sheets.iterator.flatMap { s =>
            DependencyGraph.dynamicCells(s).iterator.map(r => QualifiedRef(s.name, r))
          }.toSet -- removed
        val bucket =
          if dynamic.isEmpty then Set.empty[QualifiedRef]
          else dynamic ++ DependencyGraph.qualifiedTransitiveDependents(prunedDependents, dynamic)
        val ordered =
          if bucket.isEmpty then evalOrder
          else evalOrder.filterNot(bucket.contains) ++ evalOrder.filter(bucket.contains)

        val stripBySheet: Map[SheetName, Set[ARef]] = bucket.groupMap(_.sheet)(_.ref)
        val initialTemps: Map[SheetName, Sheet] =
          wb.sheets.map { s =>
            val toStrip = stripBySheet.getOrElse(s.name, Set.empty)
            s.name ->
              (if toStrip.isEmpty then s else SheetEvaluator.stripFormulaCaches(s, toStrip))
          }.toMap

        // Evaluate in the single global order, threading the partially evaluated workbook:
        // every reference — same-sheet or cross-sheet — reads its precedents as already-computed
        // values, so recursive re-derivation only arises on dynamic reads. Rebuilding tempWb per
        // cell is O(sheets) shallow-copy work per formula — negligible at realistic sheet
        // counts, and each evaluation must see a workbook consistent with the fold state.
        val (_, evaluated, evalErrors) = ordered.foldLeft(
          (initialTemps, Map.empty[SheetName, Map[ARef, CellValue]], Vector.empty[CellEvalError])
        ) { case ((temps, acc, errs), q) =>
          temps.get(q.sheet) match
            case None => (temps, acc, errs) // node names a sheet absent from the workbook
            case Some(tempSheet) =>
              val tempWb = wb.copy(sheets = wb.sheets.map(s => temps.getOrElse(s.name, s)))
              val evaluatedCell = rngOpt match
                case Some(rng) => tempSheet.evaluateCell(q.ref, clock, rng, Some(tempWb))
                case None => tempSheet.evaluateCell(q.ref, clock, Some(tempWb))
              evaluatedCell match
                case Right(value) =>
                  (
                    temps.updated(q.sheet, tempSheet.put(q.ref, value)),
                    acc.updated(q.sheet, acc.getOrElse(q.sheet, Map.empty) + (q.ref -> value)),
                    errs
                  )
                case Left(error) => (temps, acc, errs :+ CellEvalError(q.sheet, q.ref, error))
        }

        // Cache computed values into formula cells on the original sheets; failed cells stay
        // uncached (Excel recalculates them on open)
        val cachedSheets = wb.sheets.map { orig =>
          evaluated.getOrElse(orig.name, Map.empty).foldLeft(orig) { case (s, (ref, computed)) =>
            s.cells.get(ref).map(_.value) match
              case Some(CellValue.Formula(expr, _)) =>
                s.put(ref, CellValue.Formula(expr, Some(computed)))
              case _ => s
          }
        }

        RecalcResult(
          workbook = wb.copy(sheets = cachedSheets),
          evaluated = wb.sheets.map(s => s.name -> evaluated.getOrElse(s.name, Map.empty)).toMap,
          errors = cycleErrors ++ blockedErrors ++ evalErrors
        )
