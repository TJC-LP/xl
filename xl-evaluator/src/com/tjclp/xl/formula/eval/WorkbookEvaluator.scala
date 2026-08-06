package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.{Clock, Rng}

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

import scala.util.control.NonFatal

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
     * GH-344: a formula that COMPUTES an Excel error value (#DIV/0!, #N/A, ...) evaluates
     * successfully — the error value caches (`Formula(expr, Some(Error(code)))`), writes as a
     * `t="e"` cell, and cascades to dependents like any precedent value. Such cells report via
     * [[RecalcResult.excelErrors]], never [[RecalcResult.errors]].
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
      recalculateImpl(wb, clock, None, None)

    /**
     * Total whole-workbook recalculation with an explicit randomness source (GH-115).
     *
     * Volatile semantics: RAND/RANDBETWEEN draw fresh values on every recalculate; with
     * `Rng.seeded(seed)` the whole recalculation is reproducible.
     */
    @annotation.targetName("recalculateWithRng")
    def recalculate(clock: Clock, rng: Rng): RecalcResult =
      recalculateImpl(wb, clock, Some(rng), None)

    /**
     * GH-373: recalculation with OPT-IN bounded iterative evaluation of circular references (system
     * clock).
     *
     * Cycle members fixpoint via Jacobi iteration instead of erroring — see [[IterativeCalc]] for
     * semantics (0-seeding, |Δ| < maxChange convergence, non-convergence keeps last values with no
     * error) and for the explicit bridge from a file's `<calcPr>` settings
     * (`wb.metadata.calcPr.filter(_.iterativeCalculation).map(IterativeCalc.fromCalcPr)`). Acyclic
     * workbooks behave exactly like the default `recalculate()`.
     *
     * Note: explicit overloads instead of default parameters — extension methods with default
     * arguments crash the compiler when merged through the formulaExports wildcard export (see the
     * DependentRecalculation note in exports.scala).
     */
    @annotation.targetName("recalculateIterative")
    def recalculate(iterative: IterativeCalc): RecalcResult =
      recalculateImpl(wb, Clock.system, None, Some(iterative))

    /** GH-373: iterative recalculation with an explicit clock. */
    @annotation.targetName("recalculateIterativeWithClock")
    def recalculate(clock: Clock, iterative: IterativeCalc): RecalcResult =
      recalculateImpl(wb, clock, None, Some(iterative))

    /** GH-373: iterative recalculation with an explicit clock and randomness source. */
    @annotation.targetName("recalculateIterativeWithClockRng")
    def recalculate(clock: Clock, rng: Rng, iterative: IterativeCalc): RecalcResult =
      recalculateImpl(wb, clock, Some(rng), Some(iterative))

  // ========== recalculate internals ==========

  private def recalculateImpl(
    wb: Workbook,
    clock: Clock,
    rngOpt: Option[Rng],
    iterativeOpt: Option[IterativeCalc]
  ): RecalcResult =
    val (deps, dependents) = DependencyGraph.fromWorkbookBounded(wb)

    def formulaText(q: QualifiedRef): String =
      wb(q.sheet).toOption.flatMap(_.cells.get(q.ref)).map(_.value) match
        case Some(CellValue.Formula(expr, _, _)) => expr
        case _ => q.ref.toA1

    // GH-346: cycle isolation on the WORKBOOK-level graph — same-sheet and cross-sheet cycles
    // alike are detected by Tarjan and removed up front (previously a cross-sheet cycle fell
    // through to the evaluator's recursion depth guard).
    val cyclicCore = DependencyGraph.qualifiedCyclicNodes(deps)
    // GH-373: iterative mode inverts the cycle posture — the core fixpoints (no circular
    // errors) and its transitive dependents are NOT pre-marked blocked: they evaluate normally
    // off the converged values, AFTER the iteration (see the pre/post partition below).
    val iterating = iterativeOpt.isDefined && cyclicCore.nonEmpty
    val coreDependents = DependencyGraph.qualifiedTransitiveDependents(dependents, cyclicCore)
    val blocked = if iterating then Set.empty[QualifiedRef] else coreDependents

    val cycleErrors =
      if iterating then Vector.empty[CellEvalError]
      else
        cyclicCore.toVector.map { q =>
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

    // The cyclic core always leaves the Kahn graph (it cannot be ordered); blocked dependents
    // leave it only in the default (non-iterative) path.
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
        val dynamicAll: Set[QualifiedRef] =
          wb.sheets.iterator.flatMap { s =>
            DependencyGraph.dynamicCells(s).iterator.map(r => QualifiedRef(s.name, r))
          }.toSet
        val dynamic: Set[QualifiedRef] = dynamicAll -- removed
        val bucket =
          if dynamic.isEmpty then Set.empty[QualifiedRef]
          else dynamic ++ DependencyGraph.qualifiedTransitiveDependents(prunedDependents, dynamic)
        val ordered =
          if bucket.isEmpty then evalOrder
          else evalOrder.filterNot(bucket.contains) ++ evalOrder.filter(bucket.contains)

        // GH-492: the condensation walk evaluates the cyclic core too, so the deferral bucket must
        // be closed over the FULL graph — `prunedDependents` has the core deleted, and a cyclic
        // component downstream of a dynamic cell would otherwise be scheduled ahead of it,
        // inverting a real edge. Dependent-closure lifts to whole components (mutual reachability
        // inside an SCC), so the stable partition of the condensation stays a valid order.
        val bucketIter =
          if !iterating then bucket
          else if dynamicAll.isEmpty then Set.empty[QualifiedRef]
          else dynamicAll ++ DependencyGraph.qualifiedTransitiveDependents(dependents, dynamicAll)

        val stripBySheet: Map[SheetName, Set[ARef]] = bucketIter.groupMap(_.sheet)(_.ref)
        val initialSheets: Vector[Sheet] =
          wb.sheets.map { s =>
            val toStrip = stripBySheet.getOrElse(s.name, Set.empty)
            if toStrip.isEmpty then s else SheetEvaluator.stripFormulaCaches(s, toStrip)
          }
        val sheetIndex: Map[SheetName, Int] =
          wb.sheets.zipWithIndex.map((s, i) => s.name -> i).toMap

        // Evaluate in the single global order, threading the partially evaluated workbook:
        // every reference — same-sheet or cross-sheet — reads its precedents as already-computed
        // values, so recursive re-derivation only arises on dynamic reads. The temp sheets live
        // in a Vector parallel to wb.sheets, updated incrementally, so per-cell cost is one
        // Workbook shell copy + one Vector update rather than re-mapping every sheet.
        def evalPass(
          order: List[QualifiedRef],
          init: PassState,
          clk: Clock
        ): PassState =
          order.foldLeft(init) { case ((sheets, acc, errs), q) =>
            sheetIndex.get(q.sheet) match
              case None => (sheets, acc, errs) // node names a sheet absent from the workbook
              case Some(idx) =>
                val tempSheet = sheets(idx)
                val tempWb = wb.copy(sheets = sheets)
                // GH-388 defense in depth: recalculate is documented total — a numeric blowup
                // escaping a function implementation (e.g. BigDecimal scale overflow in a
                // diverging Newton loop) must degrade to this cell's per-cell error, never
                // unwind the whole recalculation.
                val evaluatedCell =
                  try
                    rngOpt match
                      case Some(rng) => tempSheet.evaluateCell(q.ref, clk, rng, Some(tempWb))
                      case None => tempSheet.evaluateCell(q.ref, clk, Some(tempWb))
                  catch
                    case NonFatal(e) =>
                      Left(
                        XLError.FormulaError(
                          formulaText(q),
                          s"Evaluation threw ${e.getClass.getName}"
                        )
                      )
                evaluatedCell match
                  case Right(value) =>
                    (
                      sheets.updated(idx, tempSheet.put(q.ref, value)),
                      acc.updated(q.sheet, acc.getOrElse(q.sheet, Map.empty) + (q.ref -> value)),
                      errs
                    )
                  case Left(error) => (sheets, acc, errs :+ CellEvalError(q.sheet, q.ref, error))
          }

        /**
         * GH-492: fixpoint ONE cyclic condensation node against the threaded temp sheets, then fold
         * the result into the pass state. Successful members write their value into the temp sheets
         * (later components and acyclic cells read them) and into the evaluated map (they cache at
         * the end); failed members keep their original uncached formula cell, so dependents
         * recursively evaluate and surface the underlying error — the same posture as a failed
         * acyclic cell. Exhaustion is NOT an error (Excel keeps the last values); it is reported
         * through the returned [[SccReport]].
         */
        def fixpointStep(
          component: Vector[QualifiedRef],
          iterative: IterativeCalc,
          pinnedClock: Clock,
          state: PassState
        ): (PassState, SccReport) =
          val (baseSheets, acc0, errs0) = state
          val members: List[(QualifiedRef, Int, String)] =
            component.toList.flatMap(q =>
              sheetIndex.get(q.sheet).map(idx => (q, idx, formulaText(q)))
            )
          // GH-469: Excel seeds iterative calculation from the CURRENT cell values. Read them off
          // the THREADED sheets, not the original workbook: a member whose cache the dynamic
          // bucket stripped correctly seeds 0, and no member can be read after being written
          // (each component is visited exactly once, and components are disjoint).
          val seed =
            if iterative.seedFromCaches then warmSeed(baseSheets, members)
            else Map.empty[QualifiedRef, CellValue]
          val outcome =
            jacobiFixpoint(wb, baseSheets, members, iterative, pinnedClock, rngOpt, seed)
          val folded = members.foldLeft((baseSheets, acc0, errs0)) {
            case ((sheets, acc, errs), (q, idx, _)) =>
              outcome.results.get(q) match
                case Some(Right(value)) =>
                  (
                    sheets.updated(idx, sheets(idx).put(q.ref, value)),
                    acc.updated(q.sheet, acc.getOrElse(q.sheet, Map.empty) + (q.ref -> value)),
                    errs
                  )
                case Some(Left(error)) =>
                  (sheets, acc, errs :+ CellEvalError(q.sheet, q.ref, error))
                case None => (sheets, acc, errs) // unreachable: every member evaluates every round
          }
          val report = SccReport(
            members = members.map((q, _, _) => (q.sheet, q.ref)).toVector,
            converged = outcome.converged,
            rounds = outcome.rounds,
            maxDelta = outcome.maxDelta
          )
          (folded, report)

        def runPlan(
          plan: Vector[CalcStep],
          iterative: IterativeCalc,
          pinnedClock: Clock,
          init: PassState
        ): (PassState, Vector[SccReport]) =
          plan.foldLeft((init, Vector.empty[SccReport])) {
            case ((state, reports), CalcStep.Straight(order)) =>
              (evalPass(order, state, pinnedClock), reports)
            case ((state, reports), CalcStep.Fixpoint(component)) =>
              val (next, report) = fixpointStep(component, iterative, pinnedClock, state)
              (next, reports :+ report)
          }

        val initState: PassState =
          (initialSheets, Map.empty[SheetName, Map[ARef, CellValue]], Vector.empty[CellEvalError])

        // GH-492: iterative mode walks the SCC condensation ONCE in dependency-first order — a run
        // of acyclic components is one `evalPass`, each cyclic component is one `jacobiFixpoint`
        // against the threaded temp sheets. Every precedent is therefore a freshly computed value
        // before anything reads it, so no read can fall back to a loaded cache and `converged`
        // certifies the workbook's GLOBAL fixpoint (see RecalcResult). The clock is pinned ONCE
        // for the whole walk: one iterative recalculation is one volatile generation, in the
        // fixpoints AND in the acyclic bridges between them. The default path is untouched —
        // literally the same single `evalPass(ordered, ...)` with the caller's clock.
        val (finalState, cycleReports) = iterativeOpt match
          case Some(iterative) if cyclicCore.nonEmpty =>
            val pinnedClock = Clock.fixed(clock.today(), clock.now())
            val components = DependencyGraph.qualifiedSccOrder(deps)
            val walk =
              if bucketIter.isEmpty then components
              else
                val (eager, deferred) =
                  components.partition(c => !c.members.exists(bucketIter.contains))
                eager ++ deferred
            runPlan(buildPlan(walk), iterative, pinnedClock, initState)
          case _ => (evalPass(ordered, initState, clock), Vector.empty[SccReport])

        val (_, evaluated, evalErrors) = finalState
        val cycles = cycleReports.sortBy(r =>
          r.members.headOption.fold(("", ""))((s, ref) => (s.value, ref.toA1))
        )
        val converged = cycles.forall(_.converged)
        val iterationsUsed = cycles.map(_.rounds).maxOption.getOrElse(0)

        // Cache computed values into formula cells on the original sheets; failed cells stay
        // uncached (Excel recalculates them on open). Track per sheet whether any recomputed
        // value actually differs from the pre-existing cache (a newly cached cell — prior
        // cache None — counts as a change).
        val cachedSheets: Vector[(Sheet, Boolean)] = wb.sheets.map { orig =>
          evaluated.getOrElse(orig.name, Map.empty).foldLeft((orig, false)) {
            case ((s, changed), (ref, computed)) =>
              s.cells.get(ref).map(_.value) match
                // GH-430: a data-table cache is never rewritten by recalculation — pinned
                // evaluation echoes it, and an uncached record must stay uncached, not gain
                // a synthetic Some(Empty).
                case Some(CellValue.Formula(_, _, _: FormulaKind.DataTable)) => (s, changed)
                case Some(f @ CellValue.Formula(_, cached, _)) if !cached.contains(computed) =>
                  (s.put(ref, f.copy(cachedValue = Some(computed))), true)
                case _ => (s, changed)
          }
        }

        // Reinstall changed sheets via Workbook.put — not copy — so the surgical-write
        // ModificationTracker sees every sheet whose caches actually changed (GH-352). A raw
        // copy leaves disk-read workbooks looking untouched, and the writer then preserves
        // the original worksheet XML verbatim, silently dropping the new cached values.
        // Sheets whose recomputed caches all equal the existing ones are left alone: putting
        // them would mark them modified, forcing the writer to regenerate their XML and drop
        // any unparsed parts (pivot tables, slicers, ctrlProps) that can only survive via
        // byte-for-byte preservation.
        val workbookWithCaches =
          cachedSheets.foldLeft(wb) { case (acc, (cached, changed)) =>
            if changed then acc.put(cached) else acc
          }

        RecalcResult(
          workbook = workbookWithCaches,
          evaluated = wb.sheets.map(s => s.name -> evaluated.getOrElse(s.name, Map.empty)).toMap,
          errors = cycleErrors ++ blockedErrors ++ evalErrors,
          converged = converged,
          iterationsUsed = iterationsUsed,
          cycles = cycles
        )

  /** GH-492: the threaded state of one recalculation pass (temp sheets, values, errors). */
  private[eval] type PassState =
    (Vector[Sheet], Map[SheetName, Map[ARef, CellValue]], Vector[CellEvalError])

  /**
   * GH-492: one step of the condensation walk.
   *
   * `Straight` batches a maximal run of CONSECUTIVE acyclic condensation nodes into a single
   * `evalPass` — which is why the non-iterative plan degenerates to one `Straight` holding the
   * whole topological order, i.e. to exactly the pre-GH-492 single pass.
   */
  private[eval] enum CalcStep derives CanEqual:
    case Straight(order: List[QualifiedRef])
    case Fixpoint(members: Vector[QualifiedRef])

  /** Collapse consecutive acyclic components into one `Straight`; each cyclic node is a step. */
  private[eval] def buildPlan(components: Vector[DependencyGraph.Scc]): Vector[CalcStep] =
    val (steps, trailing) =
      components.foldLeft((Vector.empty[CalcStep], Vector.empty[QualifiedRef])) {
        case ((acc, run), c) =>
          if c.cyclic then
            val flushed = if run.isEmpty then acc else acc :+ CalcStep.Straight(run.toList)
            (flushed :+ CalcStep.Fixpoint(c.members), Vector.empty[QualifiedRef])
          else (acc, run ++ c.members)
      }
    if trailing.isEmpty then steps else steps :+ CalcStep.Straight(trailing.toList)

  /**
   * GH-469: the members' currently cached values — Excel's iterative-calculation seed. Members with
   * no cache are omitted and fall back to 0 inside [[jacobiFixpoint]].
   *
   * Any cached VALUE seeds, not just numbers: convergence still requires re-evaluating the formula
   * and reproducing that value, which is a genuine fixpoint (GH-344 already blesses exactly this
   * for error values).
   */
  private[eval] def warmSeed(
    baseSheets: Vector[Sheet],
    members: List[(QualifiedRef, Int, String)]
  ): Map[QualifiedRef, CellValue] =
    members.flatMap { (q, idx, _) =>
      baseSheets.lift(idx).flatMap(_.cells.get(q.ref)).map(_.value) match
        case Some(CellValue.Formula(_, Some(cachedValue), _)) => Some(q -> cachedValue)
        case _ => None
    }.toMap

  /**
   * GH-492: the outcome of one component's Jacobi fixpoint.
   *
   * @param maxDelta
   *   the largest |Δ| among NUMERIC members in the final round (None when no member was numeric)
   */
  private[eval] final case class FixpointOutcome(
    results: Map[QualifiedRef, Either[XLError, CellValue]],
    converged: Boolean,
    rounds: Int,
    maxDelta: Option[BigDecimal]
  )

  /**
   * The shared Jacobi fixpoint engine (GH-373/GH-453/GH-454/GH-492): iterate `members` against
   * `baseSheets` until every member's |Δ| < `maxChange` (strict) or `maxIter` rounds.
   *
   * Each member is `(qualified ref, index of its sheet in baseSheets, formula text)`. Members seed
   * from `seed`, falling back to 0 for anything absent (`Map.empty` therefore reproduces the
   * original all-zero seeding exactly); each round overlays every member's cell with
   * `Formula(expr, Some(previousValue))` so every reference to a member — including self-references
   * — reads the previous round's value while `evaluateCell` re-evaluates the formula text. Callers
   * pin the clock BEFORE calling (one fixpoint is one volatile generation). A member that throws
   * holds its previous value for later rounds (GH-388 degradation) and reports its Left only from
   * the final round.
   *
   * Non-convergence KEEPS the last values with no error — Excel's semantics, and the reason
   * exhaustion surfaces through [[FixpointOutcome]] rather than as a [[CellEvalError]].
   *
   * Used by both `recalculate(IterativeCalc)` (one call per cyclic condensation node) and the
   * data-table seeder's circular what-if substitution.
   */
  private[eval] def jacobiFixpoint(
    wb: Workbook,
    baseSheets: Vector[Sheet],
    members: List[(QualifiedRef, Int, String)],
    iterative: IterativeCalc,
    pinnedClock: Clock,
    rngOpt: Option[Rng],
    seedValues: Map[QualifiedRef, CellValue]
  ): FixpointOutcome =
    val zero: CellValue = CellValue.Number(BigDecimal(0))
    val seed: Map[QualifiedRef, CellValue] =
      members.map((q, _, _) => q -> seedValues.getOrElse(q, zero)).toMap
    val maxRounds = math.max(1, iterative.maxIter)

    // Convergence: numeric values compare by |Δ| < maxChange (strict, per Excel); non-numeric
    // results converge only on exact equality — GH-344: an error VALUE arriving as a member's
    // Right converges the moment it repeats (Error(Div0) == Error(Div0)), so error-valued cycle
    // members fixpoint like any other value with no code change.
    def changeBelowThreshold(prev: CellValue, next: CellValue): Boolean =
      (prev, next) match
        case (CellValue.Number(a), CellValue.Number(b)) => (a - b).abs < iterative.maxChange
        case (a, b) => a == b

    /** |Δ| for a numeric member that stayed numeric; None otherwise. */
    def numericDelta(prev: CellValue, next: CellValue): Option[BigDecimal] =
      (prev, next) match
        case (CellValue.Number(a), CellValue.Number(b)) => Some((a - b).abs)
        case _ => None

    @annotation.tailrec
    def loop(round: Int, prev: Map[QualifiedRef, CellValue]): FixpointOutcome =
      val overlaid = members.foldLeft(baseSheets) { case (sheets, (q, idx, expr)) =>
        sheets.updated(idx, sheets(idx).put(q.ref, CellValue.Formula(expr, prev.get(q))))
      }
      val tempWb = wb.copy(sheets = overlaid)
      val results: Map[QualifiedRef, Either[XLError, CellValue]] =
        members.map { (q, idx, expr) =>
          val tempSheet = overlaid(idx)
          val evaluatedCell =
            try
              rngOpt match
                case Some(rng) => tempSheet.evaluateCell(q.ref, pinnedClock, rng, Some(tempWb))
                case None => tempSheet.evaluateCell(q.ref, pinnedClock, Some(tempWb))
            catch
              case NonFatal(e) =>
                Left(XLError.FormulaError(expr, s"Evaluation threw ${e.getClass.getName}"))
          q -> evaluatedCell
        }.toMap
      val converged = members.forall { (q, _, _) =>
        results(q) match
          case Right(next) => changeBelowThreshold(prev.getOrElse(q, zero), next)
          case Left(_) => false
      }
      if converged || round >= maxRounds then
        val maxDelta = members.flatMap { (q, _, _) =>
          results(q).toOption.flatMap(next => numericDelta(prev.getOrElse(q, zero), next))
        }.maxOption
        FixpointOutcome(results, converged, round, maxDelta)
      else
        // A throwing/failing member holds its previous value for the next round so the rest of
        // the core keeps converging (GH-388 degradation, not unwinding).
        val next = members.map { (q, _, _) =>
          q -> results(q).getOrElse(prev.getOrElse(q, zero))
        }.toMap
        loop(round + 1, next)

    loop(1, seed)
