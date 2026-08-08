package com.tjclp.xl.formula.eval

import java.util.concurrent.atomic.AtomicInteger
import java.util.concurrent.{ExecutionException, Executors, Future, ThreadFactory, TimeUnit}

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

    /**
     * GH-520: non-iterative recalculation with independent formula regions evaluated on
     * `parallelism` threads.
     *
     * The single topological order is partitioned into longest-path depth classes ("waves") of the
     * workbook graph: two cells in the same wave can have no path between them, so evaluating a
     * wave's members concurrently against the pre-wave snapshot computes exactly the values the
     * sequential pass computes, and results fold into the pass state in the sequential order — the
     * workbook, evaluated map, and error vector are identical to `recalculate()`, element for
     * element. Dynamic-reference cells (INDIRECT/OFFSET) and their static dependents keep their
     * sequential evaluate-last pass (their reads resolve at evaluation time, outside the static
     * graph), and `parallelism <= 1` is exactly `recalculate()`.
     *
     * Determinism caveats, by construction rather than by luck:
     *   - No seeded-[[Rng]] variant is offered: a seeded generator draws in evaluation-order
     *     sequence, which parallelism would reorder. RAND under this entry point uses the
     *     thread-safe system generator ([[Rng.system]] semantics), as `recalculate()` does.
     *   - No iterative variant is offered: cyclic components fixpoint sequentially (GH-492) and
     *     books that need `IterativeCalc` should use the sequential overloads.
     *   - A dependence the static graph cannot see (a defined name whose refersTo the parser
     *     rejects — multi-area or intersection names, GH-468's blind-name class) is invisible to
     *     wave placement exactly as it is invisible to the sequential Kahn order; both paths
     *     evaluate such readers at an arbitrary point. Parity is guaranteed for every dependence
     *     the graph can express.
     */
    def recalculateParallel(parallelism: Int): RecalcResult =
      recalculateImpl(wb, Clock.system, None, None, parallelism)

    /** GH-520: parallel non-iterative recalculation with an explicit clock. */
    @annotation.targetName("recalculateParallelWithClock")
    def recalculateParallel(clock: Clock, parallelism: Int): RecalcResult =
      recalculateImpl(wb, clock, None, None, parallelism)

  // ========== recalculate internals ==========

  private def recalculateImpl(
    wb: Workbook,
    clock: Clock,
    rngOpt: Option[Rng],
    iterativeOpt: Option[IterativeCalc],
    parallelism: Int = 1
  ): RecalcResult =
    // One recalculate call is one volatile calculation generation. The lazy snapshot preserves
    // the old no-volatile fast path (the supplied clock is never touched), while ensuring an
    // arbitrary Clock is never invoked concurrently by wave workers.
    val calculationClock = pinnedCalculationClock(clock)
    // One narrowly scoped cache capability per non-iterative generation. The evaluator itself is
    // immutable and AggregateMemo is thread-safe, so one instance can serve sequential cells and
    // parallel wave workers. Iterative rounds deliberately change ranges and stay memo-free.
    val generationEvaluator = iterativeOpt match
      // A fixpoint intentionally revisits and changes the same ranges over multiple rounds, so a
      // one-pass generation cache is inapplicable there.
      case Some(_) => Evaluator.instance(rngOpt.getOrElse(Rng.system))
      case None =>
        Evaluator.recalculationInstance(
          rngOpt.getOrElse(Rng.system),
          new Evaluator.AggregateMemo
        )
    // Whole-book ordering needs only formula-to-formula edges. Constant cells are read during
    // evaluation but can never participate in a cycle or constrain formula order; excluding them
    // avoids O(formulas × range-size) graph construction and every downstream traversal.
    val (deps, dependents) = DependencyGraph.fromWorkbookFormulaGraph(wb)

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
        val dynamicAll: Set[QualifiedRef] = DependencyGraph.dynamicCells(wb)
        val dynamic: Set[QualifiedRef] = dynamicAll -- removed
        val bucket =
          if dynamic.isEmpty then Set.empty[QualifiedRef]
          else dynamic ++ DependencyGraph.qualifiedTransitiveDependents(prunedDependents, dynamic)
        val orderedMain =
          if bucket.isEmpty then evalOrder else evalOrder.filterNot(bucket.contains)
        val orderedBucket =
          if bucket.isEmpty then List.empty[QualifiedRef] else evalOrder.filter(bucket.contains)

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
        // One cell against a snapshot of the threaded sheets; None when the node names a sheet
        // absent from the workbook. Shared verbatim by the sequential fold and the parallel waves
        // so the two paths cannot diverge per cell.
        def evalOne(
          sheets: Vector[Sheet],
          tempWb: Workbook,
          q: QualifiedRef,
          clk: Clock
        ): Option[Either[XLError, CellValue]] =
          sheetIndex.get(q.sheet).map { idx =>
            val tempSheet = sheets(idx)
            // GH-388 defense in depth: recalculate is documented total — a numeric blowup
            // escaping a function implementation (e.g. BigDecimal scale overflow in a
            // diverging Newton loop) must degrade to this cell's per-cell error, never
            // unwind the whole recalculation.
            try
              SheetEvaluator.evaluateCellWithEvaluator(
                tempSheet,
                q.ref,
                generationEvaluator,
                clk,
                Some(tempWb)
              )
            catch
              case NonFatal(e) =>
                Left(
                  XLError.FormulaError(
                    formulaText(q),
                    s"Evaluation threw ${e.getClass.getName}"
                  )
                )
          }

        def foldResult(
          state: PassState,
          q: QualifiedRef,
          result: Either[XLError, CellValue]
        ): PassState =
          val (sheets, acc, errs) = state
          sheetIndex.get(q.sheet) match
            case None => state
            case Some(idx) =>
              result match
                case Right(value) =>
                  (
                    sheets.updated(idx, sheets(idx).put(q.ref, value)),
                    acc.updated(q.sheet, acc.getOrElse(q.sheet, Map.empty) + (q.ref -> value)),
                    errs
                  )
                case Left(error) => (sheets, acc, errs :+ CellEvalError(q.sheet, q.ref, error))

        def evalPass(
          order: List[QualifiedRef],
          init: PassState,
          clk: Clock
        ): PassState =
          order.foldLeft(init) { case (state @ (sheets, _, _), q) =>
            evalOne(sheets, wb.copy(sheets = sheets), q, clk) match
              case None => state // node names a sheet absent from the workbook
              case Some(result) => foldResult(state, q, result)
          }

        def failPass(
          order: IterableOnce[QualifiedRef],
          init: PassState,
          message: String
        ): PassState =
          order.iterator.foldLeft(init) { (state, q) =>
            foldResult(state, q, Left(XLError.FormulaError(formulaText(q), message)))
          }

        /**
         * GH-520: the wave-parallel equivalent of one `evalPass` over `order`.
         *
         * `order` is partitioned into longest-path depth classes over the pruned graph: if u
         * depends on v then depth(u) > depth(v), so two cells of one wave can have no path between
         * them and each wave member's evaluation reads only earlier-wave state. Every member of a
         * wave therefore evaluates against the SAME pre-wave snapshot — concurrently — and computes
         * exactly the value the sequential fold computes (an intra-wave write could only be
         * observed through an edge the graph does not have; dynamic reads are excluded from `order`
         * by the caller). Results fold into the pass state in the wave's sequential order, so
         * sheets, evaluated map and error vector come out element-for-element equal to
         * `evalPass(order, init, clk)`.
         *
         * Waves below [[ParallelWaveCutoff]] run through the sequential fold — chain-shaped regions
         * would otherwise pay thread-handoff latency per cell. Worker results are published by
         * `Future.get`'s happens-before edge; the atomic cursor assigns each result slot once.
         */
        @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
        def evalWaves(
          order: List[QualifiedRef],
          init: PassState,
          clk: Clock,
          threads: Int
        ): (PassState, Option[String]) =
          val orderVec = order.toVector
          if orderVec.isEmpty then (init, None)
          else
            val depthOf = scala.collection.mutable.HashMap.empty[QualifiedRef, Int]
            val waveBuilders =
              scala.collection.mutable.ArrayBuffer.empty[
                scala.collection.mutable.ArrayBuffer[QualifiedRef]
              ]
            orderVec.foreach { q =>
              var d = 0
              prunedDeps.getOrElse(q, Set.empty).foreach { p =>
                depthOf.get(p).foreach(parentDepth => d = math.max(d, parentDepth + 1))
              }
              depthOf(q) = d
              while waveBuilders.length <= d do
                waveBuilders += scala.collection.mutable.ArrayBuffer.empty[QualifiedRef]
              waveBuilders(d) += q
            }
            val waves: Vector[Vector[QualifiedRef]] =
              waveBuilders.iterator.map(_.toVector).toVector
            val widestParallelWave =
              waves.iterator.filter(_.length >= ParallelWaveCutoff).map(_.length).maxOption
            widestParallelWave match
              case None => (evalPass(order, init, clk), None)
              case Some(width) => evalParallelWaves(waves, init, clk, math.min(threads, width))

        @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
        def evalParallelWaves(
          waves: Vector[Vector[QualifiedRef]],
          init: PassState,
          clk: Clock,
          threads: Int
        ): (PassState, Option[String]) =
          val threadFactory: ThreadFactory = runnable =>
            val thread = new Thread(
              runnable,
              s"xl-recalc-worker-${ParallelWorkerIds.incrementAndGet()}"
            )
            // Every exit path shuts the pool down. Daemon status is the final lifecycle
            // backstop for a function implementation that ignores interruption.
            thread.setDaemon(true)
            thread
          val pool = Executors.newFixedThreadPool(threads, threadFactory)
          var hardStopped = false

          def cancelOutstanding(futures: java.util.ArrayList[Future[?]]): Unit =
            var i = 0
            while i < futures.size() do
              futures.get(i).cancel(true)
              i += 1

          def stopNow(
            futures: java.util.ArrayList[Future[?]],
            restoreInterrupt: Boolean
          ): Unit =
            hardStopped = true
            cancelOutstanding(futures)
            pool.shutdownNow()
            var interrupted = restoreInterrupt
            try pool.awaitTermination(ParallelShutdownWaitSeconds, TimeUnit.SECONDS)
            catch case _: InterruptedException => interrupted = true
            if interrupted then Thread.currentThread().interrupt()

          try
            var state = init
            var waveIndex = 0
            var abortReason = Option.empty[String]
            while waveIndex < waves.length && abortReason.isEmpty do
              val wave = waves(waveIndex)
              if wave.length < ParallelWaveCutoff then state = evalPass(wave.toList, state, clk)
              else
                val (sheets, _, _) = state
                val snapshotWb = wb.copy(sheets = sheets)
                val results: Array[Option[Either[XLError, CellValue]]] =
                  Array.fill(wave.length)(None)
                // Formula costs are heterogeneous (a SUM over 5,000 cells and a scalar add can
                // share a wave). A shared cursor gives idle workers another cell instead of
                // pinning them behind the slowest static chunk; result slots and the fold order
                // remain deterministic.
                val nextIndex = AtomicInteger(0)
                val workerCount = math.min(threads, wave.length)
                val futures = new java.util.ArrayList[Future[?]]()
                try
                  var worker = 0
                  while worker < workerCount do
                    val task: Runnable = () =>
                      var i = nextIndex.getAndIncrement()
                      while i < wave.length do
                        results(i) = evalOne(sheets, snapshotWb, wave(i), clk)
                        i = nextIndex.getAndIncrement()
                    futures.add(pool.submit(task))
                    worker += 1

                  var futureIndex = 0
                  while futureIndex < futures.size() do
                    futures.get(futureIndex).get()
                    futureIndex += 1
                catch
                  case _: InterruptedException =>
                    val reason = "Parallel recalculation interrupted"
                    stopNow(futures, restoreInterrupt = true)
                    abortReason = Some(reason)
                  case execution: ExecutionException =>
                    Option(execution.getCause) match
                      // InterruptedException is not NonFatal in Scala because ordinary callers
                      // must preserve their own interrupt. Here it belongs to a worker Future;
                      // cancel the wave and report it without spuriously interrupting the caller.
                      case Some(_: InterruptedException) | Some(NonFatal(_)) | None =>
                        val reason = "Parallel recalculation worker failed"
                        stopNow(futures, restoreInterrupt = false)
                        abortReason = Some(reason)
                      case Some(fatal) =>
                        stopNow(futures, restoreInterrupt = false)
                        throw fatal
                  case NonFatal(_) =>
                    val reason = "Parallel recalculation worker failed"
                    stopNow(futures, restoreInterrupt = false)
                    abortReason = Some(reason)

                if abortReason.isEmpty then
                  state = wave.indices.foldLeft(state) { (st, i) =>
                    results(i).fold(st)(foldResult(st, wave(i), _))
                  }

              if abortReason.isEmpty then waveIndex += 1

            abortReason match
              case None => (state, None)
              case Some(reason) =>
                val unfinished = waves.drop(waveIndex).iterator.flatten
                (failPass(unfinished, state, reason), Some(reason))
          finally
            if !hardStopped then
              pool.shutdown()
              try
                if !pool.awaitTermination(ParallelShutdownWaitSeconds, TimeUnit.SECONDS) then
                  pool.shutdownNow()
                  pool.awaitTermination(ParallelShutdownWaitSeconds, TimeUnit.SECONDS)
              catch
                case _: InterruptedException =>
                  pool.shutdownNow()
                  Thread.currentThread().interrupt()

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
            jacobiFixpoint(
              wb,
              baseSheets,
              members,
              iterative,
              pinnedClock,
              rngOpt,
              seed,
              Some(generationEvaluator)
            )
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
        // fixpoints AND in the acyclic bridges between them. Non-iterative recalculation uses the
        // same pinned generation so sequential and parallel paths observe identical volatile time.
        val (finalState, cycleReports) = iterativeOpt match
          case Some(iterative) if cyclicCore.nonEmpty =>
            val components = DependencyGraph.qualifiedSccOrder(deps)
            val walk =
              if bucketIter.isEmpty then components
              else
                val (eager, deferred) =
                  components.partition(c => !c.members.exists(bucketIter.contains))
                eager ++ deferred
            runPlan(buildPlan(walk), iterative, calculationClock, initState)
          case _ =>
            // GH-520: waves parallelize only the statically ordered region; the dynamic bucket
            // keeps its sequential evaluate-last pass (its reads resolve at evaluation time).
            // A seeded Rng forces the sequential fold — its draw sequence is evaluation-order.
            val threads =
              if rngOpt.isDefined then 1 else math.min(math.max(1, parallelism), MaxParallelism)
            val (mainState, abortReason) =
              if threads <= 1 then (evalPass(orderedMain, initState, calculationClock), None)
              else evalWaves(orderedMain, initState, calculationClock, threads)
            val completedState = abortReason match
              case None => evalPass(orderedBucket, mainState, calculationClock)
              case Some(reason) => failPass(orderedBucket, mainState, reason)
            (completedState, Vector.empty[SccReport])

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
   * Snapshot an explicit clock at most once per recalculation generation.
   *
   * The shared gate serializes first access to the caller's Clock, while separate lazy fields keep
   * the capability minimal: a TODAY-only workbook never calls `now`, and a NOW-only workbook never
   * calls `today`. Scala publishes each lazy result safely to every wave worker. Workbooks without
   * volatile time functions never force either snapshot.
   */
  private[eval] def pinnedCalculationClock(clock: Clock): Clock = new Clock:
    private val gate = new Object
    private lazy val generationToday = gate.synchronized(clock.today())
    private lazy val generationNow = gate.synchronized(clock.now())

    def today(): java.time.LocalDate = generationToday
    def now(): java.time.LocalDateTime = generationNow

  private val ParallelWorkerIds = AtomicInteger(0)
  private val ParallelShutdownWaitSeconds = 5L

  /** GH-520: hard cap on requested parallelism — a pool wider than this only adds contention. */
  private val MaxParallelism = 64

  /**
   * GH-520: waves narrower than this evaluate through the sequential fold. A chain-shaped region
   * produces thousands of single-cell waves; per-cell thread handoff would cost more than the
   * evaluation itself.
   */
  private val ParallelWaveCutoff = 16

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
   * GH-469: the members' currently cached NUMERIC values — Excel's iterative-calculation seed.
   * Members with no cache are omitted and fall back to 0 inside [[jacobiFixpoint]].
   *
   * Only `Number` caches seed. A stale error or text cache is deliberately NOT trusted: arithmetic
   * propagates both (`#DIV/0! * 0.5 + 10` is `#DIV/0!`; `"junk" * 0.5` is `#VALUE!`), so seeding
   * one into a perfectly healthy cycle makes the poison its own fixpoint — the run would wedge at
   * the seed, report `converged = true` with no error, and never heal. Falling back to 0 for every
   * non-numeric shape keeps the pre-GH-469 healing behavior for poisoned books while still fixing
   * GH-469's reported failure, whose caches are numeric.
   */
  private[eval] def warmSeed(
    baseSheets: Vector[Sheet],
    members: List[(QualifiedRef, Int, String)]
  ): Map[QualifiedRef, CellValue] =
    members.flatMap { (q, idx, _) =>
      baseSheets.lift(idx).flatMap(_.cells.get(q.ref)).map(_.value) match
        case Some(CellValue.Formula(_, Some(cached @ CellValue.Number(_)), _)) => Some(q -> cached)
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
    seedValues: Map[QualifiedRef, CellValue],
    generationEvaluator: Option[Evaluator] = None
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
              generationEvaluator match
                case Some(evaluator) =>
                  SheetEvaluator.evaluateCellWithEvaluator(
                    tempSheet,
                    q.ref,
                    evaluator,
                    pinnedClock,
                    Some(tempWb)
                  )
                case None =>
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
