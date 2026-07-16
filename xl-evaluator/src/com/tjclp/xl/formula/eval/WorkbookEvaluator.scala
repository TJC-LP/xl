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
        case Some(CellValue.Formula(expr, _)) => expr
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
          init: (Vector[Sheet], Map[SheetName, Map[ARef, CellValue]], Vector[CellEvalError])
        ): (Vector[Sheet], Map[SheetName, Map[ARef, CellValue]], Vector[CellEvalError]) =
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
                      case Some(rng) => tempSheet.evaluateCell(q.ref, clock, rng, Some(tempWb))
                      case None => tempSheet.evaluateCell(q.ref, clock, Some(tempWb))
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

        // GH-373: iterative mode splits the order around the fixpoint — cells NOT downstream of
        // the cyclic core evaluate first, then the core iterates to convergence, then the core's
        // dependents evaluate off the converged values. `coreDependents` is dependent-closed, so
        // the stable partition preserves topological validity within each part (the deferDynamic
        // lemma). The default path is the identity split: everything in one pass, byte-identical
        // to pre-GH-373 behavior.
        val (preOrder, postOrder) =
          if iterating then ordered.partition(q => !coreDependents.contains(q))
          else (ordered, List.empty[QualifiedRef])

        val preState = evalPass(
          preOrder,
          (initialSheets, Map.empty[SheetName, Map[ARef, CellValue]], Vector.empty[CellEvalError])
        )
        val iterated = iterativeOpt match
          case Some(iterative) if cyclicCore.nonEmpty =>
            iterateCycles(
              wb,
              cyclicCore,
              iterative,
              clock,
              rngOpt,
              sheetIndex,
              formulaText,
              preState
            )
          case _ => preState
        val (_, evaluated, evalErrors) = evalPass(postOrder, iterated)

        // Cache computed values into formula cells on the original sheets; failed cells stay
        // uncached (Excel recalculates them on open). Track per sheet whether any recomputed
        // value actually differs from the pre-existing cache (a newly cached cell — prior
        // cache None — counts as a change).
        val cachedSheets: Vector[(Sheet, Boolean)] = wb.sheets.map { orig =>
          evaluated.getOrElse(orig.name, Map.empty).foldLeft((orig, false)) {
            case ((s, changed), (ref, computed)) =>
              s.cells.get(ref).map(_.value) match
                case Some(CellValue.Formula(expr, cached)) if !cached.contains(computed) =>
                  (s.put(ref, CellValue.Formula(expr, Some(computed))), true)
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
          errors = cycleErrors ++ blockedErrors ++ evalErrors
        )

  /**
   * GH-373: bounded Jacobi fixpoint over the cyclic core.
   *
   * Excel's iterative-calculation semantics: every member seeds to 0 (uninitialized cells), each
   * round evaluates every member against the PREVIOUS round's values (Jacobi — implemented by
   * overlaying each member's cell with `Formula(expr, Some(previousValue))`, so every reference to
   * a member, including self-references, reads the cache while `evaluateCell` re-evaluates the
   * formula text), and iteration stops when every member changes by less than `maxChange` or after
   * `maxIter` rounds — non-convergence KEEPS the last values with no error.
   *
   * The clock is pinned ONCE before the loop (`Clock.fixed`) so TODAY()/NOW() cannot re-tick
   * between rounds — one recalculation is one volatile generation, like Excel. Each member's
   * evaluation sits under the same NonFatal guard as the main pass (GH-388): a member that throws
   * holds its previous value for later rounds and, if still failing in the final round, degrades to
   * a per-cell error with its cell left uncached (dependents then surface the failure exactly like
   * any failed precedent).
   */
  private def iterateCycles(
    wb: Workbook,
    core: Set[QualifiedRef],
    iterative: IterativeCalc,
    clock: Clock,
    rngOpt: Option[Rng],
    sheetIndex: Map[SheetName, Int],
    formulaText: QualifiedRef => String,
    state: (Vector[Sheet], Map[SheetName, Map[ARef, CellValue]], Vector[CellEvalError])
  ): (Vector[Sheet], Map[SheetName, Map[ARef, CellValue]], Vector[CellEvalError]) =
    val (baseSheets, acc0, errs0) = state
    // Deterministic member order: Jacobi values are order-independent (all reads see the
    // previous round), but error vectors and sheet folds must be stable run to run.
    val members: List[(QualifiedRef, Int, String)] =
      core.toList
        .flatMap(q => sheetIndex.get(q.sheet).map(idx => (q, idx, formulaText(q))))
        .sortBy((q, _, _) => (q.sheet.value, q.ref.toA1))

    val pinnedClock = Clock.fixed(clock.today(), clock.now())
    val zero: CellValue = CellValue.Number(BigDecimal(0))
    val seed: Map[QualifiedRef, CellValue] = members.map((q, _, _) => q -> zero).toMap
    val maxRounds = math.max(1, iterative.maxIter)

    // Convergence: numeric values compare by |Δ| < maxChange (strict, per Excel); non-numeric
    // results converge only on exact equality.
    def changeBelowThreshold(prev: CellValue, next: CellValue): Boolean =
      (prev, next) match
        case (CellValue.Number(a), CellValue.Number(b)) => (a - b).abs < iterative.maxChange
        case (a, b) => a == b

    @annotation.tailrec
    def loop(
      round: Int,
      prev: Map[QualifiedRef, CellValue]
    ): Map[QualifiedRef, Either[XLError, CellValue]] =
      val overlaid = members.foldLeft(baseSheets) { case (sheets, (q, idx, expr)) =>
        sheets.updated(idx, sheets(idx).put(q.ref, CellValue.Formula(expr, prev.get(q))))
      }
      val tempWb = wb.copy(sheets = overlaid)
      val results: Map[QualifiedRef, Either[XLError, CellValue]] =
        members.map { (q, idx, _) =>
          val tempSheet = overlaid(idx)
          val evaluatedCell =
            try
              rngOpt match
                case Some(rng) => tempSheet.evaluateCell(q.ref, pinnedClock, rng, Some(tempWb))
                case None => tempSheet.evaluateCell(q.ref, pinnedClock, Some(tempWb))
            catch
              case NonFatal(e) =>
                Left(
                  XLError.FormulaError(formulaText(q), s"Evaluation threw ${e.getClass.getName}")
                )
          q -> evaluatedCell
        }.toMap
      val converged = members.forall { (q, _, _) =>
        results(q) match
          case Right(next) => changeBelowThreshold(prev.getOrElse(q, zero), next)
          case Left(_) => false
      }
      if converged || round >= maxRounds then results
      else
        // A throwing/failing member holds its previous value for the next round so the rest of
        // the core keeps converging (GH-388 degradation, not unwinding).
        val next = members.map { (q, _, _) =>
          q -> results(q).getOrElse(prev.getOrElse(q, zero))
        }.toMap
        loop(round + 1, next)

    val finalResults = loop(1, seed)
    // Fold the fixpoint into the pass state: successful members write their value into the temp
    // sheets (post-pass dependents read them) and the evaluated map (they cache at the end);
    // failed members keep their original uncached formula cell, so dependents recursively
    // evaluate and surface the underlying error — the same posture as a failed acyclic cell.
    members.foldLeft((baseSheets, acc0, errs0)) { case ((sheets, acc, errs), (q, idx, _)) =>
      finalResults.get(q) match
        case Some(Right(value)) =>
          (
            sheets.updated(idx, sheets(idx).put(q.ref, value)),
            acc.updated(q.sheet, acc.getOrElse(q.sheet, Map.empty) + (q.ref -> value)),
            errs
          )
        case Some(Left(error)) => (sheets, acc, errs :+ CellEvalError(q.sheet, q.ref, error))
        case None => (sheets, acc, errs) // unreachable: every member evaluates every round
    }
