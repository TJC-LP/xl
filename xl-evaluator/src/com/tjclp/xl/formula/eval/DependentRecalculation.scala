package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError
import com.tjclp.xl.formula.graph.{DependencyGraph, ReferenceScan}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.formula.{Clock, Rng}

/**
 * Targeted recalculation after an edit: evaluate the edited formulas and every formula the edit can
 * affect, keep every other cache.
 *
 * #891: one pass over the recalculation graph ([[RecalcGraph]], the compact graph a whole-workbook
 * recalculation orders with), in the canonical order: the lexicographically smallest topological
 * order in grid order, restricted to the cone. A seeded RAND therefore draws in the same order as a
 * full recalculation would. The per-sheet and workbook `recalculateDependents` copies that ordered
 * by a hash-seeded FIFO and dropped evaluation errors are gone; `recalculateAfterEdit` is the one
 * entry point.
 */
object DependentRecalculation:

  /**
   * #891: the dirty cone of an edit to `seeds` on `wb` — the cells [[recalculateAfterEdit]]
   * evaluates (its formulas) and the CLI's cone-scoped writes refresh: see
   * [[RecalcGraph.editCone]].
   */
  private[xl] def editCone(wb: Workbook, seeds: Set[QualifiedRef]): Set[QualifiedRef] =
    if seeds.isEmpty then Set.empty
    else
      val graph = RecalcGraph.build(wb, RecalcCounters.disabled)
      coneOf(wb, graph, seeds, DependencyGraph.dynamicCells(wb, graph.parses))

  private def coneOf(
    wb: Workbook,
    graph: RecalcGraph,
    seeds: Set[QualifiedRef],
    dynamic: Set[QualifiedRef]
  ): Set[QualifiedRef] =
    val reaches = ReferenceScan.reaches(wb, DependencyGraph.unresolvedReaders(wb))
    graph.editCone(seeds, dynamic, reaches)

  /**
   * GH-504/GH-508: evaluate authored formulas as well as affected dependents and retain outcomes.
   * The affected set is [[editCone]]: the edit, its transitive dependents, every dynamic reader,
   * and — GH-606 — an unresolved reader only when its textual reach (`ReferenceScan`) contains a
   * dirty cell, or is unbounded. A missing graph edge cannot prove independence, but the formula's
   * text often can; readers it cannot bound stay always-dirty. Unaffected formulas retain their
   * caches and do not consume clock or randomness capabilities. Cycle members and the formulas they
   * block are reported, never evaluated; a failure withdraws the caches downstream of it.
   *
   * #767: `memo` is the pass's aggregate memo (the per-pass caches live in it); the parity laws
   * supply one with the caches switched off or with small slice thresholds. #891: `rng` is the
   * pass's randomness, drawn in the canonical order.
   */
  private[xl] def recalculateAfterEdit(
    wb: Workbook,
    sheetName: SheetName,
    modifiedRefs: Set[ARef],
    clock: Clock,
    memo: Evaluator.AggregateMemo = new Evaluator.AggregateMemo,
    rng: Rng = Rng.system
  ): RecalcResult =
    if modifiedRefs.isEmpty then RecalcResult.cacheResults(wb, Map.empty, Vector.empty)
    else
      val calculationClock = WorkbookEvaluator.pinnedCalculationClock(clock)
      val counters = memo.counters
      val graph = counters.time(RecalcCounters.Phase.GraphBuild)(RecalcGraph.build(wb, counters))
      // #790: evaluation reuses the build's parse of each formula
      memo.passCache.foreach(_.useParses(graph.parses))
      val dynamic = DependencyGraph.dynamicCells(wb, graph.parses)
      val seeds = modifiedRefs.map(QualifiedRef(sheetName, _))
      val affected = coneOf(wb, graph, seeds, dynamic).filter(graph.isFormula)
      val allCyclic = graph.cyclicCore
      val allBlocked = graph.dependents(allCyclic, Set.empty)
      val cyclic = allCyclic.intersect(affected)
      val skipped = cyclic ++ allBlocked.intersect(affected)

      def expression(q: QualifiedRef): String =
        wb(q.sheet).toOption.flatMap(_.cells.get(q.ref)).map(_.value) match
          case Some(CellValue.Formula(text, _, _)) => text
          case _ => q.ref.toA1

      val cycleErrors = skipped.toVector
        .sortBy(QualifiedRef.gridKey)
        .map { q =>
          val reason =
            if cyclic.contains(q) then "Circular reference"
            else "Blocked by an upstream circular reference"
          CellEvalError(q.sheet, q.ref, XLError.FormulaError(expression(q), reason))
        }
      counters.time(RecalcCounters.Phase.Kahn)(graph.order(allCyclic ++ allBlocked)) match
        case Left(unordered) =>
          // Unreachable: the cyclic core and its dependents left the order. Stay total anyway.
          val failures = ((unordered.candidates & affected) -- skipped).toVector
            .sortBy(QualifiedRef.gridKey)
            .map { q =>
              CellEvalError(
                q.sheet,
                q.ref,
                XLError.FormulaError(
                  expression(q),
                  s"Unresolvable order: ${unordered.remaining.size} formulas left"
                )
              )
            }
          RecalcResult.cacheResultsWith(
            wb,
            Map.empty,
            cycleErrors ++ failures,
            graph.dependents,
            Vector.empty,
            Map.empty,
            graph.parses
          )
        case Right(sorted) =>
          // GH-274/GH-301: dynamic readers and every static dependent of one evaluate last. A
          // stable partition of the canonical order, so the order stays dependency-first even
          // when the dynamic chain crosses sheets.
          val dynamicClosure =
            if dynamic.isEmpty then Set.empty[QualifiedRef]
            else (dynamic ++ graph.dependents(dynamic, Set.empty)).intersect(affected) -- skipped
          val order = sorted.order.filter(affected.contains)
          val ordered =
            order.filterNot(dynamicClosure.contains) ++ order.filter(dynamicClosure.contains)
          val stripBySheet = (dynamicClosure ++ skipped ++ allCyclic).groupMap(_.sheet)(_.ref)
          val initialSheets = wb.sheets.map { sheet =>
            SheetEvaluator.stripFormulaCaches(sheet, stripBySheet.getOrElse(sheet.name, Set.empty))
          }
          val positions = wb.sheets.zipWithIndex.map((sheet, i) => sheet.name -> i).toMap
          if Evaluator.mayReadSpills(wb) then memo.trackSpills()
          // #767: one slot per cone member in evaluation order. Cells outside the cone keep caches
          // the edit cannot affect; the stripped cells this pass never evaluates (cyclic and
          // blocked ones, in or out of the cone) refuse every slice over them.
          val slots = ordered.iterator.zipWithIndex.toMap
          memo.passCache.foreach(
            _.begin(
              Readiness.build(
                slots,
                stripBySheet.iterator.flatMap { (sheet, refs) =>
                  refs.iterator.map(QualifiedRef(sheet, _))
                }.toSet
              )
            )
          )
          def evaluatorFor(q: QualifiedRef): Evaluator =
            Evaluator.recalculationInstance(
              rng,
              memo,
              WorkbookEvaluator.callerPosition(slots, dynamicClosure, q),
              Some((q.sheet, q.ref))
            )
          val initial = WorkbookEvaluator.PassState(initialSheets, Map.empty, cycleErrors)
          val WorkbookEvaluator.PassState(_, values, errors, reasons) =
            counters.time(RecalcCounters.Phase.Eval) {
              ordered.foldLeft(initial) { (state, q) =>
                positions.get(q.sheet) match
                  case None => state
                  case Some(position) =>
                    val result =
                      EvalDefect.xlGuard(expression(q), Some(q.ref)) {
                        SheetEvaluator.evaluateCellOutcome(
                          state.sheets(position),
                          q.ref,
                          evaluatorFor(q),
                          calculationClock,
                          Some(wb.copy(sheets = state.sheets))
                        )
                      }
                    result match
                      case Right(outcome) => state.computed(position, q, outcome, memo.passCache)
                      case Left(error) => state.failed(position, q, error)
              }
            }
          counters.time(RecalcCounters.Phase.Finalize)(
            RecalcResult.cacheResultsWith(
              wb,
              values,
              errors,
              graph.dependents,
              Vector.empty,
              reasons,
              graph.parses
            )
          )
