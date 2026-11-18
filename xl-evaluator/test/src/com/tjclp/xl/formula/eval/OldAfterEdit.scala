package com.tjclp.xl.formula.eval

import scala.annotation.tailrec

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.graph.{DependencyGraph, ReferenceScan}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.workbooks.Workbook

/**
 * #891 (A2 oracle, test-only): the after-edit pass as it was before it moved onto the compact graph
 * — the per-cell formula graph, the declared-range dependency index for the cone, a hash-seeded
 * FIFO order and the reverse-map invalidation walk — kept verbatim to compare the new pass against.
 * Goes when the per-cell graph builders leave the main sources.
 */
object OldAfterEdit:

  private def withReverse(
    wb: Workbook,
    successful: Map[SheetName, Map[ARef, CellValue]],
    failures: Vector[CellEvalError],
    dependents: Map[QualifiedRef, Set[QualifiedRef]],
    reasons: Map[QualifiedRef, String]
  ): RecalcResult =
    RecalcResult.cacheResultsWith(
      wb,
      successful,
      failures,
      (failed, pinned) =>
        DependencyGraph.qualifiedTransitiveDependents(
          dependents.view.mapValues(_ -- pinned).toMap,
          failed
        ),
      Vector.empty,
      reasons
    )

  def oldEditCone(
    workbook: Workbook,
    seeds: Set[QualifiedRef],
    index: DependencyGraph.QualifiedDependencyIndex,
    dynamic: Set[QualifiedRef]
  ): Set[QualifiedRef] =
    val reaches = ReferenceScan.reaches(workbook, DependencyGraph.unresolvedReaders(workbook))
    val unbounded = reaches.collect { case (reader, ReferenceScan.Reach.Unbounded) => reader }.toSet
    val bounded = reaches.toVector.collect { case (reader, areas: ReferenceScan.Reach.Areas) =>
      reader -> (areas: ReferenceScan.Reach)
    }
    val roots = seeds ++ dynamic ++ unbounded

    @tailrec
    def close(
      closure: Set[QualifiedRef],
      pending: Vector[(QualifiedRef, ReferenceScan.Reach)]
    ): Set[QualifiedRef] =
      val cellsBySheet =
        closure
          .groupMap(_.sheet)(_.ref)
          .map((sheet, cells) => sheet -> ReferenceScan.SheetCells.of(cells))
      val (joining, waiting) = pending
        .filterNot((reader, _) => closure.contains(reader))
        .partition((_, reach) => reach.touches(cellsBySheet))
      if joining.isEmpty then closure
      else
        val readers = joining.iterator.map(_._1).toSet
        // dependents of a union are the union of dependents: extend rather than recompute
        close(closure ++ readers ++ index.transitiveDependents(readers), waiting)

    close(roots ++ index.transitiveDependents(roots), bounded)

  def oldRecalculateAfterEdit(
    wb: Workbook,
    sheetName: SheetName,
    modifiedRefs: Set[ARef],
    clock: Clock,
    memo: Evaluator.AggregateMemo = new Evaluator.AggregateMemo
  ): RecalcResult =
    if modifiedRefs.isEmpty then RecalcResult.cacheResults(wb, Map.empty, Vector.empty)
    else
      val calculationClock = WorkbookEvaluator.pinnedCalculationClock(clock)
      val (graph, dependents) = DependencyGraph.fromWorkbookFormulaGraph(wb)
      val index = DependencyGraph.fromWorkbookDependencyIndex(wb)
      val dynamic = DependencyGraph.dynamicCells(wb)
      val seeds = modifiedRefs.map(QualifiedRef(sheetName, _))
      val affected = oldEditCone(wb, seeds, index, dynamic).intersect(graph.keySet)
      val affectedGraph = graph.filter((ref, _) => affected.contains(ref))
      val allCyclic = DependencyGraph.qualifiedCyclicNodes(graph)
      val cyclic = allCyclic.intersect(affected)
      val blocked =
        DependencyGraph.qualifiedTransitiveDependents(dependents, allCyclic).intersect(affected)
      val skipped = cyclic ++ blocked

      def expression(q: QualifiedRef): String =
        wb(q.sheet).toOption.flatMap(_.cells.get(q.ref)).map(_.value) match
          case Some(CellValue.Formula(text, _, _)) => text
          case _ => q.ref.toA1

      val cycleErrors = skipped.toVector
        .sortBy(q => (q.sheet.value, q.ref.row.index0, q.ref.col.index0))
        .map { q =>
          val reason =
            if cyclic.contains(q) then "Circular reference"
            else "Blocked by an upstream circular reference"
          CellEvalError(q.sheet, q.ref, XLError.FormulaError(expression(q), reason))
        }
      val orderedGraph = (affectedGraph -- skipped).view.mapValues(_ -- skipped).toMap
      val orderedDependents = (dependents.filter((ref, _) =>
        affected.contains(ref)
      ) -- skipped).view.mapValues(_ -- skipped).toMap
      DependencyGraph.qualifiedTopologicalSort(orderedGraph, orderedDependents) match
        case Left(error) =>
          val failures =
            orderedGraph.keys.toVector.sortBy(q => (q.sheet.value, q.ref.toA1)).map { q =>
              CellEvalError(
                q.sheet,
                q.ref,
                XLError.FormulaError(expression(q), s"Unresolvable order: $error")
              )
            }
          withReverse(wb, Map.empty, cycleErrors ++ failures, dependents, Map.empty)
        case Right(order) =>
          val dynamicClosure =
            (dynamic ++ DependencyGraph.qualifiedTransitiveDependents(dependents, dynamic))
              .intersect(affected) -- skipped
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
              Rng.system,
              memo,
              WorkbookEvaluator.callerPosition(slots, dynamicClosure, q),
              Some((q.sheet, q.ref))
            )
          val initial = WorkbookEvaluator.PassState(initialSheets, Map.empty, cycleErrors)
          val WorkbookEvaluator.PassState(_, values, errors, reasons) =
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
          withReverse(wb, values, errors, dependents, reasons)
