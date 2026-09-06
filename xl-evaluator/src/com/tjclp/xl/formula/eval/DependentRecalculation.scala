package com.tjclp.xl.formula.eval

import scala.util.control.NonFatal

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.formula.{Clock, Rng}

/**
 * Extension methods for eager recalculation of dependent formulas.
 *
 * When a cell value changes, all formulas that depend on it (directly or transitively) must be
 * re-evaluated to keep cached values current. This follows Excel's eager recalculation model.
 *
 * Usage:
 * {{{
 * import com.tjclp.xl.{*, given}
 *
 * // Single-sheet recalculation
 * val updatedSheet = sheet.put(ref"A1", 100).recalculateDependents(Set(ref"A1"))
 *
 * // Cross-sheet recalculation (workbook-level)
 * val updatedWb = wb.recalculateDependents(SheetName.unsafe("Sheet1"), Set(ref"A1"))
 * }}}
 */
object DependentRecalculation:

  /**
   * GH-504/GH-508: evaluate authored formulas as well as affected dependents and retain outcomes.
   * Unknown readers join the roots because a missing graph edge cannot prove independence.
   * Unaffected formulas retain their caches and do not consume clock or randomness capabilities.
   */
  private[xl] def recalculateAfterEdit(
    wb: Workbook,
    sheetName: SheetName,
    modifiedRefs: Set[ARef],
    clock: Clock
  ): RecalcResult =
    if modifiedRefs.isEmpty then RecalcResult.cacheResults(wb, Map.empty, Vector.empty, Map.empty)
    else
      val calculationClock = WorkbookEvaluator.pinnedCalculationClock(clock)
      val (graph, dependents) = DependencyGraph.fromWorkbookFormulaGraph(wb)
      val index = DependencyGraph.fromWorkbookDependencyIndex(wb)
      val dynamic = DependencyGraph.dynamicCells(wb)
      val roots = modifiedRefs.map(QualifiedRef(sheetName, _)) ++ dynamic ++
        DependencyGraph.unresolvedReaders(wb)
      val affected = (roots ++ index.transitiveDependents(roots)).intersect(graph.keySet)
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
          RecalcResult.cacheResults(wb, Map.empty, cycleErrors ++ failures, dependents)
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
          val evaluator = Evaluator.recalculationInstance(Rng.system, new Evaluator.AggregateMemo)
          val initial: WorkbookEvaluator.PassState =
            (initialSheets, Map.empty, cycleErrors)
          val (_, values, errors) = ordered.foldLeft(initial) {
            case (state @ (sheets, evaluated, failures), q) =>
              positions.get(q.sheet) match
                case None => state
                case Some(position) =>
                  val result =
                    try
                      SheetEvaluator.evaluateCellWithEvaluator(
                        sheets(position),
                        q.ref,
                        evaluator,
                        calculationClock,
                        Some(wb.copy(sheets = sheets))
                      )
                    catch
                      case NonFatal(error) =>
                        Left(
                          XLError.FormulaError(
                            expression(q),
                            s"Evaluation threw ${error.getClass.getName}"
                          )
                        )
                  result match
                    case Right(value) =>
                      (
                        sheets.updated(position, sheets(position).put(q.ref, value)),
                        evaluated.updated(
                          q.sheet,
                          evaluated.getOrElse(q.sheet, Map.empty) + (q.ref -> value)
                        ),
                        failures
                      )
                    case Left(error) =>
                      val stripped = SheetEvaluator.stripFormulaCaches(sheets(position), Set(q.ref))
                      (
                        sheets.updated(position, stripped),
                        evaluated,
                        failures :+ CellEvalError(q.sheet, q.ref, error)
                      )
          }
          RecalcResult.cacheResults(wb, values, errors, dependents)

  extension (sheet: Sheet)
    /**
     * Recalculate all formulas depending on modifiedRefs.
     *
     * Evaluates in topological order and updates cached values in Formula cells.
     *
     * GH-274: dynamic-reference cells (INDIRECT) are always dirty — the static graph cannot see
     * which cells their evaluated text names, so any edit may affect them. They and their static
     * dependents recalculate last (`DependencyGraph.deferDynamic`); earlier-refreshed caches are
     * exactly what dynamic reads should observe. This is the purity-charter encoding of Excel's
     * "volatile" marking. The entire dynamic closure has its old caches stripped before the pass,
     * so multi-hop INDIRECT/OFFSET chains cannot observe a previous calculation generation.
     *
     * @param modifiedRefs
     *   Set of cell references that have been modified
     * @param workbook
     *   Optional workbook context for cross-sheet formula references
     * @param clock
     *   Clock for date/time functions (defaults to system clock)
     * @return
     *   Sheet with updated formula caches
     */
    def recalculateDependents(
      modifiedRefs: Set[ARef],
      workbook: Option[Workbook] = None,
      clock: Clock = Clock.system
    ): Sheet =
      if modifiedRefs.isEmpty then sheet
      else
        val calculationClock = WorkbookEvaluator.pinnedCalculationClock(clock)
        val analysisWorkbook = workbook.fold(Workbook(sheet))(_.put(sheet))
        val seeds = modifiedRefs.map(ref => QualifiedRef(sheet.name, ref))
        val dynamic = DependencyGraph.dynamicCells(analysisWorkbook).filter(_.sheet == sheet.name)
        val (graph, dependents) = DependencyGraph.fromWorkbookFormulaGraph(analysisWorkbook)
        val dependencyIndex = DependencyGraph.fromWorkbookDependencyIndex(analysisWorkbook)
        val toRecalc =
          (dependencyIndex.transitiveDependents(seeds ++ dynamic) ++ dynamic)
            .filter(_.sheet == sheet.name)
        if toRecalc.isEmpty then sheet
        else
          val affectedGraph = graph.filter((q, _) => toRecalc.contains(q))
          val affectedDependents = dependents.filter((q, _) => toRecalc.contains(q))
          val cyclic = DependencyGraph.qualifiedCyclicNodes(affectedGraph)
          val blocked = DependencyGraph.qualifiedTransitiveDependents(affectedDependents, cyclic)
          val skipped = (cyclic ++ blocked) & toRecalc
          val orderGraph =
            (affectedGraph -- skipped).view.mapValues(_ -- skipped).toMap
          val orderDependents =
            (affectedDependents -- skipped).view.mapValues(_ -- skipped).toMap
          DependencyGraph.qualifiedTopologicalSort(orderGraph, orderDependents) match
            case Left(_) => sheet
            case Right(evalOrder) =>
              val dynamicClosure =
                if dynamic.isEmpty then Set.empty[QualifiedRef]
                else
                  (dynamic ++ DependencyGraph.qualifiedTransitiveDependents(dependents, dynamic))
                    .filter(_.sheet == sheet.name)
                    .diff(skipped)
              val ordered =
                evalOrder.filterNot(dynamicClosure.contains) ++
                  evalOrder.filter(dynamicClosure.contains)
              val initial = SheetEvaluator.stripFormulaCaches(
                sheet,
                dynamicClosure.iterator.map(_.ref).toSet
              )
              recalculateInOrder(
                initial,
                ordered.map(_.ref),
                Some(analysisWorkbook),
                calculationClock
              )

  extension (wb: Workbook)
    /**
     * Cross-sheet recalculation - handles formulas on other sheets.
     *
     * When Sheet1!B1 changes, this recalculates Sheet2!A1 if it references Sheet1!B1. Uses system
     * clock for date/time functions.
     *
     * @param sheetName
     *   The sheet where cells were modified
     * @param modifiedRefs
     *   Set of cell references that have been modified on that sheet
     * @return
     *   Workbook with updated formula caches across all affected sheets
     */
    def recalculateDependents(
      sheetName: SheetName,
      modifiedRefs: Set[ARef]
    ): Workbook =
      recalculateDependentsWithClock(sheetName, modifiedRefs, Clock.system)

    /**
     * Cross-sheet recalculation with explicit clock.
     *
     * Use this variant when you need deterministic date/time function results.
     */
    def recalculateDependentsWithClock(
      sheetName: SheetName,
      modifiedRefs: Set[ARef],
      clock: Clock
    ): Workbook =
      if modifiedRefs.isEmpty then wb
      else
        val calculationClock = WorkbookEvaluator.pinnedCalculationClock(clock)
        val qualifiedRefs = modifiedRefs.map(ref => QualifiedRef(sheetName, ref))
        // GH-274: dynamic-reference cells (INDIRECT) on ANY sheet are always dirty — their
        // resolved targets are invisible to the static graph, so every sheet's dynamic cells
        // join the seeds (and the recalc set) unconditionally.
        val dynamicQualified = DependencyGraph.dynamicCells(wb)
        // Ordering needs only formula-to-formula edges; dirty discovery uses symbolic declared
        // ranges so A:A never materializes one million cells and a just-cleared boundary ref still
        // reaches formulas that read it.
        val (graph, dependentsMap) = DependencyGraph.fromWorkbookFormulaGraph(wb)
        val dependencyIndex = DependencyGraph.fromWorkbookDependencyIndex(wb)
        val dynamicClosure =
          if dynamicQualified.isEmpty then Set.empty[QualifiedRef]
          else
            dynamicQualified ++ DependencyGraph.qualifiedTransitiveDependents(
              dependentsMap,
              dynamicQualified
            )
        val modifiedDependents =
          dependencyIndex.transitiveDependents(qualifiedRefs)
        // Modified seeds themselves are not recalculated unless they are dynamic (the historical
        // targeted-recalc contract); splitting the two closures avoids walking dynamic-only
        // branches twice while retaining the closure-of-union result.
        val toRecalc = ((modifiedDependents ++ dynamicClosure) -- qualifiedRefs) ++ dynamicQualified

        if toRecalc.isEmpty then wb
        else
          // Sort the affected workbook nodes once. Grouping by sheet is not valid here: a
          // downstream sheet can otherwise run before an upstream sheet and read its stale cache.
          // Restricting only the node map keeps stable precedents out of the work while Kahn still
          // sees every edge between affected formulas (it ignores successors absent from keySet).
          val affectedGraph = graph.filter((q, _) => toRecalc.contains(q))
          val affectedDependents = dependentsMap.filter((q, _) => toRecalc.contains(q))
          val cyclic = DependencyGraph.qualifiedCyclicNodes(affectedGraph)
          val blocked =
            DependencyGraph.qualifiedTransitiveDependents(affectedDependents, cyclic)
          val skipped = (cyclic ++ blocked) & toRecalc
          val orderGraph =
            (affectedGraph -- skipped).view.mapValues(_ -- skipped).toMap
          val orderDependents =
            (affectedDependents -- skipped).view.mapValues(_ -- skipped).toMap
          DependencyGraph.qualifiedTopologicalSort(orderGraph, orderDependents) match
            // Defensive totality: qualifiedCyclicNodes removed every cycle participant.
            case Left(_) => wb
            case Right(evalOrder) =>
              // GH-274/GH-301: dynamic cells and every static dependent of one evaluate last.
              // This is a workbook-wide stable partition, so the dependency-first order remains
              // valid even when the dynamic chain crosses sheet boundaries.
              val orderedMain = evalOrder.filterNot(dynamicClosure.contains)
              val orderedDynamic = evalOrder.filter(dynamicClosure.contains)
              recalculateQualifiedInOrder(
                wb,
                orderedMain ++ orderedDynamic,
                dynamicClosure -- skipped,
                calculationClock
              )

  /** Recalculate formulas in the given order, updating cached values */
  private def recalculateInOrder(
    sheet: Sheet,
    refs: List[ARef],
    workbook: Option[Workbook],
    clock: Clock
  ): Sheet =
    refs.foldLeft(sheet) { (s, ref) =>
      // A sheet-qualified self-reference resolves through the workbook context rather than the
      // direct `sheet` argument. Keep that view on the same threaded snapshot so targeted
      // recalculation cannot expose a stale copy of the sheet currently being refreshed.
      val currentWorkbook = workbook.map(_.put(s))
      recalculateOne(s, ref, currentWorkbook, clock)
    }

  /**
   * Recalculate workbook-qualified formulas in one global dependency-first pass.
   *
   * The `sheets` vector is threaded cell by cell so a cross-sheet reader sees every previously
   * refreshed precedent. Workbook/source modification tracking is applied once per changed sheet
   * after evaluation rather than rebuilding and marking the workbook for every formula.
   */
  private def recalculateQualifiedInOrder(
    workbook: Workbook,
    refs: List[QualifiedRef],
    dynamicClosure: Set[QualifiedRef],
    clock: Clock
  ): Workbook =
    val sheetIndex = workbook.sheets.zipWithIndex.map((s, i) => s.name -> i).toMap
    val changed = scala.collection.mutable.BitSet.empty
    val stripBySheet = dynamicClosure.groupMap(_.sheet)(_.ref)
    val initialSheets = workbook.sheets.map { sheet =>
      SheetEvaluator.stripFormulaCaches(
        sheet,
        stripBySheet.getOrElse(sheet.name, Set.empty)
      )
    }
    val recalculatedSheets = refs.foldLeft(initialSheets) { (sheets, q) =>
      sheetIndex.get(q.sheet) match
        case None => sheets
        case Some(index) =>
          val currentWorkbook = workbook.copy(sheets = sheets)
          val updated = recalculateOne(sheets(index), q.ref, Some(currentWorkbook), clock)
          changed += index
          sheets.updated(index, updated)
    }
    changed.foldLeft(workbook) { (current, index) =>
      current.put(recalculatedSheets(index))
    }

  /** Recalculate one formula cell, preserving pinned caches and formula record kinds. */
  private def recalculateOne(
    sheet: Sheet,
    ref: ARef,
    workbook: Option[Workbook],
    clock: Clock
  ): Sheet =
    sheet.cells
      .get(ref)
      .map { cell =>
        cell.value match
          // GH-353/GH-430: external-workbook and data-table formulas keep their Excel-written
          // cache verbatim — re-evaluating can only fail (the external workbook is not loaded;
          // TABLE(...) is a record, not evaluable text), and clearing the cache would destroy
          // the sole source of truth
          case value: CellValue.Formula if SheetEvaluator.pinnedCache(value).isDefined =>
            sheet
          case f @ CellValue.Formula(expr, _, _) =>
            // Evaluate the formula and update cache; .copy keeps the record kind (GH-430)
            val fullFormula = if expr.startsWith("=") then expr else s"=$expr"
            SheetEvaluator.evaluateFormula(sheet)(fullFormula, clock, workbook) match
              case Right(newValue) =>
                sheet.put(ref, f.copy(cachedValue = Some(newValue)))
              case Left(_) =>
                // Evaluation error - clear cache (will show error on next eval)
                sheet.put(ref, f.copy(cachedValue = None))
          case _ => sheet
      }
      .getOrElse(sheet)
