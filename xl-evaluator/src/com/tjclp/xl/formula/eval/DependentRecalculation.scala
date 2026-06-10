package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.formula.Clock

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
     * "volatile" marking. Caveat: INDIRECT→INDIRECT chains may need the full
     * `Workbook.recalculate()` for guaranteed freshness (caches are refreshed, not stripped, here).
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
        val graph = DependencyGraph.fromSheet(sheet)
        val dynamic = DependencyGraph.dynamicCells(sheet)
        val toRecalc =
          DependencyGraph.transitiveDependents(graph, modifiedRefs ++ dynamic) ++ dynamic
        if toRecalc.isEmpty then sheet
        else
          // Get topological order and filter to only affected cells
          DependencyGraph.topologicalSort(graph) match
            case Right(evalOrder) =>
              val orderedToRecalc = DependencyGraph.deferDynamic(
                evalOrder.filter(toRecalc.contains),
                DependencyGraph.dynamicClosure(graph, dynamic)
              )
              recalculateInOrder(sheet, orderedToRecalc, workbook, clock)
            case Left(_) =>
              // Cycle detected - skip recalculation (formulas will show error on eval)
              sheet

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
        val qualifiedRefs = modifiedRefs.map(ref => QualifiedRef(sheetName, ref))
        // GH-274: dynamic-reference cells (INDIRECT) on ANY sheet are always dirty — their
        // resolved targets are invisible to the static graph, so every sheet's dynamic cells
        // join the seeds (and the recalc set) unconditionally.
        val dynamicBySheet: Map[SheetName, Set[ARef]] =
          wb.sheets.map(s => s.name -> DependencyGraph.dynamicCells(s)).toMap
        val dynamicQualified: Set[QualifiedRef] =
          dynamicBySheet.toSet.flatMap { case (name, refs) =>
            refs.map(ref => QualifiedRef(name, ref))
          }
        val graph = DependencyGraph.fromWorkbook(wb)
        val dependentsMap = buildDependentsMap(graph)
        val toRecalc =
          transitiveDependentsQualified(dependentsMap, qualifiedRefs ++ dynamicQualified) ++
            dynamicQualified

        if toRecalc.isEmpty then wb
        else
          // Group by sheet and recalculate each
          toRecalc.groupBy(_.sheet).foldLeft(wb) { case (currentWb, (targetSheet, refs)) =>
            currentWb(targetSheet).toOption
              .map { sheet =>
                val sheetGraph = DependencyGraph.fromSheet(sheet)
                DependencyGraph.topologicalSort(sheetGraph) match
                  case Right(evalOrder) =>
                    val localRefs = refs.map(_.ref)
                    // GH-274: same evaluate-last partition as the sheet variant
                    val orderedToRecalc = DependencyGraph.deferDynamic(
                      evalOrder.filter(localRefs.contains),
                      DependencyGraph.dynamicClosure(
                        sheetGraph,
                        dynamicBySheet.getOrElse(sheet.name, Set.empty)
                      )
                    )
                    val updatedSheet =
                      recalculateInOrder(sheet, orderedToRecalc, Some(currentWb), clock)
                    currentWb.put(updatedSheet)
                  case Left(_) => currentWb
              }
              .getOrElse(currentWb)
          }

  /** Recalculate formulas in the given order, updating cached values */
  private def recalculateInOrder(
    sheet: Sheet,
    refs: List[ARef],
    workbook: Option[Workbook],
    clock: Clock
  ): Sheet =
    refs.foldLeft(sheet) { (s, ref) =>
      s.cells
        .get(ref)
        .map { cell =>
          cell.value match
            case CellValue.Formula(expr, _) =>
              // Evaluate the formula and update cache
              val fullFormula = if expr.startsWith("=") then expr else s"=$expr"
              SheetEvaluator.evaluateFormula(s)(fullFormula, clock, workbook) match
                case Right(newValue) =>
                  s.put(ref, CellValue.Formula(expr, Some(newValue)))
                case Left(_) =>
                  // Evaluation error - clear cache (will show error on next eval)
                  s.put(ref, CellValue.Formula(expr, None))
            case _ => s
        }
        .getOrElse(s)
    }

  /** Build reverse lookup from forward dependency graph */
  private def buildDependentsMap(
    graph: Map[QualifiedRef, Set[QualifiedRef]]
  ): Map[QualifiedRef, Set[QualifiedRef]] =
    graph.foldLeft(Map.empty[QualifiedRef, Set[QualifiedRef]]) { case (acc, (ref, deps)) =>
      deps.foldLeft(acc) { (acc2, dep) =>
        acc2.updated(dep, acc2.getOrElse(dep, Set.empty) + ref)
      }
    }

  @scala.annotation.tailrec
  private def transitiveDependentsQualified(
    dependentsMap: Map[QualifiedRef, Set[QualifiedRef]],
    refs: Set[QualifiedRef],
    visited: Set[QualifiedRef] = Set.empty
  ): Set[QualifiedRef] =
    val toVisit = refs -- visited
    if toVisit.isEmpty then visited -- refs
    else
      val directDeps = toVisit.flatMap(ref => dependentsMap.getOrElse(ref, Set.empty))
      transitiveDependentsQualified(dependentsMap, directDeps, visited ++ toVisit)
