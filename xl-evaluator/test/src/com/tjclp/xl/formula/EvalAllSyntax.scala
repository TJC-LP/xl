package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.{IterativeMode, RecalcOptions, RecalcResult}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.graph.QualifiedGraph
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * #891: a whole-sheet evaluation in the shape the removed `sheet.evaluateWithDependencyCheck` had,
 * over the one public recalculation (`wb.recalculate`), so a spec reads `sheet.evalAll()`.
 *
 * The map holds every formula cell of the sheet with its computed value. A circular reference is a
 * `Left` (the old call failed fast on one; `recalculate` reports it per cell instead), and so is
 * any other per-cell failure (a host failure: the old fold stopped at the first). Excel error
 * values are values, as before.
 *
 * The graph questions the removed sheet-level `DependencyGraph` answered (`fromSheet`,
 * `detectCycles`, `precedents`, `transitiveDependents`) go to the public `QualifiedGraph.of` of the
 * one-sheet book: `sheet.cyclicCells`, `sheet.precedentsOf(ref)`, `sheet.dependentsClosure(refs)`,
 * `sheet.dynamicRefs`.
 */
object EvalAllSyntax:

  extension (sheet: Sheet)
    def evalAll(
      clock: Clock = Clock.system,
      workbook: Option[Workbook] = None,
      rng: Rng = Rng.system
    ): XLResult[Map[ARef, CellValue]] =
      val result = sheet.evalAllResult(clock, workbook, rng)
      val circular = result.errors.find(_.error.message.contains("Circular reference"))
      circular.orElse(result.errors.headOption) match
        case Some(failure) => Left(failure.error)
        case None => Right(result.evaluated.getOrElse(sheet.name, Map.empty))

    /** The whole recalculation of `sheet` (inside `workbook` when given), cycles isolated. */
    def evalAllResult(
      clock: Clock = Clock.system,
      workbook: Option[Workbook] = None,
      rng: Rng = Rng.system
    ): RecalcResult =
      workbook
        .fold(Workbook(sheet))(_.put(sheet))
        .recalculate(RecalcOptions(clock = clock, rng = rng, iterative = IterativeMode.Off))

    /** The public inspection graph of the one-sheet book. */
    def qualifiedGraph: QualifiedGraph = QualifiedGraph.of(Workbook(sheet))

    /** The formula cells of the sheet on a circular reference (empty when acyclic). */
    def cyclicCells: Set[ARef] =
      sheet.qualifiedGraph.sccs.filter(_.cyclic).flatMap(_.members).map(_.ref).toSet

    /** The cells `ref` reads directly (ranges as their occupied cells). */
    def precedentsOf(ref: ARef): Set[ARef] =
      sheet.qualifiedGraph.precedentsOf(QualifiedRef(sheet.name, ref)).map(_.ref)

    /** The formulas downstream of `refs`, transitively, `refs` themselves excluded. */
    def dependentsClosure(refs: Set[ARef]): Set[ARef] =
      val graph = sheet.qualifiedGraph
      refs.flatMap(r => graph.dependents(QualifiedRef(sheet.name, r), 0).flatten.map(_.ref)) -- refs

    /** The dynamic (INDIRECT/OFFSET-bearing) formula cells of the one-sheet book. */
    def dynamicRefs: Set[ARef] = DependencyGraph.dynamicCells(Workbook(sheet)).map(_.ref)

    /**
     * The removed fail-fast `sheet.evaluateForRange` over `evaluateForRangePerCell`: the range's
     * formula values, or the first failure in the range's closure as a `Left`.
     */
    def evalRange(
      range: CellRange,
      clock: Clock = Clock.system,
      workbook: Option[Workbook] = None
    ): XLResult[Map[ARef, CellValue]] =
      val result = sheet.evaluateForRangePerCell(range, clock, workbook)
      result.failures.headOption.fold(Right(result.values))(failure => Left(failure.error))
