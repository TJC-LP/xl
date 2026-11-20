package com.tjclp.xl.formula.graph

import scala.annotation.tailrec

import com.tjclp.xl.CellRange
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.formula.eval.RecalcCounters
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * #891: the per-cell workbook graphs and their hash-map walks, kept as test oracles once nothing in
 * the main sources used them: `fromWorkbookFormulaGraph` (the formula-to-formula graph the compact
 * graph is pinned against, CompactGraphSpec), `fromWorkbookBounded` (every cell read, ranges
 * clipped to the used range), and the transitive walks and cyclic-node set over such maps. Moved
 * from `DependencyGraph` verbatim but for the names they reach through it.
 */
object NaiveGraphs:

  /**
   * Reverse-reachability core shared by the sheet-level and workbook-level (qualified) variants —
   * generic in the node type (GH-346). Excludes the starting refs from the result.
   */
  @SuppressWarnings(Array("org.wartremover.warts.While"))
  def transitiveDependentsOf[K](
    dependents: Map[K, Set[K]],
    originalRefs: Set[K]
  ): Set[K] =
    // A persistent-Set frontier looks elegant but is catastrophically expensive on deep graphs:
    // every round rebuilds `frontier -- visited` and `visited ++ frontier`. A 100k linear cone can
    // therefore perform billions of hash probes. Keep mutation invocation-local and freeze once.
    val visited = scala.collection.mutable.HashSet.empty[K]
    val pending = scala.collection.mutable.ArrayDeque.empty[K]
    originalRefs.foreach { ref =>
      visited += ref
      pending.append(ref)
    }
    while pending.nonEmpty do
      val ref = pending.removeHead()
      dependents.getOrElse(ref, Set.empty).foreach { dependent =>
        if visited.add(dependent) then pending.append(dependent)
      }
    visited --= originalRefs
    visited.toSet

  /**
   * GH-346: workbook-level dependency graph with ranges bounded by each TARGET sheet's used range —
   * the qualified analog of `fromSheet`'s bounded extraction (a full-column reference must not
   * enumerate 1M+ cells). Single refs stay exact. Ranges over sheets absent from the workbook (or
   * empty ones) contribute no edges: they cannot affect evaluation order, and the referencing
   * formula still errors per cell on evaluation.
   *
   * Returns forward edges (ref → its dependencies) and reverse edges (ref → its dependents).
   */
  def fromWorkbookBounded(
    workbook: com.tjclp.xl.workbooks.Workbook
  ): (Map[QualifiedRef, Set[QualifiedRef]], Map[QualifiedRef, Set[QualifiedRef]]) =
    val bounds: Map[SheetName, Option[CellRange]] =
      workbook.sheets.map(s => s.name -> s.usedRange).toMap
    // GH-522: expand any (sheet, range) once per build — see memoizedCells
    val cellsFor: (SheetName, CellRange) => Set[QualifiedRef] = memoizedCells { (sheet, range) =>
      bounds.get(sheet) match
        case Some(Some(b)) =>
          range.intersect(b) match
            case Some(clipped) => clipped.cells.map(ref => QualifiedRef(sheet, ref)).toSet
            case None => Set.empty
        case _ => Set.empty // empty or missing sheet: nothing to depend on
    }

    val nameRanges = DependencyGraph.NameRanges.empty
    val dependencies = workbook.sheets.flatMap { sheet =>
      sheet.cells.iterator.flatMap { case cell @ Cell(cellRef, _, _, _) =>
        cell.value match
          // GH-430: dataTable-kind cells are pinned value sources, never computation nodes
          case CellValue.Formula(_, _, _: FormulaKind.DataTable) => None
          case CellValue.Formula(expression, _, _) =>
            val deps = FormulaParser.parse(expression) match
              case scala.util.Right(expr) =>
                DependencyGraph.extractQualifiedDependencies(
                  expr,
                  sheet.name,
                  cellsFor,
                  Some(workbook),
                  nameRanges = nameRanges,
                  reader = Some(cellRef)
                )
              case scala.util.Left(_) => Set.empty[QualifiedRef]
            Some(QualifiedRef(sheet.name, cellRef) -> deps)
          case _ => None
      }
    }.toMap

    (dependencies, reverseEdges(dependencies))

  /**
   * Build the formula-to-formula graph needed by whole-workbook recalculation.
   *
   * The public dependency graph records every referenced cell because impact analysis must answer
   * "which formulas read this edited constant?" Whole-book topological evaluation asks a much
   * narrower question: "which FORMULAS must run before this formula?" Materializing constant cells
   * in that graph is pure waste. On a common financial-model shape — 9,900 formulas sharing
   * `SUM($A$1:$A$5000)` over a constant driver column — the value-preserving graph has no range
   * edges at all, while [[fromWorkbookBounded]] retains and repeatedly traverses 49.5 million.
   *
   * Range lookup searches the target sheet's formula cells by column and row ([[FormulaColumns]])
   * and is memoized by range. Single-cell references are filtered against the same formula-node set
   * after name resolution. Consequently this graph is exactly `fromWorkbookBounded` with every
   * non-formula successor removed, without first allocating those successors. Formula evaluation
   * itself is unchanged and still reads the complete ranges from the workbook snapshot.
   *
   * This representation is intentionally internal: dirty-cone callers still need the public graph
   * (or, ultimately, a symbolic point/range index) to discover formulas affected by edited values.
   *
   * #767: `counters` times each formula parse as the `Parse` phase and records the graph's node and
   * edge counts; the disabled default does none of it.
   */
  def fromWorkbookFormulaGraph(
    workbook: com.tjclp.xl.workbooks.Workbook,
    counters: RecalcCounters = RecalcCounters.disabled
  ): (Map[QualifiedRef, Set[QualifiedRef]], Map[QualifiedRef, Set[QualifiedRef]]) =
    val formulas: Vector[(QualifiedRef, String)] = workbook.sheets.flatMap { sheet =>
      sheet.cells.iterator.flatMap { case cell @ Cell(cellRef, _, _, _) =>
        cell.value match
          case CellValue.Formula(_, _, _: FormulaKind.DataTable) => None
          case CellValue.Formula(expression, _, _) =>
            Some(QualifiedRef(sheet.name, cellRef) -> expression)
          case _ => None
      }
    }
    val formulaNodes = formulas.iterator.map(_._1).toSet
    val formulaColumns: Map[SheetName, FormulaColumns] =
      formulas.groupMap(_._1.sheet)(_._1.ref).view.mapValues(FormulaColumns(_)).toMap
    // counted at the memo miss, not inside the memo: a build that stopped sharing expansions
    // (GH-522) would count one lookup per reader again
    val cellsFor: (SheetName, CellRange) => Set[QualifiedRef] = memoizedCells { (sheet, range) =>
      formulaColumns.get(sheet) match
        case Some(columns) => columns.within(range).map(ref => QualifiedRef(sheet, ref)).toSet
        case None => Set.empty
    }

    val canonicalSheet = DependencyGraph.sheetCanonicaliser(workbook)
    val nameRanges = DependencyGraph.NameRanges.empty
    val dependencies = formulas.iterator.map { case (ref, expression) =>
      counters.increment(RecalcCounters.Counter.FormulaParses)
      val deps = counters.time(RecalcCounters.Phase.Parse)(FormulaParser.parse(expression)) match
        case scala.util.Right(expr) =>
          DependencyGraph
            .extractQualifiedDependencies(
              expr,
              ref.sheet,
              cellsFor,
              Some(workbook),
              preciseLookups = true,
              canonicalSheet = canonicalSheet,
              nameRanges = nameRanges,
              reader = Some(ref.ref)
            )
            .filter(formulaNodes.contains)
        case scala.util.Left(_) => Set.empty[QualifiedRef]
      ref -> deps
    }.toMap

    if counters.enabled then
      counters.add(RecalcCounters.Counter.GraphNodes, dependencies.size.toLong)
      counters.add(
        RecalcCounters.Counter.GraphEdges,
        dependencies.valuesIterator.map(_.size.toLong).sum
      )
    (dependencies, reverseEdges(dependencies))

  /**
   * GH-346: every node participating in a reference cycle of the workbook-level graph — the
   * qualified analog of `cyclicNodes`, sharing the iterative (stack-safe) Tarjan engine. Detects
   * cycles that span sheets, which per-sheet analysis cannot see.
   */
  def qualifiedCyclicNodes(
    dependencies: Map[QualifiedRef, Set[QualifiedRef]]
  ): Set[QualifiedRef] =
    DependencyGraph.qualifiedSccOrder(dependencies).filter(_.cyclic).flatMap(_.members).toSet

  /**
   * GH-346: transitive dependents over the workbook-level reverse edges (excludes the starting
   * refs) — the qualified analog of `transitiveDependents`.
   */
  def qualifiedTransitiveDependents(
    dependents: Map[QualifiedRef, Set[QualifiedRef]],
    refs: Set[QualifiedRef]
  ): Set[QualifiedRef] =
    transitiveDependentsOf(dependents, refs)

  /**
   * GH-453: transitive PRECEDENTS over the workbook-level graph (excludes the starting refs) — the
   * same generic reachability BFS run over the FORWARD edge map instead of the reverse one. Used by
   * the data-table seeder to gate a table's source formula on the cyclic core.
   */
  def qualifiedTransitivePrecedents(
    dependencies: Map[QualifiedRef, Set[QualifiedRef]],
    refs: Set[QualifiedRef]
  ): Set[QualifiedRef] =
    transitiveDependentsOf(dependencies, refs)

  /**
   * GH-522: expand any given (sheet, range) exactly once per graph build.
   *
   * A financial-model shape — many formulas each aggregating the same driver column — expanded the
   * SAME range once per referencing formula: 9,900 × `SUM($A$1:$A$5000)` built ~50M QualifiedRefs
   * (54 GB of a 135 GB recalc's allocation) where 5,000 suffice. The memo makes every referencing
   * formula share one immutable Set instance; correctness is untouched because expansion is a pure
   * function of (sheet, range) within one build (bounds are computed once up front and never change
   * mid-build). The captured mutable map is intentionally build-local and unsynchronized: callers
   * must invoke the returned closure only from the single graph-building thread.
   */
  /**
   * One sheet's formula cells by column, columns and each column's rows ascending, so the formula
   * cells inside a range cost two binary searches per covered column plus the cells found. A scan
   * of every formula per distinct range made a column of running totals, each reading its own
   * `B1:B{r}`, quadratic in the rows (#743).
   */
  private final class FormulaColumns(
    columns: scala.collection.immutable.ArraySeq[Int],
    rowsOf: Map[Int, scala.collection.immutable.ArraySeq[Int]]
  ):
    /**
     * The formula cells inside `range`, counting every column and row entry the walk reads into
     * `examined` (the first entry past the range included): the work of one lookup, which a scan of
     * the sheet's formulas would make proportional to the sheet instead.
     */
    def within(range: CellRange): Iterator[ARef] =
      val firstCol = range.colStart.index0
      val lastCol = range.colEnd.index0
      val firstRow = range.rowStart.index0
      val lastRow = range.rowEnd.index0
      Iterator
        .range(columns.search(firstCol).insertionPoint, columns.length)
        .map { i =>
          columns(i)
        }
        .takeWhile(_ <= lastCol)
        .flatMap { col =>
          val rows = rowsOf.getOrElse(col, scala.collection.immutable.ArraySeq.empty[Int])
          Iterator
            .range(rows.search(firstRow).insertionPoint, rows.length)
            .map { j =>
              rows(j)
            }
            .takeWhile(_ <= lastRow)
            .map(row => ARef.from0(col, row))
        }

  private object FormulaColumns:
    def apply(refs: Vector[ARef]): FormulaColumns =
      val rowsOf = refs
        .groupMap(_.col.index0)(_.row.index0)
        .view
        .mapValues(rows => scala.collection.immutable.ArraySeq.from(rows.distinct.sorted))
        .toMap
      new FormulaColumns(
        scala.collection.immutable.ArraySeq.from(rowsOf.keys.toVector.sorted),
        rowsOf
      )

  private def memoizedCells(
    expand: (SheetName, CellRange) => Set[QualifiedRef]
  ): (SheetName, CellRange) => Set[QualifiedRef] =
    // Anchors affect formula dragging, never the cells a dependency range denotes. Keying by the
    // normalized endpoints lets `$A$1:$A$5000`, `A1:A5000`, and mixed-anchor spellings share the
    // same expansion within one graph build.
    val cache =
      scala.collection.mutable.HashMap.empty[(SheetName, ARef, ARef), Set[QualifiedRef]]
    (sheet, range) => cache.getOrElseUpdate((sheet, range.start, range.end), expand(sheet, range))

  /**
   * GH-522: reverse-edge map built with a local mutable accumulator. The immutable foldLeft
   * allocated a fresh outer-Map node chain per edge — 50M times on the shape above (8.5 GB). The A
   * mutable insertion-ordered set per precedent also avoids building one intermediate immutable Set
   * per edge. Each set receives refs in the identical sequence as the fold it replaces, then
   * freezes through the ordinary immutable Set builder, so its observable iteration order is
   * unchanged. The outer map's iteration order is deliberately not a contract; every consumer
   * performs keyed lookup or set membership.
   */
  def reverseEdges(
    dependencies: Map[QualifiedRef, Set[QualifiedRef]]
  ): Map[QualifiedRef, Set[QualifiedRef]] =
    val acc = scala.collection.mutable.HashMap.empty[
      QualifiedRef,
      scala.collection.mutable.LinkedHashSet[QualifiedRef]
    ]
    dependencies.foreach { case (ref, deps) =>
      deps.foreach { dep =>
        acc.getOrElseUpdate(dep, scala.collection.mutable.LinkedHashSet.empty) += ref
      }
    }
    acc.iterator.map((ref, refs) => ref -> refs.toSet).toMap
