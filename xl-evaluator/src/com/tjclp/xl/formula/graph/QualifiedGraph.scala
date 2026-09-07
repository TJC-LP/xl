package com.tjclp.xl.formula.graph

import scala.annotation.tailrec

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.formula.graph.DependencyGraph.{QualifiedRef, Scc}
import com.tjclp.xl.workbooks.Workbook

/**
 * ADR-017 §2.10: the bounded, cross-sheet inspection graph shared by the CLI (`cell`, `deps`,
 * `audit`) and the scripting prelude (`QualifiedGraph.of(wb)`).
 *
 * Two questions an agent tracing a broken number asks — "what does this cell read?" and "who reads
 * this cell?" — answered without ever materializing a full-column reference:
 *
 *   - `dependencies` (formula → the cells it reads) lists single-cell references exactly and
 *     expands a range only to the target sheet's OCCUPIED cells. `SUM(A:A)` over three values is
 *     three edges, not 1,048,576; an empty cell is never a node. Constants are included: they are
 *     the leaves where a wrong input lives. Every non-data-table formula is a key, an unparseable
 *     one with no edges, so [[sccs]] partitions exactly the formula cells.
 *   - `pointDependents` (cell → the formulas naming it individually) and `rangeReaders` (per sheet,
 *     each declared range and its readers) together answer the reverse question symbolically:
 *     [[dependentsOf]] unions the point readers with every range that contains the cell, so an
 *     EMPTY cell inside a summed range still reports the sum ("if I write here, what recomputes?").
 *     Keeping the reverse side symbolic is what makes the build O(formulas + distinct ranges)
 *     instead of O(readers × occupied cells).
 *
 * `precedents`/`dependents` walk the graph in layers: layer k holds the refs exactly k hops away
 * that no earlier layer (or the start cell) already listed, each layer sorted by
 * [[QualifiedGraph.byPosition]]; `depth <= 0` is unbounded. A cycle therefore terminates the walk
 * once every member has been seen. Pure and total: built from the workbook value alone.
 */
final case class QualifiedGraph(
  dependencies: Map[QualifiedRef, Set[QualifiedRef]],
  pointDependents: Map[QualifiedRef, Set[QualifiedRef]],
  rangeReaders: Map[SheetName, Vector[QualifiedGraph.RangeReaders]]
) derives CanEqual:

  /** Every formula cell the graph knows (data-table records excluded). */
  def formulas: Set[QualifiedRef] = dependencies.keySet

  /** The cells `q` reads directly: single refs exactly, ranges as their occupied cells. */
  def precedentsOf(q: QualifiedRef): Set[QualifiedRef] = dependencies.getOrElse(q, Set.empty)

  /** The formulas reading `q` directly — by name, or through a range that contains it. */
  def dependentsOf(q: QualifiedRef): Set[QualifiedRef] =
    rangeReaders
      .getOrElse(q.sheet, Vector.empty)
      .foldLeft(pointDependents.getOrElse(q, Set.empty)) { (acc, entry) =>
        if entry.range.contains(q.ref) then acc ++ entry.readers else acc
      }

  /** Precedent layers: layer k = the cells exactly k hops upstream; `depth <= 0` = unbounded. */
  def precedents(q: QualifiedRef, depth: Int): Vector[Vector[QualifiedRef]] =
    layers(q, depth, precedentsOf)

  /**
   * Dependent layers: layer k = the formulas exactly k hops downstream; `depth <= 0` = unbounded.
   */
  def dependents(q: QualifiedRef, depth: Int): Vector[Vector[QualifiedRef]] =
    layers(q, depth, dependentsOf)

  /**
   * The strongly connected components of the formula graph in dependency-first order
   * ([[DependencyGraph.qualifiedSccOrder]]); `filter(_.cyclic)` are the circular references.
   */
  lazy val sccs: Vector[Scc] = DependencyGraph.qualifiedSccOrder(dependencies)

  private def layers(
    start: QualifiedRef,
    depth: Int,
    step: QualifiedRef => Set[QualifiedRef]
  ): Vector[Vector[QualifiedRef]] =
    @tailrec
    def go(
      frontier: Vector[QualifiedRef],
      seen: Set[QualifiedRef],
      remaining: Int,
      acc: Vector[Vector[QualifiedRef]]
    ): Vector[Vector[QualifiedRef]] =
      if frontier.isEmpty || remaining == 0 then acc
      else
        val next = frontier.iterator.flatMap(step).filterNot(seen).toSet
        if next.isEmpty then acc
        else
          val layer = QualifiedGraph.sorted(next)
          go(layer, seen ++ next, remaining - 1, acc :+ layer)
    // A negative budget never reaches zero: unbounded.
    go(Vector(start), Set(start), if depth <= 0 then -1 else depth, Vector.empty)

object QualifiedGraph:

  /** One declared range on a sheet and the formulas that read it. */
  final case class RangeReaders(range: CellRange, readers: Set[QualifiedRef]) derives CanEqual

  val empty: QualifiedGraph = QualifiedGraph(Map.empty, Map.empty, Map.empty)

  /** Sheet name, then row, then column — the order every layer and every CLI listing uses. */
  val byPosition: Ordering[QualifiedRef] =
    Ordering.by(q => (q.sheet.value, q.ref.row.index0, q.ref.col.index0))

  def sorted(refs: Iterable[QualifiedRef]): Vector[QualifiedRef] =
    refs.toVector.sorted(using byPosition)

  /**
   * Build the graph from a workbook value. The forward side inverts the symbolic dependency index
   * ([[DependencyGraph.fromWorkbookDependencyIndex]]) — point references exactly, each declared
   * range expanded once (memoized by its normalized endpoints) to the target sheet's occupied
   * cells, costing min(|range|, |cells|) — and the reverse side is that index itself.
   */
  def of(wb: Workbook): QualifiedGraph =
    val index = DependencyGraph.fromWorkbookDependencyIndex(wb)
    // Reverse insertion so a duplicated sheet name keeps its FIRST position, like indexWhere.
    val cellsBySheet: Map[SheetName, Map[ARef, Cell]] =
      wb.sheets.reverseIterator.map(sheet => sheet.name -> sheet.cells).toMap
    val expansions =
      scala.collection.mutable.HashMap.empty[(SheetName, ARef, ARef), Set[QualifiedRef]]

    def occupied(sheet: SheetName, range: CellRange): Set[QualifiedRef] =
      expansions.getOrElseUpdate(
        (sheet, range.start, range.end),
        cellsBySheet.get(sheet).fold(Set.empty[QualifiedRef]) { cells =>
          val hits =
            if range.cellCount <= cells.size.toLong then range.cells.filter(cells.contains)
            else cells.keysIterator.filter(range.contains)
          hits.map(ref => QualifiedRef(sheet, ref)).toSet
        }
      )

    // formula → the cells it names individually (the point map, inverted)
    val points = index.pointDependents.foldLeft(Map.empty[QualifiedRef, Set[QualifiedRef]]) {
      case (acc, (cell, readers)) =>
        readers.foldLeft(acc) { (m, reader) =>
          m.updated(reader, m.getOrElse(reader, Set.empty) + cell)
        }
    }
    // formula → the occupied cells of every range it reads (one expansion per distinct range)
    val viaRanges = index.rangeDependents.foldLeft(Map.empty[QualifiedRef, Set[QualifiedRef]]) {
      case (acc, (sheet, entries)) =>
        entries.foldLeft(acc) { (m, entry) =>
          val cells = occupied(sheet, entry.range)
          if cells.isEmpty then m
          else
            entry.dependents.foldLeft(m) { (m2, reader) =>
              m2.updated(reader, union(m2.getOrElse(reader, Set.empty), cells))
            }
        }
    }
    val dependencies = wb.sheets.iterator.flatMap { sheet =>
      sheet.cells.iterator.collect {
        case (ref, cell) if isFormulaNode(cell.value) =>
          val q = QualifiedRef(sheet.name, ref)
          q -> union(points.getOrElse(q, Set.empty), viaRanges.getOrElse(q, Set.empty))
      }
    }.toMap
    val readers = index.rangeDependents.map { (sheet, entries) =>
      sheet -> entries
        .map(entry => RangeReaders(entry.range, entry.dependents))
        .sortBy(entry => rangeKey(entry.range))
    }
    QualifiedGraph(dependencies, index.pointDependents, readers)

  /** Data-table records are pinned value sources, never computation nodes (GH-430). */
  private def isFormulaNode(value: CellValue): Boolean = value match
    case CellValue.Formula(_, _, _: FormulaKind.DataTable) => false
    case CellValue.Formula(_, _, _) => true
    case _ => false

  private def rangeKey(range: CellRange): (Int, Int, Int, Int) =
    (range.start.row.index0, range.start.col.index0, range.end.row.index0, range.end.col.index0)

  /** Union that copies the smaller side into the larger one (a shared expansion stays shared). */
  private def union(a: Set[QualifiedRef], b: Set[QualifiedRef]): Set[QualifiedRef] =
    if a.isEmpty then b
    else if b.isEmpty then a
    else if a.size < b.size then b ++ a
    else a ++ b
