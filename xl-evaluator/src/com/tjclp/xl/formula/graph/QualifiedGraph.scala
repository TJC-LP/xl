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
 *     expands a range only to the target sheet's OCCUPIED cells, those holding a value (a
 *     style-only blank is not one). `SUM(A:A)` over three values is three edges, not 1,048,576; an
 *     empty cell inside a range is never a node. Constants are included: they are the leaves where
 *     a wrong input lives. Every non-data-table formula is a key, an unparseable one with no edges,
 *     so [[sccs]] partitions exactly the formula cells.
 *   - `pointDependents` (cell → the formulas naming it individually) and `rangeReaders` (per sheet,
 *     each declared range and its readers) together answer the reverse question symbolically:
 *     [[dependentsOf]] unions the point readers with every range that contains the cell, so an
 *     EMPTY cell inside a summed range still reports the sum ("if I write here, what recomputes?").
 *     Keeping the reverse side symbolic is what makes the build O(formulas + distinct ranges)
 *     instead of O(readers × occupied cells).
 *   - `rangeCells` (each declared range → its occupied cells) backs the DECLARED view, the way
 *     Excel's Trace Precedents draws a formula's inputs: [[declaredPrecedentsOf]] lists each cell
 *     the formula names and each range it reads as ONE [[QualifiedGraph.Node.Range]], whatever its
 *     size; [[occupiedIn]] gives that range's cells when they are wanted.
 *
 * `precedents`/`dependents` walk the graph in layers: layer k holds the refs exactly k hops away
 * that no earlier layer (or the start cell) already listed, each layer sorted by
 * [[QualifiedGraph.byPosition]]; `depth <= 0` is unbounded. A cycle therefore terminates the walk
 * once every member has been seen. [[declaredPrecedents]] walks the declared view the same way.
 * Pure and total: built from the workbook value alone.
 */
final case class QualifiedGraph(
  dependencies: Map[QualifiedRef, Set[QualifiedRef]],
  pointDependents: Map[QualifiedRef, Set[QualifiedRef]],
  rangeReaders: Map[SheetName, Vector[QualifiedGraph.RangeReaders]],
  rangeCells: Map[QualifiedGraph.DeclaredRange, Set[QualifiedRef]]
) derives CanEqual:
  import QualifiedGraph.{DeclaredRange, Node}

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

  /**
   * Formula → what it reads as declared, sorted by [[QualifiedGraph.nodeOrder]]: each cell it
   * names, and each range it reads as one node — a 1x1 range slot (`COUNTIF(C1,…)`) is a cell,
   * occupied or not; an empty range and a range on a sheet the book lacks are listed too. Anchors
   * and qualifier spellings never split a node. Derived from the symbolic index on first use, so
   * [[sccs]] and the audit never pay for it.
   */
  lazy val declared: Map[QualifiedRef, Vector[Node]] =
    val byReader =
      scala.collection.mutable.HashMap.empty[QualifiedRef, scala.collection.mutable.HashSet[Node]]
    def add(reader: QualifiedRef, node: Node): Unit =
      byReader.getOrElseUpdate(reader, scala.collection.mutable.HashSet.empty) += node
    pointDependents.foreach((cell, readers) => readers.foreach(add(_, Node.Cell(cell))))
    rangeReaders.foreach { (sheet, entries) =>
      entries.foreach { entry =>
        val node =
          if entry.range.start == entry.range.end then
            Node.Cell(QualifiedRef(sheet, entry.range.start))
          else Node.Range(DeclaredRange.of(sheet, entry.range))
        entry.readers.foreach(add(_, node))
      }
    }
    byReader.iterator.map { (reader, nodes) =>
      reader -> nodes.toVector.sorted(using QualifiedGraph.nodeOrder)
    }.toMap

  /** What `q` reads directly, as declared ([[declared]]). */
  def declaredPrecedentsOf(q: QualifiedRef): Vector[Node] = declared.getOrElse(q, Vector.empty)

  /**
   * Declared precedent layers; `depth <= 0` = unbounded. Depth still counts formula hops: a range
   * node at depth k stands for every occupied cell it covers, all k hops away, and the formulas
   * among them continue the walk at k+1 (counted in the range, not listed); constants are leaves. A
   * cell listed or covered by an EARLIER layer is never listed again, while a range reached later
   * is listed once even when an earlier layer covered all its cells; in one layer a named cell and
   * a range containing it are both listed. Each layer's covered cells are exactly the expanded
   * [[precedents]] layer's, plus the empty cells a 1x1 range slot names.
   */
  def declaredPrecedents(q: QualifiedRef, depth: Int): Vector[Vector[Node]] =
    @tailrec
    def go(
      frontier: Set[QualifiedRef],
      seenCells: Set[QualifiedRef],
      seenRanges: Set[DeclaredRange],
      remaining: Int,
      acc: Vector[Vector[Node]]
    ): Vector[Vector[Node]] =
      if frontier.isEmpty || remaining == 0 then acc
      else
        val next = frontier.iterator
          .flatMap(declaredPrecedentsOf)
          .filter {
            case Node.Cell(c) => !seenCells.contains(c)
            case Node.Range(r) => !seenRanges.contains(r)
          }
          .toSet
        if next.isEmpty then acc
        else
          val layer = next.toVector.sorted(using QualifiedGraph.nodeOrder)
          val covered = layer.iterator
            .flatMap {
              case Node.Cell(c) => Iterator.single(c)
              case Node.Range(r) => rangeCells.getOrElse(r, Set.empty).iterator
            }
            .filterNot(seenCells)
            .toSet
          go(
            covered.filter(dependencies.contains),
            seenCells ++ covered,
            seenRanges ++ layer.iterator.collect { case Node.Range(r) => r },
            remaining - 1,
            acc :+ layer
          )
    // A negative budget never reaches zero: unbounded.
    go(Set(q), Set(q), Set.empty, if depth <= 0 then -1 else depth, Vector.empty)

  /** The occupied cells of a declared range (anchors ignored); empty when no formula reads it. */
  def occupiedIn(range: DeclaredRange): Set[QualifiedRef] =
    rangeCells.getOrElse(DeclaredRange.of(range.sheet, range.range), Set.empty)

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

  /**
   * A range some formula reads, on one sheet: the corners only (anchors dropped, normalized), so
   * `$A$1:$A$3` and `A1:A3` — and `SUM(A:A)` and `SUM(A1:A1048576)` — are one range. Build one with
   * [[DeclaredRange.of]]: the constructor keeps anchors, and an anchored range neither equals the
   * graph's nodes nor keys [[QualifiedGraph.rangeCells]] (only `occupiedIn` drops them).
   */
  final case class DeclaredRange(sheet: SheetName, range: CellRange) derives CanEqual:
    /**
     * Unqualified, as Excel displays it: whole rows first (`3:5`, the whole sheet `1:1048576`),
     * then whole columns (`A:C`), a single cell as itself, else `B2:B9`.
     */
    def a1: String =
      if range.start == range.end then range.start.toA1
      else if range.isFullRow then s"${range.rowStart.index1}:${range.rowEnd.index1}"
      else if range.isFullColumn then s"${range.colStart.toLetter}:${range.colEnd.toLetter}"
      else range.toA1

    /** Qualified with the sheet quoted as a formula needs it, the way [[QualifiedRef]] prints. */
    override def toString: String = s"${SheetName.quoteForFormula(sheet.value)}!$a1"

  object DeclaredRange:
    /**
     * The range's corners on `sheet`, anchors dropped: the key [[QualifiedGraph.rangeCells]] uses.
     */
    def of(sheet: SheetName, range: CellRange): DeclaredRange =
      DeclaredRange(sheet, CellRange(range.start, range.end))

  /** One node of the declared precedent view: a cell the formula names, or a range it reads. */
  enum Node derives CanEqual:
    case Cell(ref: QualifiedRef)
    case Range(range: DeclaredRange)

    /** Qualified spelling: `Data!B2`, `Data!A:A`. */
    def label: String = this match
      // Node.Cell, never a bare Cell: the enum body sees the imported com.tjclp.xl.cells.Cell
      case Node.Cell(q) => q.toString
      case Node.Range(r) => r.toString

  val empty: QualifiedGraph = QualifiedGraph(Map.empty, Map.empty, Map.empty, Map.empty)

  /** Sheet name, then row, then column — the order every layer and every CLI listing uses. */
  val byPosition: Ordering[QualifiedRef] =
    Ordering.by(q => (q.sheet.value, q.ref.row.index0, q.ref.col.index0))

  /**
   * [[byPosition]] for declared nodes: sheet name, top row, left column, a cell before a range at
   * the same corner, then the bottom row and right column.
   */
  val nodeOrder: Ordering[Node] =
    Ordering.by[Node, (String, Int, Int, Int, Int, Int)] {
      case Node.Cell(q) => (q.sheet.value, q.ref.row.index0, q.ref.col.index0, 0, 0, 0)
      case Node.Range(r) =>
        (
          r.sheet.value,
          r.range.rowStart.index0,
          r.range.colStart.index0,
          1,
          r.range.rowEnd.index0,
          r.range.colEnd.index0
        )
    }

  def sorted(refs: Iterable[QualifiedRef]): Vector[QualifiedRef] =
    refs.toVector.sorted(using byPosition)

  /**
   * Build the graph from a workbook value. The forward side inverts the symbolic dependency index
   * ([[DependencyGraph.fromWorkbookDependencyIndex]]) — point references exactly, each declared
   * range expanded once (memoized by its normalized endpoints) to the target sheet's occupied
   * cells, costing min(|range|, |cells|) — and the reverse side is that index itself. `rangeCells`
   * keeps the memo: every declared range, an empty one or one on a missing sheet as `Set.empty`.
   */
  def of(wb: Workbook): QualifiedGraph =
    val index = DependencyGraph.fromWorkbookDependencyIndex(wb)
    // Reverse insertion so a duplicated sheet name keeps its FIRST position, like indexWhere.
    val cellsBySheet: Map[SheetName, Map[ARef, Cell]] =
      wb.sheets.reverseIterator.map(sheet => sheet.name -> sheet.cells).toMap
    val expansions =
      scala.collection.mutable.HashMap.empty[(SheetName, ARef, ARef), Set[QualifiedRef]]

    // Occupied = holds a value: a style-only record is kept in `Sheet.cells` as Empty, and is no
    // more an input than the blank beside it (the rule InMemorySource and the streaming reader use)
    def occupied(sheet: SheetName, range: CellRange): Set[QualifiedRef] =
      expansions.getOrElseUpdate(
        (sheet, range.start, range.end),
        cellsBySheet.get(sheet).fold(Set.empty[QualifiedRef]) { cells =>
          val hits =
            if range.cellCount <= cells.size.toLong then
              range.cells.filter(ref => cells.get(ref).exists(holdsValue))
            else
              cells.iterator.collect {
                case (ref, cell) if range.contains(ref) && holdsValue(cell) => ref
              }
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
    // the memo `viaRanges` filled: shared sets, no second expansion
    val rangeCells = index.rangeDependents.iterator.flatMap { (sheet, entries) =>
      entries.iterator.map(entry =>
        DeclaredRange.of(sheet, entry.range) -> occupied(sheet, entry.range)
      )
    }.toMap
    QualifiedGraph(dependencies, index.pointDependents, readers, rangeCells)

  private def holdsValue(cell: Cell): Boolean = cell.value match
    case CellValue.Empty => false
    case _ => true

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
