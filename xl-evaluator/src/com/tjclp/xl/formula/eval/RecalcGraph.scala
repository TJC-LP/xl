package com.tjclp.xl.formula.eval

import scala.collection.immutable.BitSet

import com.tjclp.xl.formula.graph.{CompactGraph, DependencyGraph, FormulaParses, ReferenceScan}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.workbooks.Workbook

/**
 * The graph questions a recalculation pass asks, answered over [[CompactGraph]] (#767): range and
 * segment nodes, O((F + R) log rows) edges. Inputs and outputs are formula cells; the questions are
 * answered over formula ids and converted at the boundary.
 */
private[eval] final class RecalcGraph private (graph: CompactGraph, counters: RecalcCounters):
  import RecalcGraph.{Ordered, Unordered}

  private val table = graph.nodes

  /** #790: the formula parses the build made, for the pass's evaluation to reuse. */
  def parses: FormulaParses = graph.parses

  /** [[CompactGraph.narrowsOnWhatIfInput]]: this graph is not safe for a what-if pass. */
  def narrowsOnWhatIfInput: Boolean = graph.narrowsOnWhatIfInput

  private def ids(refs: IterableOnce[QualifiedRef]): BitSet =
    BitSet.fromSpecific(refs.iterator.flatMap(table.lookup))

  private def refs(ids: IterableOnce[Int]): Set[QualifiedRef] =
    ids.iterator.map(table.qualified).toSet

  /** Members of reference cycles: multi-node components and formulas that read themselves. */
  lazy val cyclicCore: Set[QualifiedRef] =
    val components = graph.components
    counters.add(RecalcCounters.Counter.GraphWork, components.work)
    refs(components.cyclic)

  /**
   * Every formula reachable from `seeds` through readers, the seeds excluded, never passing through
   * a formula of `skip`.
   */
  def dependents(seeds: Set[QualifiedRef], skip: Set[QualifiedRef]): Set[QualifiedRef] =
    if seeds.isEmpty then Set.empty
    else
      val reach = graph.reach(ids(seeds), ids(skip))
      counters.add(RecalcCounters.Counter.GraphWork, reach.work)
      refs(reach.formulas)

  /**
   * The canonical order (the lexicographically smallest topological order in grid order) of every
   * formula outside `removed`, or the formulas it could not order.
   */
  def order(removed: Set[QualifiedRef]): Either[Unordered, Ordered] =
    val removedIds = ids(removed)
    val kahn = graph.kahn(removedIds)
    counters.add(RecalcCounters.Counter.GraphWork, kahn.work)
    if kahn.remaining.isEmpty then
      Right(new Ordered:
        val order: List[QualifiedRef] = kahn.formulaOrder.iterator.map(table.qualified).toList

        def waves(subset: List[QualifiedRef]): Vector[Vector[QualifiedRef]] =
          val members = ids(subset)
          val depth = graph.waveDepths(kahn, members)
          val builders = Vector.newBuilder[Vector[QualifiedRef]]
          val byDepth = subset.groupBy(q => table.lookup(q).fold(0)(depth(_)))
          val deepest = if byDepth.isEmpty then -1 else byDepth.keysIterator.max
          (0 to deepest).foreach(d => builders += byDepth.getOrElse(d, Nil).toVector)
          builders.result())
    else
      val candidates = (0 until table.formulaCount).iterator.filterNot(removedIds.contains)
      Left(Unordered(refs(candidates), refs(kahn.remaining)))

  /**
   * #891: every formula `seeds` read, directly or transitively — the seeds excluded unless a cycle
   * leads back to them — never walking past a formula of `stop` (listed, never expanded, a seed's
   * included).
   */
  def precedents(seeds: Set[QualifiedRef], stop: Set[QualifiedRef]): Set[QualifiedRef] =
    if seeds.isEmpty then Set.empty
    else
      val closure = graph.precedentClosure(ids(seeds), ids(stop))
      counters.add(RecalcCounters.Counter.GraphWork, closure.work)
      refs(closure.formulas)

  /**
   * #891: every formula downstream of the cells `cells` (formulas or not): their readers, by a
   * point reference or through a range that contains them, and everything those reach — the
   * starting cells excluded.
   */
  def cellDependents(cells: Set[QualifiedRef]): Set[QualifiedRef] =
    if cells.isEmpty then Set.empty
    else
      val roots = scala.collection.mutable.BitSet.empty
      cells.foreach { q =>
        val sheet = table.sheetOf(q.sheet)
        if sheet >= 0 then
          table.lookup(q).foreach(roots += _)
          graph.foreachReaderOfCell(sheet, q.ref.row.index0, q.ref.col.index0)(roots += _)
      }
      val starts = ids(cells)
      val reach = graph.reach(roots.toImmutable, BitSet.empty)
      counters.add(RecalcCounters.Counter.GraphWork, reach.work)
      refs((roots.toImmutable -- starts) ++ reach.formulas) -- cells

  /** Whether `q` is a formula node of the graph. */
  def isFormula(q: QualifiedRef): Boolean = table.lookup(q).isDefined

  /**
   * #891: the dirty cone of an edit to the cells `seeds` (formulas or not): the seeds, every
   * formula that reads one (by a point reference or through a range that contains it) and every
   * formula downstream of those, every dynamic reader in `dynamic` with its dependents, and
   * (GH-606) every unresolved reader in `reaches` whose textual reach is unbounded or touches the
   * cone, with its dependents, iterated until no reader joins. Reads are the recalculation graph's:
   * a fixed-column VLOOKUP/HLOOKUP reads its key and result strips only (#504), so an edit
   * elsewhere in its table leaves it clean.
   */
  def editCone(
    seeds: Set[QualifiedRef],
    dynamic: Set[QualifiedRef],
    reaches: Map[QualifiedRef, ReferenceScan.Reach]
  ): Set[QualifiedRef] =
    val readers = scala.collection.mutable.BitSet.empty
    seeds.foreach { q =>
      val sheet = table.sheetOf(q.sheet)
      if sheet >= 0 then
        table.lookup(q).foreach(readers += _)
        graph.foreachReaderOfCell(sheet, q.ref.row.index0, q.ref.col.index0)(readers += _)
    }
    val unbounded = reaches.iterator.collect { case (reader, ReferenceScan.Reach.Unbounded) =>
      reader
    }
    readers ++= ids(dynamic.iterator ++ unbounded)

    def closed(roots: BitSet): BitSet =
      if roots.isEmpty then roots
      else
        val reach = graph.reach(roots, BitSet.empty)
        counters.add(RecalcCounters.Counter.GraphWork, reach.work)
        roots ++ reach.formulas

    val bounded = reaches.toVector.collect { case (reader, areas: ReferenceScan.Reach.Areas) =>
      reader -> (areas: ReferenceScan.Reach)
    }

    @scala.annotation.tailrec
    def close(
      cone: Set[QualifiedRef],
      pending: Vector[(QualifiedRef, ReferenceScan.Reach)]
    ): Set[QualifiedRef] =
      val cellsBySheet = cone
        .groupMap(_.sheet)(_.ref)
        .map((sheet, cells) => sheet -> ReferenceScan.SheetCells.of(cells))
      val (joining, waiting) = pending
        .filterNot((reader, _) => cone.contains(reader))
        .partition((_, reach) => reach.touches(cellsBySheet))
      if joining.isEmpty then cone
      else close(cone ++ refs(closed(ids(joining.iterator.map(_._1)))), waiting)

    close(seeds ++ refs(closed(readers.toImmutable)), bounded)

  /** The SCC condensation, in `DependencyGraph.qualifiedSccOrder`'s order. */
  lazy val components: Vector[DependencyGraph.Scc] =
    graph.condensation.map { (members, cyclic) =>
      DependencyGraph.Scc(members.iterator.map(table.qualified).toVector, cyclic)
    }

  /** `DependencyGraph.withinComponentOrder` over the reads among `members`. */
  def withinComponentOrder(members: Vector[QualifiedRef]): Vector[QualifiedRef] =
    val known = ids(members)
    // #891: on formula ids when the graph knows every member (always, for a cycle's members);
    // a member the graph does not hold has no reads, which only the general form expresses
    if known.size == members.distinct.size then
      graph.withinComponentOrder(known).iterator.map(table.qualified).toVector
    else
      val reads = graph.memberPrecedents(known)
      val dependencies: Map[QualifiedRef, Set[QualifiedRef]] =
        reads.map((reader, read) => table.qualified(reader) -> refs(read))
      DependencyGraph.withinComponentOrder(members, dependencies)

private[eval] object RecalcGraph:

  /** An order of the formulas outside a removal set. */
  trait Ordered:
    def order: List[QualifiedRef]

    /**
     * `subset` (a sub-sequence of [[order]] closed under precedents within it) partitioned into
     * longest-path depth classes, each in `subset`'s order: two cells of one wave have no path
     * between them.
     */
    def waves(subset: List[QualifiedRef]): Vector[Vector[QualifiedRef]]

  /** Kahn could not order `remaining`; `candidates` are all the formulas it was given. */
  final case class Unordered(candidates: Set[QualifiedRef], remaining: Set[QualifiedRef])

  /** The graph of `wb`; `counters` receive its size and its work. */
  def build(wb: Workbook, counters: RecalcCounters, whatIfSafe: Boolean = false): RecalcGraph =
    new RecalcGraph(CompactGraph.build(wb, counters, whatIfSafe), counters)
