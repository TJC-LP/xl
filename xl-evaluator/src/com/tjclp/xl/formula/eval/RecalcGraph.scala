package com.tjclp.xl.formula.eval

import scala.collection.immutable.BitSet

import com.tjclp.xl.formula.graph.{CompactGraph, DependencyGraph, FormulaParses}
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

  /** The SCC condensation, in `DependencyGraph.qualifiedSccOrder`'s order. */
  lazy val components: Vector[DependencyGraph.Scc] =
    graph.condensation.map { (members, cyclic) =>
      DependencyGraph.Scc(members.iterator.map(table.qualified).toVector, cyclic)
    }

  /** `DependencyGraph.withinComponentOrder` over the reads among `members`. */
  def withinComponentOrder(members: Vector[QualifiedRef]): Vector[QualifiedRef] =
    val reads = graph.memberPrecedents(ids(members))
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
  def build(wb: Workbook, counters: RecalcCounters): RecalcGraph =
    new RecalcGraph(CompactGraph.build(wb, counters), counters)
