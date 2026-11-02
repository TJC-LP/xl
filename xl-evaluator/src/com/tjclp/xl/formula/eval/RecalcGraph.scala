package com.tjclp.xl.formula.eval

import scala.collection.immutable.BitSet

import com.tjclp.xl.formula.graph.{CompactGraph, DependencyGraph, FormulaParses}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.workbooks.Workbook

/**
 * #767: which graph a whole-workbook recalculation orders with. `Compact` (the default) is
 * [[CompactGraph]]: range and segment nodes, O((F + R) log rows) edges. `Legacy` is the per-cell
 * formula graph of `DependencyGraph.fromWorkbookFormulaGraph`, with O(formulas × range size) edges.
 * Both emit the same canonical order and the same results; `Legacy` stays one release as a rollback
 * and as the parity oracle, then goes.
 */
private[formula] enum RecalcEngine derives CanEqual:
  case Legacy, Compact

/**
 * The graph questions a recalculation pass asks, answered by either engine. Inputs and outputs are
 * formula cells; every order and every set is the same for both engines (the parity laws in
 * `CompactGraphSpec` and `RecalcEngineParitySpec`).
 */
private[eval] sealed trait RecalcGraph:

  /** Members of reference cycles: multi-node components and formulas that read themselves. */
  def cyclicCore: Set[QualifiedRef]

  /**
   * Every formula reachable from `seeds` through readers, the seeds excluded, never passing through
   * a formula of `skip`.
   */
  def dependents(seeds: Set[QualifiedRef], skip: Set[QualifiedRef]): Set[QualifiedRef]

  /**
   * The canonical order (the lexicographically smallest topological order in grid order) of every
   * formula outside `removed`, or the formulas it could not order.
   */
  def order(removed: Set[QualifiedRef]): Either[RecalcGraph.Unordered, RecalcGraph.Ordered]

  /** The SCC condensation, as `DependencyGraph.qualifiedSccOrder` lists it. */
  def components: Vector[DependencyGraph.Scc]

  /** `DependencyGraph.withinComponentOrder` over the reads among `members`. */
  def withinComponentOrder(members: Vector[QualifiedRef]): Vector[QualifiedRef]

  /** #790: the formula parses the build made, for the pass's evaluation to reuse. */
  def parses: FormulaParses

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

  /** The graph of `wb` under `engine`; `counters` receive its size and, for `Compact`, its work. */
  def build(wb: Workbook, engine: RecalcEngine, counters: RecalcCounters): RecalcGraph =
    engine match
      case RecalcEngine.Compact => Compact(CompactGraph.build(wb, counters), counters)
      case RecalcEngine.Legacy =>
        val (deps, dependents) = DependencyGraph.fromWorkbookFormulaGraph(wb, counters)
        Legacy(deps, dependents)

  /** `RecalcEngine.Compact`: questions answered over formula ids, converted at the boundary. */
  private final class Compact(graph: CompactGraph, counters: RecalcCounters) extends RecalcGraph:
    private val table = graph.nodes

    def parses: FormulaParses = graph.parses

    private def ids(refs: IterableOnce[QualifiedRef]): BitSet =
      BitSet.fromSpecific(refs.iterator.flatMap(table.lookup))

    private def refs(ids: IterableOnce[Int]): Set[QualifiedRef] =
      ids.iterator.map(table.qualified).toSet

    lazy val cyclicCore: Set[QualifiedRef] =
      val components = graph.components
      counters.add(RecalcCounters.Counter.GraphWork, components.work)
      refs(components.cyclic)

    def dependents(seeds: Set[QualifiedRef], skip: Set[QualifiedRef]): Set[QualifiedRef] =
      if seeds.isEmpty then Set.empty
      else
        val reach = graph.reach(ids(seeds), ids(skip))
        counters.add(RecalcCounters.Counter.GraphWork, reach.work)
        refs(reach.formulas)

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

    lazy val components: Vector[DependencyGraph.Scc] =
      graph.condensation.map { (members, cyclic) =>
        DependencyGraph.Scc(members.iterator.map(table.qualified).toVector, cyclic)
      }

    def withinComponentOrder(members: Vector[QualifiedRef]): Vector[QualifiedRef] =
      val reads = graph.memberPrecedents(ids(members))
      val dependencies: Map[QualifiedRef, Set[QualifiedRef]] =
        reads.map((reader, read) => table.qualified(reader) -> refs(read))
      DependencyGraph.withinComponentOrder(members, dependencies)

  /** `RecalcEngine.Legacy`: the per-cell formula graph, ordered canonically. */
  private final class Legacy(
    deps: Map[QualifiedRef, Set[QualifiedRef]],
    reverse: Map[QualifiedRef, Set[QualifiedRef]]
  ) extends RecalcGraph:

    /** The per-cell graph keeps no parse. */
    def parses: FormulaParses = FormulaParses.none

    lazy val cyclicCore: Set[QualifiedRef] = DependencyGraph.qualifiedCyclicNodes(deps)

    def dependents(seeds: Set[QualifiedRef], skip: Set[QualifiedRef]): Set[QualifiedRef] =
      val edges = if skip.isEmpty then reverse else reverse.view.mapValues(_ -- skip).toMap
      DependencyGraph.qualifiedTransitiveDependents(edges, seeds)

    def order(removed: Set[QualifiedRef]): Either[Unordered, Ordered] =
      val prunedDeps = (deps -- removed).view.mapValues(_ -- removed).toMap
      val prunedDependents = (reverse -- removed).view.mapValues(_ -- removed).toMap
      DependencyGraph.canonicalTopologicalSort(prunedDeps, prunedDependents) match
        case Left(remaining) => Left(Unordered(prunedDeps.keySet, remaining))
        case Right(sorted) =>
          Right(new Ordered:
            val order: List[QualifiedRef] = sorted

            @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
            def waves(subset: List[QualifiedRef]): Vector[Vector[QualifiedRef]] =
              val depthOf = scala.collection.mutable.HashMap.empty[QualifiedRef, Int]
              val waveBuilders =
                scala.collection.mutable.ArrayBuffer.empty[
                  scala.collection.mutable.ArrayBuffer[QualifiedRef]
                ]
              subset.foreach { q =>
                var d = 0
                prunedDeps.getOrElse(q, Set.empty).foreach { p =>
                  depthOf.get(p).foreach(parentDepth => d = math.max(d, parentDepth + 1))
                }
                depthOf(q) = d
                while waveBuilders.length <= d do
                  waveBuilders += scala.collection.mutable.ArrayBuffer.empty[QualifiedRef]
                waveBuilders(d) += q
              }
              waveBuilders.iterator.map(_.toVector).toVector)

    lazy val components: Vector[DependencyGraph.Scc] = DependencyGraph.qualifiedSccOrder(deps)

    def withinComponentOrder(members: Vector[QualifiedRef]): Vector[QualifiedRef] =
      DependencyGraph.withinComponentOrder(members, deps)
