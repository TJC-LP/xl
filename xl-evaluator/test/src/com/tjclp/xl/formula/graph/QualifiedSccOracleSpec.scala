package com.tjclp.xl.formula.graph

import scala.annotation.tailrec

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.formula.graph.DependencyGraph.{QualifiedRef, Scc}
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * #891: `qualifiedSccOrder` and `qualifiedCyclicNodes` over dense ids and a CSR edge array
 * ([[SccKernel]]) answer exactly as the hash-map Tarjan and immutable condensation they replaced.
 * The oracle below is that implementation, kept verbatim apart from its name.
 */
class QualifiedSccOracleSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(400)

  // ----- oracle: the pre-#891 implementation -----

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def oracleForeachScc[K](dependencies: Map[K, Set[K]])(onScc: List[K] => Unit): Unit =
    final class Frame(val node: K, val successors: Iterator[K])
    var nextIndex = 0
    val indices = scala.collection.mutable.HashMap.empty[K, Int]
    val lowLinks = scala.collection.mutable.HashMap.empty[K, Int]
    val stack = scala.collection.mutable.ArrayBuffer.empty[K]
    val onStack = scala.collection.mutable.HashSet.empty[K]
    val frames = scala.collection.mutable.ArrayBuffer.empty[Frame]
    def push(v: K): Unit =
      indices(v) = nextIndex
      lowLinks(v) = nextIndex
      nextIndex += 1
      stack += v
      onStack += v
      frames += new Frame(v, dependencies.getOrElse(v, Set.empty).iterator)
    dependencies.keySet.foreach { root =>
      if !indices.contains(root) then
        push(root)
        while frames.nonEmpty do
          val frame = frames(frames.length - 1)
          if frame.successors.hasNext then
            val successor = frame.successors.next()
            if !indices.contains(successor) then push(successor)
            else if onStack.contains(successor) then
              lowLinks(frame.node) = math.min(lowLinks(frame.node), indices(successor))
          else
            frames.remove(frames.length - 1)
            val node = frame.node
            if lowLinks(node) == indices(node) then
              val members = List.newBuilder[K]
              var complete = false
              while !complete do
                val member = stack.remove(stack.length - 1)
                onStack -= member
                members += member
                complete = member == node
              onScc(members.result())
            if frames.nonEmpty then
              val parent = frames(frames.length - 1).node
              lowLinks(parent) = math.min(lowLinks(parent), lowLinks(node))
    }

  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def oracleCyclicNodes(deps: Map[QualifiedRef, Set[QualifiedRef]]): Set[QualifiedRef] =
    var cyclic = Set.empty[QualifiedRef]
    oracleForeachScc(deps) {
      case single :: Nil =>
        if deps.get(single).exists(_.contains(single)) then cyclic = cyclic + single
      case multi => cyclic = cyclic ++ multi
    }
    cyclic

  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def oracleSccOrder(dependencies: Map[QualifiedRef, Set[QualifiedRef]]): Vector[Scc] =
    var collected = Vector.empty[Scc]
    oracleForeachScc(dependencies) { scc =>
      val cyclic = scc match
        case single :: Nil => dependencies.get(single).exists(_.contains(single))
        case _ => true
      val kept = scc.filter(dependencies.contains)
      if kept.nonEmpty then
        collected = collected :+ Scc(kept.toVector.sortBy(QualifiedRef.gridKey), cyclic)
    }
    val comps = collected
    if comps.isEmpty then Vector.empty
    else
      val compOf: Map[QualifiedRef, Int] =
        comps.zipWithIndex.flatMap((c, i) => c.members.map(_ -> i)).toMap
      val (precedentsOf, dependentsOf) =
        comps.zipWithIndex.foldLeft((Map.empty[Int, Set[Int]], Map.empty[Int, Set[Int]])) {
          case (acc, (c, i)) =>
            c.members.foldLeft(acc) { case (acc2, u) =>
              dependencies.getOrElse(u, Set.empty).foldLeft(acc2) { case ((pre, dep), v) =>
                compOf.get(v) match
                  case Some(j) if j != i =>
                    (
                      pre.updated(i, pre.getOrElse(i, Set.empty) + j),
                      dep.updated(j, dep.getOrElse(j, Set.empty) + i)
                    )
                  case _ => (pre, dep)
              }
            }
        }
      val byKey: Ordering[Int] = Ordering.by(i => (comps(i).key, i))
      val emptyFrontier = scala.collection.immutable.TreeSet.empty[Int](using byKey)
      val degree0 = comps.indices.map(i => i -> precedentsOf.getOrElse(i, Set.empty).size).toMap
      val initial =
        emptyFrontier ++ comps.indices.filter(i => precedentsOf.getOrElse(i, Set.empty).isEmpty)
      @tailrec
      def drain(
        frontier: scala.collection.immutable.TreeSet[Int],
        degree: Map[Int, Int],
        acc: Vector[Int]
      ): Vector[Int] =
        frontier.headOption match
          case None => acc
          case Some(i) =>
            val (nextDegree, newlyReady) =
              dependentsOf.getOrElse(i, Set.empty).foldLeft((degree, emptyFrontier)) {
                case ((d, ready), j) =>
                  val remaining = d.getOrElse(j, 0) - 1
                  (d.updated(j, remaining), if remaining == 0 then ready + j else ready)
              }
            drain((frontier - i) ++ newlyReady, nextDegree, acc :+ i)
      drain(initial, degree0, Vector.empty).map(comps)

  // ----- generators -----

  // Sheet names whose String order differs from their case-folded order, rows across the 9/10
  // boundary, columns across Z/AA, and two cells off Excel's grid (only an unchecked ARef makes
  // them) so the kernel's off-grid id path is exercised.
  private val universe: Vector[QualifiedRef] =
    val sheets = Vector("b", "Z", "Mid").map(SheetName.unsafe)
    val onGrid = for
      s <- sheets
      row <- Vector(0, 1, 8, 9, 1048575)
      col <- Vector(0, 1, 25, 26, 16383)
    yield QualifiedRef(s, ARef.from0(col, row))
    onGrid ++ Vector(
      QualifiedRef(sheets(0), ARef.from0(-1, 3)),
      QualifiedRef(sheets(1), ARef.from0(20000, 2))
    )

  private def genGraph(width: Int): Gen[Map[QualifiedRef, Set[QualifiedRef]]] =
    for
      nodes <- Gen.pick(width, universe).map(_.toVector)
      // a few cells that are read but are not keys: the constants of the workbook graph
      leaves <- Gen.pick(4, universe).map(_.toVector)
      entries <- Gen
        .sequence[Vector[(QualifiedRef, Set[QualifiedRef])], (QualifiedRef, Set[QualifiedRef])](
          nodes.map { k =>
            Gen
              .choose(0, 4)
              .flatMap(m => Gen.listOfN(m, Gen.oneOf(nodes ++ leaves)))
              .map(vs => k -> vs.toSet)
          }
        )
    yield entries.toMap

  private val genAny: Gen[Map[QualifiedRef, Set[QualifiedRef]]] =
    Gen.choose(0, 40).flatMap(genGraph)

  property("qualifiedSccOrder equals the hash-map Tarjan and immutable condensation") {
    Prop.forAll(genAny) { deps =>
      val expected = oracleSccOrder(deps)
      val actual = DependencyGraph.qualifiedSccOrder(deps)
      Prop(actual == expected) :| s"expected $expected\nactual   $actual"
    }
  }

  property("qualifiedCyclicNodes equals the hash-map Tarjan") {
    Prop.forAll(genAny) { deps =>
      Prop(DependencyGraph.qualifiedCyclicNodes(deps) == oracleCyclicNodes(deps))
    }
  }

  test("a 100k-long chain and a 100k-member ring stay stack-safe and match the oracle") {
    val sheet = SheetName.unsafe("S")
    def q(i: Int) = QualifiedRef(sheet, ARef.from0(i % 64, i / 64))
    val n = 100000
    val chain = (1 until n).map(i => q(i) -> Set(q(i - 1))).toMap + (q(0) -> Set.empty)
    assertEquals(DependencyGraph.qualifiedSccOrder(chain).map(_.members.head), (0 until n).map(q))
    val ring = (0 until n).map(i => q(i) -> Set(q((i + 1) % n))).toMap
    val comps = DependencyGraph.qualifiedSccOrder(ring)
    assertEquals(comps.size, 1)
    assert(comps.head.cyclic)
    assertEquals(comps.head.members, (0 until n).map(q).toVector)
    assertEquals(DependencyGraph.qualifiedCyclicNodes(ring).size, n)
  }
