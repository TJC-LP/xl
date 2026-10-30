package com.tjclp.xl.formula.graph

import scala.collection.immutable.{BitSet, TreeSet}

import com.tjclp.xl.CellRange
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.Prop
import org.scalacheck.rng.Seed

/**
 * #767: the compact graph answers every ordering question exactly as the legacy per-cell formula
 * graph (`fromWorkbookFormulaGraph`) does — reachability, cycles, the canonical order, the SCC
 * condensation, waves and within-component precedence — and finds each cell's direct readers.
 *
 * The order oracle is computed here, independently of the engines: the legacy graph pruned of the
 * cyclic core and its dependents, ordered by Kahn's algorithm with a grid-ordered (TreeSet) ready
 * set.
 */
class CompactGraphSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(150)

  private final case class Fixture(
    wb: Workbook,
    graph: CompactGraph,
    deps: Map[QualifiedRef, Set[QualifiedRef]],
    rev: Map[QualifiedRef, Set[QualifiedRef]]
  ):
    def id(q: QualifiedRef): Int = graph.nodes.lookup(q).getOrElse(-1)
    def ref(n: Int): QualifiedRef = graph.nodes.qualified(n)
    def refs(ids: IterableOnce[Int]): Set[QualifiedRef] = ids.iterator.map(ref).toSet
    def ids(qs: IterableOnce[QualifiedRef]): BitSet = BitSet.fromSpecific(qs.iterator.map(id))

  private def fixture(wb: Workbook): Fixture =
    val (deps, rev) = DependencyGraph.fromWorkbookFormulaGraph(wb)
    Fixture(wb, CompactGraph.build(wb), deps, rev)

  /** Kahn over the legacy graph with a grid-ordered ready set: the canonical order. */
  private def oracleOrder(
    deps: Map[QualifiedRef, Set[QualifiedRef]],
    removed: Set[QualifiedRef]
  ): List[QualifiedRef] =
    val nodes = deps.keySet -- removed
    val preds = nodes.iterator.map(q => q -> (deps(q) & nodes - q)).toMap
    @scala.annotation.tailrec
    def go(
      ready: TreeSet[QualifiedRef],
      left: Map[QualifiedRef, Set[QualifiedRef]],
      acc: List[QualifiedRef]
    ): List[QualifiedRef] =
      ready.headOption match
        case None => acc.reverse
        case Some(q) =>
          val waiting = left - q
          val updated = waiting.map((k, ps) => k -> (ps - q))
          val newlyReady = updated.collect {
            case (k, ps) if ps.isEmpty && waiting(k).nonEmpty => k
          }
          go(ready - q ++ newlyReady, updated, q :: acc)
    val initial = TreeSet.from(preds.collect { case (q, ps) if ps.isEmpty => q })(using
      QualifiedRef.gridOrdering
    )
    go(initial, preds, Nil)

  private def removedSet(f: Fixture): Set[QualifiedRef] =
    val core = DependencyGraph.qualifiedCyclicNodes(f.deps)
    core ++ DependencyGraph.qualifiedTransitiveDependents(f.rev, core)

  property("nodes: the legacy node set, numbered in grid order") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val f = fixture(wb)
      val listed = (0 until f.graph.formulaCount).map(f.ref).toVector
      Prop.all(
        Prop(listed.toSet == f.deps.keySet) :| "node set",
        Prop(listed == listed.sortBy(QualifiedRef.gridKey)) :| "grid order",
        Prop(listed.indices.forall(i => f.id(listed(i)) == i)) :| "lookup inverts qualified"
      )
    }
  }

  property("reachability: transitive dependents of every formula equal the legacy graph's") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val f = fixture(wb)
      Prop.all((0 until f.graph.formulaCount).map { n =>
        val q = f.ref(n)
        val compact = f.refs(f.graph.reach(BitSet(n), BitSet.empty).formulas)
        val legacy = DependencyGraph.qualifiedTransitiveDependents(f.rev, Set(q))
        Prop(compact == legacy) :| s"dependents of $q: $compact vs $legacy"
      }*)
    }
  }

  property("reachability with a skip set equals the legacy walk over pruned edges") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val f = fixture(wb)
      val all = f.deps.keySet.toVector.sortBy(QualifiedRef.gridKey)
      val seeds = all.zipWithIndex.collect { case (q, i) if i % 3 == 0 => q }.toSet
      val skip = all.zipWithIndex.collect { case (q, i) if i % 3 == 1 => q }.toSet
      val legacy = DependencyGraph.qualifiedTransitiveDependents(
        f.rev.view.mapValues(_ -- skip).toMap,
        seeds
      )
      val compact = f.refs(f.graph.reach(f.ids(seeds), f.ids(skip)).formulas)
      Prop(compact == legacy) :| s"$compact vs $legacy"
    }
  }

  property("cycles: the cyclic formulas equal the legacy graph's") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val f = fixture(wb)
      val compact = f.refs(f.graph.cyclicFormulas)
      val legacy = DependencyGraph.qualifiedCyclicNodes(f.deps)
      Prop(compact == legacy) :| s"$compact vs $legacy"
    }
  }

  property("order: Kahn after the prune equals the grid-ordered legacy oracle exactly") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val f = fixture(wb)
      val removed = removedSet(f)
      val kahn = f.graph.kahn(f.ids(removed))
      val compact = kahn.formulaOrder.iterator.map(f.ref).toList
      Prop.all(
        Prop(kahn.remaining.isEmpty) :| "everything outside the prune is ordered",
        Prop(compact == oracleOrder(f.deps, removed)) :| s"$compact"
      )
    }
  }

  property("order: with nothing removed, Kahn leaves multi-node cycles and all they reach") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val f = fixture(wb)
      val kahn = f.graph.kahn(BitSet.empty)
      val stuck = f.refs(kahn.remaining)
      // a point self-read is a flag, not an edge, so only Tarjan sees it; a range over the
      // reader's own cell is a real cycle through its range node
      val reads = readsOf(wb)
      def rangeSelf(q: QualifiedRef): Boolean =
        reads(q)._2.exists((s, r) => s == q.sheet && r.contains(q.ref))
      val deps = f.deps.map((q, ps) => q -> (if rangeSelf(q) then ps else ps - q))
      val rev = DependencyGraph.reverseEdges(deps)
      val core = DependencyGraph.qualifiedCyclicNodes(deps)
      val expected = core ++ DependencyGraph.qualifiedTransitiveDependents(rev, core)
      Prop(stuck == expected) :| s"unordered $stuck vs $expected"
    }
  }

  property("condensation: components, flags and order equal qualifiedSccOrder") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val f = fixture(wb)
      val compact = f.graph.condensation.map((members, cyclic) =>
        DependencyGraph.Scc(members.iterator.map(f.ref).toVector, cyclic)
      )
      val legacy = DependencyGraph.qualifiedSccOrder(f.deps)
      Prop(compact == legacy) :| s"$compact\nvs\n$legacy"
    }
  }

  property("within a component: the Gauss-Seidel sweep equals the legacy graph's") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val f = fixture(wb)
      val cyclic = DependencyGraph.qualifiedSccOrder(f.deps).filter(_.cyclic)
      Prop.all(cyclic.map { scc =>
        val reads = f.graph.memberPrecedents(f.ids(scc.members))
        val deps = reads.map((reader, read) => f.ref(reader) -> f.refs(read))
        val compact = DependencyGraph.withinComponentOrder(scc.members, deps)
        val legacy = DependencyGraph.withinComponentOrder(scc.members, f.deps)
        Prop(compact == legacy) :| s"$compact vs $legacy"
      }*)
    }
  }

  property("member reads: direct legacy reads among any member set, one component or many") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val f = fixture(wb)
      def expected(members: Set[QualifiedRef]): Map[QualifiedRef, Set[QualifiedRef]] =
        members.iterator
          .map(q => q -> (f.deps(q) & members - q))
          .filter(_._2.nonEmpty)
          .toMap
      def compact(members: Set[QualifiedRef]): Map[QualifiedRef, Set[QualifiedRef]] =
        f.graph.memberPrecedents(f.ids(members)).map((r, ps) => f.ref(r) -> f.refs(ps))
      val all = f.deps.keySet.toVector.sortBy(QualifiedRef.gridKey)
      val mixed = all.zipWithIndex.collect { case (q, i) if i % 2 == 0 => q }.toSet
      val sccs = DependencyGraph.qualifiedSccOrder(f.deps).map(_.members.toSet)
      Prop.all(
        (Prop(compact(mixed) == expected(mixed)) :| s"mixed $mixed") +:
          sccs.map(m => Prop(compact(m) == expected(m)) :| s"component $m")*
      )
    }
  }

  property("waves: longest-path depths equal the legacy depth over the pruned graph") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val f = fixture(wb)
      val removed = removedSet(f)
      val kahn = f.graph.kahn(f.ids(removed))
      val order = kahn.formulaOrder.iterator.map(f.ref).toVector
      // a suffix-closed subset: drop the dependents of a few formulas, as the dynamic bucket does
      val bucketSeeds = order.zipWithIndex.collect { case (q, i) if i % 7 == 3 => q }.toSet
      val bucket = bucketSeeds ++ DependencyGraph.qualifiedTransitiveDependents(f.rev, bucketSeeds)
      val main = order.filterNot(bucket.contains)
      val depth = f.graph.waveDepths(kahn, f.ids(main))
      val legacy = main
        .foldLeft(Map.empty[QualifiedRef, Int]) { (acc, q) =>
          val d = (f.deps(q) -- removed).flatMap(acc.get).map(_ + 1).maxOption.getOrElse(0)
          acc + (q -> d)
        }
      Prop.all(main.map { q =>
        Prop(depth(f.id(q)) == legacy(q)) :| s"$q: ${depth(f.id(q))} vs ${legacy(q)}"
      }*)
    }
  }

  property("readers of a cell: point and range readers, for formula, constant and empty cells") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val f = fixture(wb)
      val reads = readsOf(wb)
      val cells =
        for
          sheet <- CompactGraphGens.sheetNames
          row <- 0 to CompactGraphGens.Rows + 1
          col <- 0 to CompactGraphGens.Cols
        yield QualifiedRef(SheetName.unsafe(sheet), ARef.from0(col, row))
      val far = QualifiedRef(SheetName.unsafe("Z"), ARef.from0(90, 2))
      Prop.all((cells :+ far).map { cell =>
        val expected = reads.collect {
          case (reader, (points, ranges))
              if points.contains(cell) ||
                ranges.exists((s, r) => s == cell.sheet && r.contains(cell.ref)) =>
            reader
        }.toSet
        val s = f.graph.nodes.sheetOf(cell.sheet)
        val found = Set.newBuilder[QualifiedRef]
        f.graph.foreachReaderOfCell(s, cell.ref.row.index0, cell.ref.col.index0)(n =>
          found += f.ref(n)
        )
        Prop(found.result() == expected) :| s"$cell: ${found.result()} vs $expected"
      }*)
    }
  }

  /** Each formula's point and range reads, as the build extracts them. */
  private def readsOf(
    wb: Workbook
  ): Map[QualifiedRef, (Set[QualifiedRef], Vector[(SheetName, CellRange)])] =
    val canonical = DependencyGraph.sheetCanonicaliser(wb)
    val (nodes, texts) = NodeTable.build(wb)
    (0 until nodes.formulaCount).map { n =>
      val q = nodes.qualified(n)
      val ranges = Vector.newBuilder[(SheetName, CellRange)]
      val points = FormulaParser.parse(texts(n)) match
        case Right(expr) =>
          DependencyGraph.extractQualifiedDependencies(
            expr,
            q.sheet,
            (s, r) =>
              ranges += ((s, r))
              Set.empty
            ,
            Some(wb),
            preciseLookups = true,
            canonicalSheet = canonical,
            reader = Some(q.ref)
          )
        case Left(_) => Set.empty[QualifiedRef]
      q -> (points, ranges.result())
    }.toMap

  test("the generator exercises cycles, blocked cells, segment nodes and range nodes") {
    val books = (1 to 60).flatMap(i =>
      CompactGraphGens.genWorkbook(org.scalacheck.Gen.Parameters.default, Seed(i.toLong))
    )
    val fixtures = books.map(fixture)
    assert(books.size >= 50, s"${books.size} books")
    assert(fixtures.count(_.graph.cyclicFormulas.nonEmpty) >= 10, "cyclic books")
    assert(fixtures.count(f => removedSet(f).size > f.graph.cyclicFormulas.size) >= 10, "blocked")
    assert(fixtures.forall(_.graph.segmentCount > 0), "segment nodes")
    assert(fixtures.forall(_.graph.rangeCount > 5), "range nodes")
  }

  // ---- fixed cases --------------------------------------------------------------------------

  private def cyclic(wb: Workbook): Set[String] =
    val f = fixture(wb)
    f.refs(f.graph.cyclicFormulas).map(_.toString)

  test("SUM(A:A) in A5 is a circular reference; the other A-column formulas are not (R1-B3)") {
    val wb = CompactGraphGens.book(("S", "A1", "1+1"), ("S", "A5", "SUM(A:A)"), ("S", "A9", "A1"))
    assertEquals(cyclic(wb), Set("S!A5"))
  }

  test("SUM(A1:A10) in A5 is circular; SUM(A1:A4) in A5 is not") {
    assertEquals(
      cyclic(CompactGraphGens.book(("S", "A5", "SUM(A1:A10)"), ("S", "A2", "1"))),
      Set("S!A5")
    )
    assertEquals(cyclic(CompactGraphGens.book(("S", "A5", "SUM(A1:A4)"), ("S", "A2", "1"))), Set())
  }

  test("a running total that includes its own row is circular; one ending above it is not") {
    val rows = 2 to 30
    val own = CompactGraphGens.book(rows.map(r => ("S", s"D$r", s"SUM($$D$$2:D$r)"))*)
    assertEquals(cyclic(own).size, rows.size)
    val above = CompactGraphGens.book(rows.map(r => ("S", s"D$r", s"SUM($$D$$1:D${r - 1})"))*)
    assertEquals(cyclic(above), Set())
  }

  test("a cross-sheet range over the reader's own cell is circular; point self-reads too") {
    val wb = CompactGraphGens.book(
      ("S", "A2", "SUM(S!A:A)"),
      ("T", "B1", "SUM(S!A1:A9)"),
      ("T", "C3", "T!C3+1"),
      ("T", "C4", "C3")
    )
    // T!C4 reads only the cyclic C3: blocked, never cyclic
    assertEquals(cyclic(wb), Set("S!A2", "T!C3"))
  }

  test("running totals: O(n log n) edges, every D_n after every C_m it sums") {
    val n = 2000
    val wb = Workbook(Vector({
      val cells = (2 to n + 1).flatMap { r =>
        Vector(
          ARef.from0(2, r - 1) -> CellValue.Formula(s"B$r*2"),
          ARef.from0(3, r - 1) -> CellValue.Formula(s"SUM($$C$$2:C$r)")
        )
      }
      Sheet(SheetName.unsafe("Run"), cells = cells.map((ref, v) => ref -> Cell(ref, v)).toMap)
    }))
    val g = CompactGraph.build(wb)
    val log2 = 32 - Integer.numberOfLeadingZeros(n)
    assertEquals(g.formulaCount, 2 * n)
    assertEquals(g.rangeCount, n)
    assert(g.edgeCount <= 2 * n * (log2 + 2), s"${g.edgeCount} edges")
    val order = g.kahn(BitSet.empty).formulaOrder.iterator.map(g.nodes.qualified).toVector
    val position = order.zipWithIndex.toMap
    (2 to n + 1 by 97).foreach { r =>
      val d = position(QualifiedRef(SheetName.unsafe("Run"), ARef.from0(3, r - 1)))
      val c = position(QualifiedRef(SheetName.unsafe("Run"), ARef.from0(2, r - 1)))
      assert(c < d, s"C$r after D$r")
    }
  }

  test("a 100k-formula chain needs no call stack") {
    val n = 100000
    val cells = (1 to n).map { r =>
      val ref = ARef.from0(0, r - 1)
      ref -> Cell(ref, CellValue.Formula(if r == 1 then "1" else s"A${r - 1}+1"))
    }
    val g = CompactGraph.build(Workbook(Vector(Sheet(SheetName.unsafe("S"), cells = cells.toMap))))
    assertEquals(g.cyclicFormulas, BitSet.empty)
    assertEquals(g.kahn(BitSet.empty).formulaOrder.toVector, (0 until n).toVector)
    assertEquals(g.reach(BitSet(0), BitSet.empty).formulas.size, n - 1)
  }

  test("segment trees: for every column size up to 70, a range reaches exactly its formulas") {
    // the bottom-up layout puts leaves at m..2m-1 for any m, so for m not a power of two some
    // internal nodes cover non-adjacent leaves; the canonical pieces of a window must still be
    // exactly the formulas inside it
    val rnd = new scala.util.Random(767)
    (1 to 70).foreach { m =>
      // column A holds `1` formulas on the odd rows 1, 3, ..., 2m - 1, so windows can start and
      // end on blank rows; column C holds one reader per window
      val last = 2 * m
      val all = for a <- 1 to last; b <- a to last yield (a, b)
      val windows = if all.size <= 200 then all else Vector.fill(200)(all(rnd.nextInt(all.size)))
      val formulas = (0 until m).map(i => ("S", s"A${2 * i + 1}", "1")) ++
        windows.zipWithIndex.map { case ((a, b), k) => ("S", s"C${k + 1}", s"SUM(A$a:A$b)") }
      val g = CompactGraph.build(CompactGraphGens.book(formulas*))
      (0 until m).foreach { i =>
        val row = 2 * i + 1
        val self = g.nodes.lookup(QualifiedRef(SheetName.unsafe("S"), ARef.from0(0, row - 1)))
        val reached = self.fold(Set.empty[String])(n =>
          g.reach(BitSet(n), BitSet.empty)
            .formulas
            .iterator
            .map(g.nodes.qualified(_).ref.toA1)
            .toSet
        )
        val expected = windows.zipWithIndex.collect {
          case ((a, b), k) if a <= row && row <= b => s"C${k + 1}"
        }.toSet
        assertEquals(reached, expected, s"m=$m, A$row")
      }
    }
  }
