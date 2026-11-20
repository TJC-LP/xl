package com.tjclp.xl.formula

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*

/**
 * Tests for DependencyGraph (WI-09d).
 *
 * Tests dependency extraction, dynamic-reference detection, and the qualified SCC condensation and
 * within-component order the workbook recalculation uses.
 */
class DependencyGraphSpec extends ScalaCheckSuite:
  val emptySheet = new Sheet(name = SheetName.unsafe("Test"))

  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  /** #891: the dynamic cells of a one-sheet book, as refs on that sheet. */
  private def dynamicCellsOf(sheet: Sheet): Set[ARef] =
    DependencyGraph.dynamicCells(Workbook(sheet)).map(_.ref)

  private def parseRange(range: String): CellRange =
    CellRange.parse(range).fold(err => fail(err), identity)

  private def parseRef(ref: String): ARef =
    ARef.parse(ref).fold(err => fail(err), identity)

  // ===== Dependency Extraction Tests (10 tests) =====

  test("QualifiedRef: the cells of a sheet hash apart (#891)") {
    // A Long hashes as high ^ low, which for an ARef is row ^ column: before the override every
    // anti-diagonal of this 200 x 64 grid shared one hash (256 distinct hashes for 12,800 cells).
    val sheet = SheetName.unsafe("Model")
    val refs =
      for
        row <- 0 until 200
        col <- 0 until 64
      yield DependencyGraph.QualifiedRef(sheet, ARef.from0(col, row))
    val distinct = refs.map(_.hashCode).toSet.size
    assert(distinct > refs.size * 99 / 100, s"$distinct distinct hashes for ${refs.size} cells")
    val other = SheetName.unsafe("Other")
    assertEquals(
      DependencyGraph.QualifiedRef(sheet, ARef.from0(3, 7)).hashCode,
      DependencyGraph.QualifiedRef(sheet, ARef.from0(3, 7)).hashCode
    )
    assertNotEquals(
      DependencyGraph.QualifiedRef(sheet, ARef.from0(3, 7)).hashCode,
      DependencyGraph.QualifiedRef(other, ARef.from0(3, 7)).hashCode
    )
  }

  test("extractDependencies: Lit has no dependencies") {
    val expr = TExpr.Lit(BigDecimal(42))
    assertEquals(DependencyGraph.extractDependencies(expr), Set.empty[ARef])
  }

  test("extractDependencies: Ref has single dependency") {
    val expr = TExpr.ref(ref"A1", TExpr.decodeNumeric)
    assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1"))
  }

  test("extractDependencies: Add with two Refs") {
    val expr = TExpr.Add(
      TExpr.ref(ref"A1", TExpr.decodeNumeric),
      TExpr.ref(ref"B1", TExpr.decodeNumeric)
    )
    assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1", ref"B1"))
  }

  test("GH-374: extractDependencies walks through unary plus (=+A1+B1)") {
    FormulaParser.parse("=+A1+B1") match
      case Right(expr) =>
        assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1", ref"B1"))
        assert(DependencyGraph.containsCellReferences(expr))
      case Left(err) => fail(s"=+A1+B1 should parse: $err")
  }

  test("GH-355: extractDependencies walks through postfix percent (=A1%)") {
    FormulaParser.parse("=A1%") match
      case Right(expr) =>
        assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1"))
        assert(DependencyGraph.containsCellReferences(expr))
      case Left(err) => fail(s"=A1% should parse: $err")
  }

  test("GH-355/GH-374: dependencies survive percent nested under unary plus (=+A1%)") {
    FormulaParser.parse("=+A1%") match
      case Right(expr) =>
        assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1"))
      case Left(err) => fail(s"=+A1% should parse: $err")
  }

  test("extractDependencies: range arguments expand to all cells in range") {
    val range = parseRange("A1:A3")
    val expr = TExpr.sum(range)
    val expected = Set(ref"A1", ref"A2", ref"A3")
    assertEquals(DependencyGraph.extractDependencies(expr), expected)
  }

  test("extractDependencies: nested expressions") {
    val expr = TExpr.Mul(
      TExpr.Add(TExpr.ref(ref"A1", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(10))),
      TExpr.ref(ref"B1", TExpr.decodeNumeric)
    )
    assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1", ref"B1"))
  }

  test("extractDependencies: If with three branches") {
    val expr = TExpr.cond(
      TExpr.Gt(TExpr.ref(ref"A1", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(0))),
      TExpr.ref(ref"B1", TExpr.decodeNumeric),
      TExpr.ref(ref"C1", TExpr.decodeNumeric)
    )
    assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1", ref"B1", ref"C1"))
  }

  test("extractDependencies: Today and Now have no dependencies") {
    val today = TExpr.today()
    val now = TExpr.now()
    assertEquals(DependencyGraph.extractDependencies(today), Set.empty[ARef])
    assertEquals(DependencyGraph.extractDependencies(now), Set.empty[ARef])
  }

  test("extractDependencies: text functions with Refs") {
    val expr = TExpr.concatenate(
      List(
        TExpr.ref(ref"A1", TExpr.decodeAsString),
        TExpr.Lit(" "),
        TExpr.ref(ref"B1", TExpr.decodeAsString)
      )
    )
    assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1", ref"B1"))
  }

  test("extractDependencies: date functions with Refs") {
    val expr = TExpr.date(
      TExpr.ref(ref"A1", TExpr.decodeAsInt),
      TExpr.ref(ref"A2", TExpr.decodeAsInt),
      TExpr.ref(ref"A3", TExpr.decodeAsInt)
    )
    assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1", ref"A2", ref"A3"))
  }

  test("extractDependencies: complex nested with ranges") {
    val expr = TExpr.Add(
      TExpr.sum(parseRange("A1:A3")),
      TExpr.Mul(TExpr.ref(ref"B1", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(2)))
    )
    val expected = Set(ref"A1", ref"A2", ref"A3", ref"B1")
    assertEquals(DependencyGraph.extractDependencies(expr), expected)
  }

  // #891: the sheet-level graph (fromSheet, detectCycles, topologicalSort, precedents/dependents,
  // transitive walks, extractDependenciesBounded) is gone. Its laws live on the workbook graph:
  // CompactGraphSpec (reachability, cycles, Kahn order, range cycles, stack safety),
  // QualifiedSccOracleSpec (SCC order), QualifiedGraphSpec (precedents/dependents layers, A:A
  // bounded to occupied cells) and RecalcOrderLawSpec (evaluation order).

  // ===== GH-274: dynamic references (INDIRECT) =====

  test("containsDynamicReference: true for INDIRECT nested in IF/SUM arguments") {
    val nested = FormulaParser
      .parse("=IF(A1>0, SUM(INDIRECT(\"B1:B2\")), 0)")
      .fold(err => fail(err.toString), identity)
    assert(DependencyGraph.containsDynamicReference(nested))

    val arithmetic = FormulaParser
      .parse("=INDIRECT(\"C1\")+1")
      .fold(err => fail(err.toString), identity)
    assert(DependencyGraph.containsDynamicReference(arithmetic))
  }

  test("containsDynamicReference: false without dynamic calls") {
    // GH-301: OFFSET is now dynamicDeps-flagged (it reads a shifted window the static graph
    // cannot see), so the static example here no longer includes it — and the OFFSET form is
    // asserted dynamic below.
    val static = FormulaParser
      .parse("=SUM(A1:A3)+VLOOKUP(\"k\",B1:C3,2)")
      .fold(err => fail(err.toString), identity)
    assert(!DependencyGraph.containsDynamicReference(static))

    val offsetArithmetic = FormulaParser
      .parse("=SUM(A1:A3)+OFFSET(A1,1,0)")
      .fold(err => fail(err.toString), identity)
    assert(DependencyGraph.containsDynamicReference(offsetArithmetic))
  }

  test("extractDependencies: =INDIRECT(A1) sees the argument A1") {
    val expr = FormulaParser.parse("=INDIRECT(A1)").fold(err => fail(err.toString), identity)
    assertEquals(DependencyGraph.extractDependencies(expr), Set(parseRef("A1")))
  }

  test("extractDependencies: =INDIRECT(\"B2\") sees nothing (resolved target is invisible)") {
    // Pins the design: the static graph never gains an edge to INDIRECT's target — zero
    // Tarjan false positives; freshness is the deferred bucket's job, not the graph's.
    val expr = FormulaParser.parse("=INDIRECT(\"B2\")").fold(err => fail(err.toString), identity)
    assertEquals(DependencyGraph.extractDependencies(expr), Set.empty[ARef])
  }

  test("dynamicCells: finds INDIRECT formulas, skips static and unparseable ones") {
    val sheet = sheetWith(
      parseRef("A1") -> CellValue.Formula("=INDIRECT(\"C1\")", None),
      parseRef("A2") -> CellValue.Formula("=indirect(b1)", None), // case-insensitive
      parseRef("A3") -> CellValue.Formula("=SUM(B1:B2)", None),
      parseRef("A4") -> CellValue.Formula("=INDIRECT(", None), // unparseable: not dynamic
      parseRef("A5") -> CellValue.Text("INDIRECT(\"C1\")") // not a formula
    )
    assertEquals(dynamicCellsOf(sheet), Set(parseRef("A1"), parseRef("A2")))
  }

  test("GH-301 dynamicCells: OFFSET formulas are dynamic (parity with INDIRECT)") {
    val sheet = sheetWith(
      parseRef("A1") -> CellValue.Formula("=OFFSET(B1,1,0)", None),
      parseRef("A2") -> CellValue.Formula("=SUM(OFFSET(B1,0,0,3,1))", None), // composed
      parseRef("A3") -> CellValue.Formula("=offset(B1,1,0)*2", None), // case + operand position
      parseRef("A4") -> CellValue.Formula("=SUM(B1:B2)", None) // static
    )
    assertEquals(
      dynamicCellsOf(sheet),
      Set(parseRef("A1"), parseRef("A2"), parseRef("A3"))
    )
  }

  test("a dynamic cell's static dependents are its transitive dependents (the deferred bucket)") {
    // #891: ported from dynamicClosure — the bucket recalculation defers is the dynamic cells plus
    // what reads them, transitively; a formula reading the INDIRECT target directly is not in it.
    val sheet = sheetWith(
      parseRef("A1") -> CellValue.Formula("=INDIRECT(\"X1\")", None),
      parseRef("B1") -> CellValue.Formula("=A1+1", None),
      parseRef("C1") -> CellValue.Formula("=B1+1", None),
      parseRef("D1") -> CellValue.Formula("=X1*2", None) // not downstream of A1
    )
    val wb = Workbook(sheet)
    val dynamic = DependencyGraph.dynamicCells(wb)
    assertEquals(dynamic.map(_.ref), Set(parseRef("A1")))
    val graph = QualifiedGraph.of(wb)
    val bucket = dynamic ++ dynamic.flatMap(q => graph.dependents(q, 0).flatten)
    assertEquals(bucket.map(_.ref), Set(parseRef("A1"), parseRef("B1"), parseRef("C1")))
  }

  // ===== GH-492: SCC condensation order (qualifiedSccOrder) =====

  // A1, A10 and AA1 on two sheets: the row 9/10 and column Z/AA boundaries where the A1 STRING
  // order ("A10" < "A2", "AA1" < "B1") and the grid order disagree, so every law below is
  // exercised across them.
  private val sccUniverse: Vector[DependencyGraph.QualifiedRef] =
    (0 until 6).toVector.map { i =>
      val sheet = if i < 3 then SheetName.unsafe("S1") else SheetName.unsafe("S2")
      val ref = i % 3 match
        case 0 => ARef.from0(0, 0) // A1
        case 1 => ARef.from0(0, 9) // A10
        case _ => ARef.from0(26, 0) // AA1
      DependencyGraph.QualifiedRef(sheet, ref)
    }

  private val sccGraphGen
    : Gen[Map[DependencyGraph.QualifiedRef, Set[DependencyGraph.QualifiedRef]]] =
    Gen
      .listOf(
        for
          k <- Gen.oneOf(sccUniverse)
          vs <- Gen.someOf(sccUniverse).map(_.toSet)
        yield k -> vs
      )
      .map(_.toMap)

  property("GH-492 partition: qualifiedSccOrder covers exactly dependencies.keySet, no dupes") {
    forAll(sccGraphGen) { deps =>
      val comps = DependencyGraph.qualifiedSccOrder(deps)
      val flat = comps.flatMap(_.members)
      ((flat.toSet == deps.keySet) :| s"cover: ${flat.toSet} vs ${deps.keySet}") &&
      ((flat.size == flat.distinct.size) :| "no duplicate members") &&
      ((comps.map(_.members.size).sum == deps.keySet.size) :| "sizes sum")
    }
  }

  property("GH-492 order: every cross-component edge points strictly backwards") {
    forAll(sccGraphGen) { deps =>
      val comps = DependencyGraph.qualifiedSccOrder(deps)
      val idx = comps.zipWithIndex.flatMap((c, i) => c.members.map(_ -> i)).toMap
      val ok = deps.forall { case (u, vs) =>
        vs.filter(idx.contains).forall(v => idx(v) == idx(u) || idx(v) < idx(u))
      }
      ok :| s"condensation order violated for $deps"
    }
  }

  property("GH-492 cyclic flag: a component is cyclic iff it has two members or a self-read") {
    forAll(sccGraphGen) { deps =>
      val comps = DependencyGraph.qualifiedSccOrder(deps)
      val wrong = comps.filterNot { c =>
        c.cyclic == (c.members.size > 1 || c.members.exists(m => deps(m).contains(m)))
      }
      wrong.isEmpty :| s"cyclic flag wrong on $wrong for $deps"
    }
  }

  property("GH-492 determinism: the order depends on the graph VALUE, not on build order") {
    forAll(sccGraphGen, Gen.choose(0L, 1000000L)) { (deps, seed) =>
      val rnd = new scala.util.Random(seed)
      val rebuilt =
        rnd
          .shuffle(deps.toVector)
          .map((k, vs) => k -> rnd.shuffle(vs.toVector).toSet)
          .toMap
      val a = DependencyGraph.qualifiedSccOrder(deps)
      val b = DependencyGraph.qualifiedSccOrder(rebuilt)
      (a.map(c => (c.members, c.cyclic)) == b.map(c => (c.members, c.cyclic))) :|
        s"nondeterministic order:\n  $a\n  $b"
    }
  }

  test("GH-492: members inside a component are sorted by (sheet name, row, column)") {
    val s = SheetName.unsafe("S")
    val q = (r: String) => DependencyGraph.QualifiedRef(s, parseRef(r))
    // B2 -> A1 -> C3 -> B2: one three-member cycle
    val deps = Map(
      q("B2") -> Set(q("A1")),
      q("A1") -> Set(q("C3")),
      q("C3") -> Set(q("B2"))
    )
    val comps = DependencyGraph.qualifiedSccOrder(deps)
    assertEquals(comps.size, 1)
    assertEquals(comps.map(_.members.map(_.ref.toA1)), Vector(Vector("A1", "B2", "C3")))
    assert(comps.forall(_.cyclic))
  }

  test("GH-482: the member listing is row-major on the grid, not lexicographic on the A1 text") {
    // "A10" < "AA1" < "B2" as strings; on the grid AA1 (row 1) precedes B2 (row 2) precedes A10.
    val s = SheetName.unsafe("S")
    val q = (r: String) => DependencyGraph.QualifiedRef(s, parseRef(r))
    val deps = Map(q("B2") -> Set(q("A10")), q("A10") -> Set(q("AA1")), q("AA1") -> Set(q("B2")))
    val comps = DependencyGraph.qualifiedSccOrder(deps)
    assertEquals(comps.map(_.members.map(_.ref.toA1)), Vector(Vector("AA1", "B2", "A10")))
    // Two independent components order by their top-left member the same way.
    val two = Map(q("A10") -> Set(q("A10")), q("B2") -> Set(q("B2")))
    assertEquals(
      DependencyGraph.qualifiedSccOrder(two).map(_.members.map(_.ref.toA1)),
      Vector(Vector("B2"), Vector("A10"))
    )
  }

  test("GH-492: a self-loop singleton is cyclic; a plain singleton is not") {
    val s = SheetName.unsafe("S")
    val a1 = DependencyGraph.QualifiedRef(s, parseRef("A1"))
    val b1 = DependencyGraph.QualifiedRef(s, parseRef("B1"))
    val deps = Map(a1 -> Set(a1), b1 -> Set(a1))
    val comps = DependencyGraph.qualifiedSccOrder(deps)
    assertEquals(
      comps.map(c => (c.members.map(_.ref.toA1), c.cyclic)),
      Vector((Vector("A1"), true), (Vector("B1"), false))
    )
  }

  test("GH-492: constants (non-key successors) never leak in as components") {
    val s = SheetName.unsafe("S")
    val a1 = DependencyGraph.QualifiedRef(s, parseRef("A1"))
    val z9 = DependencyGraph.QualifiedRef(s, parseRef("Z9")) // constant: not a key
    val deps = Map(a1 -> Set(z9))
    val comps = DependencyGraph.qualifiedSccOrder(deps)
    assertEquals(comps.flatMap(_.members), Vector(a1))
  }

  // ===== GH-482: within-component evaluation order (withinComponentOrder) =====

  /**
   * The grid order: sheet name, then row, then column — Excel's left-to-right, top-to-bottom sweep.
   */
  private def keyOf(q: DependencyGraph.QualifiedRef): (String, Int, Int) =
    (q.sheet.value, q.ref.row.index0, q.ref.col.index0)

  property("GH-482 permutation: withinComponentOrder is a permutation of the component") {
    forAll(sccGraphGen) { deps =>
      DependencyGraph.qualifiedSccOrder(deps).filter(_.cyclic).forall { c =>
        DependencyGraph.withinComponentOrder(c.members, deps).sortBy(keyOf) == c.members
      } :| s"not a permutation of some component of $deps"
    }
  }

  property("GH-482 determinism: the order depends on the graph VALUE, not on build order") {
    forAll(sccGraphGen, Gen.choose(0L, 1000000L)) { (deps, seed) =>
      val rnd = new scala.util.Random(seed)
      val rebuilt =
        rnd
          .shuffle(deps.toVector)
          .map((k, vs) => k -> rnd.shuffle(vs.toVector).toSet)
          .toMap
      DependencyGraph.qualifiedSccOrder(deps).filter(_.cyclic).forall { c =>
        DependencyGraph.withinComponentOrder(c.members, deps) ==
          DependencyGraph.withinComponentOrder(rnd.shuffle(c.members), rebuilt)
      } :| s"nondeterministic within-component order for $deps"
    }
  }

  /** Acyclic graphs over `sccUniverse`: a node reads only nodes of a lower index. */
  private val dagGen: Gen[Map[DependencyGraph.QualifiedRef, Set[DependencyGraph.QualifiedRef]]] =
    Gen
      .sequence[List[Set[Int]], Set[Int]](
        sccUniverse.indices.toList.map(i => Gen.someOf(0 until i).map(_.toSet))
      )
      .map { picks =>
        sccUniverse.zipWithIndex.map((q, i) => q -> picks(i).map(sccUniverse)).toMap
      }

  /** Kahn with the smallest-key ready node first — the order a cut-free run must reproduce. */
  private def referenceKahn(
    nodes: Vector[DependencyGraph.QualifiedRef],
    deps: Map[DependencyGraph.QualifiedRef, Set[DependencyGraph.QualifiedRef]]
  ): Vector[DependencyGraph.QualifiedRef] =
    val members = nodes.toSet
    @annotation.tailrec
    def go(
      remaining: Set[DependencyGraph.QualifiedRef],
      emitted: Vector[DependencyGraph.QualifiedRef]
    ): Vector[DependencyGraph.QualifiedRef] =
      val ready = remaining.filter { q =>
        deps.getOrElse(q, Set.empty).filter(members).forall(p => p == q || emitted.contains(p))
      }
      ready.minByOption(keyOf) match
        case None => emitted
        case Some(next) => go(remaining - next, emitted :+ next)
    go(members, Vector.empty)

  property("GH-482 DAG: on an acyclic member set the order IS the key-ordered Kahn order") {
    forAll(dagGen) { deps =>
      val order = DependencyGraph.withinComponentOrder(sccUniverse, deps)
      (order == referenceKahn(sccUniverse, deps)) :| s"$order vs Kahn for $deps"
    }
  }

  test("GH-482: the cut is the TOP-LEFT remaining member — row 10 never precedes row 2") {
    // Same two-cell counter shifted across the row 9/10 and column Z/AA boundaries: the A1
    // STRING order would cut at A10 ("A10" < "B2") and at AA1 ("AA1" < "Z1"), handing the
    // trailing cell a round-old value; the grid order cuts at B2 (row 2) and Z1 (column Z).
    val s = SheetName.unsafe("S")
    val q = (r: String) => DependencyGraph.QualifiedRef(s, parseRef(r))
    val rows = Map(q("B2") -> Set(q("A10")), q("A10") -> Set(q("B2")))
    assertEquals(
      DependencyGraph.withinComponentOrder(Vector(q("A10"), q("B2")), rows).map(_.ref.toA1),
      Vector("B2", "A10")
    )
    val cols = Map(q("Z1") -> Set(q("AA1")), q("AA1") -> Set(q("Z1")))
    assertEquals(
      DependencyGraph.withinComponentOrder(Vector(q("AA1"), q("Z1")), cols).map(_.ref.toA1),
      Vector("Z1", "AA1")
    )
    // Within one column the lower row precedes, whatever its digit count.
    val same = Map(q("A2") -> Set(q("A10")), q("A10") -> Set(q("A2")))
    assertEquals(
      DependencyGraph.withinComponentOrder(Vector(q("A10"), q("A2")), same).map(_.ref.toA1),
      Vector("A2", "A10")
    )
  }

  test("GH-482: a cross-sheet cycle sweeps sheet order first, then row-major within a sheet") {
    // T!A1 reads S!A10, S!A10 reads S!B2, S!B2 reads T!A1. T!A1 sits on row 1, but sheet name
    // orders first: the cut lands on sheet S at its top-left member B2 (the A1 STRING order would
    // cut at A10), then Kahn follows the chain — A10 reads B2, T!A1 reads A10.
    val sS = SheetName.unsafe("S")
    val sT = SheetName.unsafe("T")
    val sA10 = DependencyGraph.QualifiedRef(sS, parseRef("A10"))
    val sB2 = DependencyGraph.QualifiedRef(sS, parseRef("B2"))
    val tA1 = DependencyGraph.QualifiedRef(sT, parseRef("A1"))
    val deps = Map(tA1 -> Set(sA10), sA10 -> Set(sB2), sB2 -> Set(tA1))
    assertEquals(
      DependencyGraph.withinComponentOrder(Vector(tA1, sA10, sB2), deps),
      Vector(sB2, sA10, tA1)
    )
  }

  test("GH-482: the classic chain A1 = B1+1, B1 = A1 evaluates A1 first") {
    // Cut at the top-left member (A1), then B1 (which reads A1) is ready: the counter reaches
    // 100/100 in 100 iterations only if A1 is evaluated before B1 within a round.
    val s = SheetName.unsafe("S")
    val q = (r: String) => DependencyGraph.QualifiedRef(s, parseRef(r))
    val deps = Map(q("A1") -> Set(q("B1")), q("B1") -> Set(q("A1")))
    assertEquals(
      DependencyGraph.withinComponentOrder(Vector(q("A1"), q("B1")), deps).map(_.ref.toA1),
      Vector("A1", "B1")
    )
  }

  test("GH-482: the average-balance idiom orders interest, payment, ending balance") {
    // B2 = (B1+B4)/2*A2 reads B4; B3 = A3-B2 reads B2; B4 = B1-B3 reads B3. Cutting the
    // smallest key (B2) leaves a chain B2 -> B3 -> B4: one sweep propagates the whole schedule.
    val s = SheetName.unsafe("S")
    val q = (r: String) => DependencyGraph.QualifiedRef(s, parseRef(r))
    val deps = Map(q("B2") -> Set(q("B4")), q("B3") -> Set(q("B2")), q("B4") -> Set(q("B3")))
    assertEquals(
      DependencyGraph
        .withinComponentOrder(Vector(q("B4"), q("B2"), q("B3")), deps)
        .map(_.ref.toA1),
      Vector("B2", "B3", "B4")
    )
  }

  test("GH-482: a self-edge is a back edge, never a blocker; precedents still go first") {
    val s = SheetName.unsafe("S")
    val q = (r: String) => DependencyGraph.QualifiedRef(s, parseRef(r))
    // A1 reads itself and B1; B1 reads A1 — the cycle. B1 < A1 would be Kahn's pick only if it
    // were ready, and it is not (it reads A1): the cut lands on A1, then B1 follows.
    val deps = Map(q("A1") -> Set(q("A1"), q("B1")), q("B1") -> Set(q("A1")))
    assertEquals(
      DependencyGraph.withinComponentOrder(Vector(q("A1"), q("B1")), deps).map(_.ref.toA1),
      Vector("A1", "B1")
    )
    // With B1 reading nothing in-component, B1 is ready first and A1's self-edge is ignored.
    val chain =
      Map(q("A1") -> Set(q("A1"), q("B1")), q("B1") -> Set.empty[DependencyGraph.QualifiedRef])
    assertEquals(
      DependencyGraph.withinComponentOrder(Vector(q("A1"), q("B1")), chain).map(_.ref.toA1),
      Vector("B1", "A1")
    )
  }

  test("GH-482: the three-cycle B2 -> A1 -> C3 -> B2 cuts at A1 and follows the chain") {
    val s = SheetName.unsafe("S")
    val q = (r: String) => DependencyGraph.QualifiedRef(s, parseRef(r))
    val deps = Map(q("B2") -> Set(q("A1")), q("A1") -> Set(q("C3")), q("C3") -> Set(q("B2")))
    assertEquals(
      DependencyGraph.withinComponentOrder(Vector(q("A1"), q("B2"), q("C3")), deps).map(_.ref.toA1),
      Vector("A1", "B2", "C3")
    )
  }
