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
 * Tests dependency extraction, graph construction, cycle detection (Tarjan's SCC), and topological
 * sorting (Kahn's algorithm).
 */
class DependencyGraphSpec extends ScalaCheckSuite:
  val emptySheet = new Sheet(name = SheetName.unsafe("Test"))

  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  private def parseRange(range: String): CellRange =
    CellRange.parse(range).fold(err => fail(err), identity)

  private def parseRef(ref: String): ARef =
    ARef.parse(ref).fold(err => fail(err), identity)

  private def assertActualCycle(
    dependencies: Map[ARef, Set[ARef]],
    cycle: List[ARef]
  ): Unit =
    val first = cycle.headOption.getOrElse(fail("expected a non-empty cycle"))
    val last = cycle.lastOption.getOrElse(fail("expected a non-empty cycle"))
    assertEquals(last, first, "cycle must close at its first node")
    val interior = cycle.dropRight(1)
    assertEquals(
      interior.distinct.size,
      interior.size,
      s"cycle of size ${cycle.size} must not repeat an interior node"
    )
    cycle.sliding(2).foreach {
      case from :: to :: Nil =>
        assert(
          dependencies.getOrElse(from, Set.empty).contains(to),
          s"reported non-edge ${from.toA1} -> ${to.toA1} in cycle of size ${cycle.size}"
        )
      case _ => ()
    }

  // ===== Dependency Extraction Tests (10 tests) =====

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
        assertEquals(
          DependencyGraph.extractDependenciesBounded(expr, None),
          Set(ref"A1", ref"B1")
        )
        assert(DependencyGraph.containsCellReferences(expr))
      case Left(err) => fail(s"=+A1+B1 should parse: $err")
  }

  test("GH-355: extractDependencies walks through postfix percent (=A1%)") {
    FormulaParser.parse("=A1%") match
      case Right(expr) =>
        assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1"))
        assertEquals(DependencyGraph.extractDependenciesBounded(expr, None), Set(ref"A1"))
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

  // ===== Graph Construction Tests (8 tests) =====

  test("fromSheet: empty sheet produces empty graph") {
    val graph = DependencyGraph.fromSheet(emptySheet)
    assertEquals(graph.dependencies, Map.empty[ARef, Set[ARef]])
    assertEquals(graph.dependents, Map.empty[ARef, Set[ARef]])
  }

  test("fromSheet: sheet with only constants produces empty graph") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Text("Hello")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(graph.dependencies, Map.empty[ARef, Set[ARef]])
    assertEquals(graph.dependents, Map.empty[ARef, Set[ARef]])
  }

  test("fromSheet: single formula with no deps") {
    val sheet = sheetWith(ref"A1" -> CellValue.Formula("=42"))
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(graph.dependencies, Map(ref"A1" -> Set.empty[ARef]))
    assertEquals(graph.dependents, Map.empty[ARef, Set[ARef]])
  }

  test("fromSheet: single formula with one dep") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1*2")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(graph.dependencies, Map(ref"B1" -> Set(ref"A1")))
    assertEquals(graph.dependents, Map(ref"A1" -> Set(ref"B1")))
  }

  test("fromSheet: linear chain (A1 -> B1 -> C1)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1+5"),
      ref"C1" -> CellValue.Formula("=B1*2")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(graph.dependencies, Map(ref"B1" -> Set(ref"A1"), ref"C1" -> Set(ref"B1")))
    assertEquals(graph.dependents, Map(ref"A1" -> Set(ref"B1"), ref"B1" -> Set(ref"C1")))
  }

  test("fromSheet: diamond (A1 -> B1, A1 -> C1, B1 -> D1, C1 -> D1)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1+5"),
      ref"C1" -> CellValue.Formula("=A1*2"),
      ref"D1" -> CellValue.Formula("=B1+C1")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(
      graph.dependencies,
      Map(
        ref"B1" -> Set(ref"A1"),
        ref"C1" -> Set(ref"A1"),
        ref"D1" -> Set(ref"B1", ref"C1")
      )
    )
    assertEquals(
      graph.dependents,
      Map(
        ref"A1" -> Set(ref"B1", ref"C1"),
        ref"B1" -> Set(ref"D1"),
        ref"C1" -> Set(ref"D1")
      )
    )
  }

  test("fromSheet: multiple disconnected components") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1+5"),
      ref"C1" -> CellValue.Number(BigDecimal(20)),
      ref"D1" -> CellValue.Formula("=C1*2")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(graph.dependencies, Map(ref"B1" -> Set(ref"A1"), ref"D1" -> Set(ref"C1")))
    assertEquals(graph.dependents, Map(ref"A1" -> Set(ref"B1"), ref"C1" -> Set(ref"D1")))
  }

  test("fromSheet: formula with range reference") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(1)),
      ref"A2" -> CellValue.Number(BigDecimal(2)),
      ref"A3" -> CellValue.Number(BigDecimal(3)),
      ref"B1" -> CellValue.Formula("=SUM(A1:A3)")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(graph.dependencies, Map(ref"B1" -> Set(ref"A1", ref"A2", ref"A3")))
    assertEquals(
      graph.dependents,
      Map(
        ref"A1" -> Set(ref"B1"),
        ref"A2" -> Set(ref"B1"),
        ref"A3" -> Set(ref"B1")
      )
    )
  }

  // ===== Cycle Detection Tests (12 tests) =====

  test("detectCycles: empty graph has no cycles") {
    val graph = DependencyGraph(Map.empty, Map.empty)
    assertEquals(DependencyGraph.detectCycles(graph), Right(()))
  }

  test("detectCycles: single node with no edges has no cycles") {
    val graph = DependencyGraph(Map(ref"A1" -> Set.empty[ARef]), Map.empty)
    assertEquals(DependencyGraph.detectCycles(graph), Right(()))
  }

  test("detectCycles: linear chain has no cycles") {
    val graph = DependencyGraph(
      Map(ref"B1" -> Set(ref"A1"), ref"C1" -> Set(ref"B1")),
      Map(ref"A1" -> Set(ref"B1"), ref"B1" -> Set(ref"C1"))
    )
    assertEquals(DependencyGraph.detectCycles(graph), Right(()))
  }

  test("detectCycles: diamond has no cycles") {
    val graph = DependencyGraph(
      Map(
        ref"B1" -> Set(ref"A1"),
        ref"C1" -> Set(ref"A1"),
        ref"D1" -> Set(ref"B1", ref"C1")
      ),
      Map(
        ref"A1" -> Set(ref"B1", ref"C1"),
        ref"B1" -> Set(ref"D1"),
        ref"C1" -> Set(ref"D1")
      )
    )
    assertEquals(DependencyGraph.detectCycles(graph), Right(()))
  }

  test("detectCycles: self-loop (A1 -> A1)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Formula("=A1"))
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.detectCycles(graph)
    assert(result.isLeft)
    result match
      case scala.util.Left(EvalError.CircularRef(cycle)) =>
        val first = cycle.headOption.getOrElse(fail("expected cycle with at least one ref"))
        val last = cycle.lastOption.getOrElse(fail("expected cycle with at least one ref"))
        assertEquals(first, ref"A1")
        assertEquals(last, ref"A1")
        assert(cycle.size >= 2) // At least [A1, A1]
      case _ => fail("Expected CircularRef error")
  }

  test("detectCycles: simple 2-cycle (A1 -> B1 -> A1)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=B1"),
      ref"B1" -> CellValue.Formula("=A1")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.detectCycles(graph)
    assert(result.isLeft)
    result match
      case scala.util.Left(EvalError.CircularRef(cycle)) =>
        assert(cycle.contains(ref"A1"))
        assert(cycle.contains(ref"B1"))
        val first = cycle.headOption.getOrElse(fail("expected cycle with at least one ref"))
        val last = cycle.lastOption.getOrElse(fail("expected cycle with at least one ref"))
        assertEquals(first, last) // Cycle closes
      case _ => fail("Expected CircularRef error")
  }

  test("detectCycles: 3-cycle (A1 -> B1 -> C1 -> A1)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=B1"),
      ref"B1" -> CellValue.Formula("=C1"),
      ref"C1" -> CellValue.Formula("=A1")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.detectCycles(graph)
    assert(result.isLeft)
    result match
      case scala.util.Left(EvalError.CircularRef(cycle)) =>
        assert(cycle.contains(ref"A1"))
        assert(cycle.contains(ref"B1"))
        assert(cycle.contains(ref"C1"))
        val first = cycle.headOption.getOrElse(fail("expected cycle with at least one ref"))
        val last = cycle.lastOption.getOrElse(fail("expected cycle with at least one ref"))
        assertEquals(first, last) // Cycle closes
      case _ => fail("Expected CircularRef error")
  }

  test("detectCycles: cycle with constants (A1 -> B1 -> C1 -> B1, D1 constant)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=B1"),
      ref"B1" -> CellValue.Formula("=C1"),
      ref"C1" -> CellValue.Formula("=B1"),
      ref"D1" -> CellValue.Number(BigDecimal(10))
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.detectCycles(graph)
    assert(result.isLeft)
    result match
      case scala.util.Left(EvalError.CircularRef(cycle)) =>
        assert(cycle.contains(ref"B1"))
        assert(cycle.contains(ref"C1"))
        assert(!cycle.contains(ref"D1")) // Constant not in cycle
      case _ => fail("Expected CircularRef error")
  }

  test("detectCycles: complex graph with one embedded cycle") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1+5"),
      ref"C1" -> CellValue.Formula("=B1+D1"), // Depends on D1 (cycle)
      ref"D1" -> CellValue.Formula("=C1*2") // Depends on C1 (cycle)
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.detectCycles(graph)
    assert(result.isLeft)
    result match
      case scala.util.Left(EvalError.CircularRef(cycle)) =>
        assert(cycle.contains(ref"C1"))
        assert(cycle.contains(ref"D1"))
      case _ => fail("Expected CircularRef error")
  }

  test("detectCycles: a branching SCC reports an actual directed cycle, not Tarjan pop order") {
    // All three nodes share one SCC, but B1 -> C1 and C1 -> B1 do not exist. Treating the SCC's
    // pop order as a path can therefore report A1 -> B1 -> C1 -> A1 with an invented middle edge.
    val dependencies = Map(
      ref"A1" -> Set(ref"B1", ref"C1"),
      ref"B1" -> Set(ref"A1"),
      ref"C1" -> Set(ref"A1")
    )
    val graph = DependencyGraph(
      dependencies,
      Map(
        ref"A1" -> Set(ref"B1", ref"C1"),
        ref"B1" -> Set(ref"A1"),
        ref"C1" -> Set(ref"A1")
      )
    )

    val cycle = DependencyGraph.detectCycles(graph) match
      case Left(EvalError.CircularRef(path)) => path
      case other => fail(s"expected a circular reference, got $other")
    assertActualCycle(dependencies, cycle)

    // Small immutable maps/sets retain construction order. Rebuilding the same graph in reverse
    // must not change a user-visible diagnostic.
    val rebuilt = List(
      ref"C1" -> Set(ref"A1"),
      ref"B1" -> Set(ref"A1"),
      ref"A1" -> List(ref"C1", ref"B1").toSet
    ).toMap
    val rebuiltCycle = DependencyGraph.detectCycles(DependencyGraph(rebuilt, Map.empty)) match
      case Left(EvalError.CircularRef(path)) => path
      case other => fail(s"expected a circular reference, got $other")
    assertEquals(rebuiltCycle, cycle)
  }

  test("detectCycles: cycle through range (A1 -> SUM(B1:B3) where B2 -> A1)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=SUM(B1:B3)"),
      ref"B1" -> CellValue.Number(BigDecimal(1)),
      ref"B2" -> CellValue.Formula("=A1*2"),
      ref"B3" -> CellValue.Number(BigDecimal(3))
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.detectCycles(graph)
    assert(result.isLeft)
    result match
      case scala.util.Left(EvalError.CircularRef(cycle)) =>
        assert(cycle.contains(ref"A1"))
        assert(cycle.contains(ref"B2"))
      case _ => fail("Expected CircularRef error")
  }

  test("detectCycles: multiple cycles (report first one found)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=A2"),
      ref"A2" -> CellValue.Formula("=A1"), // Cycle 1: A1 <-> A2
      ref"B1" -> CellValue.Formula("=B2"),
      ref"B2" -> CellValue.Formula("=B1") // Cycle 2: B1 <-> B2
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.detectCycles(graph)
    assert(result.isLeft) // Should detect at least one cycle
  }

  test("detectCycles: large acyclic graph (performance)") {
    // Create a large linear chain: A1 -> A2 -> A3 -> ... -> A100
    val cells = (1 to 100).map { i =>
      val ref = parseRef(s"A$i")
      if i == 1 then ref -> CellValue.Number(BigDecimal(1))
      else ref -> CellValue.Formula(s"=A${i - 1}+1")
    }
    val sheet = sheetWith(cells*)
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(DependencyGraph.detectCycles(graph), Right(()))
  }

  // ===== Topological Sort Tests (10 tests) =====

  test("topologicalSort: empty graph produces empty list") {
    val graph = DependencyGraph(Map.empty, Map.empty)
    assertEquals(DependencyGraph.topologicalSort(graph), Right(List.empty[ARef]))
  }

  test("topologicalSort: single node with no edges") {
    val graph = DependencyGraph(Map(ref"A1" -> Set.empty[ARef]), Map.empty)
    assertEquals(DependencyGraph.topologicalSort(graph), Right(List(ref"A1")))
  }

  test("topologicalSort: linear chain produces correct order") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1+5"),
      ref"C1" -> CellValue.Formula("=B1*2")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.topologicalSort(graph)
    assert(result.isRight)
    result match
      case scala.util.Right(order) =>
        assertEquals(order, List(ref"B1", ref"C1"))
        // B1 must come before C1 (B1 has no deps, C1 depends on B1)
      case _ => fail("Expected successful topological sort")
  }

  test("topologicalSort: diamond produces valid order") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1+5"),
      ref"C1" -> CellValue.Formula("=A1*2"),
      ref"D1" -> CellValue.Formula("=B1+C1")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.topologicalSort(graph)
    assert(result.isRight)
    result match
      case scala.util.Right(order) =>
        // Valid orders: [B1, C1, D1] or [C1, B1, D1]
        // D1 must come after both B1 and C1
        val b1Idx = order.indexOf(ref"B1")
        val c1Idx = order.indexOf(ref"C1")
        val d1Idx = order.indexOf(ref"D1")
        assert(b1Idx < d1Idx)
        assert(c1Idx < d1Idx)
      case _ => fail("Expected successful topological sort")
  }

  test("topologicalSort: cycle produces Left(CircularRef)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=B1"),
      ref"B1" -> CellValue.Formula("=A1")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.topologicalSort(graph)
    assert(result.isLeft)
  }

  test("topologicalSort: excludes a downstream Kahn remainder tail from the reported cycle") {
    // A1 is blocked by B1 <-> C1 but is not itself cyclic. The former single-successor walk could
    // start at A1, revisit B1, then append A1 and invent the closing edge B1 -> A1.
    val dependencies = Map(
      ref"A1" -> Set(ref"B1"),
      ref"B1" -> Set(ref"C1"),
      ref"C1" -> Set(ref"B1")
    )
    val graph = DependencyGraph(
      dependencies,
      Map(ref"B1" -> Set(ref"A1", ref"C1"), ref"C1" -> Set(ref"B1"))
    )

    DependencyGraph.topologicalSort(graph) match
      case Left(EvalError.CircularRef(cycle)) =>
        assertActualCycle(dependencies, cycle)
        assert(!cycle.contains(ref"A1"), s"downstream tail leaked into cycle: $cycle")
      case other => fail(s"expected a circular reference, got $other")
  }

  test("topologicalSort: self-loop produces Left(CircularRef)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Formula("=A1"))
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.topologicalSort(graph)
    assert(result.isLeft)
  }

  test("topologicalSort: multiple disconnected components") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1+5"),
      ref"C1" -> CellValue.Number(BigDecimal(20)),
      ref"D1" -> CellValue.Formula("=C1*2")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.topologicalSort(graph)
    assert(result.isRight)
    result match
      case scala.util.Right(order) =>
        // Both components should be present
        assert(order.contains(ref"B1"))
        assert(order.contains(ref"D1"))
      case _ => fail("Expected successful topological sort")
  }

  test("topologicalSort: complex graph with multiple levels") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(1)),
      ref"A2" -> CellValue.Number(BigDecimal(2)),
      ref"B1" -> CellValue.Formula("=A1+A2"),
      ref"B2" -> CellValue.Formula("=A1*2"),
      ref"C1" -> CellValue.Formula("=B1+B2")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.topologicalSort(graph)
    assert(result.isRight)
    result match
      case scala.util.Right(order) =>
        val b1Idx = order.indexOf(ref"B1")
        val b2Idx = order.indexOf(ref"B2")
        val c1Idx = order.indexOf(ref"C1")
        // C1 must come after both B1 and B2
        assert(b1Idx < c1Idx)
        assert(b2Idx < c1Idx)
      case _ => fail("Expected successful topological sort")
  }

  test("topologicalSort: large graph performance (100 nodes)") {
    // Create a large linear chain
    val cells = (1 to 100).map { i =>
      val ref = parseRef(s"A$i")
      if i == 1 then ref -> CellValue.Number(BigDecimal(1))
      else ref -> CellValue.Formula(s"=A${i - 1}+1")
    }
    val sheet = sheetWith(cells*)
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.topologicalSort(graph)
    assert(result.isRight)
    result match
      case scala.util.Right(order) =>
        assertEquals(order.size, 99) // All formula cells (A2 through A100)
      case _ => fail("Expected successful topological sort")
  }

  test("topologicalSort: 100k-node cycle reporting is stack-safe and returns one closed ring") {
    val n = 100000
    val dependencies = (0 until n).iterator.map { i =>
      ARef.from0(0, i) -> Set(ARef.from0(0, (i + 1) % n))
    }.toMap
    // Every node has in-degree one, so Kahn needs no reverse edges to establish the cyclic
    // remainder. Keeping this fixture direct isolates path reconstruction from graph construction.
    val graph = DependencyGraph(dependencies, Map.empty)

    DependencyGraph.topologicalSort(graph) match
      case Left(EvalError.CircularRef(cycle)) =>
        assertEquals(cycle.size, n + 1)
        assertActualCycle(dependencies, cycle)
      case other => fail(s"expected a circular reference, got $other")
  }

  test("detectCrossSheetCycles: a 20k-node cycle is stack-safe") {
    val n = 20000
    val sheet = SheetName.unsafe("Deep")
    val graph = (0 until n).iterator.map { i =>
      val node = DependencyGraph.QualifiedRef(sheet, ARef.from0(0, i))
      val next = DependencyGraph.QualifiedRef(sheet, ARef.from0(0, (i + 1) % n))
      node -> Set(next)
    }.toMap

    DependencyGraph.detectCrossSheetCycles(graph) match
      case Left(EvalError.CircularRef(cycle)) =>
        assertEquals(cycle.size, n + 1)
        assertEquals(cycle.headOption, cycle.lastOption)
      case other => fail(s"expected a circular reference, got $other")
  }

  test("topologicalSort: range dependencies maintain order") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(1)),
      ref"A2" -> CellValue.Number(BigDecimal(2)),
      ref"A3" -> CellValue.Number(BigDecimal(3)),
      ref"B1" -> CellValue.Formula("=SUM(A1:A3)"),
      ref"C1" -> CellValue.Formula("=B1*2")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.topologicalSort(graph)
    assert(result.isRight)
    result match
      case scala.util.Right(order) =>
        val b1Idx = order.indexOf(ref"B1")
        val c1Idx = order.indexOf(ref"C1")
        // B1 must come before C1
        assert(b1Idx < c1Idx)
      case _ => fail("Expected successful topological sort")
  }

  property("GH-518: optimized Kahn preserves the immutable FIFO order exactly") {
    val dagGen = for
      size <- Gen.choose(0, 30)
      rawEdges <- Gen.listOf(Gen.zip(Gen.choose(0, 29), Gen.choose(0, 29)))
      outsiderFlags <- Gen.listOfN(size, Gen.oneOf(true, false))
    yield
      val nodes = Vector.tabulate(size)(row0 => ARef.from0(0, row0))
      val dependencies = nodes.zipWithIndex.map { case (node, index) =>
        val deps = rawEdges.iterator.collect {
          case (from, to) if from == index && to < index => nodes(to)
        }.toSet
        node -> deps
      }.toMap
      val reverse = dependencies.foldLeft(Map.empty[ARef, Set[ARef]]) {
        case (acc, (dependent, precedents)) =>
          precedents.foldLeft(acc) { (seen, precedent) =>
            seen.updated(precedent, seen.getOrElse(precedent, Set.empty) + dependent)
          }
      }
      val outsiders = outsiderFlags.zip(nodes).collect { case (true, node) => node }.toSet
      val outsider = ARef.from0(1, 100)
      val dependents = outsiders.foldLeft(reverse) { (acc, node) =>
        acc.updated(node, acc.getOrElse(node, Set.empty) + outsider)
      }
      DependencyGraph(dependencies, dependents)

    forAll(dagGen) { graph =>
      val allNodes = graph.dependencies.keySet
      val initialDegree = allNodes.iterator.map { node =>
        node -> graph.dependencies.getOrElse(node, Set.empty).count(allNodes.contains)
      }.toMap
      val initialQueue = allNodes.iterator.filter(initialDegree(_) == 0).toList

      @annotation.tailrec
      def reference(
        queue: List[ARef],
        degree: Map[ARef, Int],
        ordered: List[ARef]
      ): List[ARef] =
        queue match
          case Nil => ordered
          case node :: rest =>
            val successors = graph.dependents.getOrElse(node, Set.empty).filter(allNodes.contains)
            val (nextDegree, newlyReady) =
              successors.foldLeft((degree, List.empty[ARef])) {
                case ((current, ready), successor) =>
                  val remaining = current(successor) - 1
                  (
                    current.updated(successor, remaining),
                    if remaining == 0 then ready :+ successor else ready
                  )
              }
            reference(rest ++ newlyReady, nextDegree, ordered :+ node)

      assertEquals(
        DependencyGraph.topologicalSort(graph),
        Right(reference(initialQueue, initialDegree, List.empty))
      )
    }
  }

  property("GH-522: mutable reverse-edge construction equals the immutable reference") {
    val graphGen = for
      size <- Gen.choose(0, 30)
      rawEdges <- Gen.listOf(Gen.zip(Gen.choose(0, 29), Gen.choose(0, 29)))
    yield
      val sheet = SheetName.unsafe("S")
      val nodes =
        Vector.tabulate(size)(row0 => DependencyGraph.QualifiedRef(sheet, ARef.from0(0, row0)))
      nodes.zipWithIndex.map { case (node, index) =>
        val deps = rawEdges.iterator.collect {
          case (from, to) if from == index && to < size => nodes(to)
        }.toSet
        node -> deps
      }.toMap

    forAll(graphGen) { dependencies =>
      val expected = dependencies.foldLeft(
        Map.empty[DependencyGraph.QualifiedRef, Set[DependencyGraph.QualifiedRef]]
      ) { case (acc, (dependent, precedents)) =>
        precedents.foldLeft(acc) { (seen, precedent) =>
          seen.updated(precedent, seen.getOrElse(precedent, Set.empty) + dependent)
        }
      }

      assertEquals(DependencyGraph.reverseEdges(dependencies), expected)
    }
  }

  // ===== Query Tests (4 tests) =====

  test("precedents: returns empty for non-existent node") {
    val graph = DependencyGraph(Map.empty, Map.empty)
    assertEquals(DependencyGraph.precedents(graph, ref"A1"), Set.empty[ARef])
  }

  test("precedents: returns dependencies") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1*2")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(DependencyGraph.precedents(graph, ref"B1"), Set(ref"A1"))
  }

  test("dependents: returns empty for non-existent node") {
    val graph = DependencyGraph(Map.empty, Map.empty)
    assertEquals(DependencyGraph.dependents(graph, ref"A1"), Set.empty[ARef])
  }

  test("dependents: returns cells that depend on this cell") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1*2"),
      ref"C1" -> CellValue.Formula("=A1+5")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(DependencyGraph.dependents(graph, ref"A1"), Set(ref"B1", ref"C1"))
  }

  // ===== Transitive Dependencies Tests (6 tests) =====

  test("transitiveDependencies: empty set returns empty") {
    val graph = DependencyGraph(Map.empty, Map.empty)
    assertEquals(DependencyGraph.transitiveDependencies(graph, Set.empty), Set.empty[ARef])
  }

  test("transitiveDependencies: single node with no deps") {
    val sheet = sheetWith(ref"A1" -> CellValue.Formula("=42"))
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(DependencyGraph.transitiveDependencies(graph, Set(ref"A1")), Set(ref"A1"))
  }

  test("transitiveDependencies: single hop") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1*2")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.transitiveDependencies(graph, Set(ref"B1"))
    // B1 depends on A1
    assertEquals(result, Set(ref"B1", ref"A1"))
  }

  test("transitiveDependencies: multi-hop chain") {
    // A1 <- B1 <- C1 <- D1 (D1 transitively depends on A1)
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(5)),
      ref"B1" -> CellValue.Formula("=A1*2"),
      ref"C1" -> CellValue.Formula("=B1+10"),
      ref"D1" -> CellValue.Formula("=C1*3")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.transitiveDependencies(graph, Set(ref"D1"))
    // D1 -> C1 -> B1 -> A1 (A1 is constant, not in graph dependencies)
    assertEquals(result, Set(ref"D1", ref"C1", ref"B1", ref"A1"))
  }

  test("transitiveDependencies: diamond dependency") {
    // D1 depends on both B1 and C1, both depend on A1
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1+5"),
      ref"C1" -> CellValue.Formula("=A1*2"),
      ref"D1" -> CellValue.Formula("=B1+C1")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.transitiveDependencies(graph, Set(ref"D1"))
    // D1 -> B1, C1 -> A1
    assertEquals(result, Set(ref"D1", ref"B1", ref"C1", ref"A1"))
  }

  test("transitiveDependencies: multiple starting points") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1*2"),
      ref"C1" -> CellValue.Number(BigDecimal(20)),
      ref"D1" -> CellValue.Formula("=C1+5")
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val result = DependencyGraph.transitiveDependencies(graph, Set(ref"B1", ref"D1"))
    // B1 -> A1, D1 -> C1
    assertEquals(result, Set(ref"B1", ref"A1", ref"D1", ref"C1"))
  }

  // ===== Bounded Dependency Extraction Tests =====

  test("extractDependenciesBounded: bounds limits full column range") {
    val bounds = Some(parseRange("A1:Z10"))
    val expr = TExpr.sum(parseRange("A:A")) // Full column A
    val deps = DependencyGraph.extractDependenciesBounded(expr, bounds)
    // Should only include A1:A10, not all 1M+ cells
    assertEquals(deps.size, 10)
    assert(deps.contains(ref"A1"))
    assert(deps.contains(ref"A10"))
    assert(!deps.contains(ref"A11"))
  }

  test("extractDependenciesBounded: bounds limits full row range") {
    val bounds = Some(parseRange("A1:D100"))
    val expr = TExpr.sum(parseRange("1:1")) // Full row 1
    val deps = DependencyGraph.extractDependenciesBounded(expr, bounds)
    // Should only include A1:D1, not all 16K+ cells
    assertEquals(deps.size, 4)
    assert(deps.contains(ref"A1"))
    assert(deps.contains(ref"D1"))
    assert(!deps.contains(ref"E1"))
  }

  test("extractDependenciesBounded: no bounds behaves like original") {
    val expr = TExpr.sum(parseRange("A1:A3"))
    val bounded = DependencyGraph.extractDependenciesBounded(expr, None)
    val original = DependencyGraph.extractDependencies(expr)
    assertEquals(bounded, original)
  }

  test("fromSheet: full column formula uses bounded extraction") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(1)),
      ref"A2" -> CellValue.Number(BigDecimal(2)),
      ref"A3" -> CellValue.Number(BigDecimal(3)),
      ref"B1" -> CellValue.Formula("=SUM(A:A)") // Full column reference
    )
    val graph = DependencyGraph.fromSheet(sheet)
    // Should only include cells in the used range (A1:B1 expanded to A1:A3, B1)
    // Not all 1M+ cells in column A
    val deps = graph.dependencies(ref"B1")
    assertEquals(deps, Set(ref"A1", ref"A2", ref"A3"))
    // Verify it didn't include cells outside used range
    assert(!deps.contains(ref"A100"))
  }

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
    assertEquals(DependencyGraph.dynamicCells(sheet), Set(parseRef("A1"), parseRef("A2")))
  }

  test("GH-301 dynamicCells: OFFSET formulas are dynamic (parity with INDIRECT)") {
    val sheet = sheetWith(
      parseRef("A1") -> CellValue.Formula("=OFFSET(B1,1,0)", None),
      parseRef("A2") -> CellValue.Formula("=SUM(OFFSET(B1,0,0,3,1))", None), // composed
      parseRef("A3") -> CellValue.Formula("=offset(B1,1,0)*2", None), // case + operand position
      parseRef("A4") -> CellValue.Formula("=SUM(B1:B2)", None) // static
    )
    assertEquals(
      DependencyGraph.dynamicCells(sheet),
      Set(parseRef("A1"), parseRef("A2"), parseRef("A3"))
    )
  }

  test("dynamicClosure includes the dynamic cells AND their transitive static dependents") {
    val sheet = sheetWith(
      parseRef("A1") -> CellValue.Formula("=INDIRECT(\"X1\")", None),
      parseRef("B1") -> CellValue.Formula("=A1+1", None),
      parseRef("C1") -> CellValue.Formula("=B1+1", None),
      parseRef("D1") -> CellValue.Formula("=X1*2", None) // not downstream of A1
    )
    val graph = DependencyGraph.fromSheet(sheet)
    val closure = DependencyGraph.dynamicClosure(graph, Set(parseRef("A1")))
    assertEquals(closure, Set(parseRef("A1"), parseRef("B1"), parseRef("C1")))
  }

  test("deferDynamic: empty closure is the identity") {
    val order = List(parseRef("A1"), parseRef("B1"), parseRef("C1"))
    assertEquals(DependencyGraph.deferDynamic(order, Set.empty), order)
  }

  property("deferDynamic lemma: stable partition of a topo order is a topo order, closure last") {
    // Random DAG: node i may only depend on node j < i (acyclic by construction).
    val genCase: Gen[(DependencyGraph, Set[ARef])] =
      for
        n <- Gen.choose(2, 12)
        edges <- Gen.listOf(
          for
            i <- Gen.choose(1, n - 1)
            j <- Gen.choose(0, n - 1).map(_ % math.max(1, i)) // j < i
          yield (ARef.from0(0, i), ARef.from0(0, j))
        )
        seedIdx <- Gen.choose(0, n - 1)
      yield
        val nodes = (0 until n).map(k => ARef.from0(0, k))
        val deps0 = nodes.map(_ -> Set.empty[ARef]).toMap
        val deps = edges.foldLeft(deps0) { case (m, (u, v)) => m.updated(u, m(u) + v) }
        val dependents = deps.foldLeft(Map.empty[ARef, Set[ARef]]) { case (acc, (u, vs)) =>
          vs.foldLeft(acc)((a, v) => a.updated(v, a.getOrElse(v, Set.empty) + u))
        }
        (DependencyGraph(deps, dependents), Set(ARef.from0(0, seedIdx)))

    forAll(genCase) { case (graph, seeds) =>
      DependencyGraph.topologicalSort(graph) match
        case Left(err) => falsified :| s"generated DAG reported cyclic: $err"
        case Right(order) =>
          val closure = DependencyGraph.dynamicClosure(graph, seeds)
          val deferred = DependencyGraph.deferDynamic(order, closure)
          val pos = deferred.zipWithIndex.toMap
          val key = (r: ARef) => (r.row.index0, r.col.index0)
          val isPermutation =
            deferred.sortBy(key) == order.sortBy(key) && deferred.size == order.size
          val isTopoOrder = graph.dependencies.forall { case (u, deps) =>
            deps.filter(pos.contains).forall(v => pos(v) < pos(u))
          }
          val closureIsTail =
            deferred.dropWhile(r => !closure.contains(r)).forall(closure.contains)
          (isPermutation :| "permutation") &&
          (isTopoOrder :| "valid topological order") &&
          (closureIsTail :| "closure occupies the tail")
    }
  }

  // ===== GH-492: SCC condensation order (qualifiedSccOrder) =====

  private val sccUniverse: Vector[DependencyGraph.QualifiedRef] =
    (0 until 6).toVector.map { i =>
      val sheet = if i < 3 then SheetName.unsafe("S1") else SheetName.unsafe("S2")
      DependencyGraph.QualifiedRef(sheet, ARef.from0(0, i % 3))
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

  property("GH-492 cyclic-view: the cyclic components flatten to exactly qualifiedCyclicNodes") {
    forAll(sccGraphGen) { deps =>
      val comps = DependencyGraph.qualifiedSccOrder(deps)
      val fromComps = comps.filter(_.cyclic).flatMap(_.members).toSet
      (fromComps == DependencyGraph.qualifiedCyclicNodes(deps)) :|
        s"cyclic view mismatch: $fromComps vs ${DependencyGraph.qualifiedCyclicNodes(deps)}"
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

  test("GH-492: members inside a component are sorted by (sheet name, A1)") {
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
