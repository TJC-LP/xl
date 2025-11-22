package com.tjclp.xl.formula

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

/**
 * Tests for DependencyGraph (WI-09d).
 *
 * Tests dependency extraction, graph construction, cycle detection (Tarjan's SCC), and topological
 * sorting (Kahn's algorithm).
 */
class DependencyGraphSpec extends FunSuite:
  val emptySheet = new Sheet(name = SheetName.unsafe("Test"))

  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  private def parseRange(range: String): CellRange =
    CellRange.parse(range).fold(err => fail(err), identity)

  private def parseRef(ref: String): ARef =
    ARef.parse(ref).fold(err => fail(err), identity)

  // ===== Dependency Extraction Tests (10 tests) =====

  test("extractDependencies: Lit has no dependencies") {
    val expr = TExpr.Lit(BigDecimal(42))
    assertEquals(DependencyGraph.extractDependencies(expr), Set.empty[ARef])
  }

  test("extractDependencies: Ref has single dependency") {
    val expr = TExpr.Ref(ref"A1", TExpr.decodeNumeric)
    assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1"))
  }

  test("extractDependencies: Add with two Refs") {
    val expr = TExpr.Add(
      TExpr.Ref(ref"A1", TExpr.decodeNumeric),
      TExpr.Ref(ref"B1", TExpr.decodeNumeric)
    )
    assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1", ref"B1"))
  }

  test("extractDependencies: FoldRange expands to all cells in range") {
    val range = parseRange("A1:A3")
    val expr = TExpr.sum(range)
    val expected = Set(ref"A1", ref"A2", ref"A3")
    assertEquals(DependencyGraph.extractDependencies(expr), expected)
  }

  test("extractDependencies: nested expressions") {
    val expr = TExpr.Mul(
      TExpr.Add(TExpr.Ref(ref"A1", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(10))),
      TExpr.Ref(ref"B1", TExpr.decodeNumeric)
    )
    assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1", ref"B1"))
  }

  test("extractDependencies: If with three branches") {
    val expr = TExpr.If(
      TExpr.Gt(TExpr.Ref(ref"A1", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(0))),
      TExpr.Ref(ref"B1", TExpr.decodeNumeric),
      TExpr.Ref(ref"C1", TExpr.decodeNumeric)
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
    val expr = TExpr.Concatenate(List(
      TExpr.Ref(ref"A1", TExpr.decodeString),
      TExpr.Lit(" "),
      TExpr.Ref(ref"B1", TExpr.decodeString)
    ))
    assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1", ref"B1"))
  }

  test("extractDependencies: date functions with Refs") {
    val expr = TExpr.Date(
      TExpr.Ref(ref"A1", TExpr.decodeInt),
      TExpr.Ref(ref"A2", TExpr.decodeInt),
      TExpr.Ref(ref"A3", TExpr.decodeInt)
    )
    assertEquals(DependencyGraph.extractDependencies(expr), Set(ref"A1", ref"A2", ref"A3"))
  }

  test("extractDependencies: complex nested with ranges") {
    val expr = TExpr.Add(
      TExpr.sum(parseRange("A1:A3")),
      TExpr.Mul(TExpr.Ref(ref"B1", TExpr.decodeNumeric), TExpr.Lit(BigDecimal(2)))
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
      ref"D1" -> CellValue.Formula("=C1*2")   // Depends on C1 (cycle)
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
      ref"B2" -> CellValue.Formula("=B1")  // Cycle 2: B1 <-> B2
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
