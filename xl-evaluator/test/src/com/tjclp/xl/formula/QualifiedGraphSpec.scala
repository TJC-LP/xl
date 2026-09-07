package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.graph.QualifiedGraph
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * ADR-017 §2.10: the bounded, cross-sheet inspection graph behind `cell`, `deps` and `audit`.
 *
 * Precedents are exact for single-cell references and expand a range only to the target sheet's
 * OCCUPIED cells, so a full-column reader never materializes a million nodes; dependents come from
 * the symbolic range index, so an empty cell inside a summed range still names its readers. Layers
 * are exactly k hops away, deduplicated against earlier layers, and sorted.
 */
class QualifiedGraphSpec extends FunSuite:

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def formula(expr: String): CellValue = CellValue.Formula(expr, None)
  private def a1(s: String): ARef = ARef.parse(s).fold(err => fail(err), identity)
  private def q(sheet: String, ref: String): QualifiedRef =
    QualifiedRef(SheetName.unsafe(sheet), a1(ref))

  private def sheetWith(name: String, cells: (String, CellValue)*): Sheet =
    cells.foldLeft(Sheet(SheetName.unsafe(name))) { case (s, (ref, value)) =>
      s.put(a1(ref), value)
    }

  /** A → B → C across three sheets: C!A1 reads B!A1, which reads A!A1 (a constant). */
  private val chain: Workbook = Workbook(
    Vector(
      sheetWith("A", "A1" -> num(5)),
      sheetWith("B", "A1" -> formula("A!A1*2")),
      sheetWith("C", "A1" -> formula("B!A1+1"), "B1" -> formula("A1*10"))
    )
  )

  test("precedents(C, 2) is two layers: the formula it reads, then that formula's constant input") {
    val graph = QualifiedGraph.of(chain)
    assertEquals(
      graph.precedents(q("C", "A1"), 2),
      Vector(Vector(q("B", "A1")), Vector(q("A", "A1")))
    )
  }

  test("precedents at depth 1 stop after one hop; depth 0 is unbounded") {
    val graph = QualifiedGraph.of(chain)
    assertEquals(graph.precedents(q("C", "B1"), 1), Vector(Vector(q("C", "A1"))))
    assertEquals(
      graph.precedents(q("C", "B1"), 0),
      Vector(Vector(q("C", "A1")), Vector(q("B", "A1")), Vector(q("A", "A1")))
    )
  }

  test("dependents(A, 0) reaches every reader transitively: {B!A1, C!A1, C!B1} in layers") {
    val graph = QualifiedGraph.of(chain)
    val layers = graph.dependents(q("A", "A1"), 0)
    assertEquals(layers, Vector(Vector(q("B", "A1")), Vector(q("C", "A1")), Vector(q("C", "B1"))))
    assertEquals(layers.flatten.toSet, Set(q("B", "A1"), q("C", "A1"), q("C", "B1")))
  }

  test("a constant has no precedents and a leaf formula has no dependents") {
    val graph = QualifiedGraph.of(chain)
    assertEquals(graph.precedents(q("A", "A1"), 0), Vector.empty)
    assertEquals(graph.dependents(q("C", "B1"), 0), Vector.empty)
    assertEquals(graph.precedentsOf(q("A", "A1")), Set.empty)
    assertEquals(graph.dependentsOf(q("C", "B1")), Set.empty)
  }

  test("an A:A reader expands only to the occupied cells of the column, never to 1,048,576") {
    val wb = Workbook(
      Vector(
        sheetWith("Data", "A1" -> num(1), "A2" -> num(2), "A3" -> num(3), "B1" -> num(9)),
        sheetWith("Calc", "A1" -> formula("SUM(Data!A:A)"))
      )
    )
    val graph = QualifiedGraph.of(wb)
    val direct = graph.precedentsOf(q("Calc", "A1"))
    assertEquals(direct, Set(q("Data", "A1"), q("Data", "A2"), q("Data", "A3")))
    val nodes = graph.dependencies.iterator.map { case (k, v) => v.size + 1 }.sum
    assert(nodes <= 8, s"graph must stay bounded by occupancy, saw $nodes nodes")
  }

  test("an empty cell inside a read range still reports its readers (symbolic range index)") {
    val wb = Workbook(
      Vector(sheetWith("S", "A1" -> num(1), "A3" -> num(3), "B1" -> formula("SUM(A1:A5)")))
    )
    val graph = QualifiedGraph.of(wb)
    // A2 and A5 are empty: no forward edge names them, yet the range covers them
    assertEquals(graph.dependentsOf(q("S", "A2")), Set(q("S", "B1")))
    assertEquals(graph.dependents(q("S", "A5"), 1), Vector(Vector(q("S", "B1"))))
    // outside the range: nothing
    assertEquals(graph.dependentsOf(q("S", "A6")), Set.empty)
    assertEquals(graph.dependentsOf(q("S", "C1")), Set.empty)
    // and the forward side lists only the occupied cells
    assertEquals(graph.precedentsOf(q("S", "B1")), Set(q("S", "A1"), q("S", "A3")))
  }

  test("sccs finds a two-cycle across sheets and marks the rest acyclic") {
    val wb = Workbook(
      Vector(
        sheetWith("X", "A1" -> formula("Y!A1+1"), "B1" -> num(4)),
        sheetWith("Y", "A1" -> formula("X!A1+1"), "B1" -> formula("X!B1*2"))
      )
    )
    val graph = QualifiedGraph.of(wb)
    val cyclic = graph.sccs.filter(_.cyclic)
    assertEquals(cyclic.map(_.members.toSet), Vector(Set(q("X", "A1"), q("Y", "A1"))))
    assertEquals(graph.sccs.filterNot(_.cyclic).flatMap(_.members), Vector(q("Y", "B1")))
    // a self-loop is a cyclic singleton
    val self = QualifiedGraph.of(Workbook(Vector(sheetWith("Z", "A1" -> formula("A1+1")))))
    assertEquals(self.sccs.map(s => (s.members, s.cyclic)), Vector((Vector(q("Z", "A1")), true)))
  }

  test("layers are deduplicated: a diamond reports each node once, at its shortest distance") {
    //     D
    //    / \
    //   B   C
    //    \ /
    //     A
    val wb = Workbook(
      Vector(
        sheetWith(
          "S",
          "A1" -> num(1),
          "B1" -> formula("A1*2"),
          "C1" -> formula("A1*3"),
          "D1" -> formula("B1+C1+A1")
        )
      )
    )
    val graph = QualifiedGraph.of(wb)
    assertEquals(
      graph.precedents(q("S", "D1"), 0),
      Vector(Vector(q("S", "A1"), q("S", "B1"), q("S", "C1")))
    )
    assertEquals(
      graph.dependents(q("S", "A1"), 0),
      Vector(Vector(q("S", "B1"), q("S", "C1"), q("S", "D1")))
    )
  }

  test("layers are sorted by sheet, row, column regardless of insertion order") {
    val wb = Workbook(
      Vector(
        sheetWith("S", "A1" -> formula("B!C3+B!A10+B!A2+A2"), "A2" -> num(1)),
        sheetWith("B", "A2" -> num(1), "A10" -> num(2), "C3" -> num(3))
      )
    )
    val graph = QualifiedGraph.of(wb)
    assertEquals(
      graph.precedents(q("S", "A1"), 1),
      Vector(Vector(q("B", "A2"), q("B", "C3"), q("B", "A10"), q("S", "A2")))
    )
  }

  test("unparseable and data-table formulas are nodes without edges; empty book is empty") {
    val wb = Workbook(Vector(sheetWith("S", "A1" -> formula("UNSUPPORTED(1)"), "A2" -> num(1))))
    val graph = QualifiedGraph.of(wb)
    assertEquals(graph.dependencies, Map(q("S", "A1") -> Set.empty[QualifiedRef]))
    assertEquals(graph.precedents(q("S", "A1"), 0), Vector.empty)
    assertEquals(QualifiedGraph.of(Workbook(Vector.empty)), QualifiedGraph.empty)
  }

  test("a cycle does not loop the layer walk: every node appears once") {
    val wb = Workbook(Vector(sheetWith("S", "A1" -> formula("B1+1"), "B1" -> formula("A1+1"))))
    val graph = QualifiedGraph.of(wb)
    assertEquals(graph.precedents(q("S", "A1"), 0), Vector(Vector(q("S", "B1"))))
    assertEquals(graph.dependents(q("S", "A1"), 0), Vector(Vector(q("S", "B1"))))
  }
