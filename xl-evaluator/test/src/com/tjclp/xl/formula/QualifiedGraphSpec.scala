package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, Anchor, CellRange, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.graph.QualifiedGraph
import com.tjclp.xl.formula.graph.QualifiedGraph.{DeclaredRange, Node}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.units.StyleId
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * ADR-017 §2.10: the bounded, cross-sheet inspection graph behind `cell`, `deps` and `audit`.
 *
 * Precedents are exact for single-cell references and expand a range only to the target sheet's
 * OCCUPIED cells, so a full-column reader never materializes a million nodes; dependents come from
 * the symbolic range index, so an empty cell inside a summed range still names its readers. Layers
 * are exactly k hops away, deduplicated against earlier layers, and sorted.
 *
 * The declared view (`declaredPrecedents*`) lists what a formula reads the way Excel's Trace
 * Precedents draws it: each named cell, and each range as ONE node whatever its size, spelled as
 * Excel displays it; the walk continues through the formulas inside a range.
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

  test("a qualifier spelled in another case resolves to the workbook's sheet, as in Excel") {
    // `sheet1!A1` on a book whose tab is `Sheet1`: one node, not a phantom per spelling
    val wb = Workbook(
      Vector(
        sheetWith("Sheet1", "A1" -> num(5), "A2" -> num(7)),
        sheetWith("Summary", "B1" -> formula("sheet1!A1+SUM(SHEET1!A:A)"))
      )
    )
    val graph = QualifiedGraph.of(wb)
    assertEquals(
      graph.precedentsOf(q("Summary", "B1")),
      Set(q("Sheet1", "A1"), q("Sheet1", "A2"))
    )
    assertEquals(graph.dependentsOf(q("Sheet1", "A1")), Set(q("Summary", "B1")))
    assertEquals(graph.dependentsOf(q("Sheet1", "A2")), Set(q("Summary", "B1")))
    // the range reader is indexed under the tab's spelling too
    assertEquals(graph.rangeReaders.keySet, Set(SheetName.unsafe("Sheet1")))
    assertEquals(
      graph.declaredPrecedentsOf(q("Summary", "B1")).map(_.label),
      Vector("Sheet1!A1", "Sheet1!A:A")
    )
    // a cycle routed through a re-spelled qualifier is still a cycle
    val loop = Workbook(
      Vector(
        sheetWith("Alpha", "A1" -> formula("beta!A1+1")),
        sheetWith("Beta", "A1" -> formula("ALPHA!A1+1"))
      )
    )
    assertEquals(QualifiedGraph.of(loop).sccs.count(_.cyclic), 1)
    // a qualifier naming no sheet at all is left as written (no edges: nothing to expand)
    val unknown = Workbook(Vector(sheetWith("Only", "A1" -> formula("Nowhere!A1+1"))))
    assertEquals(
      QualifiedGraph.of(unknown).precedentsOf(q("Only", "A1")),
      Set(q("Nowhere", "A1"))
    )
  }

  test("unparseable and data-table formulas are nodes without edges; empty book is empty") {
    val wb = Workbook(Vector(sheetWith("S", "A1" -> formula("UNSUPPORTED(1)"), "A2" -> num(1))))
    val graph = QualifiedGraph.of(wb)
    assertEquals(graph.dependencies, Map(q("S", "A1") -> Set.empty[QualifiedRef]))
    assertEquals(graph.precedents(q("S", "A1"), 0), Vector.empty)
    assertEquals(graph.declaredPrecedentsOf(q("S", "A1")), Vector.empty)
    assertEquals(graph.declaredPrecedents(q("S", "A1"), 0), Vector.empty)
    assertEquals(QualifiedGraph.of(Workbook(Vector.empty)), QualifiedGraph.empty)
    assertEquals(QualifiedGraph.empty.declaredPrecedentsOf(q("S", "A1")), Vector.empty)
  }

  test("a cycle does not loop the layer walk: every node appears once") {
    val wb = Workbook(Vector(sheetWith("S", "A1" -> formula("B1+1"), "B1" -> formula("A1+1"))))
    val graph = QualifiedGraph.of(wb)
    assertEquals(graph.precedents(q("S", "A1"), 0), Vector(Vector(q("S", "B1"))))
    assertEquals(graph.dependents(q("S", "A1"), 0), Vector(Vector(q("S", "B1"))))
  }

  // ---------------------------------------------------------------------------------------------
  // The declared view: a range is one node
  // ---------------------------------------------------------------------------------------------

  private def text(s: String): CellValue = CellValue.Text(s)
  private def span(from: String, to: String): CellRange = CellRange(a1(from), a1(to))
  private def declared(sheet: String, from: String, to: String): DeclaredRange =
    DeclaredRange(SheetName.unsafe(sheet), span(from, to))
  private def cellNode(sheet: String, ref: String): Node = Node.Cell(q(sheet, ref))
  private def rangeNode(sheet: String, from: String, to: String): Node =
    Node.Range(declared(sheet, from, to))
  private val lastRow = "1048576"

  /**
   * The worked example: a whole-column SUMIFS, a block of formulas, a point, a same-sheet criterion
   * — and a style-only blank at Data!B6 that no range may count.
   */
  private val ranges: Workbook = Workbook(
    Vector(
      sheetWith(
        "Data",
        "A1" -> text("Region"),
        "B1" -> text("Amount"),
        "C1" -> text("Double"),
        "E1" -> num(2),
        "A2" -> text("North"),
        "A3" -> text("South"),
        "A4" -> text("North"),
        "B2" -> num(10),
        "B3" -> num(20),
        "B4" -> num(30),
        "C2" -> formula("B2*$E$1"),
        "C3" -> formula("B3*$E$1"),
        "C4" -> formula("B4*$E$1")
      ).put(Cell(a1("B6"), CellValue.Empty, Some(StyleId(1)))),
      sheetWith(
        "Summary",
        "A1" -> text("North"),
        "B1" -> formula("SUMIFS(Data!B:B,Data!A:A,A1)+SUM(Data!C2:C4)+Data!B2")
      )
    )
  )

  test("a whole-column range is ONE declared precedent, spelled as Excel displays it") {
    val wb = Workbook(
      Vector(
        sheetWith("Data", "A1" -> num(1), "A2" -> num(2), "A3" -> num(3), "B1" -> num(9)),
        sheetWith(
          "Calc",
          "A1" -> formula("SUM(Data!A:A)"),
          "A2" -> formula(s"SUM(Data!A1:A$lastRow)")
        )
      )
    )
    val graph = QualifiedGraph.of(wb)
    val col = declared("Data", "A1", s"A$lastRow")
    assertEquals(graph.declaredPrecedentsOf(q("Calc", "A1")), Vector(Node.Range(col)))
    assertEquals(graph.declaredPrecedentsOf(q("Calc", "A1")).map(_.label), Vector("Data!A:A"))
    assertEquals(graph.occupiedIn(col), Set(q("Data", "A1"), q("Data", "A2"), q("Data", "A3")))
    // the corner spelling of a whole column is the same node
    assertEquals(graph.declaredPrecedentsOf(q("Calc", "A2")), Vector(Node.Range(col)))
    assertEquals(graph.rangeCells.keySet, Set(col))
    // the expanded view is untouched
    assertEquals(graph.precedentsOf(q("Calc", "A1")), graph.occupiedIn(col))
  }

  test("DeclaredRange spelling: whole rows first, then whole columns, a 1x1 as its cell") {
    assertEquals(declared("Data", "A1", s"A$lastRow").a1, "A:A")
    assertEquals(declared("Data", "A1", s"C$lastRow").a1, "A:C")
    assertEquals(declared("Data", "A2", "XFD3").a1, "2:3")
    assertEquals(declared("Data", "A1", s"XFD$lastRow").a1, s"1:$lastRow")
    assertEquals(declared("Data", "B2", "C9").a1, "B2:C9")
    assertEquals(declared("Data", "C1", "C1").a1, "C1")
    assertEquals(declared("Data", "A2", "B9").toString, "Data!A2:B9")
    assertEquals(declared("On-Premise", "A1", s"A$lastRow").toString, "'On-Premise'!A:A")
    assertEquals(declared("Q1 Data", "A1", "B2").toString, "'Q1 Data'!A1:B2")
    // a sheet named like a cell is quoted, as QualifiedRef quotes it
    assertEquals(declared("A1", "B2", "B3").toString, "'A1'!B2:B3")
    assertEquals(rangeNode("Data", "A3", "XFD5").label, "Data!3:5")
    assertEquals(cellNode("Data", "B2").label, "Data!B2")
  }

  test("anchors and qualifier case never split a node") {
    val wb = Workbook(
      Vector(
        sheetWith(
          "S",
          "A1" -> num(1),
          "A2" -> num(2),
          "A3" -> num(3),
          "B1" -> formula("SUM($A$1:$A$3)"),
          "B2" -> formula("SUM(A1:A3)"),
          "B3" -> formula("SUM(s!A$1:$A3)+SUM(A1:A3)")
        )
      )
    )
    val graph = QualifiedGraph.of(wb)
    val one = Vector(rangeNode("S", "A1", "A3"))
    assertEquals(graph.declaredPrecedentsOf(q("S", "B1")), one)
    assertEquals(graph.declaredPrecedentsOf(q("S", "B2")), one)
    assertEquals(graph.declaredPrecedentsOf(q("S", "B3")), one)
    assertEquals(graph.rangeCells.keySet, Set(declared("S", "A1", "A3")))
    // an anchored key finds the same cells
    val anchored = DeclaredRange(
      SheetName.unsafe("S"),
      CellRange(a1("A1"), a1("A3"), Anchor.Absolute, Anchor.Absolute)
    )
    assertEquals(graph.occupiedIn(anchored), Set(q("S", "A1"), q("S", "A2"), q("S", "A3")))
  }

  test("a 1x1 range slot is a cell precedent, occupied or not") {
    val wb = Workbook(Vector(sheetWith("S", "B1" -> formula("COUNTIF(Z9,\"x\")"))))
    val graph = QualifiedGraph.of(wb)
    assertEquals(graph.declaredPrecedentsOf(q("S", "B1")), Vector(cellNode("S", "Z9")))
    // the expanded view lists only the value-holding cells of a range
    assertEquals(graph.precedentsOf(q("S", "B1")), Set.empty)
  }

  test("an empty range and a range on a missing sheet are listed, with no occupied cells") {
    val wb = Workbook(
      Vector(
        sheetWith(
          "S",
          "A1" -> num(1),
          "B1" -> formula("SUM(M1:M9)"),
          "B2" -> formula("SUM(Missing!A1:A3)")
        )
      )
    )
    val graph = QualifiedGraph.of(wb)
    assertEquals(graph.declaredPrecedentsOf(q("S", "B1")), Vector(rangeNode("S", "M1", "M9")))
    assertEquals(graph.occupiedIn(declared("S", "M1", "M9")), Set.empty)
    assertEquals(graph.declaredPrecedentsOf(q("S", "B2")).map(_.label), Vector("Missing!A1:A3"))
    assertEquals(graph.occupiedIn(declared("Missing", "A1", "A3")), Set.empty)
    assertEquals(
      graph.declaredPrecedents(q("S", "B2"), 0),
      Vector(Vector(rangeNode("Missing", "A1", "A3")))
    )
  }

  test("a style-only or explicitly empty cell inside a range is not occupied") {
    val wb = Workbook(
      Vector(
        sheetWith("Data", "B1" -> num(1), "B2" -> num(2), "B3" -> num(3), "B4" -> num(4))
          .put(Cell(a1("B6"), CellValue.Empty, Some(StyleId(1))))
          .put(a1("B7"), CellValue.Empty),
        sheetWith("Calc", "A1" -> formula("SUM(Data!B:B)"), "A2" -> formula("Data!B6"))
      )
    )
    val graph = QualifiedGraph.of(wb)
    val values = Set(q("Data", "B1"), q("Data", "B2"), q("Data", "B3"), q("Data", "B4"))
    assertEquals(graph.occupiedIn(declared("Data", "B1", s"B$lastRow")), values)
    assertEquals(graph.precedentsOf(q("Calc", "A1")), values)
    // a cell named individually is a precedent whatever it holds
    assertEquals(graph.precedentsOf(q("Calc", "A2")), Set(q("Data", "B6")))
    assertEquals(graph.declaredPrecedentsOf(q("Calc", "A2")), Vector(cellNode("Data", "B6")))
  }

  test("declaredPrecedents walks through a range's formulas; its constants are leaves") {
    val graph = QualifiedGraph.of(ranges)
    val layer1 = Vector(
      rangeNode("Data", "A1", s"A$lastRow"),
      rangeNode("Data", "B1", s"B$lastRow"),
      cellNode("Data", "B2"),
      rangeNode("Data", "C2", "C4"),
      cellNode("Summary", "A1")
    )
    // B2..B4 are covered by Data!B:B at depth 1: only E1 is new at depth 2
    val layer2 = Vector(cellNode("Data", "E1"))
    assertEquals(graph.declaredPrecedents(q("Summary", "B1"), 0), Vector(layer1, layer2))
    assertEquals(graph.declaredPrecedents(q("Summary", "B1"), 1), Vector(layer1))
    assertEquals(graph.declaredPrecedentsOf(q("Summary", "B1")), layer1)
    assertEquals(
      layer1.map(_.label),
      Vector("Data!A:A", "Data!B:B", "Data!B2", "Data!C2:C4", "Summary!A1")
    )
    assertEquals(graph.occupiedIn(declared("Data", "B1", s"B$lastRow")).size, 4)
    assertEquals(graph.occupiedIn(declared("Data", "C2", "C4")).count(graph.formulas.contains), 3)
  }

  test("each declared layer covers exactly the cells of the expanded layer") {
    // D1 reads a block of two formulas and a constant; the block's formulas read a range and a point
    val diamond = Workbook(
      Vector(
        sheetWith(
          "S",
          "A1" -> num(1),
          "A2" -> num(2),
          "B1" -> formula("SUM(A1:A2)"),
          "C1" -> formula("A1*3"),
          "D1" -> formula("SUM(B1:C1)+A2")
        )
      )
    )
    val diamondGraph = QualifiedGraph.of(diamond)
    assertEquals(
      diamondGraph.declaredPrecedents(q("S", "D1"), 0),
      Vector(
        Vector(rangeNode("S", "B1", "C1"), cellNode("S", "A2")),
        Vector(cellNode("S", "A1"), rangeNode("S", "A1", "A2"))
      )
    )
    Vector((diamondGraph, q("S", "D1")), (QualifiedGraph.of(ranges), q("Summary", "B1"))).foreach {
      (graph, start) =>
        val covered = graph
          .declaredPrecedents(start, 0)
          .foldLeft((Set(start), Vector.empty[Set[QualifiedRef]])) { case ((seen, acc), layer) =>
            val cells = layer.iterator.flatMap {
              case Node.Cell(c) => Iterator.single(c)
              case Node.Range(r) => graph.occupiedIn(r).iterator
            }.toSet -- seen
            (seen ++ cells, acc :+ cells)
          }
          ._2
          .filter(_.nonEmpty)
        assertEquals(covered, graph.precedents(start, 0).map(_.toSet))
    }
  }

  test("a cell inside an earlier range is not relisted; a range reached later is still listed") {
    val wb = Workbook(
      Vector(
        sheetWith(
          "S",
          "A1" -> num(1),
          "A2" -> num(2),
          "A3" -> num(3),
          "B1" -> formula("SUM(A1:A3)+C1"),
          "C1" -> formula("A2*2"),
          "D1" -> formula("A2+E1"),
          "E1" -> formula("SUM(A1:A3)")
        )
      )
    )
    val graph = QualifiedGraph.of(wb)
    assertEquals(
      graph.declaredPrecedents(q("S", "B1"), 0),
      Vector(Vector(rangeNode("S", "A1", "A3"), cellNode("S", "C1")))
    )
    assertEquals(
      graph.declaredPrecedents(q("S", "D1"), 0),
      // row-major within the layer: E1 (row 1) before A2 (row 2)
      Vector(Vector(cellNode("S", "E1"), cellNode("S", "A2")), Vector(rangeNode("S", "A1", "A3")))
    )
  }

  test("a range reached later is listed once though earlier ranges covered it; it walks nothing") {
    val wb = Workbook(
      Vector(
        sheetWith(
          "S",
          "A1" -> num(1),
          "A2" -> num(2),
          "A3" -> num(3),
          "A4" -> num(4),
          "A5" -> formula("SUM(A1:A4)"),
          "B1" -> formula("SUM(A:A)"),
          "B2" -> formula("SUM(A:A)+SUM(A1:A4)")
        )
      )
    )
    val graph = QualifiedGraph.of(wb)
    // A:A covers A1..A5 at depth 1; A5's A1:A4 is the author's reference, listed at depth 2
    assertEquals(
      graph.declaredPrecedents(q("S", "B1"), 0),
      Vector(Vector(rangeNode("S", "A1", s"A$lastRow")), Vector(rangeNode("S", "A1", "A4")))
    )
    // at its shortest distance only: B2 names A1:A4 itself, so A5 adds no layer
    assertEquals(
      graph.declaredPrecedents(q("S", "B2"), 0),
      Vector(Vector(rangeNode("S", "A1", "A4"), rangeNode("S", "A1", s"A$lastRow")))
    )
  }

  test("overlapping and nested ranges are distinct nodes") {
    val wb = Workbook(
      Vector(
        sheetWith(
          "S",
          "A1" -> num(1),
          "A2" -> num(2),
          "A3" -> num(3),
          "A4" -> num(4),
          "B2" -> num(5),
          "B3" -> num(6),
          "B4" -> num(7),
          "C1" -> formula("SUM(A:A)+SUM(A1:A3)+SUM(A2:B4)")
        )
      )
    )
    val graph = QualifiedGraph.of(wb)
    val nodes = Vector(
      rangeNode("S", "A1", "A3"),
      rangeNode("S", "A1", s"A$lastRow"),
      rangeNode("S", "A2", "B4")
    )
    assertEquals(graph.declaredPrecedentsOf(q("S", "C1")), nodes)
    assertEquals(graph.declaredPrecedents(q("S", "C1"), 0), Vector(nodes))
    // each range counts its own cells, shared ones included
    assertEquals(
      Vector(
        declared("S", "A1", "A3"),
        declared("S", "A1", s"A$lastRow"),
        declared("S", "A2", "B4")
      ).map(graph.occupiedIn(_).size),
      Vector(3, 4, 6)
    )
  }

  test("in one layer a named cell and a range containing it are both listed") {
    val wb = Workbook(
      Vector(
        sheetWith(
          "S",
          "A1" -> num(1),
          "A2" -> num(2),
          "A3" -> num(3),
          "F1" -> formula("A1+SUM(A1:A3)")
        )
      )
    )
    assertEquals(
      QualifiedGraph.of(wb).declaredPrecedents(q("S", "F1"), 0),
      Vector(Vector(cellNode("S", "A1"), rangeNode("S", "A1", "A3")))
    )
  }

  test("a cycle through a range terminates the declared walk") {
    val wb = Workbook(
      Vector(
        sheetWith("S", "A1" -> formula("A2+1"), "A2" -> formula("SUM(A1:A3)"), "A3" -> num(5))
      )
    )
    assertEquals(
      QualifiedGraph.of(wb).declaredPrecedents(q("S", "A2"), 0),
      Vector(Vector(rangeNode("S", "A1", "A3")))
    )
  }

  test("nodeOrder: sheet name, top-left corner, a cell before a range there, then bottom-right") {
    val expected = Vector(
      rangeNode("Data", "A1", s"A$lastRow"),
      cellNode("Data", "B1"),
      rangeNode("Data", "B1", "B3"),
      rangeNode("Data", "B1", s"B$lastRow"),
      cellNode("Data", "A2"),
      rangeNode("Data", "A2", "C2"),
      cellNode("Data", "A10"),
      cellNode("Summary", "A1")
    )
    val shuffled = Vector(7, 2, 5, 0, 6, 3, 1, 4).flatMap(expected.lift)
    assertEquals(shuffled.sorted(using QualifiedGraph.nodeOrder), expected)
  }
