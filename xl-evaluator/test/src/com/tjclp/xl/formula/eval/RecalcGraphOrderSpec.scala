package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.graph.{CompactGraphGens, DependencyGraph, NaiveGraphs}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.rng.Seed

/**
 * #891: [[RecalcGraph.withinComponentOrder]], the sweep order a Gauss–Seidel fixpoint evaluates a
 * cycle's members in. It answers on formula ids when the graph holds every member, and through
 * `DependencyGraph.withinComponentOrder` over the members' reads otherwise. Law, on both sides of
 * that guard: the answer is `DependencyGraph.withinComponentOrder(members, deps)` over the per-cell
 * graph of the workbook ([[NaiveGraphs.fromWorkbookFormulaGraph]]).
 *
 * A member the graph does not hold (a constant, a blank, a cell of a sheet the workbook lacks) is
 * kept, and sweeps before any cycle is cut: it reads nothing, so Kahn's frontier holds it from the
 * start, in grid order among the other members that read no member. That matters to the caller,
 * which pairs every member of the order with its formula: a member dropped from the order would
 * never be evaluated.
 */
class RecalcGraphOrderSpec extends ScalaCheckSuite:

  private def q(sheet: String, a1: String): QualifiedRef =
    QualifiedRef(SheetName.unsafe(sheet), ARef.parse(a1).fold(fail(_), identity))

  private def book(sheets: (String, List[(String, String)])*): Workbook =
    Workbook(sheets.toVector.map { (name, cells) =>
      cells.foldLeft(Sheet(SheetName.unsafe(name))) { case (sheet, (a1, text)) =>
        val value =
          if text.startsWith("=") then CellValue.Formula(text, None)
          else CellValue.Number(BigDecimal(text))
        sheet.put(ARef.parse(a1).fold(fail(_), identity), value)
      }
    })

  private def graph(wb: Workbook): RecalcGraph = RecalcGraph.build(wb, RecalcCounters.disabled)

  /** The Map form over the workbook's per-cell graph: the oracle for any member vector. */
  private def oracle(wb: Workbook, members: Vector[QualifiedRef]): Vector[QualifiedRef] =
    DependencyGraph.withinComponentOrder(members, NaiveGraphs.fromWorkbookFormulaGraph(wb)._1)

  // A1 reads B1, B1 reads C1, C1 reads A1; D1 is a constant. The cut falls on A1 (top-left), which
  // frees its reader C1, which frees B1: the sweep is A1, C1, B1, not the grid order A1, B1, C1.
  private val ring: Workbook =
    book("S" -> List("A1" -> "=B1+1", "B1" -> "=C1*0.5", "C1" -> "=A1*0.5", "D1" -> "7"))
  private val ringSweep = Vector(q("S", "A1"), q("S", "C1"), q("S", "B1"))

  test("every member known: the sweep follows the reads, not the grid") {
    val members = Vector(q("S", "A1"), q("S", "B1"), q("S", "C1"))
    assertEquals(graph(ring).withinComponentOrder(members), ringSweep)
    assertEquals(oracle(ring, members), ringSweep)
    // the order is a function of the member set, not of the vector's order
    members.permutations.foreach(p => assertEquals(graph(ring).withinComponentOrder(p), ringSweep))
  }

  test("an empty member vector has an empty order") {
    assertEquals(graph(ring).withinComponentOrder(Vector.empty), Vector.empty)
  }

  test("duplicate members are swept once") {
    val members = Vector(q("S", "B1"), q("S", "A1"), q("S", "B1"), q("S", "C1"), q("S", "A1"))
    assertEquals(graph(ring).withinComponentOrder(members), ringSweep)
    assertEquals(oracle(ring, members), ringSweep)
  }

  test("a constant member is kept and swept first, wherever it sits in the grid") {
    // D1 is last in grid order. It reads nothing, so it is ready before the cycle is cut.
    val members = Vector(q("S", "A1"), q("S", "B1"), q("S", "C1"), q("S", "D1"))
    val expected = q("S", "D1") +: ringSweep
    assertEquals(graph(ring).withinComponentOrder(members), expected)
    assertEquals(oracle(ring, members), expected)
  }

  test("members on a sheet the workbook lacks, and blank cells, are kept and swept first") {
    // none of the three is a node of the graph; among themselves they come out in grid order
    // (sheet name, then row, then column): Ghost!B2, S!E9, Zed!A1
    val unknown = Vector(q("Zed", "A1"), q("S", "E9"), q("Ghost", "B2"))
    val members = unknown ++ Vector(q("S", "C1"), q("S", "B1"), q("S", "A1"))
    val expected = Vector(q("Ghost", "B2"), q("S", "E9"), q("Zed", "A1")) ++ ringSweep
    assertEquals(graph(ring).withinComponentOrder(members), expected)
    assertEquals(oracle(ring, members), expected)
    // with duplicates of a known and of an unknown member
    val repeated = members ++ Vector(q("S", "A1"), q("Zed", "A1"))
    assertEquals(graph(ring).withinComponentOrder(repeated), expected)
  }

  test("only unknown members: grid order") {
    val members = Vector(q("S", "D1"), q("Ghost", "B2"), q("S", "E9"))
    assertEquals(
      graph(ring).withinComponentOrder(members),
      Vector(q("Ghost", "B2"), q("S", "D1"), q("S", "E9"))
    )
  }

  // One cycle over two sheets held in the order "b", "Z": Z!B2 reads b!A1, b!A1 reads b!C1, b!C1
  // reads Z!B2. Grid order compares sheet NAMES ("Z" sorts before "b"), so the cut falls on Z!B2,
  // which frees its reader b!C1, which frees b!A1. Were sheets ranked by their place in the
  // workbook, the cut would fall on b!A1 and the sweep would be b!A1, Z!B2, b!C1.
  private val twoSheets: Workbook = book(
    "b" -> List("A1" -> "=C1+1", "C1" -> "='Z'!B2*0.5", "B5" -> "3"),
    "Z" -> List("B2" -> "='b'!A1*0.5", "A1" -> "=B2+1")
  )
  private val twoSheetSweep = Vector(q("Z", "B2"), q("b", "C1"), q("b", "A1"))

  test("a cycle over sheets whose workbook order is not their name order") {
    val members = Vector(q("b", "A1"), q("b", "C1"), q("Z", "B2"))
    assertEquals(graph(twoSheets).withinComponentOrder(members), twoSheetSweep)
    assertEquals(oracle(twoSheets, members), twoSheetSweep)
    // the same book with its sheets the other way round: the order does not move
    val swapped = Workbook(twoSheets.sheets.reverse)
    assertEquals(graph(swapped).withinComponentOrder(members), twoSheetSweep)
  }

  test("two sheets with unknown members and a reader outside the cycle") {
    // b!B5 is a constant and b!D4 a blank: ready from the start, in grid order (row 4, row 5).
    // Z!A1 reads the cycle's Z!B2 and is not part of the cycle. Once the ready members are out,
    // the frontier is dry and the cut falls on the top-left member that remains, which is Z!A1
    // itself: the cut is positional and does not ask whether a member is on a cycle, so Z!A1
    // sweeps before the member it reads. Then the cycle as above.
    val members =
      Vector(q("b", "A1"), q("b", "B5"), q("b", "C1"), q("b", "D4"), q("Z", "A1"), q("Z", "B2"))
    val expected =
      Vector(q("b", "D4"), q("b", "B5"), q("Z", "A1"), q("Z", "B2"), q("b", "C1"), q("b", "A1"))
    assertEquals(oracle(twoSheets, members), expected)
    assertEquals(graph(twoSheets).withinComponentOrder(members), expected)
  }

  property("any member vector, known or not, repeated or not: the Map form over the cell graph") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val (deps, _) = NaiveGraphs.fromWorkbookFormulaGraph(wb)
      val recalc = graph(wb)
      val formulas = deps.keySet.toVector.sortBy(QualifiedRef.gridKey)
      // off the generated grid, on a sheet of the book and on a sheet it lacks
      val strangers = Vector(q("Mid", "Z99"), q("Ghost", "A1"), q("b", "Y50"))
      val components = DependencyGraph.qualifiedSccOrder(deps).map(_.members)
      val known: Vector[Vector[QualifiedRef]] =
        Vector(formulas, formulas.zipWithIndex.collect { case (m, i) if i % 2 == 0 => m }) ++
          components
      val vectors =
        known ++ known.map(_.reverse) ++ known.map(m => m ++ m.take(2)) ++
          known.map(_ ++ strangers) ++ known.map(m => strangers.take(1) ++ m ++ m.take(1))
      Prop.all(vectors.map { members =>
        val expected = DependencyGraph.withinComponentOrder(members, deps)
        val actual = recalc.withinComponentOrder(members)
        Prop(actual == expected) :| s"$members: $actual vs $expected"
      }*)
    }
  }

  test("the generator behind the property reaches cut cycles, on one sheet and across sheets") {
    // fixed seeds: what the property above can see, counted so it cannot go vacuous unnoticed
    var cyclic = 0
    var cutMoved = 0
    var acrossSheets = 0
    (0 until 60).foreach { i =>
      CompactGraphGens.genWorkbook(Gen.Parameters.default, Seed(i.toLong)).foreach { wb =>
        val recalc = graph(wb)
        recalc.components.filter(c => c.cyclic && c.members.size > 1).foreach { component =>
          val order = recalc.withinComponentOrder(component.members)
          assertEquals(order, oracle(wb, component.members))
          cyclic += 1
          if order != component.members then cutMoved += 1
          if component.members.map(_.sheet).distinct.size > 1 then acrossSheets += 1
        }
      }
    }
    def clue = s"cyclic=$cyclic cutMoved=$cutMoved acrossSheets=$acrossSheets"
    assert(cyclic >= 20, clue)
    assert(cutMoved > 0, clue)
    assert(acrossSheets > 0, clue)
  }
