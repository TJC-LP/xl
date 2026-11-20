package com.tjclp.xl.formula

import scala.collection.immutable.BitSet

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.RecalcCounters
import com.tjclp.xl.formula.eval.RecalcCounters.Counter
import com.tjclp.xl.formula.graph.{CompactGraph, DependencyGraph, QualifiedGraph}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * Regression gates for the graphs whole-workbook recalculation (the compact graph) and inspection
 * (`QualifiedGraph`, the symbolic dependency index) build.
 *
 * The performance gates pin the build's work counters and graph sizes, never its wall time: a
 * shared CI runner that is running three other test JVMs stretches a 2 s build past 5 s (#843),
 * while a counter is an exact function of the workbook and the algorithm. Wall time appears only in
 * the assertion clues.
 *
 * #891: the per-cell formula graph these gates were first written against is a test oracle now
 * (`NaiveGraphs`); its own range-lookup counters left with it, and the compact graph's equivalence
 * to it is CompactGraphSpec's.
 */
class FormulaGraphPerfSpec extends FunSuite:

  private def number(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def formula(text: String): CellValue = CellValue.Formula(text, None)

  /** `rows` rows of a running total: B holds the driver, C `=SUM($B$1:INDEX(B:B,ROW()))`. */
  private def runningTotals(rows: Int): Workbook =
    Workbook((1 to rows).foldLeft(Sheet(SheetName.unsafe("S"))) { (current, row) =>
      current
        .put(ARef.from0(1, row - 1), number(row))
        .put(ARef.from0(2, row - 1), formula("=SUM($B$1:INDEX(B:B,ROW()))"))
    })

  /**
   * The recalculation graph's direct formula-to-formula reads: every formula (a key even when it
   * reads none) and the other formulas it reads. Self-reads are left out (`memberPrecedents` drops
   * them; a cell reading itself is a cycle, CompactGraphSpec's).
   */
  private def formulaReads(
    graph: CompactGraph
  ): Map[DependencyGraph.QualifiedRef, Set[DependencyGraph.QualifiedRef]] =
    val reads = graph.memberPrecedents(BitSet.fromSpecific(0 until graph.formulaCount))
    (0 until graph.formulaCount).iterator.map { n =>
      graph.nodes.qualified(n) -> reads.getOrElse(n, Set.empty).map(graph.nodes.qualified)
    }.toMap

  /**
   * The inspection graph's reads restricted to the other formula cells: the topology recalc needs.
   */
  private def inspectionFormulaReads(
    wb: Workbook
  ): Map[DependencyGraph.QualifiedRef, Set[DependencyGraph.QualifiedRef]] =
    val dependencies = QualifiedGraph.of(wb).dependencies
    dependencies.map((formula, reads) => formula -> (reads & dependencies.keySet - formula))

  /** The compact graph of `wb` with its build counters; the wall time is diagnostic only. */
  private def measured(wb: Workbook): (CompactGraph, RecalcCounters.Snapshot, Long) =
    val counters = RecalcCounters.create()
    val started = System.nanoTime()
    val graph = CompactGraph.build(wb, counters)
    val elapsedMs = (System.nanoTime() - started) / 1000000L
    (graph, counters.snapshot, elapsedMs)

  test("the recalculation graph's formula reads equal the inspection graph's, on formula nodes"):
    val data = Sheet(SheetName.unsafe("Data"))
      .put(ref"A1", number(1))
      .put(ref"A2", number(2))
      .put(ref"A3", formula("=A1+A2"))
      .put(ref"B3", formula("=SUM(A1:A3)"))
    val summary = Sheet(SheetName.unsafe("Summary"))
      .put(ref"A1", formula("=SUM(Data!A:A)+Data!B3"))
    val wb = Workbook(data, summary)

    assertEquals(formulaReads(CompactGraph.build(wb)), inspectionFormulaReads(wb))

  test("constant driver ranges contribute no topology edges and are one range node (GH-522)"):
    val driverRows = 5000
    val formulaCount = 1000
    val data = (1 to driverRows).foldLeft(Sheet(SheetName.unsafe("Data"))) { (sheet, row) =>
      sheet.put(ARef.from0(0, row - 1), number(row))
    }
    val model = (1 to formulaCount).foldLeft(Sheet(SheetName.unsafe("Model"))) { (sheet, row) =>
      sheet.put(ARef.from0(1, row - 1), formula("=SUM(Data!A:A)"))
    }
    val wb = Workbook(data, model)

    val (graph, snap, elapsedMs) = measured(wb)
    val clue = s"${snap.render} (${elapsedMs} ms)"

    val reads = formulaReads(graph)
    assertEquals(reads.size, formulaCount)
    assert(reads.valuesIterator.forall(_.isEmpty))
    // one range node shared by the 1,000 readers (GH-522: it was one expansion per reader, ~5M
    // refs); Data holds no formula, so no segment node
    assertEquals(graph.rangeCount, 1, clue)
    assertEquals(snap(Counter.RangeNodes), 1L, clue)
    assertEquals(snap(Counter.VirtualNodes), 0L, clue)

  test("range reads find formula cells across columns, rows and whole-row/column ranges"):
    // formulas in A..E rows 1..20 (every other cell), read through ranges of each shape
    val grid = (for col <- 0 until 5; row <- 0 until 20 if (col + row) % 2 == 0 yield (col, row))
      .foldLeft(Sheet(SheetName.unsafe("S"))) { case (sheet, (col, row)) =>
        sheet.put(ARef.from0(col, row), formula(s"=${col + row}"))
      }
    val readers = Vector(
      "=SUM(B3:D9)",
      "=SUM(2:4)",
      "=SUM(B:C)",
      "=SUM(A20:E20)",
      "=SUM(E1:E1)",
      "=SUM(F1:H30)",
      "=SUM(A21:E40)"
    )
    val sheet = readers.zipWithIndex.foldLeft(grid) { case (current, (text, i)) =>
      current.put(ARef.from0(10, i), formula(text))
    }
    val wb = Workbook(sheet)
    val actual = formulaReads(CompactGraph.build(wb))
    assertEquals(actual, inspectionFormulaReads(wb))
    // the reader of B3:D9 depends on exactly the grid formulas inside it
    val inBox =
      for col <- 1 to 3; row <- 2 to 8 if (col + row) % 2 == 0
      yield DependencyGraph.QualifiedRef(sheet.name, ARef.from0(col, row))
    assertEquals(actual(DependencyGraph.QualifiedRef(sheet.name, ARef.from0(10, 0))), inBox.toSet)

  test("#743: an INDEX(…,ROW()) running total builds a recalculation graph linear in its rows"):
    // each row reads its own B1:B{r} (and INDEX's own B{r}), a distinct range per row, all over
    // constants: no formula reads another, and the graph holds a bounded number of range nodes per
    // row. Two sizes: the graph must grow with the rows, never with their square.
    def size(rows: Int): (Long, Long) =
      val (graph, snap, elapsedMs) = measured(runningTotals(rows))
      val clue = s"$rows rows: ${snap.render} (${elapsedMs} ms)"
      val reads = formulaReads(graph)
      assertEquals(reads.size, rows)
      assert(reads.valuesIterator.forall(_.isEmpty), clue)
      assertEquals(snap(Counter.VirtualNodes), 0L, clue)
      assert(snap(Counter.RangeNodes) >= rows.toLong, clue)
      assert(snap(Counter.RangeNodes) <= 2L * rows + 2L, clue)
      (snap(Counter.RangeNodes), snap(Counter.GraphEdges))
    val (rangesSmall, edgesSmall) = size(2000)
    val (rangesLarge, edgesLarge) = size(8000)
    // four times the rows: at most about four times the graph (a quadratic one would be sixteen)
    assert(rangesLarge <= 4.5 * rangesSmall, s"range nodes $rangesSmall -> $rangesLarge")
    assert(edgesLarge <= 4.5 * edgesSmall, s"edges $edgesSmall -> $edgesLarge")

  test("formula cells inside ranges remain topology precedents"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", number(1))
      .put(ref"A2", formula("=A1+1"))
      .put(ref"A3", formula("=SUM(A1:A2)"))
    val dependencies = formulaReads(CompactGraph.build(Workbook(sheet)))
    val qA2 = DependencyGraph.QualifiedRef(sheet.name, ref"A2")
    val qA3 = DependencyGraph.QualifiedRef(sheet.name, ref"A3")

    assertEquals(dependencies(qA2), Set.empty)
    assertEquals(dependencies(qA3), Set(qA2))

  test("the inspection graph reuses one expansion across anchor spellings of the same range"):
    val sheet = (1 to 1000)
      .foldLeft(Sheet(SheetName.unsafe("S"))) { (current, row) =>
        current.put(ARef.from0(0, row - 1), number(row))
      }
      .put(ref"B1", formula("=SUM(A1:A1000)"))
      .put(ref"B2", formula("=SUM($A$1:$A$1000)"))
    val wb = Workbook(sheet)

    val graph = QualifiedGraph.of(wb)
    val relative = graph.precedentsOf(DependencyGraph.QualifiedRef(sheet.name, ref"B1"))
    val anchored = graph.precedentsOf(DependencyGraph.QualifiedRef(sheet.name, ref"B2"))

    assertEquals(relative.size, 1000)
    assert(anchored eq relative, "anchor-only spelling differences must hit the same range cache")
    assertEquals(graph.rangeCells.size, 1)
    // and the recalculation graph holds the range once
    assertEquals(CompactGraph.build(wb).rangeCount, 1)

  test("symbolic range index survives clearing the old used-range boundary"):
    val sheetName = SheetName.unsafe("S")
    val sum = DependencyGraph.QualifiedRef(sheetName, ref"B1")
    val cleared = DependencyGraph.QualifiedRef(sheetName, ref"A100")
    // This is the POST-edit sheet: A100 has disappeared, so usedRange no longer intersects A:A.
    val sheet = Sheet(sheetName)
      .put(ref"B1", CellValue.Formula("=SUM(A:A)", Some(number(5))))
    val index = DependencyGraph.fromWorkbookDependencyIndex(Workbook(sheet))

    assertEquals(index.transitiveDependents(Set(cleared)), Set(sum))

  test("symbolic index stores a shared range once instead of per covered cell"):
    val formulaCount = 1000
    val dataName = SheetName.unsafe("Data")
    val modelName = SheetName.unsafe("Model")
    val data = Sheet(dataName).put(ref"A1", number(1)).put(ref"A1000", number(1000))
    val model = (1 to formulaCount).foldLeft(Sheet(modelName)) { (sheet, row) =>
      sheet.put(ARef.from0(1, row - 1), formula("=SUM(Data!A1:A1000)"))
    }

    val index = DependencyGraph.fromWorkbookDependencyIndex(Workbook(data, model))
    val entries = index.rangeDependents.getOrElse(dataName, Vector.empty)

    assertEquals(entries.size, 1)
    assertEquals(entries(0).range, CellRange(ref"A1", ref"A1000"))
    assertEquals(entries(0).dependents.size, formulaCount)
    assertEquals(
      index.transitiveDependents(Set(DependencyGraph.QualifiedRef(dataName, ref"A500"))).size,
      formulaCount
    )
