package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.RecalcCounters
import com.tjclp.xl.formula.eval.RecalcCounters.Counter
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * Regression gates for the formula-only graph used by whole-workbook recalculation.
 *
 * The performance gates pin the build's work counters (`RangeLookups`, `RangeLookupWork`,
 * `GraphEdges`), never its wall time: a shared CI runner that is running three other test JVMs
 * stretches a 2 s build past 5 s (#843), while a counter is an exact function of the workbook and
 * the algorithm. Wall time appears only in the assertion clues.
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

  /** The formula graph of `wb` with its build counters; the wall time is diagnostic only. */
  private def measured(wb: Workbook): (
    Map[DependencyGraph.QualifiedRef, Set[DependencyGraph.QualifiedRef]],
    Map[DependencyGraph.QualifiedRef, Set[DependencyGraph.QualifiedRef]],
    RecalcCounters.Snapshot,
    Long
  ) =
    val counters = RecalcCounters.create()
    val started = System.nanoTime()
    val (dependencies, dependents) = DependencyGraph.fromWorkbookFormulaGraph(wb, counters)
    val elapsedMs = (System.nanoTime() - started) / 1000000L
    (dependencies, dependents, counters.snapshot, elapsedMs)

  test("formula-only graph equals the bounded graph restricted to formula nodes"):
    val data = Sheet(SheetName.unsafe("Data"))
      .put(ref"A1", number(1))
      .put(ref"A2", number(2))
      .put(ref"A3", formula("=A1+A2"))
      .put(ref"B3", formula("=SUM(A1:A3)"))
    val summary = Sheet(SheetName.unsafe("Summary"))
      .put(ref"A1", formula("=SUM(Data!A:A)+Data!B3"))
    val wb = Workbook(data, summary)

    val (bounded, _) = DependencyGraph.fromWorkbookBounded(wb)
    val formulaNodes = bounded.keySet
    val expected = bounded.view.mapValues(_ & formulaNodes).toMap
    val (actual, actualReverse) = DependencyGraph.fromWorkbookFormulaGraph(wb)

    assertEquals(actual, expected)
    assertEquals(actualReverse, DependencyGraph.reverseEdges(expected))

  test("constant driver ranges contribute no topology edges and expand once (GH-522)"):
    val driverRows = 5000
    val formulaCount = 1000
    val data = (1 to driverRows).foldLeft(Sheet(SheetName.unsafe("Data"))) { (sheet, row) =>
      sheet.put(ARef.from0(0, row - 1), number(row))
    }
    val model = (1 to formulaCount).foldLeft(Sheet(SheetName.unsafe("Model"))) { (sheet, row) =>
      sheet.put(ARef.from0(1, row - 1), formula("=SUM(Data!A:A)"))
    }
    val wb = Workbook(data, model)

    val (dependencies, dependents, snap, elapsedMs) = measured(wb)
    val clue = s"${snap.render} (${elapsedMs} ms)"

    assertEquals(dependencies.size, formulaCount)
    assert(dependencies.valuesIterator.forall(_.isEmpty))
    assertEquals(dependents, Map.empty)
    // one expansion shared by the 1,000 readers (GH-522: it was one per reader, ~5M refs); Data
    // holds no formula, so the lookup examines no index entry
    assertEquals(snap(Counter.RangeLookups), 1L, clue)
    assertEquals(snap(Counter.RangeLookupWork), 0L, clue)

  test("range lookup finds formula cells across columns, rows and whole-row/column ranges"):
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
    val (bounded, _) = DependencyGraph.fromWorkbookBounded(wb)
    val expected = bounded.view.mapValues(_ & bounded.keySet).toMap
    val (actual, _) = DependencyGraph.fromWorkbookFormulaGraph(wb)
    assertEquals(actual, expected)
    // the reader of B3:D9 depends on exactly the grid formulas inside it
    val inBox =
      for col <- 1 to 3; row <- 2 to 8 if (col + row) % 2 == 0
      yield DependencyGraph.QualifiedRef(sheet.name, ARef.from0(col, row))
    assertEquals(actual(DependencyGraph.QualifiedRef(sheet.name, ARef.from0(10, 0))), inBox.toSet)

  test("#743: an INDEX(…,ROW()) running total builds the formula graph in linear work"):
    // each row reads its own B1:B{r} (and INDEX's own B{r}), a distinct range per row; a scan of
    // every formula per range was quadratic: 2 x rows^2 entries, about 1.3e8 at 8,000 rows, where
    // the column index examines one entry per lookup (the one column holding formulas, C, which
    // lies past B). Two sizes: the work must grow with the rows, never with their square.
    def work(rows: Int): (Long, Long) =
      val (dependencies, _, snap, elapsedMs) = measured(runningTotals(rows))
      val clue = s"$rows rows: ${snap.render} (${elapsedMs} ms)"
      assertEquals(dependencies.size, rows)
      assert(dependencies.valuesIterator.forall(_.isEmpty), clue)
      assertEquals(snap(Counter.GraphEdges), 0L, clue)
      assert(snap(Counter.RangeLookups) >= rows.toLong, clue)
      assert(snap(Counter.RangeLookups) <= 2L * rows + 2L, clue)
      assert(snap(Counter.RangeLookupWork) <= 2L * rows + 2L, clue)
      (snap(Counter.RangeLookups), snap(Counter.RangeLookupWork))
    val (lookupsSmall, workSmall) = work(2000)
    val (lookupsLarge, workLarge) = work(8000)
    // four times the rows: at most about four times the work (a scan per lookup would be sixteen)
    assert(lookupsLarge <= 4.5 * lookupsSmall, s"lookups $lookupsSmall -> $lookupsLarge")
    assert(workLarge <= 4.5 * workSmall, s"index work $workSmall -> $workLarge")

  test("the lookup work counter grows with the formula cells a range covers"):
    // D{n} = SUM($C$1:C{n}) over formulas in C: every covered cell is an entry examined and an
    // edge, so the work is quadratic exactly as the per-cell graph's edges are (RecalcPerfGatesSpec
    // pins rows*(rows+1)/2 edges for the legacy engine on this shape). The #743 shape above covers
    // no formula and stays flat: a regression to a scan could not hide behind an empty result.
    val rows = 200
    val sheet = (1 to rows).foldLeft(Sheet(SheetName.unsafe("S"))) { (current, row) =>
      current
        .put(ARef.from0(2, row - 1), formula(s"=$row*2"))
        .put(ARef.from0(3, row - 1), formula(s"=SUM($$C$$1:C$row)"))
    }
    val (_, _, snap, elapsedMs) = measured(Workbook(sheet))
    val clue = s"${snap.render} (${elapsedMs} ms)"
    val edges = rows.toLong * (rows + 1) / 2
    assertEquals(snap(Counter.GraphEdges), edges, clue)
    assertEquals(snap(Counter.RangeLookups), rows.toLong, clue)
    // every covered formula cell, plus the column entries and the one row entry past each range
    assert(snap(Counter.RangeLookupWork) >= edges, clue)
    assert(snap(Counter.RangeLookupWork) <= edges + 4L * rows, clue)

  test("formula cells inside ranges remain topology precedents"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", number(1))
      .put(ref"A2", formula("=A1+1"))
      .put(ref"A3", formula("=SUM(A1:A2)"))
    val (dependencies, _) = DependencyGraph.fromWorkbookFormulaGraph(Workbook(sheet))
    val qA2 = DependencyGraph.QualifiedRef(sheet.name, ref"A2")
    val qA3 = DependencyGraph.QualifiedRef(sheet.name, ref"A3")

    assertEquals(dependencies(qA2), Set.empty)
    assertEquals(dependencies(qA3), Set(qA2))

  test("sheet graph reuses one expansion across anchor spellings of the same range"):
    val sheet = (1 to 1000)
      .foldLeft(Sheet(SheetName.unsafe("S"))) { (current, row) =>
        current.put(ARef.from0(0, row - 1), number(row))
      }
      .put(ref"B1", formula("=SUM(A1:A1000)"))
      .put(ref"B2", formula("=SUM($A$1:$A$1000)"))

    val graph = DependencyGraph.fromSheet(sheet)
    val relative = graph.dependencies(ref"B1")
    val anchored = graph.dependencies(ref"B2")

    assertEquals(relative.size, 1000)
    assert(anchored eq relative, "anchor-only spelling differences must hit the same range cache")

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
