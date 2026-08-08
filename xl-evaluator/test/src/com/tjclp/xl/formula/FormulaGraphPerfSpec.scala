package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/** Regression gates for the formula-only graph used by whole-workbook recalculation. */
class FormulaGraphPerfSpec extends FunSuite:

  private def number(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def formula(text: String): CellValue = CellValue.Formula(text, None)

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

  test("constant driver ranges contribute no topology edges"):
    val driverRows = 5000
    val formulaCount = 1000
    val data = (1 to driverRows).foldLeft(Sheet(SheetName.unsafe("Data"))) { (sheet, row) =>
      sheet.put(ARef.from0(0, row - 1), number(row))
    }
    val model = (1 to formulaCount).foldLeft(Sheet(SheetName.unsafe("Model"))) { (sheet, row) =>
      sheet.put(ARef.from0(1, row - 1), formula("=SUM(Data!A:A)"))
    }
    val wb = Workbook(data, model)

    val started = System.nanoTime()
    val (dependencies, dependents) = DependencyGraph.fromWorkbookFormulaGraph(wb)
    val elapsedMs = (System.nanoTime() - started) / 1000000L

    assertEquals(dependencies.size, formulaCount)
    assert(dependencies.valuesIterator.forall(_.isEmpty))
    assertEquals(dependents, Map.empty)
    assert(
      elapsedMs < 5000L,
      s"formula-only graph took ${elapsedMs}ms for $formulaCount shared driver ranges"
    )

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
