package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.StructuralEditor
import com.tjclp.xl.formula.eval.DependentRecalculation.*
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.printer.{FormulaPrinter, FormulaShifter}
import com.tjclp.xl.tables.{TableColumn, TableSpec}

import munit.FunSuite

class TableColumnReferenceSpec extends FunSuite:
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private val table = TableSpec(
    "Internal",
    "Table1",
    ref"B2:C6",
    Vector(TableColumn(1, "Col"), TableColumn(2, "Other Column")),
    showTotalsRow = true
  )
  private val data = Sheet("Data")
    .withTable(table)
    .put(
      ref"B2" -> 900,
      ref"C2" -> 900,
      ref"B3" -> 1,
      ref"B5" -> 3,
      ref"C3" -> 10,
      ref"C4" -> 20,
      ref"C5" -> 30,
      ref"B6" -> 1000,
      ref"C6" -> 1000
    )
  private def cached(wb: Workbook, sheet: String, ref: ARef): Option[CellValue] =
    wb(SheetName.unsafe(sheet)).toOption.flatMap { s =>
      s(ref).value match
        case CellValue.Formula(_, cache, _) => cache
        case value => Some(value)
    }

  test("GH-669: qualified table column is accepted by the formula gate") {
    val parsed = FormulaParser.parse("=SUM(Table1[Col])")
    assert(parsed.isRight, parsed.toString)
    assertEquals(data.evaluateFormula("=SUM(Table1[Col])"), Right(num(4)))
  }

  for formula -> expected <- List(
      "SUM(Table1[Col])" -> 4,
      "COUNT(Table1[Col])" -> 2,
      "COUNTBLANK(Table1[Col])" -> 1,
      "COUNTIF(Table1[Col],\">1\")" -> 1,
      "SUMIFS(Table1[Other Column],Table1[Col],\">1\")" -> 30,
      "SUMIF(Table1[Col],\">0\",Table1[Other Column])" -> 40,
      "ROWS(Table1[Col])" -> 3,
      "COLUMNS(Table1[Col])" -> 1,
      "ROW(INDEX(Table1[Col],1))" -> 3,
      "SUM(INDEX(Table1[Col],0))" -> 4,
      "SUM(OFFSET(Table1[Col],0,1))" -> 60,
      "LET(r,Table1[Col],COUNTIF(r,\">1\"))" -> 1,
      "SUM(IF(TRUE,Table1[Col],C3:C5))" -> 4,
      "SUM((Table1[Col],C3:C5))" -> 64,
      "SUM((Table1[Col] B4:B5))" -> 3
    )
  do
    test(s"GH-669: reference consumer $formula") {
      assertEquals(data.evaluateFormula(formula), Right(num(expected)))
    }

  test("GH-669: lookup uses formula displayName and folds table/column case") {
    assertEquals(data.evaluateFormula("SUM(tAbLe1[cOL])"), Right(num(4)))
    assertEquals(data.evaluateFormula("SUM(Internal[Col])"), Right(CellValue.Error(CellError.Ref)))
  }

  test("GH-669: workbook lookup finds the table on another sheet") {
    val wb = Workbook(Sheet("Reader"), data)
    assertEquals(wb.evaluateFormula("SUM(Table1[Other Column])", "Reader"), Right(num(60)))
    assertEquals(wb.evaluateFormula("COUNTIF(Table1[Col],\">1\")", "Reader"), Right(num(1)))
  }

  test("GH-669: direct array has the declared height and blank positions") {
    val result = data.evaluateArrayFormula("Table1[Col]", ref"F1").toOption.get
    assertEquals(result._2, ref"F1:F3")
    assertEquals(result._1(ref"F1").value, num(1))
    assertEquals(result._1(ref"F2").value, CellValue.Empty)
    assertEquals(result._1(ref"F3").value, num(3))
  }

  test("GH-669: plain cell implicit intersection follows table geometry") {
    assertEquals(data.evaluateFormula("Table1[Col]", currentCell = Some(ref"F5")), Right(num(3)))
    assertEquals(
      data.evaluateFormula("Table1[Col]", currentCell = Some(ref"F8")),
      Right(CellValue.Error(CellError.Value))
    )
  }

  test("GH-669: referenced errors propagate through aggregate and range-only slots") {
    val errors = data.put(ref"B4", CellValue.Error(CellError.Div0))
    assertEquals(errors.evaluateFormula("SUM(Table1[Col])"), Right(CellValue.Error(CellError.Div0)))
    assertEquals(
      errors.evaluateFormula("INDEX(Table1[Col],2)"),
      Right(CellValue.Error(CellError.Div0))
    )
  }

  for formula <- List("SUM(Missing[Col])", "COUNTIF(Table1[Missing],1)") do
    test(s"GH-669: unresolved metadata is a reference error: $formula") {
      assertEquals(data.evaluateFormula(formula), Right(CellValue.Error(CellError.Ref)))
    }

  test("GH-669: folded duplicate tables refuse an exact-spelling preference") {
    val duplicate = Sheet("Other").withTable(table.copy(name = "Other", displayName = "TABLE1"))
    val wb = Workbook(data, duplicate)
    assertEquals(
      wb.evaluateFormula("SUM(Table1[Col])", "Data"),
      Right(CellValue.Error(CellError.Ref))
    )
  }

  for invalid <- List(
      table.copy(columns = Vector(TableColumn(1, "Col"), TableColumn(2, "COL"))),
      table.copy(columns = Vector(TableColumn(1, "Col"))),
      table.copy(range = ref"B2:C2", showTotalsRow = false),
      table.copy(range = ref"B2:C3"),
      table.copy(range = ref"B1048576:C1048576", showTotalsRow = false),
      table.copy(range = ref"B1:C1", showHeaderRow = false)
    )
  do
    test(s"GH-669: invalid/empty metadata never normalizes into data: $invalid") {
      assertEquals(
        data.withTable(invalid).evaluateFormula("SUM(Table1[Col])"),
        Right(CellValue.Error(CellError.Ref))
      )
    }

  test("GH-669: hidden header and totals flags define data boundaries") {
    val noHeader = table.copy(showHeaderRow = false, showTotalsRow = false)
    assertEquals(data.withTable(noHeader).evaluateFormula("ROWS(Table1[Col])"), Right(num(5)))
    assertEquals(data.withTable(noHeader).evaluateFormula("SUM(Table1[Col])"), Right(num(1904)))
  }

  test("GH-669: identifiers round-trip and survive copy, structural shift and sheet rename") {
    val text = "COUNTIF(tAbLe1[Other Column],\">0\")"
    val parsed = FormulaParser.parse(text).toOption.get
    assertEquals(FormulaPrinter.printFileForm(parsed), text)
    assertEquals(FormulaParser.parse(FormulaPrinter.print(parsed)), Right(parsed))
    assertEquals(FormulaPrinter.printFileForm(FormulaShifter.shift(parsed, 3, 4)), text)
    assertEquals(
      FormulaPrinter.printFileForm(
        FormulaShifter.shiftStructural(parsed, true, "Data", true, 0, 1)
      ),
      text
    )
    assertEquals(
      FormulaPrinter.printFileForm(
        FormulaShifter.renameSheet(parsed, SheetName.unsafe("Data"), SheetName.unsafe("New"))
      ),
      text
    )
    assert(DependencyGraph.containsCellReferences(parsed))
  }

  test("GH-669: sheet graph orders table formulas and detects a table self-cycle") {
    val sheet = data
      .put(ref"B3", CellValue.Formula("C3+1", Some(num(99))))
      .put(ref"A1", CellValue.Formula("SUM(Table1[Col])", Some(num(-1))))
    assertEquals(
      DependencyGraph.fromSheet(sheet).dependencies(ref"A1"),
      Set(ref"B3", ref"B4", ref"B5")
    )
    assertEquals(sheet.evaluateWithDependencyCheck().toOption.get(ref"A1"), num(14))
    val cycle = data.put(ref"B3", CellValue.Formula("SUM(Table1[Col])"))
    assert(DependencyGraph.detectCycles(DependencyGraph.fromSheet(cycle)).isLeft)
  }

  test("GH-669: workbook graph and dirty recalc refresh table precedents first") {
    val source = data.put(ref"B3", CellValue.Formula("C3+1", Some(num(99))))
    val reader = Sheet("Reader")
      .put(ref"A1", CellValue.Formula("SUM(Table1[Col])", Some(num(-1))))
      .put(ref"A2", CellValue.Formula("A1*2", Some(num(-2))))
    val wb = Workbook(reader, source)
    val result = wb.recalculate()
    assertEquals(result.errors, Vector.empty)
    assertEquals(cached(result.workbook, "Reader", ref"A1"), Some(num(14)))
    val updated =
      result.workbook.put(result.workbook(SheetName.unsafe("Data")).toOption.get.put(ref"C3", 20))
    val dirty = updated.recalculateDependents(SheetName.unsafe("Data"), Set(ref"C3"))
    assertEquals(cached(dirty, "Reader", ref"A1"), Some(num(24)))
    assertEquals(cached(dirty, "Reader", ref"A2"), Some(num(48)))
  }

  test("GH-669: resizing and removing table metadata is observed by explicit recalculation") {
    val reader = Sheet("Reader")
      .put(ref"A1", CellValue.Formula("SUM(Table1[Col])", Some(num(999))))
      .put(ref"A2", CellValue.Formula("A1*2", Some(num(999))))
    val wb = Workbook(reader, data).recalculate().workbook
    val smaller = wb.put(data.withTable(table.copy(range = ref"B2:C5"))).recalculate().workbook
    assertEquals(cached(smaller, "Reader", ref"A1"), Some(num(1)))
    assertEquals(cached(smaller, "Reader", ref"A2"), Some(num(2)))
    val removed = smaller.put(data.removeTable("Internal")).recalculate().workbook
    assertEquals(cached(removed, "Reader", ref"A1"), Some(CellValue.Error(CellError.Ref)))
    assertEquals(cached(removed, "Reader", ref"A2"), Some(CellValue.Error(CellError.Ref)))
  }

  test("GH-669: structural edits preserve table text, invalidate caches and follow moved rows") {
    val reader = Sheet("Reader").put(ref"A1", CellValue.Formula("SUM(Table1[Col])", Some(num(999))))
    val wb = Workbook(reader, data)
    val shifted = StructuralEditor.insertRows(wb, data.name, 3, 1)
    assertEquals(cached(shifted, "Reader", ref"A1"), None)
    assertEquals(cached(shifted.recalculate().workbook, "Reader", ref"A1"), Some(num(4)))
  }

  for text <- List(
      "[@Col]",
      "[Col]",
      "Table1[@Col]",
      "Table1[#All]",
      "Table1[#Headers]",
      "Table1[#Totals]",
      "Table1[[Col]]",
      "Table1[[Col]:[Other]]",
      "Table1[Col:Other]",
      "Table1['#Col]",
      "Table1[]",
      "SUM(Sheet1:Sheet3!A1)",
      "LAMBDA(x,x+1)(2)"
    )
  do
    test(s"GH-669: excluded grammar remains refused: $text") {
      assert(FormulaParser.parse(text).isLeft, text)
    }

  for position <- List(ref"F5", ref"F8") do
    test(s"GH-669: lifted table references match explicit data ranges at ${position.toA1}") {
      val plain = data.evaluateFormula("SUM(ABS(B3:B5))", currentCell = Some(position))
      assertEquals(
        data.evaluateFormula("SUM(ABS(Table1[Col]))", currentCell = Some(position)),
        plain
      )
    }

  test("GH-669: range-refusing lifted slots distinguish references from values") {
    assertEquals(
      data.evaluateFormula("EOMONTH(Table1[Col],0)", currentCell = Some(ref"F5")),
      data.evaluateFormula("EOMONTH(B3:B5,0)", currentCell = Some(ref"F5"))
    )
  }

  for name <- List("NOT", "TRUE", "FALSE", "SUM") do
    test(s"GH-669: a table name $name precedes keyword/function interpretation") {
      val sheet = data.withTable(table.copy(displayName = name))
      assertEquals(sheet.evaluateFormula(s"SUM($name[Col])"), Right(num(4)))
    }
