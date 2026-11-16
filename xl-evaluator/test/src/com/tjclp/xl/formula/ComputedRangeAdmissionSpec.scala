package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.FunSuite

class ComputedRangeAdmissionSpec extends FunSuite:
  private val data = Sheet("Data").put(ref"A1" -> 1, ref"A2" -> 3, ref"A3" -> 5)

  test("GH-723: range-only slots reject known value INDEX results at parse time") {
    List(
      "MATCH(1,INDEX((A:A=x)*(B:B=y),0),0)",
      "LARGE(INDEX(A1:A3*2,0),1)",
      "COUNTIF(INDEX({1},1,1),1)",
      "COUNTIF(INDEX(5,1),1)",
      "COUNTIF(INDEX(INDEX({1},1),1),1)",
      "COUNTIF(INDEX({1},1):A3,1)",
      "COUNTIF(INDEX((A1:A3,INDEX({1},1)),0),1)",
      "COUNTIF(INDEX((A1:A3 INDEX({1},1)),0),1)",
      "COUNTIF(XLOOKUP(1,{1},{2}),1)"
    ).foreach(text => assert(FormulaParser.parse(text).isLeft, text))
  }

  test(
    "GH-723: unsupported value-array formula reports an error without caching a fabricated value"
  ) {
    val value =
      CellValue.Formula("LARGE(INDEX(A1:A3*2,0),1)", Some(CellValue.Number(BigDecimal(10))))
    val result = Workbook(data.put(ref"D1", value)).recalculate()
    assert(result.errors.nonEmpty)
    val observed = result.workbook.sheets.find(_.name == data.name).getOrElse(fail("missing Data"))
    assertEquals(observed(ref"D1").value, CellValue.Formula("LARGE(INDEX(A1:A3*2,0),1)", None))
  }

  test("GH-723: reference INDEX and nested reference producers remain accepted") {
    List(
      """COUNTIF(INDEX(A1:A3,0),">2")""",
      """COUNTIF(INDEX(OFFSET(A1,0,0,3,1),0),">2")""",
      """COUNTIF(INDEX(INDEX(A1:A3,0),0),">2")""",
      """COUNTIF(A1:INDEX(A1:A3,3),">2")"""
    ).foreach(text =>
      assertEquals(data.evaluateFormula(text), Right(CellValue.Number(BigDecimal(2))), text)
    )
  }

  test("GH-723: value INDEX remains supported outside reference-only slots") {
    assertEquals(data.evaluateFormula("INDEX({7},1,1)"), Right(CellValue.Number(BigDecimal(7))))
    assertEquals(data.evaluateFormula("INDEX(A1:A3*2,2)"), Right(CellValue.Number(BigDecimal(6))))
  }

  test("GH-723: LET value INDEX reports unsupported while reference LET still works") {
    val text = "LET(k,{1},LARGE(INDEX(k,0),1))"
    assert(data.evaluateFormula(text).isLeft)
    val result =
      Workbook(data.put(ref"D1", CellValue.Formula(text, Some(CellValue.Number(1))))).recalculate()
    assert(result.errors.nonEmpty)
    val observed = result.workbook.sheets.find(_.name == data.name).getOrElse(fail("missing Data"))
    assertEquals(observed(ref"D1").value, CellValue.Formula(text, None))
    assertEquals(
      data.evaluateFormula("LET(k,A1:A3,LARGE(INDEX(k,0),1))"),
      Right(CellValue.Number(5))
    )
  }

  test("GH-723: defined value names do not become cached errors; reference names resolve") {
    List("{1}", "5").foreach(definition =>
      val text = "LARGE(INDEX(Values,0),1)"
      val wb = Workbook(data.put(ref"D1", CellValue.Formula(text, Some(CellValue.Number(5)))))
      val named = wb.copy(metadata =
        wb.metadata.copy(definedNames = Vector(DefinedName("Values", definition)))
      )
      val result = named.recalculate()
      assert(result.errors.nonEmpty, definition)
      val observed =
        result.workbook.sheets.find(_.name == data.name).getOrElse(fail("missing Data"))
      assertEquals(observed(ref"D1").value, CellValue.Formula(text, None), definition)
    )
    val text = "LARGE(INDEX(SourceRange,0),1)"
    val wb = Workbook(data.put(ref"D1", CellValue.Formula(text, None)))
    val named = wb.copy(metadata =
      wb.metadata.copy(definedNames = Vector(DefinedName("SourceRange", "Data!A1:A3")))
    )
    val result = named.recalculate()
    assertEquals(result.errors.size, 0)
    val observed = result.workbook.sheets.find(_.name == data.name).getOrElse(fail("missing Data"))
    assertEquals(observed(ref"D1").value, CellValue.Formula(text, Some(CellValue.Number(5))))
  }

  test("GH-723: runtime INDEX errors retain their spreadsheet error code") {
    assertEquals(
      data.evaluateFormula("LET(k,#N/A,LARGE(INDEX(k,1),1))"),
      Right(CellValue.Error(CellError.NA))
    )
  }
