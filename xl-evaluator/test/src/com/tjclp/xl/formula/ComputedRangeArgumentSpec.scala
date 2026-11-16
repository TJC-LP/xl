package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.formula.printer.{FormulaPrinter, FormulaShifter}
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

class ComputedRangeArgumentSpec extends FunSuite:
  test("GH-723: COUNTIF accepts a range computed directly by OFFSET") {
    val sheet = Sheet("Data").put(ref"A1" -> 1, ref"A2" -> 3, ref"A3" -> 5)
    assertEquals(
      sheet.evaluateFormula("=COUNTIF(OFFSET(A1,0,0,3,1),\">2\")"),
      Right(CellValue.Number(BigDecimal(2)))
    )
  }

  private val data = Sheet("Data").put(
    ref"A1" -> 1,
    ref"A2" -> 3,
    ref"A3" -> 5,
    ref"B1" -> 10,
    ref"B2" -> 20,
    ref"B3" -> 30
  )

  for source <- List("INDIRECT(\"A1:A3\")", "INDEX(A1:A3,0)", "A1:INDEX(A1:A3,3)") do
    test(s"GH-723: COUNTIF accepts $source as a reference") {
      assertEquals(
        data.evaluateFormula(s"=COUNTIF($source,\">2\")"),
        Right(CellValue.Number(BigDecimal(2)))
      )
    }

  test("GH-723: computed reference errors retain their code") {
    assertEquals(
      data.evaluateFormula("=COUNTIF(OFFSET(A1,-1,0),1)"),
      Right(CellValue.Error(CellError.Ref))
    )
  }

  test("GH-723: a computed scalar fallback is not a reference") {
    assertEquals(
      data.evaluateFormula("=COUNTIF(XLOOKUP(9,A1:A3,B1:B3,7),\">0\")"),
      Right(CellValue.Error(CellError.Value))
    )
  }

  test("GH-723: SUMIF resizes a computed target from the resolved criteria shape") {
    assertEquals(
      data.evaluateFormula("=SUMIF(OFFSET(A1,0,0,3,1),\">2\",OFFSET(B1,0,0))"),
      Right(CellValue.Number(BigDecimal(50)))
    )
  }

  test("GH-723: computed references preserve LET bindings") {
    assertEquals(
      data.evaluateFormula("=LET(r,A1,COUNTIF(OFFSET(r,0,0,3,1),\">2\"))"),
      Right(CellValue.Number(BigDecimal(2)))
    )
  }

  test("GH-723: parser and both printers round-trip computed slots") {
    val text = "COUNTIF(OFFSET(A1,0,0,3,1),\">2\")"
    val parsed = FormulaParser.parse(text).toOption.get
    assertEquals(FormulaPrinter.printFileForm(parsed), text)
    assertEquals(FormulaParser.parse(FormulaPrinter.print(parsed)), Right(parsed))
  }

  test("GH-723: computed slots shift their inner references") {
    val parsed = FormulaParser.parse("COUNTIF(OFFSET(A1,0,0,3,1),\">2\")").toOption.get
    assertEquals(
      FormulaPrinter.printFileForm(FormulaShifter.shift(parsed, 1, 1)),
      "COUNTIF(OFFSET(B2,0,0,3,1),\">2\")"
    )
  }

  test("GH-723: computed slots rename qualified inner references") {
    val parsed = FormulaParser.parse("COUNTIF(OFFSET(Data!A1,0,0,3,1),\">2\")").toOption.get
    assert(FormulaShifter.mentionsSheet(parsed, "Data"))
    val renamed =
      FormulaShifter.renameSheet(parsed, SheetName.unsafe("Data"), SheetName.unsafe("New"))
    assertEquals(FormulaPrinter.printFileForm(renamed), "COUNTIF(OFFSET(New!A1,0,0,3,1),\">2\")")
  }

  test("GH-723: direct INDEX range slots are scheduled dynamically") {
    val s = data.put(ref"D1", CellValue.Formula("COUNTIF(INDEX(A1:A3,0),\">2\")", None))
    assert(DependencyGraph.dynamicCells(s).contains(ref"D1"))
  }

  test("GH-723: computed slot traversals retain external references") {
    val parsed = FormulaParser.parse("COUNTIF(OFFSET([1]Other!A1,0,0),1)").toOption.get
    assert(com.tjclp.xl.formula.ast.TExpr.containsExternalRef(parsed))
  }

  test("GH-723: recalculation reads fresh cross-sheet computed ranges") {
    val source =
      Sheet("Source").put(ref"A1", CellValue.Formula("1+1", Some(CellValue.Number(BigDecimal(99)))))
    val reader =
      Sheet("Reader").put(ref"A1", CellValue.Formula("COUNTIF(INDEX(Source!A1:A1,0),\">2\")", None))
    val result = Workbook(source, reader).recalculate()
    assertEquals(result.errors, Vector.empty)
    assertEquals(
      result.evaluated(SheetName.unsafe("Reader"))(ref"A1"),
      CellValue.Number(BigDecimal(0))
    )
  }

  test("GH-723: structural row edits reach computed range arguments") {
    val parsed = FormulaParser.parse("COUNTIF(OFFSET(A1,0,0,3,1),1)").toOption.get
    val shifted = FormulaShifter.shiftStructural(parsed, true, "Data", true, 0, 1)
    assertEquals(FormulaPrinter.printFileForm(shifted), "COUNTIF(OFFSET(A2,0,0,3,1),1)")
  }

  test("GH-723: INDEX value inputs remain outside reference-only slot admission") {
    assert(FormulaParser.parse("COUNTIF(INDEX({1},1,1),1)").isLeft)
    assert(FormulaParser.parse("COUNTIF(INDEX({1},1,2),1)").isLeft)
  }

  test("GH-723: the computed-call pre-filter ignores literals and quoted payloads") {
    import com.tjclp.xl.formula.functions.ReferenceFunctionNames
    val rejected =
      List("\"a:b\"&\"INDEX(\"", "SUM('INDEX('!A1)", "SUM(Table1[INDEX(])", "MY_INDEX(A1)")
    rejected.foreach(text =>
      assert(
        !ReferenceFunctionNames
          .mayContainComputingCall(text, text.toUpperCase(java.util.Locale.ROOT)),
        text
      )
    )
    val accepted = List(
      "COUNTIF(INDEX(A:A,0),1)",
      "COUNTIF(_xlfn.XLOOKUP (1,A:A,B:B),1)",
      "\"ß\"&COUNTIF(INDEX(A:A,0),1)"
    )
    accepted.foreach(text =>
      assert(
        ReferenceFunctionNames
          .mayContainComputingCall(text, text.toUpperCase(java.util.Locale.ROOT)),
        text
      )
    )
  }
