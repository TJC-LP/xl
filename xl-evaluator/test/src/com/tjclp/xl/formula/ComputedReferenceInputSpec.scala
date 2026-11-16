package com.tjclp.xl.formula
import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.FunSuite
class ComputedReferenceInputSpec extends FunSuite:
  private val data = Sheet("Data").put(
    ref"A1" -> 1,
    ref"A2" -> 3,
    ref"B1" -> 10,
    ref"B2" -> 20,
    ref"C1" -> 0,
    ref"C2" -> 1
  )
  private val names = Vector(
    DefinedName("Steps", "{0,1}"),
    DefinedName("TextRefs", "{\"A1\",\"A2\"}"),
    DefinedName("BadOffset", "OFFSET(A1,{0,1},0)")
  )
  private def recalc(text: String) =
    val wb = Workbook(
      data.put(ref"F1", CellValue.Formula(text, Some(CellValue.Number(99)))),
      Sheet("Other").put(ref"A1" -> 5)
    )
    wb.copy(metadata = wb.metadata.copy(definedNames = names)).recalculate()
  private def unsupported(text: String): Unit =
    val result = recalc(text)
    val observed = result.workbook.sheets
      .find(_.name == data.name)
      .getOrElse(fail("missing Data"))(ref"F1")
      .value
    assert(result.errors.nonEmpty, s"$text: cached=$observed")
    assertEquals(observed, CellValue.Formula(text, None), text)

  test("GH-723 reference inputs: OFFSET rows") {
    unsupported("COUNTIF(OFFSET(A1,{0,1},0),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(OFFSET(A1,{0,1},0),\">0\"))")
  }

  test("GH-723 reference inputs: OFFSET cols") {
    unsupported("COUNTIF(OFFSET(A1,0,{0,1}),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(OFFSET(A1,0,{0,1}),\">0\"))")
  }

  test("GH-723 reference inputs: OFFSET height") {
    unsupported("COUNTIF(OFFSET(A1,0,0,{1,2}),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(OFFSET(A1,0,0,{1,2}),\">0\"))")
  }

  test("GH-723 reference inputs: OFFSET width") {
    unsupported("COUNTIF(OFFSET(A1,0,0,1,{1,2}),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(OFFSET(A1,0,0,1,{1,2}),\">0\"))")
  }

  test("GH-723 reference inputs: OFFSET vertical rows") {
    unsupported("COUNTIF(OFFSET(A1,{0;1},0),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(OFFSET(A1,{0;1},0),\">0\"))")
  }

  test("GH-723 reference inputs: INDIRECT text") {
    unsupported("COUNTIF(INDIRECT({\"Data!A1\",\"Other!A1\"}),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(INDIRECT({\"Data!A1\",\"Other!A1\"}),\">0\"))")
  }

  test("GH-723 reference inputs: INDIRECT style") {
    unsupported("COUNTIF(INDIRECT(\"A1\",{TRUE,FALSE}),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(INDIRECT(\"A1\",{TRUE,FALSE}),\">0\"))")
  }

  test("GH-723 reference inputs: INDIRECT concat") {
    unsupported("COUNTIF(INDIRECT(\"A\"&{1,2}),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(INDIRECT(\"A\"&{1,2}),\">0\"))")
  }

  test("GH-723 reference inputs: LET OFFSET rows") {
    unsupported("LET(k,{0,1},COUNTIF(OFFSET(A1,k,0),\">0\"))")
    unsupported("SUMPRODUCT(LET(k,{0,1},COUNTIF(OFFSET(A1,k,0),\">0\")))")
  }

  test("GH-723 reference inputs: LET INDIRECT text") {
    unsupported("LET(k,{\"A1\",\"A2\"},COUNTIF(INDIRECT(k),\">0\"))")
    unsupported("SUMPRODUCT(LET(k,{\"A1\",\"A2\"},COUNTIF(INDIRECT(k),\">0\")))")
  }

  test("GH-723 reference inputs: named OFFSET rows") {
    unsupported("COUNTIF(OFFSET(A1,Steps,0),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(OFFSET(A1,Steps,0),\">0\"))")
  }

  test("GH-723 reference inputs: named INDIRECT text") {
    unsupported("COUNTIF(INDIRECT(TextRefs),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(INDIRECT(TextRefs),\">0\"))")
  }

  test("GH-723 reference inputs: nested OFFSET base") {
    unsupported("COUNTIF(OFFSET(OFFSET(A1,{0,1},0),0,0),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(OFFSET(OFFSET(A1,{0,1},0),0,0),\">0\"))")
  }

  test("GH-723 reference inputs: nested INDEX source") {
    unsupported("COUNTIF(INDEX(OFFSET(A1,{0,1},0),0),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(INDEX(OFFSET(A1,{0,1},0),0),\">0\"))")
  }

  test("GH-723 reference inputs: INDEX rows") {
    unsupported("COUNTIF(INDEX(A1:A2,{1,2}),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(INDEX(A1:A2,{1,2}),\">0\"))")
  }

  test("GH-723 reference inputs: INDEX columns") {
    unsupported("COUNTIF(INDEX(A1:B2,1,{1,2}),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(INDEX(A1:B2,1,{1,2}),\">0\"))")
  }

  test("GH-723 reference inputs: XLOOKUP lookup") {
    unsupported("COUNTIF(XLOOKUP({1,3},A1:A2,B1:B2),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(XLOOKUP({1,3},A1:A2,B1:B2),\">0\"))")
  }

  test("GH-723 reference inputs: XLOOKUP match mode") {
    unsupported("COUNTIF(XLOOKUP(1,A1:A2,B1:B2,0,{0,1}),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(XLOOKUP(1,A1:A2,B1:B2,0,{0,1}),\">0\"))")
  }

  test("GH-723 reference inputs: XLOOKUP search mode") {
    unsupported("COUNTIF(XLOOKUP(1,A1:A2,B1:B2,0,0,{1,-1}),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(XLOOKUP(1,A1:A2,B1:B2,0,0,{1,-1}),\">0\"))")
  }

  test("GH-723 reference inputs: scalar and singleton references remain supported") {
    List(
      "COUNTIF(OFFSET(A1,1,0),\">0\")",
      "COUNTIF(OFFSET(A1,{1},0),\">0\")",
      "COUNTIF(INDIRECT(\"A1\"),\">0\")",
      "COUNTIF(INDIRECT({\"A1\"}),\">0\")",
      "COUNTIF(INDEX(A1:A2,1),\">0\")",
      "COUNTIF(XLOOKUP(1,A1:A2,B1:B2),\">0\")",
      "LET(k,1,COUNTIF(OFFSET(A1,k,0),\">0\"))"
    ).foreach(text => assertEquals(data.evaluateFormula(text), Right(CellValue.Number(1)), text))
  }
  test("GH-723 reference inputs: genuine reference and scalar errors propagate") {
    List(
      "COUNTIF(OFFSET(A1,-1,0),1)" -> CellError.Ref,
      "COUNTIF(INDIRECT(\"not a reference\"),1)" -> CellError.Ref,
      "LET(k,#N/A,COUNTIF(OFFSET(A1,k,0),1))" -> CellError.NA,
      "COUNTIF(OFFSET(A1,{#N/A},0),1)" -> CellError.NA
    ).foreach((text, code) =>
      assertEquals(data.evaluateFormula(text), Right(CellValue.Error(code)), text)
    )
  }
  test(
    "GH-723 reference inputs: plain reference intersection and standalone values remain supported"
  ) {
    val text = "COUNTIF(OFFSET(A1,C1:C2,0),\">0\")"
    val result = recalc(text)
    assertEquals(result.errors.size, 0)
    val observed = result.workbook.sheets
      .find(_.name == data.name)
      .getOrElse(fail("missing Data"))(ref"F1")
      .value
    assertEquals(observed, CellValue.Formula(text, Some(CellValue.Number(1))))
    unsupported("SUMPRODUCT(" + text + ")")
    assertEquals(data.evaluateFormula("INDEX({7},1)"), Right(CellValue.Number(7)))
    assertEquals(data.evaluateFormula("SUM(ABS({-1,-2}))"), Right(CellValue.Number(3)))
  }
  test("GH-723 reference inputs: volatile scalar selectors evaluate once") {
    val draws = new java.util.concurrent.atomic.AtomicInteger(0)
    val rng = new Rng:
      def nextDouble(): Double =
        draws.incrementAndGet()
        0.25
    assertEquals(
      data.evaluateFormula("COUNTIF(OFFSET(A1,INT(RAND()),0),\">0\")", Clock.system, rng),
      Right(CellValue.Number(1))
    )
    assertEquals(draws.get(), 1)
  }

  test("GH-723 reference inputs: numeric coercion wrappers cannot hide arrays") {
    List(
      "COUNTIF(OFFSET(A1,{0,1}+0,0),\">0\")",
      "COUNTIF(OFFSET(A1,ABS({0,1}),0),\">0\")",
      "COUNTIF(OFFSET(A1,+({0,1}+0),0),\">0\")",
      "LET(k,{0,1},COUNTIF(OFFSET(A1,k+0,0),\">0\"))"
    ).foreach(text => unsupported("SUMPRODUCT(" + text + ")"))
  }

  test(
    "GH-723 reference inputs: computed names and reference-valued LET bindings cannot hide arrays"
  ) {
    List(
      "COUNTIF(BadOffset,\">0\")",
      "LET(r,OFFSET(A1,{0,1},0),COUNTIF(r,\">0\"))",
      "COUNTIF(INDEX(IF(TRUE,OFFSET(A1,{0,1},0),A1),0),\">0\")"
    ).foreach(text => unsupported("SUMPRODUCT(" + text + ")"))
  }
  test("GH-723 reference inputs: genuine spreadsheet errors retain normal cache behavior") {
    val text = "COUNTIF(OFFSET(A1,-1,0),1)"
    val result = recalc(text)
    assertEquals(result.errors.size, 0)
    val observed = result.workbook.sheets
      .find(_.name == data.name)
      .getOrElse(fail("missing Data"))(ref"F1")
      .value
    assertEquals(observed, CellValue.Formula(text, Some(CellValue.Error(CellError.Ref))))
  }

  test("GH-723 selected input: OFFSET through IF") {
    unsupported("COUNTIF(OFFSET(A1,IF(TRUE,{0,1},0),0),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(OFFSET(A1,IF(TRUE,{0,1},0),0),\">0\"))")
  }

  test("GH-723 selected input: OFFSET through IFS") {
    unsupported("COUNTIF(OFFSET(A1,IFS(TRUE,{0,1}),0),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(OFFSET(A1,IFS(TRUE,{0,1}),0),\">0\"))")
  }

  test("GH-723 selected input: OFFSET through CHOOSE") {
    unsupported("COUNTIF(OFFSET(A1,CHOOSE(1,{0,1},0),0),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(OFFSET(A1,CHOOSE(1,{0,1},0),0),\">0\"))")
  }

  test("GH-723 selected input: OFFSET through SWITCH") {
    unsupported("COUNTIF(OFFSET(A1,SWITCH(1,1,{0,1},0),0),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(OFFSET(A1,SWITCH(1,1,{0,1},0),0),\">0\"))")
  }

  test("GH-723 selected input: OFFSET through LET") {
    unsupported("COUNTIF(OFFSET(A1,LET(k,0,IF(TRUE,{0,1},k)),0),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(OFFSET(A1,LET(k,0,IF(TRUE,{0,1},k)),0),\">0\"))")
  }

  test("GH-723 selected input: INDIRECT through IF") {
    unsupported("COUNTIF(INDIRECT(IF(TRUE,{\"A1\",\"A2\"},\"A1\")),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(INDIRECT(IF(TRUE,{\"A1\",\"A2\"},\"A1\")),\">0\"))")
  }

  test("GH-723 selected input: INDIRECT through IFS") {
    unsupported("COUNTIF(INDIRECT(IFS(TRUE,{\"A1\",\"A2\"})),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(INDIRECT(IFS(TRUE,{\"A1\",\"A2\"})),\">0\"))")
  }

  test("GH-723 selected input: INDIRECT through CHOOSE") {
    unsupported("COUNTIF(INDIRECT(CHOOSE(1,{\"A1\",\"A2\"},\"A1\")),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(INDIRECT(CHOOSE(1,{\"A1\",\"A2\"},\"A1\")),\">0\"))")
  }

  test("GH-723 selected input: INDIRECT through SWITCH") {
    unsupported("COUNTIF(INDIRECT(SWITCH(1,1,{\"A1\",\"A2\"},\"A1\")),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(INDIRECT(SWITCH(1,1,{\"A1\",\"A2\"},\"A1\")),\">0\"))")
  }

  test("GH-723 selected input: INDIRECT through LET") {
    unsupported("COUNTIF(INDIRECT(LET(k,\"A1\",IF(TRUE,{\"A1\",\"A2\"},k))),\">0\")")
    unsupported("SUMPRODUCT(COUNTIF(INDIRECT(LET(k,\"A1\",IF(TRUE,{\"A1\",\"A2\"},k))),\">0\"))")
  }

  test("GH-723 selected input: unused array branches and scalar reductions remain supported") {
    List(
      "COUNTIF(OFFSET(A1,IF(FALSE,{0,1},1),0),\">0\")",
      "COUNTIF(OFFSET(A1,SUM({0,1}),0),\">0\")",
      "COUNTIF(INDIRECT(IF(FALSE,{\"A1\",\"A2\"},\"A1\")),\">0\")"
    ).foreach(text =>
      val result = recalc(text)
      assertEquals(result.errors.size, 0, text)
      val observed = result.workbook.sheets
        .find(_.name == data.name)
        .getOrElse(fail("missing Data"))(ref"F1")
        .value
      assertEquals(observed, CellValue.Formula(text, Some(CellValue.Number(1))), text)
    )
  }
  test("GH-723 selected input: true selected references keep plain intersection") {
    val text = "COUNTIF(OFFSET(A1,IF(TRUE,C1:C2,99),0),\">0\")"
    val result = recalc(text)
    assertEquals(result.errors.size, 0)
    val observed = result.workbook.sheets
      .find(_.name == data.name)
      .getOrElse(fail("missing Data"))(ref"F1")
      .value
    assertEquals(observed, CellValue.Formula(text, Some(CellValue.Number(1))))
    unsupported("SUMPRODUCT(" + text + ")")
  }

  test("GH-723 selected input: element-wise operators cannot erase bound or selected arrays") {
    List(
      "LET(k,{0,1},COUNTIF(OFFSET(A1,k+0,0),\">0\"))",
      "COUNTIF(OFFSET(A1,IF(TRUE,{0,1},0)+0,0),\">0\")",
      "LET(k,{1,2},COUNTIF(INDIRECT(\"A\"&k),\">0\"))",
      "COUNTIF(OFFSET(A1,IF(TRUE,{0,100},0)%,0),\">0\")"
    ).foreach(text =>
      unsupported(text)
      unsupported("SUMPRODUCT(" + text + ")")
    )
  }
  test("GH-723 selected input: scalar operator operands keep values and single reads") {
    List(
      "COUNTIF(OFFSET(A1,C1+1,0),\">0\")",
      "COUNTIF(OFFSET(A1,IF(TRUE,0,1)+1,0),\">0\")",
      "COUNTIF(INDIRECT(\"A\"&1),\">0\")"
    ).foreach(text => assertEquals(data.evaluateFormula(text), Right(CellValue.Number(1)), text))
  }

  test("GH-723 selected input: element-wise calls cannot erase nested or bound arrays") {
    List(
      "COUNTIF(OFFSET(A1,ABS(IF(TRUE,{0,1},0)),0),\">0\")",
      "LET(k,{0,1},COUNTIF(OFFSET(A1,ABS(k),0),\">0\"))",
      "COUNTIF(OFFSET(A1,LET(k,{0,1},ABS(k)),0),\">0\")"
    ).foreach(text =>
      unsupported(text)
      unsupported("SUMPRODUCT(" + text + ")")
    )
  }
  test("GH-723 selected input: nested scalar functions retain one random evaluation") {
    val draws = new java.util.concurrent.atomic.AtomicInteger(0)
    val rng = new Rng:
      def nextDouble(): Double =
        draws.incrementAndGet()
        0.25
    val text = "COUNTIF(OFFSET(A1,ABS(IF(TRUE,INT(RAND()),99)),0),\">0\")"
    assertEquals(data.evaluateFormula(text, Clock.system, rng), Right(CellValue.Number(1)))
    assertEquals(draws.get(), 1)
  }

  test("GH-723 input policy: supported producer lifting and reductions keep working") {
    List(
      "SUM(INDEX(A1:A2,{1;2}))" -> BigDecimal(4),
      "SUM(XLOOKUP({1,3},A1:A2,B1:B2))" -> BigDecimal(30)
    ).foreach((text, expected) =>
      assertEquals(data.evaluateFormula(text), Right(CellValue.Number(expected)), text)
    )
  }
  test("GH-723 input policy: an uncached precedent keeps its own plain-cell semantics") {
    val sheet = data.put(ref"C1", CellValue.Formula("IF(TRUE,{0,1},0)", None))
    assertEquals(
      sheet.evaluateFormula("COUNTIF(OFFSET(A1,C1,0),\">0\")"),
      Right(CellValue.Number(1))
    )
  }
