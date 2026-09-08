package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.graph.DependencyGraph
import munit.FunSuite

/**
 * GH-612: Excel error literals (`#REF!`, `#N/A`, …) are first-class formula tokens. They are what a
 * drag writes when a reference falls off the grid (`=A1` copied from B2 to B1 is `=#REF!`), what
 * Excel itself writes after a delete, and what users type (`=IF(x, #N/A, 1)`). They parse to
 * `TExpr.ErrorLit`, print back verbatim, evaluate to the error value they name, and contribute no
 * dependency edges.
 */
class ErrorLiteralSpec extends FunSuite:

  private val sheet = Sheet("T").put(ref"A1", CellValue.Number(BigDecimal(5)))

  private def eval(formula: String): CellValue =
    sheet.evaluateFormula(formula).fold(e => fail(e.message), identity)

  test("an error literal evaluates to the error value it names") {
    CellError.values.foreach { err =>
      assertEquals(eval(s"=${err.toExcel}"), CellValue.Error(err))
    }
  }

  test("an error literal propagates through arithmetic, functions and range slots") {
    assertEquals(eval("=#REF!+A1"), CellValue.Error(CellError.Ref))
    assertEquals(eval("=ABS(#REF!)"), CellValue.Error(CellError.Ref))
    assertEquals(eval("=SUM(#REF!)"), CellValue.Error(CellError.Ref))
    assertEquals(eval("=IF(A1>0,#N/A,1)"), CellValue.Error(CellError.NA))
    assertEquals(eval("=IF(A1<0,#N/A,1)"), CellValue.Number(BigDecimal(1)))
  }

  test("an error literal in a range-typed slot evaluates to the error it names (GH-612 S1)") {
    assertEquals(eval("=COUNTIF(#REF!,1)"), CellValue.Error(CellError.Ref))
    assertEquals(eval("=SUMIF(#REF!,\">0\")"), CellValue.Error(CellError.Ref))
    assertEquals(eval("=VLOOKUP(1,#REF!,1,FALSE)"), CellValue.Error(CellError.Ref))
    assertEquals(eval("=SUMPRODUCT(#N/A,A1:A1)"), CellValue.Error(CellError.NA))
    assertEquals(eval("=IFERROR(COUNTIF(#REF!,1),-1)"), CellValue.Number(BigDecimal(-1)))
  }

  test("ROW/COLUMN/ROWS/COLUMNS of an error literal are that error, not a host failure (N5)") {
    assertEquals(eval("=ROWS(#REF!)"), CellValue.Error(CellError.Ref))
    assertEquals(eval("=COLUMNS(#REF!)"), CellValue.Error(CellError.Ref))
    assertEquals(eval("=ROW(#REF!)"), CellValue.Error(CellError.Ref))
    assertEquals(eval("=COLUMN(#N/A)"), CellValue.Error(CellError.NA))
  }

  test("error literals are case-insensitive on entry, like Excel") {
    assertEquals(eval("=#ref!"), CellValue.Error(CellError.Ref))
    assertEquals(eval("=#n/a+1"), CellValue.Error(CellError.NA))
  }

  test("#N/A does not swallow an operator that follows it: #N/A/2 is #N/A divided by 2") {
    assertEquals(eval("=#N/A/2"), CellValue.Error(CellError.NA))
    assertEquals(eval("=1/#N/A"), CellValue.Error(CellError.NA))
    assertEquals(eval("=IFERROR(#N/A/2,7)"), CellValue.Number(BigDecimal(7)))
    assertEquals(eval("=#NAME?*2"), CellValue.Error(CellError.Name))
  }

  test("IFERROR and ISERROR see an error literal as an error value") {
    assertEquals(eval("=IFERROR(#DIV/0!,0)"), CellValue.Number(BigDecimal(0)))
    assertEquals(eval("=ISERROR(#VALUE!)"), CellValue.Bool(true))
  }

  private def parsed(formula: String): TExpr[?] =
    FormulaParser.parse(formula) match
      case Right(expr) => expr
      case Left(err) => fail(s"$formula should parse: $err")

  test("an error literal contributes no dependency edges") {
    assertEquals(DependencyGraph.extractDependencies(parsed("=#REF!+A1")), Set(ref"A1"))
    assertEquals(DependencyGraph.extractDependencies(parsed("=#N/A")), Set.empty[ARef])
  }

  test("an unknown # token is a parse error naming the whole token, not a literal") {
    assert(FormulaParser.parse("=#BOGUS!").isLeft)
    assert(FormulaParser.parse("=#").isLeft)
    FormulaParser.parse("=#NOT_AN_ERROR") match
      case Left(err) => assert(err.toString.contains("#NOT_AN_ERROR"), err.toString)
      case Right(expr) => fail(s"expected a parse error, got $expr")
  }

  // ===== GH-630: the modern error codes and ERROR.TYPE =====

  test("GH-630: the modern codes parse (case-insensitively), print canonically and evaluate") {
    val modern = List(
      "#SPILL!" -> CellError.Spill,
      "#CALC!" -> CellError.Calc,
      "#FIELD!" -> CellError.Field,
      "#CONNECT!" -> CellError.Connect,
      "#BLOCKED!" -> CellError.Blocked,
      "#UNKNOWN!" -> CellError.Unknown,
      "#GETTING_DATA" -> CellError.GettingData
    )
    modern.foreach { (code, err) =>
      assertEquals(CellError.parse(code), Right(err), code)
      assertEquals(CellError.parse(code.toLowerCase), Right(err), code)
      assertEquals(err.toExcel, code)
      assertEquals(eval(s"=$code"), CellValue.Error(err), code)
      assertEquals(eval(s"=${code.toLowerCase}"), CellValue.Error(err), code)
      assertEquals(FormulaParser.parse(s"=$code").map(FormulaPrinter.print(_)), Right(s"=$code"))
    }
    // #GETTING_DATA has no terminator: it is matched whole, and what follows it is an operator
    assertEquals(eval("=#GETTING_DATA+1"), CellValue.Error(CellError.GettingData))
    assertEquals(eval("=IFERROR(#SPILL!,0)"), CellValue.Number(BigDecimal(0)))
    assertEquals(eval("=ISERROR(#CALC!)"), CellValue.Bool(true))
  }

  test("GH-630: ERROR.TYPE returns Microsoft's number for each error, #N/A for anything else") {
    val expected = Map(
      CellError.Null -> 1,
      CellError.Div0 -> 2,
      CellError.Value -> 3,
      CellError.Ref -> 4,
      CellError.Name -> 5,
      CellError.Num -> 6,
      CellError.NA -> 7,
      CellError.GettingData -> 8,
      CellError.Spill -> 9,
      CellError.Connect -> 10,
      CellError.Blocked -> 11,
      CellError.Unknown -> 12,
      CellError.Field -> 13,
      CellError.Calc -> 14
    )
    CellError.values.foreach { err =>
      val n = expected(err)
      assertEquals(err.errorTypeNumber, n)
      assertEquals(
        eval(s"=ERROR.TYPE(${err.toExcel})"),
        CellValue.Number(BigDecimal(n)),
        err.toExcel
      )
    }
    // a computed error, a cell holding one, a cached formula error — and non-errors
    val errors = Sheet("T")
      .put(ref"A1", CellValue.Error(CellError.Div0))
      .put(ref"A2", CellValue.Formula("1/0", Some(CellValue.Error(CellError.Div0))))
      .put(ref"A3", CellValue.Number(BigDecimal(3)))
    def evalIn(formula: String): CellValue =
      errors.evaluateFormula(formula).fold(e => fail(e.message), identity)
    assertEquals(evalIn("=ERROR.TYPE(1/0)"), CellValue.Number(BigDecimal(2)))
    assertEquals(evalIn("=ERROR.TYPE(A1)"), CellValue.Number(BigDecimal(2)))
    assertEquals(evalIn("=ERROR.TYPE(A2)"), CellValue.Number(BigDecimal(2)))
    assertEquals(evalIn("=ERROR.TYPE(A3)"), CellValue.Error(CellError.NA))
    assertEquals(evalIn("=ERROR.TYPE(1)"), CellValue.Error(CellError.NA))
    assertEquals(evalIn("=ERROR.TYPE(\"x\")"), CellValue.Error(CellError.NA))
    assertEquals(evalIn("=ERROR.TYPE(A9)"), CellValue.Error(CellError.NA))
    assertEquals(evalIn("=IF(ERROR.TYPE(A1)=2,\"div\",\"other\")"), CellValue.Text("div"))
  }
