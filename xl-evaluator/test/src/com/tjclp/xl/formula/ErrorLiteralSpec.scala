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

  test("an unknown # token is a parse error, not a literal") {
    assert(FormulaParser.parse("=#BOGUS!").isLeft)
    assert(FormulaParser.parse("=#").isLeft)
  }
