package com.tjclp.xl.formula

import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

/** GH-605: RRI(nper, pv, fv) — the equivalent interest rate for the growth of an investment. */
class RriSpec extends FunSuite:

  private val sheet = Sheet("Test")

  private def number(formula: String): Double =
    sheet.evaluateFormula(formula) match
      case Right(CellValue.Number(n)) => n.toDouble
      case other => fail(s"$formula: expected a number, got $other")

  private def error(formula: String): EvalError =
    FormulaParser
      .parse(formula)
      .fold(e => fail(e.toString), expr => Evaluator.eval(expr, sheet))
      .fold(identity, v => fail(s"$formula: expected an error, got $v"))

  test("RRI is (fv/pv)^(1/nper) - 1") {
    assertEqualsDouble(number("=RRI(10,100,150)"), 0.0413797439, 1e-9)
    // Microsoft's documented example
    assertEqualsDouble(number("=RRI(96,10000,11000)"), 0.0009933, 1e-7)
    assertEqualsDouble(number("=RRI(10,100,0)"), -1.0, 1e-12)
    assertEqualsDouble(number("=RRI(1,100,110)"), 0.1, 1e-12)
  }

  test("Excel's #NUM! rules: nper ≤ 0, pv = 0, or a negative ratio") {
    List("=RRI(0,100,150)", "=RRI(-2,100,150)", "=RRI(10,0,150)", "=RRI(10,-100,150)").foreach {
      f =>
        error(f) match
          case EvalError.ErrorValue(CellError.Num, _) => ()
          case other => fail(s"$f: expected #NUM!, got $other")
    }
  }

  test("RRI round-trips through the printer and is in the registry") {
    val parsed = FormulaParser.parse("=RRI(10,100,150)")
    assertEquals(parsed.map(FormulaPrinter.print(_)), Right("=RRI(10, 100, 150)"))
    assert(com.tjclp.xl.formula.functions.FunctionRegistry.allNames.contains("RRI"))
  }
