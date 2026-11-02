package com.tjclp.xl.formula

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.formula.eval.ArrayResult
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

class DepreciationSpec extends FunSuite:
  private val sheet = Sheet("Test")
  private def number(formula: String, expected: Double): Unit =
    sheet.evaluateFormula(formula) match
      case Right(CellValue.Number(n)) => assertEqualsDouble(n.toDouble, expected, 1e-9)
      case other => fail(s"$formula: $other")

  test("GH-684: Microsoft's SLN and SYD examples") {
    number("SLN(30000,7500,10)", 2250)
    number("SYD(30000,7500,10,1)", 4090.90909090909)
    number("SYD(30000,7500,10,10)", 409.090909090909)
  }

  test("GH-684: fractions, signed inputs and out-of-life periods match the LO oracle") {
    List(
      "SLN(100,10,-2)" -> -45.0,
      "SLN(-100,10,2)" -> -55.0,
      "SLN(100,-10,2)" -> 55.0,
      "SLN(10,100,2)" -> -45.0,
      "SYD(100,10,2,0.5)" -> 75.0,
      "SYD(100,10,2.5,1.5)" -> 41.1428571428571,
      "SYD(100,10,-2,1)" -> -180.0,
      "SYD(100,10,2,0)" -> 90.0,
      "SYD(100,10,2,3)" -> 0.0,
      "SYD(10,100,2,1)" -> -60.0,
      "SYD(-100,10,2,1)" -> -73.3333333333333
    ).foreach((f, n) => number(f, n))
  }

  test("GH-684: zero denominators, overflow and propagated errors stay cell errors") {
    List(
      "SLN(100,10,0)" -> CellError.Div0,
      "SYD(100,10,0,1)" -> CellError.Num,
      "SYD(100,10,-1,1)" -> CellError.Num,
      "SLN(1E308,-1E308,1)" -> CellError.Num,
      "SYD(1E308,-1E308,2,0)" -> CellError.Num,
      "SLN(#N/A,10,2)" -> CellError.NA,
      "SLN(\"abc\",10,2)" -> CellError.Value
    ).foreach { (f, error) =>
      assertEquals(sheet.evaluateFormula(f), Right(CellValue.Error(error)), f)
    }
    number("IFERROR(SLN(100,10,0),42)", 42)
  }

  test("GH-684: scalar text, boolean, omitted and blank arguments use numeric coercion") {
    number("SLN(\"100\",10,2)", 45)
    number("SLN(TRUE(),10,2)", -4.5)
    number("SYD(100,10,2,TRUE())", 60)
    number("SLN(100,,2)", 50)
    val blank = sheet.evaluateFormula("SLN(A1,10,2)")
    assertEquals(blank, Right(CellValue.Number(BigDecimal(-5))))
  }

  test("GH-684: arity, registry and file-form round trips") {
    List("SLN(1,2)", "SLN(1,2,3,4)", "SYD(1,2,3)", "SYD(1,2,3,4,5)")
      .foreach(f => assert(FormulaParser.parse(f).isLeft, f))
    List("SLN(100,10,2)", "SYD(100,10,2,1)").foreach { f =>
      val parsed = FormulaParser.parse(f).fold(e => fail(e.toString), identity)
      val stored = FormulaPrinter.printFileForm(parsed)
      assertEquals(stored, f)
      assert(FormulaParser.parse(stored).isRight)
      assertEquals(com.tjclp.xl.ooxml.FormulaStorage.toStored(stored), stored)
    }
    val names = com.tjclp.xl.formula.functions.FunctionRegistry.allNames
    assert(names.contains("SLN") && names.contains("SYD"))
  }

  test("GH-684: array lifting retains per-element errors") {
    val expr = FormulaParser
      .parse("SLN({100;200},10,{2;0})")
      .fold(e => fail(e.toString), identity)
    val expected = ArrayResult(
      Vector(Vector(CellValue.Number(BigDecimal(45))), Vector(CellValue.Error(CellError.Div0)))
    )
    val observed: Either[EvalError, Any] = Evaluator.arrayInstance.eval(expr, sheet)
    assertEquals(observed, Right(expected))
  }

  test("GH-684: every scalar reference remains a dependency and evaluation reads changed inputs") {
    val refs = (0 to 3).map(i => ARef.from0(i, 0)).toSet
    val expr = FormulaParser
      .parse("SYD(A1,B1,C1,D1)")
      .fold(e => fail(e.toString), identity)
    assertEquals(DependencyGraph.extractDependencies(expr), refs)
    val inputs = sheet
      .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(100)))
      .put(ARef.from0(1, 0), CellValue.Number(BigDecimal(10)))
      .put(ARef.from0(2, 0), CellValue.Number(BigDecimal(2)))
      .put(ARef.from0(3, 0), CellValue.Number(BigDecimal(1)))
    assertEquals(
      inputs.evaluateFormula("SYD(A1,B1,C1,D1)"),
      Right(CellValue.Number(BigDecimal(60)))
    )
    assertEquals(
      inputs
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(200)))
        .evaluateFormula("SYD(A1,B1,C1,D1)"),
      Right(CellValue.Number(BigDecimal("126.66666666666667")))
    )
  }
