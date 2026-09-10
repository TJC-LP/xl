package com.tjclp.xl.formula

import com.tjclp.xl.XLResult
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-604: `@x` — Excel 365's implicit-intersection operator, stored as `_xlfn.SINGLE(x)`. A single
 * cell or scalar is itself; a column vector yields the cell in the formula's row, a row vector the
 * cell in the formula's column; anything else is `#VALUE!`. Defined names bound to ranges intersect
 * the same way. The parser accepts `@x` and `SINGLE(x)`; the printer emits `@x`.
 */
class ImplicitIntersectionSpec extends FunSuite:

  private val sheet = Sheet("Test")
    .put(ref"A1", CellValue.Number(BigDecimal(10)))
    .put(ref"A2", CellValue.Number(BigDecimal(20)))
    .put(ref"A3", CellValue.Number(BigDecimal(30)))
    .put(ref"B1", CellValue.Number(BigDecimal(1)))
    .put(ref"C1", CellValue.Number(BigDecimal(2)))
    .put(ref"B2", CellValue.Formula("A2*2", Some(CellValue.Number(BigDecimal(40)))))

  private val wb = Workbook(sheet)
    .withDefinedName("rng", "Test!$A$1:$A$3")
    .withDefinedName("flag", "Test!$B$1")
    .withDefinedName("k", "5")

  private def at(formula: String, cell: ARef): XLResult[CellValue] =
    sheet.evaluateFormula(formula, Clock.system, Some(wb), Some(cell))

  private def num(n: Int): XLResult[CellValue] = Right(CellValue.Number(BigDecimal(n)))

  private def errorAt(formula: String, cell: ARef): EvalError =
    val expr = FormulaParser.parse(formula).fold(e => fail(e.toString), identity)
    Evaluator
      .eval(expr, sheet, Clock.system, Some(wb), Some(cell))
      .fold(identity, v => fail(s"$formula at ${cell.toA1}: expected an error, got $v"))

  test("a column vector intersects the formula's row") {
    assertEquals(at("=@A1:A3", ref"D2"), num(20))
    assertEquals(at("=@A1:A3", ref"D3"), num(30))
    assertEquals(at("=SINGLE(A1:A3)", ref"D2"), num(20), "the stored spelling is accepted too")
  }

  test("a row vector intersects the formula's column") {
    assertEquals(at("=@A1:C1", ref"B4"), num(1))
    assertEquals(at("=@A1:C1", ref"C4"), num(2))
  }

  test("a single cell, a scalar and an array value are themselves (top-left for arrays)") {
    assertEquals(at("=@A1", ref"Z9"), num(10))
    assertEquals(at("=@(1+2)", ref"Z9"), num(3))
    assertEquals(at("=@SEQUENCE(3)", ref"Z9"), num(1))
    assertEquals(at("=@B2", ref"Z9"), num(40), "a cached formula cell reads its cache")
  }

  test("no intersection and a 2-D range are #VALUE!") {
    errorAt("=@A1:A3", ref"D5") match
      case EvalError.ErrorValue(CellError.Value, _) => ()
      case other => fail(s"expected #VALUE!, got $other")
    errorAt("=@A1:B2", ref"D5") match
      case EvalError.ErrorValue(CellError.Value, _) => ()
      case other => fail(s"expected #VALUE!, got $other")
  }

  test("a multi-cell range needs the formula's position") {
    val expr = FormulaParser.parse("=@A1:A3").fold(e => fail(e.toString), identity)
    Evaluator.eval(expr, sheet) match
      case Left(EvalError.EvalFailed(msg, _)) => assert(msg.contains("cell position"), msg)
      case other => fail(s"expected EvalFailed, got $other")
  }

  test("defined names: a range name intersects, a cell name and a constant are themselves") {
    assertEquals(at("=@rng", ref"D3"), num(30))
    assertEquals(at("=@flag", ref"D3"), num(1))
    assertEquals(at("=@k", ref"D3"), num(5))
    // the dogfood shape: IF(@name=1, other!cell, 0)
    assertEquals(at("=IF(@flag=1,A2,0)", ref"D3"), num(20))
  }

  test("@ binds as a primary: @A1:A3*2 is (@A1:A3)*2, @A1:A3% is (@A1:A3)%") {
    assertEquals(at("=@A1:A3*2", ref"D2"), num(40))
    assertEquals(at("=@A1:A3%", ref"D2"), Right(CellValue.Number(BigDecimal("0.2"))))
    assertEquals(at("=SUM(@A1:A3,1)", ref"D1"), num(11))
  }

  test("printing: SINGLE(x) renders as @x, operands parenthesize only when they must") {
    def printed(f: String): String =
      FormulaParser.parse(f).map(FormulaPrinter.print(_)).fold(e => fail(e.toString), identity)
    assertEquals(printed("=@acq"), "=@acq")
    assertEquals(printed("=SINGLE(acq)"), "=@acq")
    assertEquals(printed("=@'My Sheet'!A1:A10"), "=@'My Sheet'!A1:A10")
    assertEquals(printed("=@INDEX(A1:A3,2,1)"), "=@INDEX(A1:A3, 2, 1)")
    assertEquals(printed("=@(A1+A2)"), "=@(A1+A2)")
    assertEquals(printed("=@A1:A3*2"), "=@A1:A3*2")
    assertEquals(printed("=@A1:A3%"), "=@A1:A3%")
    assertEquals(printed("=IF(@acq=1,'M&A'!I12,0)"), "=IF(@acq=1, 'M&A'!I12, 0)")
    val parsed = FormulaParser.parse("=IF(@acq=1,'M&A'!I12,0)")
    assertEquals(parsed.map(FormulaPrinter.print(_)).flatMap(FormulaParser.parse), parsed)
    assertEquals(
      FormulaPrinter.printFileForm(parsed.fold(e => fail(e.toString), identity)),
      "IF(@acq=1,'M&A'!I12,0)"
    )
  }

  test("@ with no operand is a parse error, not a crash") {
    assert(FormulaParser.parse("=@").isLeft)
    assert(FormulaParser.parse("=@+1").isLeft)
  }
