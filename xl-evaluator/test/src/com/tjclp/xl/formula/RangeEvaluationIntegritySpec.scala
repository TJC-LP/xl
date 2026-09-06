package com.tjclp.xl.formula

import java.time.LocalDate

import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.eval.ArrayResult

import munit.FunSuite

class RangeEvaluationIntegritySpec extends FunSuite:
  private def num(value: Int): CellValue = CellValue.Number(BigDecimal(value))

  private def cashflows: Sheet = Sheet("S")
    .put(ref"A1", CellValue.Formula("100"))
    .put(ref"A2", num(121))

  private def assertRejected(result: XLResult[CellValue]): Unit = result match
    case Left(_) => ()
    case Right(CellValue.Error(_)) => ()
    case other => fail(s"Expected an explicit failure, got $other")

  test("GH-499: uncached cash flows agree for local, qualified, and named ranges") {
    val sheet = cashflows
    val wb = Workbook(sheet).withDefinedName("Flows", "S!$A$1:$A$2")
    List("=NPV(0,A1:A2)", "=NPV(0,S!A1:A2)", "=NPV(0,Flows)").foreach { formula =>
      assertEquals(sheet.evaluateFormula(formula, workbook = Some(wb)), Right(num(221)))
    }
    assertEquals(sheet(ref"A1").value, CellValue.Formula("100"))
  }

  test("GH-499: a failed cash-flow formula cannot disappear from NPV or IRR") {
    val sheet = cashflows.put(ref"A1", CellValue.Formula("UNSUPPORTED(1)"))
    List("=NPV(0,A1:A2)", "=IRR(A1:A2)").foreach { formula =>
      assert(sheet.evaluateFormula(formula).isLeft, formula)
    }
  }

  test("GH-499: existing numeric caches remain authoritative in direct evaluation") {
    val sheet = cashflows.put(ref"A1", CellValue.Formula("999", Some(num(100))))
    assertEquals(sheet.evaluateFormula("=NPV(0,A1:A2)"), Right(num(221)))
  }

  test("GH-499: cash-flow text and blanks retain their skip semantics") {
    val sheet = Sheet("S")
      .put(ref"A1", CellValue.Formula("100"))
      .put(ref"A2", CellValue.Text("annotation"))
      .put(ref"A4", num(121))
    assertEquals(sheet.evaluateFormula("=NPV(0,A1:A4)"), Right(num(221)))
  }

  private def datedFlows: Sheet = Sheet("S")
    .put(ref"A1", CellValue.Formula("-100"))
    .put(ref"A3", CellValue.Formula("121"))
    .put(ref"B1", CellValue.Formula("DATE(2021,1,1)"))
    .put(ref"B3", CellValue.Formula("DATE(2022,1,1)"))

  private def materializedDatedFlows: Sheet = datedFlows
    .put(ref"A1", num(-100))
    .put(ref"A3", num(121))
    .put(ref"B1", CellValue.DateTime(LocalDate.of(2021, 1, 1).atStartOfDay()))
    .put(ref"B3", CellValue.DateTime(LocalDate.of(2022, 1, 1).atStartOfDay()))

  test("GH-499: XNPV and XIRR resolve cash-flow and date formulas with aligned holes") {
    val sheet = datedFlows
    assertEquals(sheet.evaluateFormula("=XNPV(0,A1:A3,B1:B3)"), Right(num(21)))
    sheet.evaluateFormula("=XIRR(A1:A3,B1:B3)") match
      case Right(CellValue.Number(rate)) =>
        assert((rate - BigDecimal("0.21")).abs < BigDecimal("0.000001"))
      case other => fail(s"Expected a numeric rate, got $other")
  }

  test("GH-499: equal filtered lengths cannot conceal differently positioned blank cells") {
    val sheet = materializedDatedFlows
      .put(ref"B2", CellValue.DateTime(LocalDate.of(2022, 1, 1).atStartOfDay()))
      .remove(ref"B3")
    assertRejected(sheet.evaluateFormula("=XNPV(0,A1:A3,B1:B3)"))
    assertRejected(sheet.evaluateFormula("=XIRR(A1:A3,B1:B3)"))
  }

  test("GH-499: an invalid date cannot be dropped to match a shorter values range") {
    val sheet = materializedDatedFlows.put(ref"B2", num(-5))
    assertRejected(sheet.evaluateFormula("=XNPV(0,A1:A3,B1:B3)"))
    assertRejected(sheet.evaluateFormula("=XIRR(A1:A3,B1:B3)"))
  }

  private def arraySheet: Sheet = Sheet("S")
    .put(ref"A1", CellValue.Formula("2"))
    .put(ref"A2", CellValue.Formula("1"))
    .put(ref"B1", CellValue.Formula("A1*10"))
    .put(ref"B2", CellValue.Formula("A2*10"))

  private def evalArray(
    sheet: Sheet,
    formula: String,
    workbook: Option[Workbook] = None
  ): Either[EvalError, ArrayResult] =
    FormulaParser
      .parse(formula)
      .left
      .map(error => EvalError.EvalFailed(error.toString, None))
      .flatMap(expr => Evaluator.arrayInstance.eval(expr, sheet, workbook = workbook))
      .flatMap {
        case result: ArrayResult => Right(result)
        case other => Left(EvalError.TypeMismatch("array", "ArrayResult", other.toString))
      }

  private def array(sheet: Sheet, formula: String): ArrayResult =
    evalArray(sheet, formula, Some(Workbook(sheet))).fold(error => fail(error.toString), identity)

  test("GH-499: array arithmetic resolves formulas and preserves the rectangular shape") {
    assertEquals(
      array(arraySheet, "=A1:B2*2").values,
      Vector(Vector(num(4), num(40)), Vector(num(2), num(20)))
    )
    assertEquals(
      array(arraySheet.remove(ref"B2"), "=S!A1:B2*2").values,
      Vector(Vector(num(4), num(40)), Vector(num(2), num(0)))
    )
  }

  test("GH-499: TRANSPOSE and SORT consume evaluated range values") {
    assertEquals(
      array(arraySheet, "=TRANSPOSE(A1:B2)").values,
      Vector(Vector(num(2), num(1)), Vector(num(20), num(10)))
    )
    assertEquals(
      array(arraySheet, "=SORT(A1:B2)").values,
      Vector(Vector(num(1), num(10)), Vector(num(2), num(20)))
    )
  }

  test("GH-499: UNIQUE and FILTER evaluate keys and include flags") {
    val sheet = arraySheet
      .put(ref"A3", CellValue.Formula("2"))
      .put(ref"C1", CellValue.Formula("FALSE"))
      .put(ref"C2", CellValue.Formula("TRUE"))
    assertEquals(array(sheet, "=UNIQUE(A1:A3)").values, Vector(Vector(num(2)), Vector(num(1))))
    assertEquals(array(sheet, "=FILTER(A1:B2,C1:C2)").values, Vector(Vector(num(1), num(10))))
  }

  test("GH-499: a failed array element produces a diagnostic instead of a raw Formula value") {
    val sheet = arraySheet.put(ref"A1", CellValue.Formula("UNSUPPORTED(1)"))
    List("=A1:B2*2", "=TRANSPOSE(A1:B2)", "=SORT(A1:B2)").foreach { formula =>
      assert(evalArray(sheet, formula).isLeft, formula)
    }
  }

  test("GH-499: Excel error values retain their array positions") {
    val sheet = arraySheet.put(ref"A1", CellValue.Formula("1/0"))
    val result = array(sheet, "=A1:A2*2")
    assertEquals(result.values, Vector(Vector(CellValue.Error(CellError.Div0)), Vector(num(2))))
  }

  test("GH-499: range-valued LET and defined names resolve uncached formulas") {
    val sheet = arraySheet
    val wb = Workbook(sheet).withDefinedName("Values", "S!$A$1:$A$2")
    List("=LET(r,A1:A2,r*2)", "=Values*2").foreach { formula =>
      val result = evalArray(sheet, formula, Some(wb))
      assertEquals(result.map(_.values), Right(Vector(Vector(num(4)), Vector(num(2)))))
    }
  }

  test("GH-499: an uncached formula retains its own ROW context when read through a range") {
    val sheet = Sheet("S").put(ref"A5", CellValue.Formula("ROW()"))
    assertEquals(sheet.evaluateFormula("=NPV(0,A5:A5)"), Right(num(5)))
  }

  test("GH-499: typed aggregate nodes agree with function calls on uncached inputs") {
    val sheet = cashflows
    val expr = TExpr.Aggregate("SUM", TExpr.RangeLocation.Local(ref"A1:A2"))
    assertEquals(Evaluator.instance.eval(expr, sheet), Right(BigDecimal(221)))
  }

  test("GH-499: range recursion shares the caller's RNG and memo across input cells") {
    final class CountingRng extends Rng:
      var calls = 0
      def nextDouble(): Double =
        calls += 1
        0.25
    val rng = new CountingRng
    val sheet = Sheet("S")
      .put(ref"A1", CellValue.Formula("RAND()"))
      .put(ref"A2", CellValue.Formula("A1+A1"))
    val result = sheet.evaluateFormula("=NPV(0,A1:A2)", Clock.system, rng)
    assertEquals(result, Right(CellValue.Number(BigDecimal("0.75"))))
    assertEquals(rng.calls, 1)
  }

  test("GH-499: recursive range formulas terminate with an explicit cycle failure") {
    val sheet = Sheet("S").put(ref"A1", CellValue.Formula("NPV(0,A1:A1)"))
    assert(sheet.evaluateFormula("=NPV(0,A1:A1)").isLeft)
  }
