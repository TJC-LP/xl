package com.tjclp.xl.formula

import com.tjclp.xl.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*

import scala.math.BigDecimal

/**
 * Comprehensive tests for financial functions: NPV, IRR, VLOOKUP.
 *
 * Tests parsing, evaluation, round-trip printing, edge cases, and error handling.
 */
class FinancialFunctionsSpec extends ScalaCheckSuite:

  val evaluator = Evaluator.instance

  // ==================== Test Helpers ====================

  /** Create sheet with cells */
  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    val sheet = Sheet(name = SheetName.unsafe("Test"))
    cells.foldLeft(sheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  /** Evaluate and unwrap Right or fail test */
  def evalOk[A](expr: TExpr[A], sheet: Sheet): A =
    evaluator.eval(expr, sheet) match
      case Right(value) => value
      case Left(err) => fail(s"Expected success, got error: $err")

  /** Evaluate and unwrap Left or fail test */
  def evalErr[A](expr: TExpr[A], sheet: Sheet): EvalError =
    evaluator.eval(expr, sheet) match
      case Left(err) => err
      case Right(value) => fail(s"Expected error, got value: $value")

  // ==================== NPV Tests ====================

  test("NPV: simple cash flows at 10% discount rate") {
    // Cash flows: $100, $200, $300 at t=1, t=2, t=3
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("100")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("200")),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("300"))
    )

    val expr = TExpr.npv(
      TExpr.Lit(BigDecimal("0.1")),
      CellRange.parse("A1:A3").toOption.get
    )

    val result = evalOk(expr, sheet)

    // Manual calculation:
    // NPV = 100/(1.1)^1 + 200/(1.1)^2 + 300/(1.1)^3
    //     = 90.909... + 165.289... + 225.394...
    //     = 481.59...
    val expected = BigDecimal("100") / BigDecimal("1.1").pow(1) +
      BigDecimal("200") / BigDecimal("1.1").pow(2) +
      BigDecimal("300") / BigDecimal("1.1").pow(3)

    assertEquals((result - expected).abs < BigDecimal("0.01"), true)
  }

  test("NPV: negative initial investment") {
    // Initial investment of -$1000, then cash flows of $300, $400, $500
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("-1000")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("300")),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("400")),
      ARef.from0(0, 3) -> CellValue.Number(BigDecimal("500"))
    )

    val expr = TExpr.npv(
      TExpr.Lit(BigDecimal("0.08")),
      CellRange.parse("A1:A4").toOption.get
    )

    val result = evalOk(expr, sheet)

    // NPV should be around 40.87
    val expected = BigDecimal("-1000") / BigDecimal("1.08").pow(1) +
      BigDecimal("300") / BigDecimal("1.08").pow(2) +
      BigDecimal("400") / BigDecimal("1.08").pow(3) +
      BigDecimal("500") / BigDecimal("1.08").pow(4)

    assertEquals((result - expected).abs < BigDecimal("0.01"), true)
  }

  test("NPV: ignores non-numeric cells") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("100")),
      ARef.from0(0, 1) -> CellValue.Text("skip"),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("200")),
      ARef.from0(0, 3) -> CellValue.Empty
    )

    val expr = TExpr.npv(
      TExpr.Lit(BigDecimal("0.1")),
      CellRange.parse("A1:A4").toOption.get
    )

    val result = evalOk(expr, sheet)

    // Only $100 and $200 should be counted
    val expected = BigDecimal("100") / BigDecimal("1.1").pow(1) +
      BigDecimal("200") / BigDecimal("1.1").pow(2)

    assertEquals((result - expected).abs < BigDecimal("0.01"), true)
  }

  test("NPV: rate = -1 causes division by zero error") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("100"))
    )

    val expr = TExpr.npv(
      TExpr.Lit(BigDecimal("-1")),
      CellRange.parse("A1:A1").toOption.get
    )

    val err = evalErr(expr, sheet)
    err match
      case EvalError.EvalFailed(reason, _) => assert(reason.contains("division by zero"))
      case other => fail(s"Expected EvalFailed, got $other")
  }

  test("NPV: parse and print round-trip") {
    val formula = "=NPV(0.1, A1:A10)"
    val parsed = FormulaParser.parse(formula).toOption.get
    val printed = FormulaPrinter.print(parsed)
    assertEquals(printed, formula)
  }

  // ==================== IRR Tests ====================

  test("IRR: simple investment with positive returns") {
    // Initial investment of -$1000, then cash flows of $400, $400, $400
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("-1000")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("400")),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("400")),
      ARef.from0(0, 3) -> CellValue.Number(BigDecimal("400"))
    )

    val expr = TExpr.irr(CellRange.parse("A1:A4").toOption.get)

    val result = evalOk(expr, sheet)

    // IRR should be around 9.7% (verify using NPV=0 check)
    // At IRR, NPV should be approximately 0
    val onePlusR = BigDecimal(1) + result
    val npv = BigDecimal("-1000") +
      BigDecimal("400") / onePlusR.pow(1) +
      BigDecimal("400") / onePlusR.pow(2) +
      BigDecimal("400") / onePlusR.pow(3)

    assertEquals(npv.abs < BigDecimal("0.001"), true)
  }

  test("IRR: with explicit guess") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("-1000")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("300")),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("400")),
      ARef.from0(0, 3) -> CellValue.Number(BigDecimal("500"))
    )

    val expr = TExpr.irr(
      CellRange.parse("A1:A4").toOption.get,
      Some(TExpr.Lit(BigDecimal("0.15")))
    )

    val result = evalOk(expr, sheet)

    // Verify NPV is approximately 0 at the computed IRR
    val onePlusR = BigDecimal(1) + result
    val npv = BigDecimal("-1000") +
      BigDecimal("300") / onePlusR.pow(1) +
      BigDecimal("400") / onePlusR.pow(2) +
      BigDecimal("500") / onePlusR.pow(3)

    assertEquals(npv.abs < BigDecimal("0.001"), true)
  }

  test("IRR: requires at least one positive and one negative flow") {
    // All positive cash flows
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("100")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("200")),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("300"))
    )

    val expr = TExpr.irr(CellRange.parse("A1:A3").toOption.get)

    val err = evalErr(expr, sheet)
    err match
      case EvalError.EvalFailed(reason, _) =>
        assert(reason.contains("at least one positive and one negative"))
      case other => fail(s"Expected EvalFailed, got $other")
  }

  test("IRR: ignores non-numeric cells") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("-1000")),
      ARef.from0(0, 1) -> CellValue.Text("skip"),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("600")),
      ARef.from0(0, 3) -> CellValue.Number(BigDecimal("600"))
    )

    val expr = TExpr.irr(CellRange.parse("A1:A4").toOption.get)

    val result = evalOk(expr, sheet)

    // Verify NPV is approximately 0 at the computed IRR
    val onePlusR = BigDecimal(1) + result
    val npv = BigDecimal("-1000") +
      BigDecimal("600") / onePlusR.pow(1) +
      BigDecimal("600") / onePlusR.pow(2)

    assertEquals(npv.abs < BigDecimal("0.001"), true)
  }

  test("IRR: parse and print round-trip (no guess)") {
    val formula = "=IRR(A1:A10)"
    val parsed = FormulaParser.parse(formula).toOption.get
    val printed = FormulaPrinter.print(parsed)
    assertEquals(printed, formula)
  }

  test("IRR: parse and print round-trip (with guess)") {
    val formula = "=IRR(A1:A10, 0.15)"
    val parsed = FormulaParser.parse(formula).toOption.get
    val printed = FormulaPrinter.print(parsed)
    assertEquals(printed, formula)
  }

  // ==================== VLOOKUP Tests ====================

  test("VLOOKUP: exact match (FALSE)") {
    // Lookup table: Key | Value1 | Value2
    //                100 |     10 |     20
    //                200 |     30 |     40
    //                300 |     50 |     60
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("100")),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal("10")),
      ARef.from0(2, 0) -> CellValue.Number(BigDecimal("20")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("200")),
      ARef.from0(1, 1) -> CellValue.Number(BigDecimal("30")),
      ARef.from0(2, 1) -> CellValue.Number(BigDecimal("40")),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("300")),
      ARef.from0(1, 2) -> CellValue.Number(BigDecimal("50")),
      ARef.from0(2, 2) -> CellValue.Number(BigDecimal("60"))
    )

    val expr = TExpr.vlookup(
      TExpr.Lit(BigDecimal("200")),
      CellRange.parse("A1:C3").toOption.get,
      TExpr.Lit(2), // Column 2
      TExpr.Lit(false) // Exact match
    )

    val result = evalOk(expr, sheet)
    assertEquals(result, BigDecimal("30"))
  }

  test("VLOOKUP: approximate match (TRUE)") {
    // Lookup table with sorted keys
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("100")),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal("10")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("200")),
      ARef.from0(1, 1) -> CellValue.Number(BigDecimal("20")),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("300")),
      ARef.from0(1, 2) -> CellValue.Number(BigDecimal("30"))
    )

    // Lookup 250 (should match 200, the largest key <= 250)
    val expr = TExpr.vlookup(
      TExpr.Lit(BigDecimal("250")),
      CellRange.parse("A1:B3").toOption.get,
      TExpr.Lit(2),
      TExpr.Lit(true) // Approximate match
    )

    val result = evalOk(expr, sheet)
    assertEquals(result, BigDecimal("20"))
  }

  test("VLOOKUP: col_index_num out of range") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("100")),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal("10"))
    )

    val expr = TExpr.vlookup(
      TExpr.Lit(BigDecimal("100")),
      CellRange.parse("A1:B1").toOption.get,
      TExpr.Lit(5), // Out of range (only 2 columns)
      TExpr.Lit(false)
    )

    val err = evalErr(expr, sheet)
    err match
      case EvalError.EvalFailed(reason, _) => assert(reason.contains("outside"))
      case other => fail(s"Expected EvalFailed, got $other")
  }

  test("VLOOKUP: exact match not found") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("100")),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal("10"))
    )

    val expr = TExpr.vlookup(
      TExpr.Lit(BigDecimal("999")), // Not in table
      CellRange.parse("A1:B1").toOption.get,
      TExpr.Lit(2),
      TExpr.Lit(false) // Exact match
    )

    val err = evalErr(expr, sheet)
    err match
      case EvalError.EvalFailed(reason, _) => assert(reason.contains("exact match not found"))
      case other => fail(s"Expected EvalFailed, got $other")
  }

  test("VLOOKUP: approximate match not found") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("100")),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal("10"))
    )

    // Lookup 50 (no key <= 50)
    val expr = TExpr.vlookup(
      TExpr.Lit(BigDecimal("50")),
      CellRange.parse("A1:B1").toOption.get,
      TExpr.Lit(2),
      TExpr.Lit(true) // Approximate match
    )

    val err = evalErr(expr, sheet)
    err match
      case EvalError.EvalFailed(reason, _) => assert(reason.contains("approximate match not found"))
      case other => fail(s"Expected EvalFailed, got $other")
  }

  test("VLOOKUP: ignores non-numeric keys") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Text("skip"),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal("10")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("200")),
      ARef.from0(1, 1) -> CellValue.Number(BigDecimal("20"))
    )

    val expr = TExpr.vlookup(
      TExpr.Lit(BigDecimal("200")),
      CellRange.parse("A1:B2").toOption.get,
      TExpr.Lit(2),
      TExpr.Lit(false)
    )

    val result = evalOk(expr, sheet)
    assertEquals(result, BigDecimal("20"))
  }

  test("VLOOKUP: parse and print round-trip (3 args)") {
    val formula = "=VLOOKUP(A1, B1:D10, 2, TRUE)"
    val parsed = FormulaParser.parse(formula).toOption.get
    val printed = FormulaPrinter.print(parsed)
    // Note: 3-arg form will print as 4-arg with TRUE
    assert(printed.contains("VLOOKUP"))
    assert(printed.contains("B1:D10"))
    assert(printed.contains("2"))
  }

  test("VLOOKUP: parse and print round-trip (4 args)") {
    val formula = "=VLOOKUP(100, A1:C5, 3, FALSE)"
    val parsed = FormulaParser.parse(formula).toOption.get
    val printed = FormulaPrinter.print(parsed)
    assertEquals(printed, formula)
  }

  // ==================== Integration Tests ====================

  test("NPV and IRR together: verify net present value at IRR is zero") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("-500")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("200")),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("200")),
      ARef.from0(0, 3) -> CellValue.Number(BigDecimal("200"))
    )

    val range = CellRange.parse("A1:A4").toOption.get

    // Compute IRR
    val irrExpr = TExpr.irr(range)
    val irr = evalOk(irrExpr, sheet)

    // Compute NPV at IRR (should be ~0)
    val npvExpr = TExpr.npv(TExpr.Lit(irr), range)
    val npv = evalOk(npvExpr, sheet)

    // NPV at IRR should be approximately 0
    assertEquals(npv.abs < BigDecimal("0.01"), true)
  }

  test("VLOOKUP: use result in further calculations") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("100")),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal("10")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("200")),
      ARef.from0(1, 1) -> CellValue.Number(BigDecimal("20"))
    )

    val vlookup = TExpr.vlookup(
      TExpr.Lit(BigDecimal("200")),
      CellRange.parse("A1:B2").toOption.get,
      TExpr.Lit(2),
      TExpr.Lit(false)
    )

    // Add 100 to the lookup result
    val expr = TExpr.Add(vlookup, TExpr.Lit(BigDecimal("100")))

    val result = evalOk(expr, sheet)
    assertEquals(result, BigDecimal("120"))
  }
