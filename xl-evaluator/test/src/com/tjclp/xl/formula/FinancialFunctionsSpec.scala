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
    val sheet = new Sheet(name = SheetName.unsafe("Test"))
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
      CellRange.parse("A1:A3").getOrElse(fail("Invalid range"))
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
      CellRange.parse("A1:A4").getOrElse(fail("Invalid range"))
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
      CellRange.parse("A1:A4").getOrElse(fail("Invalid range"))
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
      CellRange.parse("A1:A1").getOrElse(fail("Invalid range"))
    )

    val err = evalErr(expr, sheet)
    err match
      case EvalError.EvalFailed(reason, _) => assert(reason.contains("division by zero"))
      case other => fail(s"Expected EvalFailed, got $other")
  }

  test("NPV: parse and print round-trip") {
    val formula = "=NPV(0.1, A1:A10)"
    val parsed = FormulaParser.parse(formula).getOrElse(fail("Parse failed"))
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

    val expr = TExpr.irr(CellRange.parse("A1:A4").getOrElse(fail("Invalid range")))

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
      CellRange.parse("A1:A4").getOrElse(fail("Invalid range")),
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

    val expr = TExpr.irr(CellRange.parse("A1:A3").getOrElse(fail("Invalid range")))

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

    val expr = TExpr.irr(CellRange.parse("A1:A4").getOrElse(fail("Invalid range")))

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
    val parsed = FormulaParser.parse(formula).getOrElse(fail("Parse failed"))
    val printed = FormulaPrinter.print(parsed)
    assertEquals(printed, formula)
  }

  test("IRR: parse and print round-trip (with guess)") {
    val formula = "=IRR(A1:A10, 0.15)"
    val parsed = FormulaParser.parse(formula).getOrElse(fail("Parse failed"))
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
      CellRange.parse("A1:C3").getOrElse(fail("Invalid range")),
      TExpr.Lit(2), // Column 2
      TExpr.Lit(false) // Exact match
    )

    val result = evalOk(expr, sheet)
    assertEquals(result, CellValue.Number(BigDecimal("30")))
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
      CellRange.parse("A1:B3").getOrElse(fail("Invalid range")),
      TExpr.Lit(2),
      TExpr.Lit(true) // Approximate match
    )

    val result = evalOk(expr, sheet)
    assertEquals(result, CellValue.Number(BigDecimal("20")))
  }

  test("VLOOKUP: col_index_num out of range") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("100")),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal("10"))
    )

    val expr = TExpr.vlookup(
      TExpr.Lit(BigDecimal("100")),
      CellRange.parse("A1:B1").getOrElse(fail("Invalid range")),
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
      CellRange.parse("A1:B1").getOrElse(fail("Invalid range")),
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
      CellRange.parse("A1:B1").getOrElse(fail("Invalid range")),
      TExpr.Lit(2),
      TExpr.Lit(true) // Approximate match
    )

    val err = evalErr(expr, sheet)
    err match
      case EvalError.EvalFailed(reason, _) => assert(reason.contains("approximate match not found"))
      case other => fail(s"Expected EvalFailed, got $other")
  }

  test("VLOOKUP: ignores non-numeric keys for numeric lookup") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Text("skip"),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal("10")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("200")),
      ARef.from0(1, 1) -> CellValue.Number(BigDecimal("20"))
    )

    val expr = TExpr.vlookup(
      TExpr.Lit(BigDecimal("200")),
      CellRange.parse("A1:B2").getOrElse(fail("Invalid range")),
      TExpr.Lit(2),
      TExpr.Lit(false)
    )

    val result = evalOk(expr, sheet)
    assertEquals(result, CellValue.Number(BigDecimal("20")))
  }

  test("VLOOKUP: parse and print round-trip (3 args)") {
    val formula = "=VLOOKUP(A1, B1:D10, 2, TRUE)"
    val parsed = FormulaParser.parse(formula).getOrElse(fail("Parse failed"))
    val printed = FormulaPrinter.print(parsed)
    // Note: 3-arg form will print as 4-arg with TRUE
    assert(printed.contains("VLOOKUP"))
    assert(printed.contains("B1:D10"))
    assert(printed.contains("2"))
  }

  test("VLOOKUP: parse and print round-trip (4 args)") {
    val formula = "=VLOOKUP(100, A1:C5, 3, FALSE)"
    val parsed = FormulaParser.parse(formula).getOrElse(fail("Parse failed"))
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

    val range = CellRange.parse("A1:A4").getOrElse(fail("Invalid range"))

    // Compute IRR
    val irrExpr = TExpr.irr(range)
    val irr = evalOk(irrExpr, sheet)

    // Compute NPV at IRR (should be ~0)
    val npvExpr = TExpr.npv(TExpr.Lit(irr), range)
    val npv = evalOk(npvExpr, sheet)

    // NPV at IRR should be approximately 0
    assertEquals(npv.abs < BigDecimal("0.01"), true)
  }

  test("VLOOKUP: returns numeric CellValue for numeric data") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("100")),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal("10")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("200")),
      ARef.from0(1, 1) -> CellValue.Number(BigDecimal("20"))
    )

    val expr = TExpr.vlookup(
      TExpr.Lit(BigDecimal("200")),
      CellRange.parse("A1:B2").getOrElse(fail("Invalid range")),
      TExpr.Lit(2),
      TExpr.Lit(false)
    )

    val result = evalOk(expr, sheet)
    assertEquals(result, CellValue.Number(BigDecimal("20")))

    // Verify we can extract numeric value from CellValue
    result match
      case CellValue.Number(n) => assertEquals(n + 100, BigDecimal("120"))
      case other => fail(s"Expected Number, got $other")
  }

  test("VLOOKUP: text lookup (case-insensitive)") {
    // Table with text keys in first column
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Text("Widget A"),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal("100")),
      ARef.from0(0, 1) -> CellValue.Text("Widget B"),
      ARef.from0(1, 1) -> CellValue.Number(BigDecimal("200")),
      ARef.from0(0, 2) -> CellValue.Text("Widget C"),
      ARef.from0(1, 2) -> CellValue.Number(BigDecimal("300"))
    )

    // Lookup "widget b" (case-insensitive)
    val expr = TExpr.vlookup(
      TExpr.Lit("widget b"), // lowercase to test case-insensitivity
      CellRange.parse("A1:B3").getOrElse(fail("Invalid range")),
      TExpr.Lit(2),
      TExpr.Lit(false) // exact match
    )

    val result = evalOk(expr, sheet)
    assertEquals(result, CellValue.Number(BigDecimal("200")))
  }

  test("VLOOKUP: text lookup returns text result") {
    // Table with text keys and text results
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Text("ID001"),
      ARef.from0(1, 0) -> CellValue.Text("Product A"),
      ARef.from0(0, 1) -> CellValue.Text("ID002"),
      ARef.from0(1, 1) -> CellValue.Text("Product B")
    )

    val expr = TExpr.vlookup(
      TExpr.Lit("ID002"),
      CellRange.parse("A1:B2").getOrElse(fail("Invalid range")),
      TExpr.Lit(2),
      TExpr.Lit(false)
    )

    val result = evalOk(expr, sheet)
    assertEquals(result, CellValue.Text("Product B"))
  }

  // ==================== PMT Tests ====================

  /** Helper to evaluate numeric formulas */
  def evalNumeric(formula: String, sheet: Sheet = sheetWith()): Double =
    FormulaParser.parse(formula) match
      case Right(expr) =>
        evaluator.eval(expr, sheet) match
          case Right(v: BigDecimal) => v.toDouble
          case Right(other) => fail(s"Expected BigDecimal, got: $other")
          case Left(err) => fail(s"Eval error: $err")
      case Left(err) => fail(s"Parse error: $err")

  /** Helper to assert approximately equal with tolerance */
  def assertApprox(actual: Double, expected: Double, tolerance: Double = 0.01): Unit =
    assert(
      math.abs(actual - expected) < tolerance,
      s"Expected $expected ± $tolerance, got $actual"
    )

  test("PMT: basic loan payment (Excel example)") {
    // $10,000 loan, 5% annual rate (0.05/12 per month), 24 months
    // Excel: =PMT(0.05/12, 24, 10000) ≈ -438.71
    val result = evalNumeric("=PMT(0.05/12, 24, 10000)")
    assertApprox(result, -438.71, 0.01)
  }

  test("PMT: with future value parameter") {
    // Saving for $10,000 target, 6% annual rate (0.06/12 per month), 60 months, $0 pv
    // Excel: =PMT(0.06/12, 60, 0, 10000) ≈ -143.33
    val result = evalNumeric("=PMT(0.06/12, 60, 0, 10000)")
    assertApprox(result, -143.33, 0.01)
  }

  test("PMT: beginning of period (type=1)") {
    // Same loan, but payments at beginning of period
    // Excel: =PMT(0.05/12, 24, 10000, 0, 1) ≈ -436.89
    val result = evalNumeric("=PMT(0.05/12, 24, 10000, 0, 1)")
    assertApprox(result, -436.89, 0.01)
  }

  test("PMT: zero interest rate") {
    // No interest loan
    // PMT(0, 24, 10000) = -10000/24 ≈ -416.67
    val result = evalNumeric("=PMT(0, 24, 10000)")
    assertApprox(result, -416.67, 0.01)
  }

  test("PMT: parse and print round-trip") {
    val formula = "=PMT(0.05, 12, 10000)"
    val parsed = FormulaParser.parse(formula).getOrElse(fail("Parse failed"))
    val printed = FormulaPrinter.print(parsed)
    assertEquals(printed, formula)
  }

  // ==================== FV Tests ====================

  test("FV: future value of investment (Excel example)") {
    // $200/month deposits, 6% annual rate (0.06/12 per month), 60 months, $0 starting
    // Excel: =FV(0.06/12, 60, -200) ≈ 13954.01
    val result = evalNumeric("=FV(0.06/12, 60, -200)")
    assertApprox(result, 13954.01, 0.01)
  }

  test("FV: with present value parameter") {
    // $100/month deposits, 5% annual rate, 12 months, $1000 starting
    // Implementation gives ≈ 2279.05
    val result = evalNumeric("=FV(0.05/12, 12, -100, -1000)")
    assertApprox(result, 2279.05, 1.0)
  }

  test("FV: beginning of period (type=1)") {
    // Same investment, but deposits at beginning of period
    // Implementation gives ≈ 14023.78 (type=1 adjustment)
    val result = evalNumeric("=FV(0.06/12, 60, -200, 0, 1)")
    assertApprox(result, 14023.78, 1.0)
  }

  test("FV: zero interest rate") {
    // Simple accumulation
    // FV(0, 12, -100) = 100 * 12 = 1200
    val result = evalNumeric("=FV(0, 12, -100)")
    assertApprox(result, 1200.0, 0.01)
  }

  test("FV: parse and print round-trip") {
    val formula = "=FV(0.05, 12, -100, -1000, 1)"
    val parsed = FormulaParser.parse(formula).getOrElse(fail("Parse failed"))
    val printed = FormulaPrinter.print(parsed)
    assertEquals(printed, formula)
  }

  // ==================== PV Tests ====================

  test("PV: present value of annuity (Excel example)") {
    // $500/month payments, 5% annual rate (0.05/12 per month), 60 months
    // Implementation gives ≈ 26495.35
    val result = evalNumeric("=PV(0.05/12, 60, -500)")
    assertApprox(result, 26495.35, 1.0)
  }

  test("PV: with future value parameter") {
    // How much to invest now for $10,000 in 10 years at 5%?
    // PV(0.05, 10, 0, 10000) ≈ -6139.13
    val result = evalNumeric("=PV(0.05, 10, 0, 10000)")
    assertApprox(result, -6139.13, 0.01)
  }

  test("PV: beginning of period (type=1)") {
    // Same annuity, but payments at beginning of period
    // Implementation gives ≈ 26605.75
    val result = evalNumeric("=PV(0.05/12, 60, -500, 0, 1)")
    assertApprox(result, 26605.75, 1.0)
  }

  test("PV: zero interest rate") {
    // Simple summation
    // PV(0, 12, -100) = 100 * 12 = 1200
    val result = evalNumeric("=PV(0, 12, -100)")
    assertApprox(result, 1200.0, 0.01)
  }

  test("PV: parse and print round-trip") {
    val formula = "=PV(0.05, 12, -100)"
    val parsed = FormulaParser.parse(formula).getOrElse(fail("Parse failed"))
    val printed = FormulaPrinter.print(parsed)
    assertEquals(printed, formula)
  }

  // ==================== NPER Tests ====================

  test("NPER: how long to pay off loan (Excel example)") {
    // $10,000 loan, 8% annual rate, $200/month payments
    // Implementation gives ≈ 61.02 periods
    val result = evalNumeric("=NPER(0.08/12, -200, 10000)")
    assertApprox(result, 61.02, 0.1)
  }

  test("NPER: with future value parameter") {
    // Saving $100/month at 5% to reach $5,000
    // Implementation gives ≈ 45.51 periods
    val result = evalNumeric("=NPER(0.05/12, -100, 0, 5000)")
    assertApprox(result, 45.51, 0.1)
  }

  test("NPER: beginning of period (type=1)") {
    // Same loan, but payments at beginning of period
    // Implementation gives ≈ 60.52 periods
    val result = evalNumeric("=NPER(0.08/12, -200, 10000, 0, 1)")
    assertApprox(result, 60.52, 0.1)
  }

  test("NPER: zero interest rate") {
    // Simple division
    // NPER(0, -100, 1000) = 1000 / 100 = 10
    val result = evalNumeric("=NPER(0, -100, 1000)")
    assertApprox(result, 10.0, 0.01)
  }

  test("NPER: parse and print round-trip") {
    val formula = "=NPER(0.05, -100, 1000)"
    val parsed = FormulaParser.parse(formula).getOrElse(fail("Parse failed"))
    val printed = FormulaPrinter.print(parsed)
    assertEquals(printed, formula)
  }

  // ==================== RATE Tests ====================

  test("RATE: find interest rate for loan (Excel example)") {
    // 24 months, $500/month payments, $10,000 loan
    // Implementation gives ≈ 0.01513 (1.513% per month)
    val result = evalNumeric("=RATE(24, -500, 10000)")
    assertApprox(result, 0.01513, 0.0001)
  }

  test("RATE: with future value parameter") {
    // 60 months, $100/month deposits, $0 pv, $10,000 fv target
    // Implementation gives ≈ 0.01615 (1.615% per month)
    val result = evalNumeric("=RATE(60, -100, 0, 10000)")
    assertApprox(result, 0.01615, 0.0001)
  }

  test("RATE: with explicit guess") {
    // Same as first test, but with explicit guess
    val result = evalNumeric("=RATE(24, -500, 10000, 0, 0, 0.01)")
    assertApprox(result, 0.01513, 0.0001)
  }

  test("RATE: parse and print round-trip (3 args)") {
    val formula = "=RATE(24, -500, 10000)"
    val parsed = FormulaParser.parse(formula).getOrElse(fail("Parse failed"))
    val printed = FormulaPrinter.print(parsed)
    assertEquals(printed, formula)
  }

  test("RATE: parse and print round-trip (6 args)") {
    val formula = "=RATE(24, -500, 10000, 0, 0, 0.1)"
    val parsed = FormulaParser.parse(formula).getOrElse(fail("Parse failed"))
    val printed = FormulaPrinter.print(parsed)
    assertEquals(printed, formula)
  }

  // ==================== TVM Cross-Validation Tests ====================

  test("TVM: PMT and FV consistency - FV of payments equals target") {
    // If we know PMT for a target FV, then FV of those payments should equal target
    // PMT to reach $10,000 in 60 months at 0.5% per month
    val pmt = evalNumeric("=PMT(0.005, 60, 0, 10000)") // ≈ -143.33

    // FV of those payments
    val fvFormula = s"=FV(0.005, 60, $pmt, 0)"
    val fv = evalNumeric(fvFormula)
    assertApprox(fv, 10000.0, 0.1)
  }

  test("TVM: PV and FV consistency - PV grows to FV") {
    // If we know PV for a target FV, then FV of that PV should equal target
    // PV returns negative (investment outflow), FV returns negative (future value)
    val pv = evalNumeric("=PV(0.05, 10, 0, 10000)") // ≈ -6139.13

    // FV of that PV investment should equal the target (with sign inversion)
    val fvFormula = s"=FV(0.05, 10, 0, ${-pv})"
    val fv = evalNumeric(fvFormula)
    // FV returns -10000 (same sign convention: negative = outflow/investment result)
    assertApprox(math.abs(fv), 10000.0, 0.1)
  }

  test("TVM: NPER and PMT consistency - loan paid off in NPER periods") {
    // NPER to pay off $10,000 at 0.5% per month with $200 payments
    val nper = evalNumeric("=NPER(0.005, -200, 10000)") // ≈ 54.7

    // FV after NPER periods should be ≈ 0
    val fvFormula = s"=FV(0.005, $nper, -200, 10000)"
    val fv = evalNumeric(fvFormula)
    assertApprox(fv, 0.0, 1.0)
  }
