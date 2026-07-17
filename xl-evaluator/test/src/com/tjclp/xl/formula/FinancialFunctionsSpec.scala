package com.tjclp.xl.formula

import com.tjclp.xl.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*

import java.time.LocalDate
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

  // ==================== GH-388: Newton Divergence Containment ====================

  // Issue repro: a thin one-entry/one-exit strip whose exit approaches zero. Newton
  // overshoots below -1 and diverges doubly-exponentially until BigDecimal's int scale
  // overflows (java.lang.ArithmeticException from checkScale). Totality requires the same
  // contained EvalFailed as the documented non-convergence path — never a thrown exception.
  private def divergingStrip: Sheet = sheetWith(
    ARef.from0(0, 0) -> CellValue.Number(BigDecimal("-228")),
    ARef.from0(0, 1) -> CellValue.Number(BigDecimal("0")),
    ARef.from0(0, 2) -> CellValue.Number(BigDecimal("0")),
    ARef.from0(0, 3) -> CellValue.Number(BigDecimal("0")),
    ARef.from0(0, 4) -> CellValue.Number(BigDecimal("0")),
    ARef.from0(0, 5) -> CellValue.Number(BigDecimal("0.0001"))
  )

  test("IRR: Newton divergence returns contained error, never throws (GH-388)") {
    val expr = TExpr.irr(CellRange.parse("A1:A6").getOrElse(fail("Invalid range")))
    evalErr(expr, divergingStrip) match
      case EvalError.EvalFailed(reason, _) =>
        assert(reason.contains("IRR"), s"unexpected reason: $reason")
      case other => fail(s"Expected EvalFailed, got $other")
  }

  test("IRR: Newton divergence inside recalculate() is a per-cell error, never throws (GH-388)") {
    val s = divergingStrip
      .put(ARef.from0(1, 0), CellValue.Formula("=IRR(A1:A6)", None))
      .put(ARef.from0(1, 1), CellValue.Formula("=SUM(A1:A6)", None))
    // Documented contract (see also LetFunctionSpec): recalculate is total — per-cell
    // error reporting, never a thrown exception.
    val result = Workbook(Vector(s)).recalculate()
    assert(
      result.errors.exists(e => e.ref == ARef.from0(1, 0)),
      s"expected a per-cell error for B1, got: ${result.errors}"
    )
    val evaluated = result.evaluated.getOrElse(SheetName.unsafe("Test"), Map.empty)
    assert(evaluated.get(ARef.from0(1, 0)).isEmpty, "diverging IRR must not cache a value")
    assertEquals(
      evaluated.get(ARef.from0(1, 1)),
      Some(CellValue.Number(BigDecimal("-227.9999")))
    )
  }

  // ==================== XNPV / XIRR Evaluation Tests ====================

  /** Two cash flows anchored at 2020-01-01 with a parameterized exit flow and date. */
  private def xSchedule(entry: BigDecimal, exit: BigDecimal, exitDate: LocalDate): Sheet =
    sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(entry),
      ARef.from0(0, 1) -> CellValue.Number(exit),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDate.of(2020, 1, 1).atStartOfDay),
      ARef.from0(1, 1) -> CellValue.DateTime(exitDate.atStartOfDay)
    )

  test("XNPV: discounts irregular cash flows by actual/365 year fractions") {
    val expr = TExpr.xnpv(
      TExpr.Lit(BigDecimal("0.1")),
      CellRange.parse("A1:A2").getOrElse(fail("Invalid range")),
      CellRange.parse("B1:B2").getOrElse(fail("Invalid range"))
    )
    val sheet = xSchedule(BigDecimal("-1000"), BigDecimal("1200"), LocalDate.of(2021, 1, 1))
    val result = evalOk(expr, sheet)
    // -1000 + 1200 / 1.1^(366/365) — 2020 is a leap year, 366 days between the dates
    val expected = -1000.0 + 1200.0 / math.pow(1.1, 366.0 / 365.0)
    assert((result.toDouble - expected).abs < 0.01, s"got $result, expected ~$expected")
  }

  test("XIRR: two-flow investment rate makes XNPV zero") {
    val expr = TExpr.xirr(
      CellRange.parse("A1:A2").getOrElse(fail("Invalid range")),
      CellRange.parse("B1:B2").getOrElse(fail("Invalid range"))
    )
    val sheet = xSchedule(BigDecimal("-1000"), BigDecimal("1100"), LocalDate.of(2021, 1, 1))
    val result = evalOk(expr, sheet)
    val npvAtResult = -1000.0 + 1100.0 / math.pow(1.0 + result.toDouble, 366.0 / 365.0)
    assert(
      math.abs(npvAtResult) < 0.001,
      s"XNPV at computed XIRR should be ~0, got $npvAtResult (rate: $result)"
    )
  }

  test("XIRR: Newton divergence returns contained error, never throws (GH-388)") {
    // Same shape as the IRR repro: near-zero exit five years out. Newton overshoots to a
    // hugely negative rate; math.pow(negative base, fractional exponent) yields NaN, and
    // BigDecimal(NaN) must not escape as a NumberFormatException.
    val expr = TExpr.xirr(
      CellRange.parse("A1:A2").getOrElse(fail("Invalid range")),
      CellRange.parse("B1:B2").getOrElse(fail("Invalid range"))
    )
    val sheet = xSchedule(BigDecimal("-228"), BigDecimal("0.0001"), LocalDate.of(2025, 1, 1))
    evalErr(expr, sheet) match
      case EvalError.EvalFailed(reason, _) =>
        assert(reason.contains("XIRR"), s"unexpected reason: $reason")
      case other => fail(s"Expected EvalFailed, got $other")
  }

  test("XNPV: non-finite discount factor is a contained error, never throws (GH-388)") {
    // Rate 1e300 over a ~5-year fraction: math.pow(1e300, ~5) = Infinity, and
    // BigDecimal(Infinity) must not escape as a NumberFormatException.
    val expr = TExpr.xnpv(
      TExpr.Lit(BigDecimal("1e300")),
      CellRange.parse("A1:A2").getOrElse(fail("Invalid range")),
      CellRange.parse("B1:B2").getOrElse(fail("Invalid range"))
    )
    val sheet = xSchedule(BigDecimal("-1000"), BigDecimal("1200"), LocalDate.of(2025, 1, 1))
    evalErr(expr, sheet) match
      case EvalError.EvalFailed(reason, _) =>
        assert(reason.contains("XNPV"), s"unexpected reason: $reason")
      case other => fail(s"Expected EvalFailed, got $other")
  }

  // ==================== GH-405: serial-number dates in XIRR/XNPV ranges ====================

  /** Excel serial for a date, exactly as OOXML stores (and XlsxReader re-reads) it. */
  private def dateSerial(d: LocalDate): BigDecimal =
    BigDecimal(CellValue.dateTimeToExcelSerial(d.atStartOfDay))

  test("GH-405: XIRR accepts a MIXED serial/DateTime date range (issue repro)") {
    // Post-round-trip state: authored DateTime anchors re-read as serial Numbers while
    // EDATE-recomputed dependents are DateTime — the range decode must coerce uniformly
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("-1000")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("500")),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("700")),
      ARef.from0(1, 0) -> CellValue.Number(dateSerial(LocalDate.of(2020, 1, 1))),
      ARef.from0(1, 1) -> CellValue.DateTime(LocalDate.of(2020, 7, 1).atStartOfDay),
      ARef.from0(1, 2) -> CellValue.Number(dateSerial(LocalDate.of(2021, 1, 1)))
    )
    val expr = TExpr.xirr(
      CellRange.parse("A1:A3").getOrElse(fail("Invalid range")),
      CellRange.parse("B1:B3").getOrElse(fail("Invalid range"))
    )
    val rate = evalOk(expr, sheet)
    // XIRR definition: XNPV at the computed rate is ~0 (2020-01-01 → 2020-07-01 is 182
    // days, → 2021-01-01 is 366 — leap year)
    val npvAtRate = -1000.0 +
      500.0 / math.pow(1.0 + rate.toDouble, 182.0 / 365.0) +
      700.0 / math.pow(1.0 + rate.toDouble, 366.0 / 365.0)
    assert(math.abs(npvAtRate) < 0.001, s"XNPV at computed XIRR should be ~0, got $npvAtRate")
  }

  test("GH-405: XNPV accepts a MIXED serial/DateTime date range") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("-1000")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("500")),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("700")),
      ARef.from0(1, 0) -> CellValue.Number(dateSerial(LocalDate.of(2020, 1, 1))),
      ARef.from0(1, 1) -> CellValue.DateTime(LocalDate.of(2020, 7, 1).atStartOfDay),
      ARef.from0(1, 2) -> CellValue.Number(dateSerial(LocalDate.of(2021, 1, 1)))
    )
    val expr = TExpr.xnpv(
      TExpr.Lit(BigDecimal("0.1")),
      CellRange.parse("A1:A3").getOrElse(fail("Invalid range")),
      CellRange.parse("B1:B3").getOrElse(fail("Invalid range"))
    )
    val result = evalOk(expr, sheet)
    val expected = -1000.0 +
      500.0 / math.pow(1.1, 182.0 / 365.0) +
      700.0 / math.pow(1.1, 366.0 / 365.0)
    assert((result.toDouble - expected).abs < 0.01, s"got $result, expected ~$expected")
  }

  test("GH-405: sparse XIRR block — aligned blank (value, date) pairs still drop together") {
    // Range folds keep blank-SKIP semantics: the scalar Empty -> 1900-01-01 arm (GH-396) must
    // not turn a blank date cell into a phantom date, or a sparse block's values/dates counts
    // would diverge (values skip blanks) and the whole formula would error
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("-1000")),
      // A2/B2 deliberately blank — an unfilled row inside the block
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal("1100")),
      ARef.from0(1, 0) -> CellValue.Number(dateSerial(LocalDate.of(2020, 1, 1))),
      ARef.from0(1, 2) -> CellValue.DateTime(LocalDate.of(2021, 1, 1).atStartOfDay)
    )
    val expr = TExpr.xirr(
      CellRange.parse("A1:A3").getOrElse(fail("Invalid range")),
      CellRange.parse("B1:B3").getOrElse(fail("Invalid range"))
    )
    val rate = evalOk(expr, sheet)
    assert((rate.toDouble - 0.1).abs < 0.01, s"expected ~0.1, got $rate")
  }

  test("GH-405: out-of-range serials in a date range still drop (MaxExcelDateSerial guard)") {
    // A negative serial is not a date; the coercion guard (0..MaxExcelDateSerial) keeps it
    // out, so the length mismatch surfaces as a clean error rather than a bogus 1899 date
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("-1000")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("1100")),
      ARef.from0(1, 0) -> CellValue.Number(dateSerial(LocalDate.of(2020, 1, 1))),
      ARef.from0(1, 1) -> CellValue.Number(BigDecimal(-5))
    )
    val expr = TExpr.xirr(
      CellRange.parse("A1:A2").getOrElse(fail("Invalid range")),
      CellRange.parse("B1:B2").getOrElse(fail("Invalid range"))
    )
    val err = evalErr(expr, sheet)
    assert(err.toString.contains("XIRR"), s"expected an XIRR-scoped error, got $err")
  }

  test("GH-405: XIRR returns block survives the OOXML value round trip (field chain)") {
    // Field chain (FinAgent QA): a date anchor authored as CellValue.DateTime WRITES as a
    // serial (correct xlsx) and RE-READS as CellValue.Number — typed DateTime is not
    // recoverable from OOXML. The EDATE dependent recomputes to DateTime, so post-round-trip
    // recalculation hands XIRR a mixed Number/DateTime date range. Simulate the reader's
    // exact output (every DateTime becomes its serial; numFmt styling is invisible to the
    // evaluator), then recalculate and assert the returns block yields the rate — not the
    // house IF(ISERROR(...)) "NM " cache.
    val authored = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal("-1000")),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal("1100")),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDate.of(2020, 1, 1).atStartOfDay),
      ARef.from0(1, 1) -> CellValue.Formula("=EDATE(B1, 12)", None),
      ARef.from0(2, 0) -> CellValue.Formula("=XIRR(A1:A2, B1:B2)", None),
      ARef.from0(2, 1) -> CellValue.Formula(
        "=IF(ISERROR(XIRR(A1:A2, B1:B2)), \"NM \", XIRR(A1:A2, B1:B2))",
        None
      )
    )
    val reread = authored.cells.foldLeft(authored) { case (s, (ref, cell)) =>
      cell.value match
        case CellValue.DateTime(dt) =>
          s.put(ref, CellValue.Number(BigDecimal(CellValue.dateTimeToExcelSerial(dt))))
        case _ => s
    }
    val result = Workbook(Vector(reread)).recalculate()
    assert(
      result.errors.isEmpty,
      s"post-round-trip recalc must not error the returns block, got: ${result.errors}"
    )
    val evaluated = result.evaluated.getOrElse(SheetName.unsafe("Test"), Map.empty)
    evaluated.get(ARef.from0(2, 0)) match
      case Some(CellValue.Number(rate)) =>
        // -1000 → 1100 over ~1 year ⇒ ~10%
        assert((rate.toDouble - 0.1).abs < 0.01, s"expected ~0.1, got $rate")
      case other => fail(s"expected XIRR rate at C1, got $other")
    evaluated.get(ARef.from0(2, 1)) match
      case Some(CellValue.Number(_)) => () // the NM guard passes the rate through
      case other => fail(s"the IF(ISERROR(...)) guard must yield the rate, not $other")
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

  // ==================== GH-388: TVM Numeric Containment ====================

  /**
   * Expect a contained failure naming the function — the eval must never throw. GH-344: a
   * non-convergence now classifies as the #NUM! error VALUE (still contained, still Left on the
   * direct-eval channel); NumericGuard blowups stay EvalFailed.
   */
  private def assertContained(expr: TExpr[BigDecimal], fn: String): Unit =
    evalErr(expr, sheetWith()) match
      case EvalError.EvalFailed(reason, _) =>
        assert(reason.contains(fn), s"unexpected reason: $reason")
      case EvalError.ErrorValue(com.tjclp.xl.cells.CellError.Num, Some(ctx)) =>
        assert(ctx.contains(fn), s"unexpected context: $ctx")
      case other => fail(s"Expected contained failure, got $other")

  test("PMT: zero rate with zero periods is a contained error, never throws (GH-388)") {
    // rate ≈ 0, nper = 0 hits the BigDecimal(Double.NaN) branch
    val expr = TExpr.pmt(
      TExpr.Lit(BigDecimal(0)),
      TExpr.Lit(BigDecimal(0)),
      TExpr.Lit(BigDecimal(10000))
    )
    assertContained(expr, "PMT")
  }

  test("PMT: non-finite intermediate (huge rate and nper) is a contained error (GH-388)") {
    // pow(1e300, 1e300) = Infinity, -Inf/Inf = NaN → BigDecimal(NaN) must be contained
    val expr = TExpr.pmt(
      TExpr.Lit(BigDecimal("1e300")),
      TExpr.Lit(BigDecimal("1e300")),
      TExpr.Lit(BigDecimal(10000))
    )
    assertContained(expr, "PMT")
  }

  test("FV: non-finite intermediate (huge rate and nper) is a contained error (GH-388)") {
    val expr = TExpr.fv(
      TExpr.Lit(BigDecimal("1e300")),
      TExpr.Lit(BigDecimal("1e300")),
      TExpr.Lit(BigDecimal(100))
    )
    assertContained(expr, "FV")
  }

  test("PV: non-finite intermediate (huge rate and nper) is a contained error (GH-388)") {
    val expr = TExpr.pv(
      TExpr.Lit(BigDecimal("1e300")),
      TExpr.Lit(BigDecimal("1e300")),
      TExpr.Lit(BigDecimal(100))
    )
    assertContained(expr, "PV")
  }

  test("NPER: zero rate with zero payment is a contained error, never throws (GH-388)") {
    val expr = TExpr.nper(
      TExpr.Lit(BigDecimal(0)),
      TExpr.Lit(BigDecimal(0)),
      TExpr.Lit(BigDecimal(10000))
    )
    assertContained(expr, "NPER")
  }

  test("RATE: NaN-producing inputs yield a contained non-convergence error (GH-388)") {
    // pow(1.1, 1e10) = Infinity → f and df both NaN; the Newton loop must exhaust its
    // iterations with a contained EvalFailed rather than surfacing NaN or throwing.
    val expr = TExpr.rate(
      TExpr.Lit(BigDecimal("1e10")),
      TExpr.Lit(BigDecimal(-500)),
      TExpr.Lit(BigDecimal(10000))
    )
    assertContained(expr, "RATE")
  }
