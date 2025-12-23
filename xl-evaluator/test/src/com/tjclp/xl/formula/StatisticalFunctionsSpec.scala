package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.addressing.SheetName
import munit.FunSuite

/**
 * Comprehensive tests for statistical functions: MEDIAN, STDEV, STDEVP, VAR, VARP.
 *
 * MEDIAN - Returns the middle value in a sorted list
 * STDEV  - Sample standard deviation (divides by n-1)
 * STDEVP - Population standard deviation (divides by n)
 * VAR    - Sample variance (divides by n-1)
 * VARP   - Population variance (divides by n)
 */
class StatisticalFunctionsSpec extends FunSuite:
  val emptySheet = new Sheet(name = SheetName.unsafe("Test"))
  val evaluator = Evaluator.instance

  /** Helper to create sheet with cells */
  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  /** Helper to assert approximate equality for floating point results */
  def assertApprox(result: Either[EvalError, BigDecimal], expected: Double, tolerance: Double = 0.001)(
      implicit loc: munit.Location
  ): Unit =
    result match
      case Right(value) =>
        assert(
          math.abs(value.toDouble - expected) < tolerance,
          s"Expected $expected ± $tolerance but got $value"
        )
      case Left(err) =>
        fail(s"Expected successful evaluation but got error: $err")

  // ===== MEDIAN Tests =====

  test("MEDIAN: odd count - returns middle value") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(2),
      ref"A2" -> CellValue.Number(4),
      ref"A3" -> CellValue.Number(6)
    )
    val range = CellRange.parse("A1:A3").toOption.get
    val expr = TExpr.Aggregate("MEDIAN", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(4)))
  }

  test("MEDIAN: even count - returns average of two middle values") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(1),
      ref"A2" -> CellValue.Number(2),
      ref"A3" -> CellValue.Number(3),
      ref"A4" -> CellValue.Number(4)
    )
    val range = CellRange.parse("A1:A4").toOption.get
    val expr = TExpr.Aggregate("MEDIAN", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(2.5)))
  }

  test("MEDIAN: single value") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(42))
    val range = CellRange.parse("A1:A1").toOption.get
    val expr = TExpr.Aggregate("MEDIAN", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(42)))
  }

  test("MEDIAN: unsorted values - returns correct median") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(9),
      ref"A2" -> CellValue.Number(2),
      ref"A3" -> CellValue.Number(7),
      ref"A4" -> CellValue.Number(4),
      ref"A5" -> CellValue.Number(5)
    )
    val range = CellRange.parse("A1:A5").toOption.get
    val expr = TExpr.Aggregate("MEDIAN", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    // Sorted: 2, 4, 5, 7, 9 -> median = 5
    assertEquals(result, Right(BigDecimal(5)))
  }

  test("MEDIAN: empty range returns error") {
    val sheet = emptySheet
    val range = CellRange.parse("A1:A5").toOption.get
    val expr = TExpr.Aggregate("MEDIAN", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assert(result.isLeft, "Expected error for empty range")
  }

  test("MEDIAN: skips non-numeric cells") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(1),
      ref"A2" -> CellValue.Text("text"),
      ref"A3" -> CellValue.Number(3),
      ref"A4" -> CellValue.Number(5)
    )
    val range = CellRange.parse("A1:A4").toOption.get
    val expr = TExpr.Aggregate("MEDIAN", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    // Only numeric: 1, 3, 5 -> median = 3
    assertEquals(result, Right(BigDecimal(3)))
  }

  // ===== STDEV Tests (Sample Standard Deviation) =====

  test("STDEV: standard case") {
    // Data: 2, 4, 4, 4, 5, 5, 7, 9
    // Mean: 5, Sample variance: 4.571..., Sample stdev: 2.138...
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(2),
      ref"A2" -> CellValue.Number(4),
      ref"A3" -> CellValue.Number(4),
      ref"A4" -> CellValue.Number(4),
      ref"A5" -> CellValue.Number(5),
      ref"A6" -> CellValue.Number(5),
      ref"A7" -> CellValue.Number(7),
      ref"A8" -> CellValue.Number(9)
    )
    val range = CellRange.parse("A1:A8").toOption.get
    val expr = TExpr.Aggregate("STDEV", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertApprox(result, 2.138, 0.001)
  }

  test("STDEV: two values") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(10),
      ref"A2" -> CellValue.Number(20)
    )
    val range = CellRange.parse("A1:A2").toOption.get
    val expr = TExpr.Aggregate("STDEV", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    // Mean: 15, Variance: (25+25)/1 = 50, Stdev: 7.071...
    assertApprox(result, 7.071, 0.001)
  }

  test("STDEV: single value returns error (requires n >= 2)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(42))
    val range = CellRange.parse("A1:A1").toOption.get
    val expr = TExpr.Aggregate("STDEV", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assert(result.isLeft, "Expected error for single value")
  }

  test("STDEV: empty range returns error") {
    val sheet = emptySheet
    val range = CellRange.parse("A1:A5").toOption.get
    val expr = TExpr.Aggregate("STDEV", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assert(result.isLeft, "Expected error for empty range")
  }

  // ===== STDEVP Tests (Population Standard Deviation) =====

  test("STDEVP: standard case") {
    // Data: 2, 4, 4, 4, 5, 5, 7, 9
    // Mean: 5, Population variance: 4, Population stdev: 2
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(2),
      ref"A2" -> CellValue.Number(4),
      ref"A3" -> CellValue.Number(4),
      ref"A4" -> CellValue.Number(4),
      ref"A5" -> CellValue.Number(5),
      ref"A6" -> CellValue.Number(5),
      ref"A7" -> CellValue.Number(7),
      ref"A8" -> CellValue.Number(9)
    )
    val range = CellRange.parse("A1:A8").toOption.get
    val expr = TExpr.Aggregate("STDEVP", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertApprox(result, 2.0, 0.001)
  }

  test("STDEVP: single value returns 0") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(42))
    val range = CellRange.parse("A1:A1").toOption.get
    val expr = TExpr.Aggregate("STDEVP", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(0)))
  }

  test("STDEVP: empty range returns error") {
    val sheet = emptySheet
    val range = CellRange.parse("A1:A5").toOption.get
    val expr = TExpr.Aggregate("STDEVP", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assert(result.isLeft, "Expected error for empty range")
  }

  // ===== VAR Tests (Sample Variance) =====

  test("VAR: standard case") {
    // Data: 2, 4, 4, 4, 5, 5, 7, 9
    // Mean: 5, Sample variance: 32/7 = 4.571...
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(2),
      ref"A2" -> CellValue.Number(4),
      ref"A3" -> CellValue.Number(4),
      ref"A4" -> CellValue.Number(4),
      ref"A5" -> CellValue.Number(5),
      ref"A6" -> CellValue.Number(5),
      ref"A7" -> CellValue.Number(7),
      ref"A8" -> CellValue.Number(9)
    )
    val range = CellRange.parse("A1:A8").toOption.get
    val expr = TExpr.Aggregate("VAR", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertApprox(result, 4.571, 0.001)
  }

  test("VAR: two values") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(10),
      ref"A2" -> CellValue.Number(20)
    )
    val range = CellRange.parse("A1:A2").toOption.get
    val expr = TExpr.Aggregate("VAR", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    // Mean: 15, Variance: (25+25)/1 = 50
    assertEquals(result, Right(BigDecimal(50)))
  }

  test("VAR: single value returns error (requires n >= 2)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(42))
    val range = CellRange.parse("A1:A1").toOption.get
    val expr = TExpr.Aggregate("VAR", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assert(result.isLeft, "Expected error for single value")
  }

  // ===== VARP Tests (Population Variance) =====

  test("VARP: standard case") {
    // Data: 2, 4, 4, 4, 5, 5, 7, 9
    // Mean: 5, Population variance: 32/8 = 4
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(2),
      ref"A2" -> CellValue.Number(4),
      ref"A3" -> CellValue.Number(4),
      ref"A4" -> CellValue.Number(4),
      ref"A5" -> CellValue.Number(5),
      ref"A6" -> CellValue.Number(5),
      ref"A7" -> CellValue.Number(7),
      ref"A8" -> CellValue.Number(9)
    )
    val range = CellRange.parse("A1:A8").toOption.get
    val expr = TExpr.Aggregate("VARP", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertApprox(result, 4.0, 0.0001)
  }

  test("VARP: single value returns 0") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(42))
    val range = CellRange.parse("A1:A1").toOption.get
    val expr = TExpr.Aggregate("VARP", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(0)))
  }

  test("VARP: empty range returns error") {
    val sheet = emptySheet
    val range = CellRange.parse("A1:A5").toOption.get
    val expr = TExpr.Aggregate("VARP", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    assert(result.isLeft, "Expected error for empty range")
  }

  // ===== Parser Roundtrip Tests =====

  test("MEDIAN: parse from string") {
    val result = FormulaParser.parse("=MEDIAN(A1:A10)")
    assert(result.isRight, s"Failed to parse MEDIAN: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=MEDIAN(A1:A10)")
    }
  }

  test("STDEV: parse from string") {
    val result = FormulaParser.parse("=STDEV(B1:B20)")
    assert(result.isRight, s"Failed to parse STDEV: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=STDEV(B1:B20)")
    }
  }

  test("STDEVP: parse from string") {
    val result = FormulaParser.parse("=STDEVP(C1:C5)")
    assert(result.isRight, s"Failed to parse STDEVP: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=STDEVP(C1:C5)")
    }
  }

  test("VAR: parse from string") {
    val result = FormulaParser.parse("=VAR(D1:D100)")
    assert(result.isRight, s"Failed to parse VAR: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=VAR(D1:D100)")
    }
  }

  test("VARP: parse from string") {
    val result = FormulaParser.parse("=VARP(E1:E50)")
    assert(result.isRight, s"Failed to parse VARP: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=VARP(E1:E50)")
    }
  }

  // ===== Cross-Sheet Tests =====

  test("MEDIAN: cross-sheet reference parses correctly") {
    val result = FormulaParser.parse("=MEDIAN(Sales!A1:A10)")
    assert(result.isRight, s"Failed to parse cross-sheet MEDIAN: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=MEDIAN(Sales!A1:A10)")
    }
  }

  test("STDEV: cross-sheet reference parses correctly") {
    val result = FormulaParser.parse("=STDEV(Data!B1:B100)")
    assert(result.isRight, s"Failed to parse cross-sheet STDEV: $result")
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=STDEV(Data!B1:B100)")
    }
  }

  // ===== Numerical Stability Tests (Welford's Algorithm) =====

  test("STDEV: numerical stability with large numbers") {
    // Classic catastrophic cancellation case for naive algorithm
    // Values: 10000000000000, 10000000000001, 10000000000002
    // Naive formula (sumSq - sum²/n) would lose precision
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal("10000000000000")),
      ref"A2" -> CellValue.Number(BigDecimal("10000000000001")),
      ref"A3" -> CellValue.Number(BigDecimal("10000000000002"))
    )
    val range = CellRange.parse("A1:A3").toOption.get
    val expr = TExpr.Aggregate("STDEV", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    // Sample stdev of [0, 1, 2] = 1.0, so same for shifted values
    assertApprox(result, 1.0, 0.0001)
  }

  test("VAR: numerical stability with numbers of similar magnitude") {
    // Another classic case: values very close together relative to magnitude
    // [1000000.0, 1000001.0, 1000002.0] - variance should be 1.0
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal("1000000.0")),
      ref"A2" -> CellValue.Number(BigDecimal("1000001.0")),
      ref"A3" -> CellValue.Number(BigDecimal("1000002.0"))
    )
    val range = CellRange.parse("A1:A3").toOption.get
    val expr = TExpr.Aggregate("VAR", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    // Sample variance of [0, 1, 2] = 1.0
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("STDEVP: handles negative values correctly") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(-10),
      ref"A2" -> CellValue.Number(-5),
      ref"A3" -> CellValue.Number(0),
      ref"A4" -> CellValue.Number(5),
      ref"A5" -> CellValue.Number(10)
    )
    val range = CellRange.parse("A1:A5").toOption.get
    val expr = TExpr.Aggregate("STDEVP", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    // Mean = 0, Population variance = (100+25+0+25+100)/5 = 50
    // Population stdev = sqrt(50) ≈ 7.071
    assertApprox(result, 7.071, 0.001)
  }

  test("VARP: handles many decimal places") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal("0.123456789")),
      ref"A2" -> CellValue.Number(BigDecimal("0.123456790")),
      ref"A3" -> CellValue.Number(BigDecimal("0.123456791"))
    )
    val range = CellRange.parse("A1:A3").toOption.get
    val expr = TExpr.Aggregate("VARP", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    // Tiny variance from tiny differences - Welford handles this correctly
    result match
      case Right(value) =>
        // Should be approximately 6.67e-19 (very small but non-zero)
        assert(value > BigDecimal(0), s"Expected positive variance, got $value")
        assert(value < BigDecimal("0.000001"), s"Variance too large: $value")
      case Left(err) => fail(s"Expected success but got: $err")
  }

  test("MEDIAN: handles larger dataset") {
    // Create a sheet with values 1 to 10 in column A
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(1),
      ref"A2" -> CellValue.Number(2),
      ref"A3" -> CellValue.Number(3),
      ref"A4" -> CellValue.Number(4),
      ref"A5" -> CellValue.Number(5),
      ref"A6" -> CellValue.Number(6),
      ref"A7" -> CellValue.Number(7),
      ref"A8" -> CellValue.Number(8),
      ref"A9" -> CellValue.Number(9),
      ref"A10" -> CellValue.Number(10)
    )
    val range = CellRange.parse("A1:A10").toOption.get
    val expr = TExpr.Aggregate("MEDIAN", TExpr.RangeLocation.Local(range))
    val result = evaluator.eval(expr, sheet)
    // Median of 1..10 = (5 + 6) / 2 = 5.5
    assertEquals(result, Right(BigDecimal(5.5)))
  }
