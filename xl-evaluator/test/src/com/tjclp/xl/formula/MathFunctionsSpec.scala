package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.addressing.SheetName
import munit.FunSuite

/**
 * Comprehensive tests for math functions.
 *
 * Tests SQRT, MOD, POWER, LOG, LN, EXP, FLOOR, CEILING, TRUNC, SIGN, INT functions.
 * Covers basic operations, edge cases, and error conditions.
 */
class MathFunctionsSpec extends FunSuite:
  val emptySheet = new Sheet(name = SheetName.unsafe("Test"))
  val evaluator = Evaluator.instance

  /** Helper to create sheet with cells */
  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  // ===== SQRT Tests =====

  test("SQRT: square root of 16") {
    val expr = TExpr.Sqrt(TExpr.Lit(BigDecimal(16)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(4.0)))
  }

  test("SQRT: square root of 0") {
    val expr = TExpr.Sqrt(TExpr.Lit(BigDecimal(0)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(0.0)))
  }

  test("SQRT: square root of 2 (irrational)") {
    val expr = TExpr.Sqrt(TExpr.Lit(BigDecimal(2)))
    val result = evaluator.eval(expr, emptySheet)
    result match
      case Right(value) => assertEqualsDouble(value.toDouble, 1.4142135623730951, 0.0000001)
      case Left(err) => fail(s"Expected success, got $err")
  }

  test("SQRT: negative number returns error") {
    val expr = TExpr.Sqrt(TExpr.Lit(BigDecimal(-4)))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, "SQRT of negative number should return error")
  }

  // ===== MOD Tests =====

  test("MOD: 5 mod 3 = 2") {
    val expr = TExpr.Mod(TExpr.Lit(BigDecimal(5)), TExpr.Lit(BigDecimal(3)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(2)))
  }

  test("MOD: -5 mod 3 = 1 (Excel semantics: result has sign of divisor)") {
    val expr = TExpr.Mod(TExpr.Lit(BigDecimal(-5)), TExpr.Lit(BigDecimal(3)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("MOD: 5 mod -3 = -1 (Excel semantics: result has sign of divisor)") {
    val expr = TExpr.Mod(TExpr.Lit(BigDecimal(5)), TExpr.Lit(BigDecimal(-3)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(-1)))
  }

  test("MOD: 10 mod 10 = 0") {
    val expr = TExpr.Mod(TExpr.Lit(BigDecimal(10)), TExpr.Lit(BigDecimal(10)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(0)))
  }

  test("MOD: division by zero returns error") {
    val expr = TExpr.Mod(TExpr.Lit(BigDecimal(5)), TExpr.Lit(BigDecimal(0)))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, "MOD by zero should return error")
  }

  // ===== POWER Tests =====

  test("POWER: 2^3 = 8") {
    val expr = TExpr.Power(TExpr.Lit(BigDecimal(2)), TExpr.Lit(BigDecimal(3)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(8.0)))
  }

  test("POWER: 4^0.5 = 2 (square root)") {
    val expr = TExpr.Power(TExpr.Lit(BigDecimal(4)), TExpr.Lit(BigDecimal(0.5)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(2.0)))
  }

  test("POWER: 10^0 = 1") {
    val expr = TExpr.Power(TExpr.Lit(BigDecimal(10)), TExpr.Lit(BigDecimal(0)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1.0)))
  }

  test("POWER: 2^-1 = 0.5") {
    val expr = TExpr.Power(TExpr.Lit(BigDecimal(2)), TExpr.Lit(BigDecimal(-1)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(0.5)))
  }

  // ===== LOG Tests =====

  test("LOG: log10(100) = 2") {
    val expr = TExpr.Log(TExpr.Lit(BigDecimal(100)), TExpr.Lit(BigDecimal(10)))
    val result = evaluator.eval(expr, emptySheet)
    result match
      case Right(value) => assertEqualsDouble(value.toDouble, 2.0, 0.0000001)
      case Left(err) => fail(s"Expected success, got $err")
  }

  test("LOG: log2(8) = 3") {
    val expr = TExpr.Log(TExpr.Lit(BigDecimal(8)), TExpr.Lit(BigDecimal(2)))
    val result = evaluator.eval(expr, emptySheet)
    result match
      case Right(value) => assertEqualsDouble(value.toDouble, 3.0, 0.0000001)
      case Left(err) => fail(s"Expected success, got $err")
  }

  test("LOG: log of 0 returns error") {
    val expr = TExpr.Log(TExpr.Lit(BigDecimal(0)), TExpr.Lit(BigDecimal(10)))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, "LOG of 0 should return error")
  }

  test("LOG: log of negative returns error") {
    val expr = TExpr.Log(TExpr.Lit(BigDecimal(-1)), TExpr.Lit(BigDecimal(10)))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, "LOG of negative should return error")
  }

  test("LOG: log base 1 returns error") {
    val expr = TExpr.Log(TExpr.Lit(BigDecimal(10)), TExpr.Lit(BigDecimal(1)))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, "LOG base 1 should return error")
  }

  // ===== LN Tests =====

  test("LN: ln(e) ≈ 1") {
    val e = BigDecimal(Math.E)
    val expr = TExpr.Ln(TExpr.Lit(e))
    val result = evaluator.eval(expr, emptySheet)
    result match
      case Right(value) => assertEqualsDouble(value.toDouble, 1.0, 0.0000001)
      case Left(err) => fail(s"Expected success, got $err")
  }

  test("LN: ln(1) = 0") {
    val expr = TExpr.Ln(TExpr.Lit(BigDecimal(1)))
    val result = evaluator.eval(expr, emptySheet)
    result match
      case Right(value) => assertEqualsDouble(value.toDouble, 0.0, 0.0000001)
      case Left(err) => fail(s"Expected success, got $err")
  }

  test("LN: ln(0) returns error") {
    val expr = TExpr.Ln(TExpr.Lit(BigDecimal(0)))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, "LN of 0 should return error")
  }

  test("LN: ln of negative returns error") {
    val expr = TExpr.Ln(TExpr.Lit(BigDecimal(-1)))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, "LN of negative should return error")
  }

  // ===== EXP Tests =====

  test("EXP: e^0 = 1") {
    val expr = TExpr.Exp(TExpr.Lit(BigDecimal(0)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1.0)))
  }

  test("EXP: e^1 ≈ 2.718") {
    val expr = TExpr.Exp(TExpr.Lit(BigDecimal(1)))
    val result = evaluator.eval(expr, emptySheet)
    result match
      case Right(value) => assertEqualsDouble(value.toDouble, Math.E, 0.0000001)
      case Left(err) => fail(s"Expected success, got $err")
  }

  test("EXP: e^-1 ≈ 0.368") {
    val expr = TExpr.Exp(TExpr.Lit(BigDecimal(-1)))
    val result = evaluator.eval(expr, emptySheet)
    result match
      case Right(value) => assertEqualsDouble(value.toDouble, 1.0 / Math.E, 0.0000001)
      case Left(err) => fail(s"Expected success, got $err")
  }

  // ===== FLOOR Tests =====

  test("FLOOR: floor(2.5, 1) = 2") {
    val expr = TExpr.Floor(TExpr.Lit(BigDecimal(2.5)), TExpr.Lit(BigDecimal(1)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(2)))
  }

  test("FLOOR: floor(-2.5, -1) = -2") {
    val expr = TExpr.Floor(TExpr.Lit(BigDecimal(-2.5)), TExpr.Lit(BigDecimal(-1)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(-2)))
  }

  test("FLOOR: floor(1.5, 0.5) = 1.5") {
    val expr = TExpr.Floor(TExpr.Lit(BigDecimal(1.5)), TExpr.Lit(BigDecimal(0.5)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1.5)))
  }

  test("FLOOR: floor(3.7, 2) = 2") {
    val expr = TExpr.Floor(TExpr.Lit(BigDecimal(3.7)), TExpr.Lit(BigDecimal(2)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(2)))
  }

  test("FLOOR: significance 0 returns error") {
    val expr = TExpr.Floor(TExpr.Lit(BigDecimal(2.5)), TExpr.Lit(BigDecimal(0)))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, "FLOOR with significance 0 should return error")
  }

  test("FLOOR: positive number with negative significance returns error") {
    val expr = TExpr.Floor(TExpr.Lit(BigDecimal(2.5)), TExpr.Lit(BigDecimal(-1)))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, "FLOOR with mismatched signs should return error")
  }

  // ===== CEILING Tests =====

  test("CEILING: ceiling(2.5, 1) = 3") {
    val expr = TExpr.Ceiling(TExpr.Lit(BigDecimal(2.5)), TExpr.Lit(BigDecimal(1)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(3)))
  }

  test("CEILING: ceiling(-2.5, -1) = -3") {
    val expr = TExpr.Ceiling(TExpr.Lit(BigDecimal(-2.5)), TExpr.Lit(BigDecimal(-1)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(-3)))
  }

  test("CEILING: ceiling(1.2, 0.5) = 1.5") {
    val expr = TExpr.Ceiling(TExpr.Lit(BigDecimal(1.2)), TExpr.Lit(BigDecimal(0.5)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1.5)))
  }

  test("CEILING: significance 0 returns error") {
    val expr = TExpr.Ceiling(TExpr.Lit(BigDecimal(2.5)), TExpr.Lit(BigDecimal(0)))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, "CEILING with significance 0 should return error")
  }

  test("FLOOR: negative number with positive significance returns error") {
    val expr = TExpr.Floor(TExpr.Lit(BigDecimal(-2.5)), TExpr.Lit(BigDecimal(1)))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, "FLOOR(-2.5, 1) should return error (mismatched signs)")
  }

  test("CEILING: negative number with positive significance returns error") {
    val expr = TExpr.Ceiling(TExpr.Lit(BigDecimal(-2.5)), TExpr.Lit(BigDecimal(1)))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, "CEILING(-2.5, 1) should return error (mismatched signs)")
  }

  test("FLOOR: negative number with negative significance works") {
    // FLOOR(-4.1, -2) rounds down (toward -∞) to nearest multiple of -2
    // -4.1 / -2 = 2.05, floor(2.05) = 2, 2 * -2 = -4
    val expr = TExpr.Floor(TExpr.Lit(BigDecimal(-4.1)), TExpr.Lit(BigDecimal(-2)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(-4)))
  }

  test("CEILING: negative number with negative significance works") {
    // CEILING(-4.1, -2) rounds up (toward +∞) to nearest multiple of -2
    // -4.1 / -2 = 2.05, ceil(2.05) = 3, 3 * -2 = -6
    val expr = TExpr.Ceiling(TExpr.Lit(BigDecimal(-4.1)), TExpr.Lit(BigDecimal(-2)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(-6)))
  }

  // ===== TRUNC Tests =====

  test("TRUNC: trunc(8.9) = 8") {
    val expr = TExpr.Trunc(TExpr.Lit(BigDecimal(8.9)), TExpr.Lit(BigDecimal(0)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(8)))
  }

  test("TRUNC: trunc(-8.9) = -8") {
    val expr = TExpr.Trunc(TExpr.Lit(BigDecimal(-8.9)), TExpr.Lit(BigDecimal(0)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(-8)))
  }

  test("TRUNC: trunc(3.14159, 2) = 3.14") {
    val expr = TExpr.Trunc(TExpr.Lit(BigDecimal(3.14159)), TExpr.Lit(BigDecimal(2)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(3.14)))
  }

  test("TRUNC: trunc(1234.5678, -2) = 1200") {
    val expr = TExpr.Trunc(TExpr.Lit(BigDecimal(1234.5678)), TExpr.Lit(BigDecimal(-2)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1200)))
  }

  // ===== SIGN Tests =====

  test("SIGN: sign(5) = 1") {
    val expr = TExpr.Sign(TExpr.Lit(BigDecimal(5)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("SIGN: sign(-5) = -1") {
    val expr = TExpr.Sign(TExpr.Lit(BigDecimal(-5)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(-1)))
  }

  test("SIGN: sign(0) = 0") {
    val expr = TExpr.Sign(TExpr.Lit(BigDecimal(0)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(0)))
  }

  // ===== INT Tests =====

  test("INT: int(8.9) = 8") {
    val expr = TExpr.Int_(TExpr.Lit(BigDecimal(8.9)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(8)))
  }

  test("INT: int(-8.9) = -9 (floor behavior)") {
    val expr = TExpr.Int_(TExpr.Lit(BigDecimal(-8.9)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(-9)))
  }

  test("INT: int(0) = 0") {
    val expr = TExpr.Int_(TExpr.Lit(BigDecimal(0)))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(0)))
  }

  test("INT vs TRUNC: INT rounds toward negative infinity, TRUNC toward zero") {
    // For -8.9: INT gives -9, TRUNC gives -8
    val intExpr = TExpr.Int_(TExpr.Lit(BigDecimal(-8.9)))
    val truncExpr = TExpr.Trunc(TExpr.Lit(BigDecimal(-8.9)), TExpr.Lit(BigDecimal(0)))

    val intResult = evaluator.eval(intExpr, emptySheet)
    val truncResult = evaluator.eval(truncExpr, emptySheet)

    assertEquals(intResult, Right(BigDecimal(-9)))
    assertEquals(truncResult, Right(BigDecimal(-8)))
  }

  // ===== Parser Integration Tests (using evaluateFormula) =====

  test("SQRT parser: parse and evaluate =SQRT(16)") {
    val result = emptySheet.evaluateFormula("=SQRT(16)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(4.0))))
  }

  test("MOD parser: parse and evaluate =MOD(7, 3)") {
    val result = emptySheet.evaluateFormula("=MOD(7, 3)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(1))))
  }

  test("POWER parser: parse and evaluate =POWER(2, 10)") {
    val result = emptySheet.evaluateFormula("=POWER(2, 10)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(1024.0))))
  }

  test("LOG parser: parse and evaluate =LOG(1000)") {
    val result = emptySheet.evaluateFormula("=LOG(1000)")
    result match
      case Right(CellValue.Number(value)) => assertEqualsDouble(value.toDouble, 3.0, 0.0000001)
      case other => fail(s"Expected Number, got $other")
  }

  test("LOG parser with base: parse and evaluate =LOG(64, 2)") {
    val result = emptySheet.evaluateFormula("=LOG(64, 2)")
    result match
      case Right(CellValue.Number(value)) => assertEqualsDouble(value.toDouble, 6.0, 0.0000001)
      case other => fail(s"Expected Number, got $other")
  }

  test("LN parser: parse and evaluate =LN(2.718281828)") {
    val result = emptySheet.evaluateFormula("=LN(2.718281828)")
    result match
      case Right(CellValue.Number(value)) => assertEqualsDouble(value.toDouble, 1.0, 0.0001)
      case other => fail(s"Expected Number, got $other")
  }

  test("EXP parser: parse and evaluate =EXP(2)") {
    val result = emptySheet.evaluateFormula("=EXP(2)")
    result match
      case Right(CellValue.Number(value)) => assertEqualsDouble(value.toDouble, Math.E * Math.E, 0.0001)
      case other => fail(s"Expected Number, got $other")
  }

  test("FLOOR parser: parse and evaluate =FLOOR(5.7, 2)") {
    val result = emptySheet.evaluateFormula("=FLOOR(5.7, 2)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(4))))
  }

  test("CEILING parser: parse and evaluate =CEILING(5.1, 2)") {
    val result = emptySheet.evaluateFormula("=CEILING(5.1, 2)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(6))))
  }

  test("TRUNC parser: parse and evaluate =TRUNC(8.999)") {
    val result = emptySheet.evaluateFormula("=TRUNC(8.999)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(8))))
  }

  test("SIGN parser: parse and evaluate =SIGN(-42)") {
    val result = emptySheet.evaluateFormula("=SIGN(-42)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(-1))))
  }

  test("INT parser: parse and evaluate =INT(-8.9)") {
    val result = emptySheet.evaluateFormula("=INT(-8.9)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(-9))))
  }

  // ===== Combined Formula Tests =====

  test("Combined: =SQRT(POWER(3, 2) + POWER(4, 2)) = 5 (Pythagorean theorem)") {
    val result = emptySheet.evaluateFormula("=SQRT(POWER(3, 2) + POWER(4, 2))")
    result match
      case Right(CellValue.Number(value)) => assertEqualsDouble(value.toDouble, 5.0, 0.0000001)
      case other => fail(s"Expected Number, got $other")
  }

  test("Combined: =EXP(LN(10)) = 10 (inverse relationship)") {
    val result = emptySheet.evaluateFormula("=EXP(LN(10))")
    result match
      case Right(CellValue.Number(value)) => assertEqualsDouble(value.toDouble, 10.0, 0.0000001)
      case other => fail(s"Expected Number, got $other")
  }

  test("Combined: =LN(EXP(5)) = 5 (inverse relationship)") {
    val result = emptySheet.evaluateFormula("=LN(EXP(5))")
    result match
      case Right(CellValue.Number(value)) => assertEqualsDouble(value.toDouble, 5.0, 0.0000001)
      case other => fail(s"Expected Number, got $other")
  }
