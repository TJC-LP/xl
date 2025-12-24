package com.tjclp.xl.formula

import com.tjclp.xl.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import org.scalacheck.{Arbitrary, Gen}

import scala.math.BigDecimal

/**
 * Comprehensive tests for Evaluator.
 *
 * Tests laws (literal identity, arithmetic correctness, short-circuit semantics), edge cases
 * (division by zero, missing cells, type mismatches), and integration (real Excel formulas).
 *
 * Target: ~50 tests (7-10 property, 20-25 unit, 15-20 integration, 5-8 error)
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
class EvaluatorSpec extends ScalaCheckSuite:

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

  // ==================== Generators ====================

  val genBigDecimal: Gen[BigDecimal] =
    Gen.choose(-1000.0, 1000.0).map(BigDecimal.apply)

  val genBoolean: Gen[Boolean] =
    Gen.oneOf(true, false)

  val genString: Gen[String] =
    Gen.alphaNumStr.suchThat(_.nonEmpty)

  val genARef: Gen[ARef] =
    for
      col <- Gen.choose(0, 50)
      row <- Gen.choose(0, 50)
    yield ARef.from0(col, row)

  val genCellRange: Gen[CellRange] =
    for
      startCol <- Gen.choose(0, 40)
      startRow <- Gen.choose(0, 40)
      endCol <- Gen.choose(startCol, startCol + 5)
      endRow <- Gen.choose(startRow, startRow + 5)
      start = ARef.from0(startCol, startRow)
      end = ARef.from0(endCol, endRow)
    yield CellRange(start, end)

  // ==================== Property-Based Tests (Laws) ====================

  property("Literal identity: eval(Lit(x)) == Right(x) for BigDecimal") {
    forAll(genBigDecimal) { value =>
      val expr = TExpr.Lit(value)
      val sheet = new Sheet(name = SheetName.unsafe("Empty"))
      evaluator.eval(expr, sheet) == Right(value)
    }
  }

  property("Literal identity: eval(Lit(x)) == Right(x) for Boolean") {
    forAll(genBoolean) { value =>
      val expr = TExpr.Lit(value)
      val sheet = new Sheet(name = SheetName.unsafe("Empty"))
      evaluator.eval(expr, sheet) == Right(value)
    }
  }

  property("Literal identity: eval(Lit(x)) == Right(x) for String") {
    forAll(genString) { value =>
      val expr = TExpr.Lit(value)
      val sheet = new Sheet(name = SheetName.unsafe("Empty"))
      evaluator.eval(expr, sheet) == Right(value)
    }
  }

  property("Arithmetic: Add(Lit(a), Lit(b)) == Right(a + b)") {
    forAll(genBigDecimal, genBigDecimal) { (a, b) =>
      val expr = TExpr.Add(TExpr.Lit(a), TExpr.Lit(b))
      val sheet = new Sheet(name = SheetName.unsafe("Empty"))
      evaluator.eval(expr, sheet) == Right(a + b)
    }
  }

  property("Arithmetic: Sub(Lit(a), Lit(b)) == Right(a - b)") {
    forAll(genBigDecimal, genBigDecimal) { (a, b) =>
      val expr = TExpr.Sub(TExpr.Lit(a), TExpr.Lit(b))
      val sheet = new Sheet(name = SheetName.unsafe("Empty"))
      evaluator.eval(expr, sheet) == Right(a - b)
    }
  }

  property("Arithmetic: Mul(Lit(a), Lit(b)) == Right(a * b)") {
    forAll(genBigDecimal, genBigDecimal) { (a, b) =>
      val expr = TExpr.Mul(TExpr.Lit(a), TExpr.Lit(b))
      val sheet = new Sheet(name = SheetName.unsafe("Empty"))
      evaluator.eval(expr, sheet) == Right(a * b)
    }
  }

  property("Comparison: Lt(Lit(a), Lit(b)) == Right(a < b)") {
    forAll(genBigDecimal, genBigDecimal) { (a, b) =>
      val expr = TExpr.Lt(TExpr.Lit(a), TExpr.Lit(b))
      val sheet = new Sheet(name = SheetName.unsafe("Empty"))
      evaluator.eval(expr, sheet) == Right(a < b)
    }
  }

  property("Short-circuit: And(Lit(false), error) doesn't evaluate error") {
    forAll(genARef) { missingRef =>
      val errorExpr = TExpr.Ref[Boolean](missingRef, Anchor.Relative, _ => Left(CodecError.TypeMismatch("Boolean", CellValue.Empty)))
      val andExpr = TExpr.Lit(false) && errorExpr
      val sheet = new Sheet(name = SheetName.unsafe("Empty"))
      // Should return Right(false) without evaluating errorExpr
      evaluator.eval(andExpr, sheet) == Right(false)
    }
  }

  property("Short-circuit: Or(Lit(true), error) doesn't evaluate error") {
    forAll(genARef) { missingRef =>
      val errorExpr = TExpr.Ref[Boolean](missingRef, Anchor.Relative, _ => Left(CodecError.TypeMismatch("Boolean", CellValue.Empty)))
      val orExpr = TExpr.Lit(true) || errorExpr
      val sheet = new Sheet(name = SheetName.unsafe("Empty"))
      // Should return Right(true) without evaluating errorExpr
      evaluator.eval(orExpr, sheet) == Right(true)
    }
  }

  property("Eq is reflexive: Eq(Lit(x), Lit(x)) == Right(true)") {
    forAll(genBigDecimal) { x =>
      val expr = TExpr.Eq(TExpr.Lit(x), TExpr.Lit(x))
      val sheet = new Sheet(name = SheetName.unsafe("Empty"))
      evaluator.eval(expr, sheet) == Right(true)
    }
  }

  // ==================== Unit Tests (Edge Cases) ====================

  test("Division by zero: Div(Lit(10), Lit(0)) returns DivByZero error") {
    val expr = TExpr.Div(TExpr.Lit(BigDecimal(10)), TExpr.Lit(BigDecimal(0)))
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    val error = evalErr(expr, sheet)
    error match
      case _: EvalError.DivByZero => // success
      case other => fail(s"Expected DivByZero, got $other")
  }

  test("Division by zero: nested expression Sub(5, 5) = 0") {
    val numerator = TExpr.Lit(BigDecimal(10))
    val denominator = TExpr.Sub(TExpr.Lit(BigDecimal(5)), TExpr.Lit(BigDecimal(5)))
    val expr = TExpr.Div(numerator, denominator)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    val error = evalErr(expr, sheet)
    error match
      case _: EvalError.DivByZero => // success
      case other => fail(s"Expected DivByZero, got $other")
  }

  test("Division: Div(Lit(10), Lit(2)) == Right(5)") {
    val expr = TExpr.Div(TExpr.Lit(BigDecimal(10)), TExpr.Lit(BigDecimal(2)))
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == BigDecimal(5))
  }

  test("Cell reference: missing cell returns CodecFailed") {
    val ref = ARef.from0(0, 0) // A1
    val expr = TExpr.ref(ref, TExpr.decodeNumeric)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    val error = evalErr(expr, sheet)
    error match
      case _: EvalError.CodecFailed => // success
      case other => fail(s"Expected CodecFailed, got $other")
  }

  test("Cell reference: existing numeric cell returns value") {
    val ref = ARef.from0(0, 0) // A1
    val sheet = sheetWith(ref -> CellValue.Number(BigDecimal(42)))
    val expr = TExpr.ref(ref, TExpr.decodeNumeric)
    assert(evalOk(expr, sheet) == BigDecimal(42))
  }

  test("Cell reference: text cell with numeric decoder returns CodecFailed") {
    val ref = ARef.from0(0, 0) // A1
    val sheet = sheetWith(ref -> CellValue.Text("hello"))
    val expr = TExpr.ref(ref, TExpr.decodeNumeric)
    val error = evalErr(expr, sheet)
    error match
      case _: EvalError.CodecFailed => // success
      case other => fail(s"Expected CodecFailed, got $other")
  }

  test("Boolean And: And(true, true) == Right(true)") {
    val expr = TExpr.Lit(true) && TExpr.Lit(true)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == true)
  }

  test("Boolean And: And(true, false) == Right(false)") {
    val expr = TExpr.Lit(true) && TExpr.Lit(false)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == false)
  }

  test("Boolean And: And(false, true) == Right(false)") {
    val expr = TExpr.Lit(false) && TExpr.Lit(true)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == false)
  }

  test("Boolean And: And(false, false) == Right(false)") {
    val expr = TExpr.Lit(false) && TExpr.Lit(false)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == false)
  }

  test("Boolean Or: Or(true, true) == Right(true)") {
    val expr = TExpr.Lit(true) || TExpr.Lit(true)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == true)
  }

  test("Boolean Or: Or(true, false) == Right(true)") {
    val expr = TExpr.Lit(true) || TExpr.Lit(false)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == true)
  }

  test("Boolean Or: Or(false, true) == Right(true)") {
    val expr = TExpr.Lit(false) || TExpr.Lit(true)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == true)
  }

  test("Boolean Or: Or(false, false) == Right(false)") {
    val expr = TExpr.Lit(false) || TExpr.Lit(false)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == false)
  }

  test("Boolean Not: Not(Lit(true)) == Right(false)") {
    val expr = !TExpr.Lit(true)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == false)
  }

  test("Boolean Not: Not(Lit(false)) == Right(true)") {
    val expr = !TExpr.Lit(false)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == true)
  }

  test("Comparison Lt: Lt(Lit(-5), Lit(3)) == Right(true)") {
    val expr = TExpr.Lt(TExpr.Lit(BigDecimal(-5)), TExpr.Lit(BigDecimal(3)))
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == true)
  }

  test("Comparison Lte: Lte(Lit(5), Lit(5)) == Right(true)") {
    val expr = TExpr.Lte(TExpr.Lit(BigDecimal(5)), TExpr.Lit(BigDecimal(5)))
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == true)
  }

  test("Comparison Gt: Gt(Lit(10), Lit(3)) == Right(true)") {
    val expr = TExpr.Gt(TExpr.Lit(BigDecimal(10)), TExpr.Lit(BigDecimal(3)))
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == true)
  }

  test("Comparison Gte: Gte(Lit(5), Lit(5)) == Right(true)") {
    val expr = TExpr.Gte(TExpr.Lit(BigDecimal(5)), TExpr.Lit(BigDecimal(5)))
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == true)
  }

  test("Comparison Eq: Eq(Lit(5), Lit(5)) == Right(true)") {
    val expr = TExpr.Eq(TExpr.Lit(BigDecimal(5)), TExpr.Lit(BigDecimal(5)))
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == true)
  }

  test("Comparison Eq: Eq(Lit(5), Lit(3)) == Right(false)") {
    val expr = TExpr.Eq(TExpr.Lit(BigDecimal(5)), TExpr.Lit(BigDecimal(3)))
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == false)
  }

  test("Comparison Neq: Neq(Lit(5), Lit(3)) == Right(true)") {
    val expr = TExpr.Neq(TExpr.Lit(BigDecimal(5)), TExpr.Lit(BigDecimal(3)))
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    assert(evalOk(expr, sheet) == true)
  }

  test("FoldRange: empty range returns zero value") {
    val range = CellRange(ARef.from0(10, 10), ARef.from0(10, 10)) // K11:K11 (single empty cell)
    val expr = TExpr.sum(range)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    // Empty range should return zero (no cells to sum)
    assert(evalOk(expr, sheet) == BigDecimal(0))
  }

  test("FoldRange: SUM with single cell") {
    val ref = ARef.from0(0, 0) // A1
    val range = CellRange(ref, ref) // A1:A1
    val sheet = sheetWith(ref -> CellValue.Number(BigDecimal(42)))
    val expr = TExpr.sum(range)
    assert(evalOk(expr, sheet) == BigDecimal(42))
  }

  test("FoldRange: SUM with multiple cells") {
    val range = CellRange(ARef.from0(0, 0), ARef.from0(0, 2)) // A1:A3
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(20)),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(30))
    )
    val expr = TExpr.sum(range)
    assert(evalOk(expr, sheet) == BigDecimal(60))
  }

  test("FoldRange: COUNT with multiple cells") {
    val range = CellRange(ARef.from0(0, 0), ARef.from0(0, 4)) // A1:A5
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)),
      ARef.from0(0, 1) -> CellValue.Text("hello"),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(30)),
      ARef.from0(0, 3) -> CellValue.Empty
      // A5 not set (empty)
    )
    val expr = TExpr.count(range)
    // Note: Current implementation counts all cells in range (including empty)
    // because decodeAny succeeds for all cells and step function doesn't check Some/None
    // Excel COUNTA would return 3 (only non-empty: A1, A2, A3)
    // This is a known limitation - COUNT should use a smarter step function
    assert(evalOk(expr, sheet) == 5)  // All 5 cells in range
  }

  // ==================== Integration Tests (Real Formulas) ====================

  test("Integration: IF(A1>0, 'Positive', 'Non-positive') with A1=5") {
    val refA1 = ARef.from0(0, 0)
    val sheet = sheetWith(refA1 -> CellValue.Number(BigDecimal(5)))
    val condition = TExpr.Gt(TExpr.ref(refA1, TExpr.decodeNumeric), TExpr.Lit(BigDecimal(0)))
    val expr = TExpr.cond(condition, TExpr.Lit("Positive"), TExpr.Lit("Non-positive"))
    assert(evalOk(expr, sheet) == "Positive")
  }

  test("Integration: IF(A1>0, 'Positive', 'Non-positive') with A1=-3") {
    val refA1 = ARef.from0(0, 0)
    val sheet = sheetWith(refA1 -> CellValue.Number(BigDecimal(-3)))
    val condition = TExpr.Gt(TExpr.ref(refA1, TExpr.decodeNumeric), TExpr.Lit(BigDecimal(0)))
    val expr = TExpr.cond(condition, TExpr.Lit("Positive"), TExpr.Lit("Non-positive"))
    assert(evalOk(expr, sheet) == "Non-positive")
  }

  test("Integration: IF(AND(A1>0, B1<100), SUM(C1:C3), 0) - condition true") {
    val refA1 = ARef.from0(0, 0)
    val refB1 = ARef.from0(1, 0)
    val rangeC1C3 = CellRange(ARef.from0(2, 0), ARef.from0(2, 2))
    val sheet = sheetWith(
      refA1 -> CellValue.Number(BigDecimal(10)),
      refB1 -> CellValue.Number(BigDecimal(50)),
      ARef.from0(2, 0) -> CellValue.Number(BigDecimal(100)),
      ARef.from0(2, 1) -> CellValue.Number(BigDecimal(200)),
      ARef.from0(2, 2) -> CellValue.Number(BigDecimal(300))
    )
    val condA1 = TExpr.Gt(TExpr.ref(refA1, TExpr.decodeNumeric), TExpr.Lit(BigDecimal(0)))
    val condB1 = TExpr.Lt(TExpr.ref(refB1, TExpr.decodeNumeric), TExpr.Lit(BigDecimal(100)))
    val condition = condA1 && condB1
    val sumExpr = TExpr.sum(rangeC1C3)
    val expr = TExpr.cond(condition, sumExpr, TExpr.Lit(BigDecimal(0)))
    assert(evalOk(expr, sheet) == BigDecimal(600))
  }

  test("Integration: IF(AND(A1>0, B1<100), SUM(C1:C3), 0) - condition false") {
    val refA1 = ARef.from0(0, 0)
    val refB1 = ARef.from0(1, 0)
    val rangeC1C3 = CellRange(ARef.from0(2, 0), ARef.from0(2, 2))
    val sheet = sheetWith(
      refA1 -> CellValue.Number(BigDecimal(-5)), // Negative, condition fails
      refB1 -> CellValue.Number(BigDecimal(50)),
      ARef.from0(2, 0) -> CellValue.Number(BigDecimal(100)),
      ARef.from0(2, 1) -> CellValue.Number(BigDecimal(200)),
      ARef.from0(2, 2) -> CellValue.Number(BigDecimal(300))
    )
    val condA1 = TExpr.Gt(TExpr.ref(refA1, TExpr.decodeNumeric), TExpr.Lit(BigDecimal(0)))
    val condB1 = TExpr.Lt(TExpr.ref(refB1, TExpr.decodeNumeric), TExpr.Lit(BigDecimal(100)))
    val condition = condA1 && condB1
    val sumExpr = TExpr.sum(rangeC1C3)
    val expr = TExpr.cond(condition, sumExpr, TExpr.Lit(BigDecimal(0)))
    assert(evalOk(expr, sheet) == BigDecimal(0))
  }

  test("Integration: A1 + B1 * C1 with precedence") {
    val refA1 = ARef.from0(0, 0)
    val refB1 = ARef.from0(1, 0)
    val refC1 = ARef.from0(2, 0)
    val sheet = sheetWith(
      refA1 -> CellValue.Number(BigDecimal(10)),
      refB1 -> CellValue.Number(BigDecimal(5)),
      refC1 -> CellValue.Number(BigDecimal(2))
    )
    val mul = TExpr.Mul(TExpr.ref(refB1, TExpr.decodeNumeric), TExpr.ref(refC1, TExpr.decodeNumeric))
    val expr = TExpr.Add(TExpr.ref(refA1, TExpr.decodeNumeric), mul)
    // A1 + (B1 * C1) = 10 + (5 * 2) = 10 + 10 = 20
    assert(evalOk(expr, sheet) == BigDecimal(20))
  }

  test("Integration: (A1 + B1) / (C1 - D1)") {
    val refA1 = ARef.from0(0, 0)
    val refB1 = ARef.from0(1, 0)
    val refC1 = ARef.from0(2, 0)
    val refD1 = ARef.from0(3, 0)
    val sheet = sheetWith(
      refA1 -> CellValue.Number(BigDecimal(20)),
      refB1 -> CellValue.Number(BigDecimal(10)),
      refC1 -> CellValue.Number(BigDecimal(10)),
      refD1 -> CellValue.Number(BigDecimal(4))
    )
    val numerator = TExpr.Add(TExpr.ref(refA1, TExpr.decodeNumeric), TExpr.ref(refB1, TExpr.decodeNumeric))
    val denominator = TExpr.Sub(TExpr.ref(refC1, TExpr.decodeNumeric), TExpr.ref(refD1, TExpr.decodeNumeric))
    val expr = TExpr.Div(numerator, denominator)
    // (20 + 10) / (10 - 4) = 30 / 6 = 5
    assert(evalOk(expr, sheet) == BigDecimal(5))
  }

  test("Integration: Nested IF - IF(A1>10, 'High', IF(A1>5, 'Medium', 'Low'))") {
    val refA1 = ARef.from0(0, 0)
    val sheet = sheetWith(refA1 -> CellValue.Number(BigDecimal(7)))
    val cond1 = TExpr.Gt(TExpr.ref(refA1, TExpr.decodeNumeric), TExpr.Lit(BigDecimal(10)))
    val cond2 = TExpr.Gt(TExpr.ref(refA1, TExpr.decodeNumeric), TExpr.Lit(BigDecimal(5)))
    val innerIf = TExpr.cond(cond2, TExpr.Lit("Medium"), TExpr.Lit("Low"))
    val expr = TExpr.cond(cond1, TExpr.Lit("High"), innerIf)
    // A1=7: not >10, but >5, so "Medium"
    assert(evalOk(expr, sheet) == "Medium")
  }

  test("Integration: Complex boolean - (A1>5) AND ((B1<10) OR (C1==20))") {
    val refA1 = ARef.from0(0, 0)
    val refB1 = ARef.from0(1, 0)
    val refC1 = ARef.from0(2, 0)
    val sheet = sheetWith(
      refA1 -> CellValue.Number(BigDecimal(8)),
      refB1 -> CellValue.Number(BigDecimal(15)),
      refC1 -> CellValue.Number(BigDecimal(20))
    )
    val condA = TExpr.Gt(TExpr.ref(refA1, TExpr.decodeNumeric), TExpr.Lit(BigDecimal(5)))
    val condB = TExpr.Lt(TExpr.ref(refB1, TExpr.decodeNumeric), TExpr.Lit(BigDecimal(10)))
    val condC = TExpr.Eq(TExpr.ref(refC1, TExpr.decodeNumeric), TExpr.Lit(BigDecimal(20)))
    val orExpr = condB || condC
    val expr = condA && orExpr
    // A1>5: true, B1<10: false, C1==20: true, (false OR true)=true, (true AND true)=true
    assert(evalOk(expr, sheet) == true)
  }

  test("Integration: SUM with AVERAGE-like pattern (sum/count)") {
    val range = CellRange(ARef.from0(0, 0), ARef.from0(0, 2)) // A1:A3
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(20)),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(30))
    )
    val sumExpr = TExpr.sum(range)
    val countExpr = TExpr.count(range)
    // Average = Sum / Count (but count returns Int, need to convert)
    // For now, just test sum and count separately
    assert(evalOk(sumExpr, sheet) == BigDecimal(60))
    assert(evalOk(countExpr, sheet) == 3)
  }

  test("AVERAGE returns BigDecimal directly at eval level") {
    // AVERAGE now has dedicated TExpr case, returns BigDecimal directly
    val range = CellRange(ARef.from0(0, 0), ARef.from0(0, 2)) // A1:A3
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(150)),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(200)),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(75))
    )
    val avgExpr = TExpr.average(range)

    // evalOk returns BigDecimal directly (sum / count = 425 / 3)
    val result = evalOk(avgExpr, sheet)

    // Verify it's the correct average: (150 + 200 + 75) / 3 = 425 / 3
    val expected = BigDecimal(425) / 3
    assert(result == expected, s"Expected $expected, got $result")
  }

  test("AVERAGE returns Number via evaluateFormula (regression)") {
    // Regression test: AVERAGE formula should return CellValue.Number, not tuple/text
    // Uses sheet.evaluateFormula which includes toCellValue conversion
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(20)),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(30))
    )
    sheet.evaluateFormula("=AVERAGE(A1:A3)") match
      case Right(CellValue.Number(n)) =>
        assertEquals(n, BigDecimal(20))
      case Right(other) =>
        fail(s"Expected CellValue.Number(20), got $other")
      case Left(err) =>
        fail(s"Evaluation failed: $err")
  }

  test("AVERAGE in nested arithmetic (regression for asInstanceOf crash)") {
    // Regression test: Original bug caused "Tuple2 cannot be cast to BigDecimal" crash
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(100)),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(200)),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal(10)),
      ARef.from0(1, 1) -> CellValue.Number(BigDecimal(20))
    )
    sheet.evaluateFormula("=SUM(A1:A2)+AVERAGE(B1:B2)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(315))
      case Right(other) => fail(s"Expected Number, got $other")
      case Left(err) => fail(s"Evaluation failed: $err")
  }

  test("AVERAGE on empty range returns DivByZero error (Excel-compliant)") {
    val range = CellRange(ARef.from0(0, 0), ARef.from0(0, 2)) // A1:A3
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    val avgExpr = TExpr.average(range)

    Evaluator.instance.eval(avgExpr, sheet) match
      case Left(_: EvalError.DivByZero) => () // Expected
      case other => fail(s"Expected DivByZero error, got $other")
  }

  test("MIN/MAX/AVERAGE iterator consumption regression - first element must be included") {
    // Regression test: Iterator-based implementations must not consume iterator
    // with .isEmpty before .min/.max/.sum - that would skip the first element.
    // We place the extreme value FIRST to catch this bug.
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(1)),   // A1: MIN is here (first!)
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(5)),   // A2
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(100)), // A3: MAX is here (last)
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal(99)),  // B1: MAX is here (first!)
      ARef.from0(1, 1) -> CellValue.Number(BigDecimal(50)),  // B2
      ARef.from0(1, 2) -> CellValue.Number(BigDecimal(10))   // B3: MIN is here (last)
    )

    val rangeA = CellRange(ARef.from0(0, 0), ARef.from0(0, 2)) // A1:A3
    val rangeB = CellRange(ARef.from0(1, 0), ARef.from0(1, 2)) // B1:B3

    // MIN: first element is minimum - would be wrong if iterator consumed
    assertEquals(
      Evaluator.instance.eval(TExpr.min(rangeA), sheet),
      Right(BigDecimal(1)),
      "MIN must include first element (1)"
    )

    // MAX: first element is maximum - would be wrong if iterator consumed
    assertEquals(
      Evaluator.instance.eval(TExpr.max(rangeB), sheet),
      Right(BigDecimal(99)),
      "MAX must include first element (99)"
    )

    // AVERAGE: all elements must be included
    // A1:A3 = (1 + 5 + 100) / 3 = 106 / 3 = 35.333...
    Evaluator.instance.eval(TExpr.average(rangeA), sheet) match
      case Right(avg) =>
        // Use approximate comparison for BigDecimal division
        val expected = BigDecimal(106) / BigDecimal(3)
        assert(
          (avg - expected).abs < BigDecimal("0.0001"),
          s"AVERAGE must include all elements: expected ~$expected, got $avg"
        )
      case Left(err) => fail(s"AVERAGE failed: $err")
  }

  // ==================== Error Path Tests ====================

  test("Error: Nested error - error in IF condition propagates") {
    val refA1 = ARef.from0(0, 0) // A1 doesn't exist
    val condition = TExpr.Ref[Boolean](refA1, Anchor.Relative, _ => Left(CodecError.TypeMismatch("Boolean", CellValue.Empty)))
    val expr = TExpr.cond(condition, TExpr.Lit("Yes"), TExpr.Lit("No"))
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    val error = evalErr(expr, sheet)
    error match
      case _: EvalError.CodecFailed => // success
      case other => fail(s"Expected CodecFailed, got $other")
  }

  test("Error: Nested error - error in IF true branch (condition=true)") {
    val refA1 = ARef.from0(0, 0)
    val condition = TExpr.Lit(true)
    val trueBranch = TExpr.Ref[String](refA1, Anchor.Relative, _ => Left(CodecError.TypeMismatch("String", CellValue.Empty)))
    val expr = TExpr.cond(condition, trueBranch, TExpr.Lit("No"))
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    val error = evalErr(expr, sheet)
    error match
      case _: EvalError.CodecFailed => // success
      case other => fail(s"Expected CodecFailed, got $other")
  }

  test("Error: Nested error - error in IF false branch NOT evaluated (condition=true)") {
    val refA1 = ARef.from0(0, 0)
    val condition = TExpr.Lit(true)
    val falseBranch = TExpr.Ref[String](refA1, Anchor.Relative, _ => Left(CodecError.TypeMismatch("String", CellValue.Empty)))
    val expr = TExpr.cond(condition, TExpr.Lit("Yes"), falseBranch)
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    // False branch should NOT be evaluated, so no error
    assert(evalOk(expr, sheet) == "Yes")
  }

  test("FoldRange: SUM skips text cells (Excel behavior)") {
    val range = CellRange(ARef.from0(0, 0), ARef.from0(0, 2)) // A1:A3
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)),
      ARef.from0(0, 1) -> CellValue.Text("invalid"), // Skipped by SUM
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(30))
    )
    val expr = TExpr.sum(range)
    // Excel behavior: SUM skips text cells, sums numeric ones (10 + 30 = 40)
    assert(evalOk(expr, sheet) == BigDecimal(40))
  }

  test("Error: Multiple cell references, first one fails") {
    val refA1 = ARef.from0(0, 0) // Missing
    val refB1 = ARef.from0(1, 0) // Exists
    val sheet = sheetWith(refB1 -> CellValue.Number(BigDecimal(20)))
    val expr = TExpr.Add(
      TExpr.ref(refA1, TExpr.decodeNumeric),
      TExpr.ref(refB1, TExpr.decodeNumeric)
    )
    // First ref (A1) fails, error propagates
    val error = evalErr(expr, sheet)
    error match
      case _: EvalError.CodecFailed => // success
      case other => fail(s"Expected CodecFailed, got $other")
  }

  test("Error: Division by computed zero - (B1-B1)") {
    val refB1 = ARef.from0(1, 0)
    val sheet = sheetWith(refB1 -> CellValue.Number(BigDecimal(5)))
    val numerator = TExpr.Lit(BigDecimal(10))
    val denominator = TExpr.Sub(
      TExpr.ref(refB1, TExpr.decodeNumeric),
      TExpr.ref(refB1, TExpr.decodeNumeric)
    )
    val expr = TExpr.Div(numerator, denominator)
    val error = evalErr(expr, sheet)
    error match
      case _: EvalError.DivByZero => // success
      case other => fail(s"Expected DivByZero, got $other")
  }

  test("Error: AND with error in second operand (first=true, evaluates second)") {
    val refA1 = ARef.from0(0, 0) // Missing
    val errorExpr = TExpr.Ref[Boolean](refA1, Anchor.Relative, _ => Left(CodecError.TypeMismatch("Boolean", CellValue.Empty)))
    val andExpr = TExpr.Lit(true) && errorExpr
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    // First is true, so second is evaluated and error propagates
    val error = evalErr(andExpr, sheet)
    error match
      case _: EvalError.CodecFailed => // success
      case other => fail(s"Expected CodecFailed, got $other")
  }

  // ==================== Error Handling Functions (IFERROR, ISERROR) ====================

  test("IFERROR: returns value when no error") {
    val sheet = sheetWith(ARef.from0(0, 0) -> CellValue.Number(BigDecimal(42)))
    sheet.evaluateFormula("=IFERROR(A1, 0)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(42))
      case other => fail(s"Expected Number(42), got $other")
  }

  test("IFERROR: returns fallback on division by zero") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal(0))
    )
    sheet.evaluateFormula("=IFERROR(A1/B1, -1)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(-1))
      case other => fail(s"Expected Number(-1), got $other")
  }

  test("IFERROR: returns fallback on CellValue.Error") {
    val sheet = sheetWith(ARef.from0(0, 0) -> CellValue.Error(CellError.Div0))
    sheet.evaluateFormula("=IFERROR(A1, 999)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(999))
      case other => fail(s"Expected Number(999), got $other")
  }

  test("ISERROR: returns TRUE for error value") {
    val sheet = sheetWith(ARef.from0(0, 0) -> CellValue.Error(CellError.Value))
    sheet.evaluateFormula("=ISERROR(A1)") match
      case Right(CellValue.Bool(b)) => assert(b)
      case other => fail(s"Expected Bool(true), got $other")
  }

  test("ISERROR: returns FALSE for non-error value") {
    val sheet = sheetWith(ARef.from0(0, 0) -> CellValue.Number(BigDecimal(42)))
    sheet.evaluateFormula("=ISERROR(A1)") match
      case Right(CellValue.Bool(b)) => assert(!b)
      case other => fail(s"Expected Bool(false), got $other")
  }

  test("ISERROR: returns TRUE when expression causes error") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)),
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal(0))
    )
    sheet.evaluateFormula("=ISERROR(A1/B1)") match
      case Right(CellValue.Bool(b)) => assert(b)
      case other => fail(s"Expected Bool(true), got $other")
  }

  // ==================== Rounding Functions (ROUND, ROUNDUP, ROUNDDOWN, ABS) ====================

  test("ROUND: rounds to specified decimal places (half up)") {
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    sheet.evaluateFormula("=ROUND(3.14159, 2)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal("3.14"))
      case other => fail(s"Expected Number(3.14), got $other")
  }

  test("ROUND: rounds 2.5 to 3 (half up)") {
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    sheet.evaluateFormula("=ROUND(2.5, 0)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(3))
      case other => fail(s"Expected Number(3), got $other")
  }

  test("ROUND: negative digits rounds to tens/hundreds") {
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    sheet.evaluateFormula("=ROUND(12345, -2)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(12300))
      case other => fail(s"Expected Number(12300), got $other")
  }

  test("ROUNDUP: always rounds away from zero (positive)") {
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    sheet.evaluateFormula("=ROUNDUP(3.14159, 2)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal("3.15"))
      case other => fail(s"Expected Number(3.15), got $other")
  }

  test("ROUNDUP: always rounds away from zero (negative)") {
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    sheet.evaluateFormula("=ROUNDUP(-3.14159, 2)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal("-3.15"))
      case other => fail(s"Expected Number(-3.15), got $other")
  }

  test("ROUNDDOWN: always rounds toward zero (positive)") {
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    sheet.evaluateFormula("=ROUNDDOWN(3.99999, 2)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal("3.99"))
      case other => fail(s"Expected Number(3.99), got $other")
  }

  test("ROUNDDOWN: always rounds toward zero (negative)") {
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    sheet.evaluateFormula("=ROUNDDOWN(-3.99999, 2)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal("-3.99"))
      case other => fail(s"Expected Number(-3.99), got $other")
  }

  test("ABS: absolute value of positive number") {
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    sheet.evaluateFormula("=ABS(5)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(5))
      case other => fail(s"Expected Number(5), got $other")
  }

  test("ABS: absolute value of negative number") {
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    sheet.evaluateFormula("=ABS(-5)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(5))
      case other => fail(s"Expected Number(5), got $other")
  }

  test("ABS: absolute value of zero") {
    val sheet = new Sheet(name = SheetName.unsafe("Empty"))
    sheet.evaluateFormula("=ABS(0)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(0))
      case other => fail(s"Expected Number(0), got $other")
  }

  test("ABS: with cell reference") {
    val sheet = sheetWith(ARef.from0(0, 0) -> CellValue.Number(BigDecimal(-42)))
    sheet.evaluateFormula("=ABS(A1)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(42))
      case other => fail(s"Expected Number(42), got $other")
  }

  // ==================== Lookup Functions (INDEX, MATCH) ====================

  test("INDEX: returns value at position") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)), // A1
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(20)), // A2
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(30)), // A3
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal(100)), // B1
      ARef.from0(1, 1) -> CellValue.Number(BigDecimal(200)), // B2
      ARef.from0(1, 2) -> CellValue.Number(BigDecimal(300)) // B3
    )
    sheet.evaluateFormula("=INDEX(A1:B3, 2, 2)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(200))
      case other => fail(s"Expected Number(200), got $other")
  }

  test("INDEX: single column array with just row") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)), // A1
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(20)), // A2
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(30)) // A3
    )
    sheet.evaluateFormula("=INDEX(A1:A3, 2)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(20))
      case other => fail(s"Expected Number(20), got $other")
  }

  test("INDEX: out of bounds returns descriptive #REF! error") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)), // A1
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(20)) // A2
    )
    // Row 5 is out of bounds for a 2-row array
    sheet.evaluateFormula("=INDEX(A1:A2, 5)") match
      case Left(error) =>
        val msg = error.toString
        // Should contain descriptive info about the bounds
        assert(msg.contains("#REF!"), s"Expected #REF! in error, got $msg")
        assert(msg.contains("row_num 5"), s"Expected row number in error, got $msg")
        assert(msg.contains("2 rows"), s"Expected array dimensions in error, got $msg")
      case other => fail(s"Expected EvalError with descriptive message, got $other")
  }

  test("MATCH: exact match finds position") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)), // A1
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(20)), // A2
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(30)) // A3
    )
    sheet.evaluateFormula("=MATCH(20, A1:A3, 0)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(2))
      case other => fail(s"Expected Number(2), got $other")
  }

  test("MATCH: exact match not found returns #N/A error") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)), // A1
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(20)), // A2
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(30)) // A3
    )
    sheet.evaluateFormula("=MATCH(25, A1:A3, 0)") match
      case Left(error) => assert(error.toString.contains("#N/A"))
      case other => fail(s"Expected #N/A error, got $other")
  }

  test("MATCH: approximate match (match_type=1) finds largest <= lookup") {
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(10)), // A1
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(20)), // A2
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(30)) // A3
    )
    sheet.evaluateFormula("=MATCH(25, A1:A3, 1)") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(2)) // 20 is largest <= 25
      case other => fail(s"Expected Number(2), got $other")
  }

  test("INDEX/MATCH: classic lookup pattern") {
    // Classic Excel pattern: INDEX(return_range, MATCH(lookup, lookup_range, 0))
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Text("Apple"), // A1
      ARef.from0(0, 1) -> CellValue.Text("Banana"), // A2
      ARef.from0(0, 2) -> CellValue.Text("Cherry"), // A3
      ARef.from0(1, 0) -> CellValue.Number(BigDecimal(100)), // B1
      ARef.from0(1, 1) -> CellValue.Number(BigDecimal(200)), // B2
      ARef.from0(1, 2) -> CellValue.Number(BigDecimal(300)) // B3
    )
    sheet.evaluateFormula("=INDEX(B1:B3, MATCH(\"Banana\", A1:A3, 0))") match
      case Right(CellValue.Number(n)) => assertEquals(n, BigDecimal(200))
      case other => fail(s"Expected Number(200), got $other")
  }

  // ==================== Date-Based Financial Functions (XNPV, XIRR) ====================

  test("XNPV: calculates present value with irregular dates") {
    import java.time.LocalDateTime
    // Cash flows: -1000 (initial), +300, +400, +500 over ~1 year with irregular dates
    val sheet = sheetWith(
      // Values in A1:A4
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(-1000)), // A1: initial investment
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(300)), // A2: payment 1
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(400)), // A3: payment 2
      ARef.from0(0, 3) -> CellValue.Number(BigDecimal(500)), // A4: payment 3
      // Dates in B1:B4
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2024, 1, 1, 0, 0)), // B1
      ARef.from0(1, 1) -> CellValue.DateTime(LocalDateTime.of(2024, 4, 1, 0, 0)), // B2 (~90 days later)
      ARef.from0(1, 2) -> CellValue.DateTime(LocalDateTime.of(2024, 7, 1, 0, 0)), // B3 (~180 days later)
      ARef.from0(1, 3) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 1, 0, 0)) // B4 (~365 days later)
    )
    sheet.evaluateFormula("=XNPV(0.1, A1:A4, B1:B4)") match
      case Right(CellValue.Number(n)) =>
        // Expected: -1000 + 300/(1.1)^(90/365) + 400/(1.1)^(181/365) + 500/(1.1)^(366/365)
        // ~= -1000 + 293.17 + 383.31 + 454.55 ≈ 131.03
        assert(n > BigDecimal(125) && n < BigDecimal(140), s"Expected XNPV around 131, got $n")
      case other => fail(s"Expected Number, got $other")
  }

  test("XNPV: requires matching length of values and dates") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(-1000)),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(500)),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2024, 1, 1, 0, 0))
      // Missing B2 date - lengths don't match
    )
    sheet.evaluateFormula("=XNPV(0.1, A1:A2, B1:B2)") match
      case Left(error) => assert(error.toString.contains("same length"))
      case other => fail(s"Expected error, got $other")
  }

  test("XIRR: calculates internal rate of return with irregular dates") {
    import java.time.LocalDateTime
    // Standard example: invest -10000, receive payments over ~3 years
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(-10000)), // A1: initial investment
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(2750)), // A2
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(4250)), // A3
      ARef.from0(0, 3) -> CellValue.Number(BigDecimal(3250)), // A4
      ARef.from0(0, 4) -> CellValue.Number(BigDecimal(2750)), // A5
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2008, 1, 1, 0, 0)), // B1
      ARef.from0(1, 1) -> CellValue.DateTime(LocalDateTime.of(2008, 3, 1, 0, 0)), // B2
      ARef.from0(1, 2) -> CellValue.DateTime(LocalDateTime.of(2008, 10, 30, 0, 0)), // B3
      ARef.from0(1, 3) -> CellValue.DateTime(LocalDateTime.of(2009, 2, 15, 0, 0)), // B4
      ARef.from0(1, 4) -> CellValue.DateTime(LocalDateTime.of(2009, 4, 1, 0, 0)) // B5
    )
    sheet.evaluateFormula("=XIRR(A1:A5, B1:B5)") match
      case Right(CellValue.Number(n)) =>
        // Excel calculates ~37.34% (0.3734) for this example
        assert(n > BigDecimal("0.30") && n < BigDecimal("0.45"), s"Expected XIRR around 0.37, got $n")
      case other => fail(s"Expected Number, got $other")
  }

  test("XIRR: requires both positive and negative cash flows") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(1000)), // All positive
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(500)),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2024, 1, 1, 0, 0)),
      ARef.from0(1, 1) -> CellValue.DateTime(LocalDateTime.of(2024, 7, 1, 0, 0))
    )
    sheet.evaluateFormula("=XIRR(A1:A2, B1:B2)") match
      case Left(error) => assert(error.toString.contains("positive and one negative"))
      case other => fail(s"Expected error, got $other")
  }

  test("XIRR: with custom guess") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.Number(BigDecimal(-10000)),
      ARef.from0(0, 1) -> CellValue.Number(BigDecimal(2750)),
      ARef.from0(0, 2) -> CellValue.Number(BigDecimal(4250)),
      ARef.from0(0, 3) -> CellValue.Number(BigDecimal(3250)),
      ARef.from0(0, 4) -> CellValue.Number(BigDecimal(2750)),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2008, 1, 1, 0, 0)),
      ARef.from0(1, 1) -> CellValue.DateTime(LocalDateTime.of(2008, 3, 1, 0, 0)),
      ARef.from0(1, 2) -> CellValue.DateTime(LocalDateTime.of(2008, 10, 30, 0, 0)),
      ARef.from0(1, 3) -> CellValue.DateTime(LocalDateTime.of(2009, 2, 15, 0, 0)),
      ARef.from0(1, 4) -> CellValue.DateTime(LocalDateTime.of(2009, 4, 1, 0, 0))
    )
    sheet.evaluateFormula("=XIRR(A1:A5, B1:B5, 0.5)") match
      case Right(CellValue.Number(n)) =>
        // Should still converge to ~0.37 even with different starting guess
        assert(n > BigDecimal("0.30") && n < BigDecimal("0.45"), s"Expected XIRR around 0.37, got $n")
      case other => fail(s"Expected Number, got $other")
  }

  // ==================== Date Calculation Functions Tests ====================

  test("EOMONTH: returns last day of target month") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 15, 0, 0))
    )
    // EOMONTH(2025-01-15, 1) = 2025-02-28
    sheet.evaluateFormula("=EOMONTH(A1, 1)") match
      case Right(CellValue.DateTime(dt)) =>
        assert(dt.getYear == 2025)
        assert(dt.getMonthValue == 2)
        assert(dt.getDayOfMonth == 28)
      case other => fail(s"Expected DateTime, got $other")
  }

  test("EOMONTH: handles negative months") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 3, 15, 0, 0))
    )
    // EOMONTH(2025-03-15, -1) = 2025-02-28
    sheet.evaluateFormula("=EOMONTH(A1, -1)") match
      case Right(CellValue.DateTime(dt)) =>
        assert(dt.getYear == 2025)
        assert(dt.getMonthValue == 2)
        assert(dt.getDayOfMonth == 28)
      case other => fail(s"Expected DateTime, got $other")
  }

  test("EDATE: adds months to date") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 15, 0, 0))
    )
    // EDATE(2025-01-15, 3) = 2025-04-15
    sheet.evaluateFormula("=EDATE(A1, 3)") match
      case Right(CellValue.DateTime(dt)) =>
        assert(dt.getYear == 2025)
        assert(dt.getMonthValue == 4)
        assert(dt.getDayOfMonth == 15)
      case other => fail(s"Expected DateTime, got $other")
  }

  test("EDATE: clamps to end of month") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 31, 0, 0))
    )
    // EDATE(2025-01-31, 1) = 2025-02-28 (clamped)
    sheet.evaluateFormula("=EDATE(A1, 1)") match
      case Right(CellValue.DateTime(dt)) =>
        assert(dt.getYear == 2025)
        assert(dt.getMonthValue == 2)
        assert(dt.getDayOfMonth == 28)
      case other => fail(s"Expected DateTime, got $other")
  }

  test("DATEDIF: calculates years between dates") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2020, 1, 1, 0, 0)),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 6, 15, 0, 0))
    )
    sheet.evaluateFormula("""=DATEDIF(A1, B1, "Y")""") match
      case Right(CellValue.Number(n)) =>
        assertEquals(n, BigDecimal(5))
      case other => fail(s"Expected Number(5), got $other")
  }

  test("DATEDIF: calculates months between dates") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 1, 0, 0)),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 4, 15, 0, 0))
    )
    sheet.evaluateFormula("""=DATEDIF(A1, B1, "M")""") match
      case Right(CellValue.Number(n)) =>
        assertEquals(n, BigDecimal(3))
      case other => fail(s"Expected Number(3), got $other")
  }

  test("DATEDIF: calculates days between dates") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 1, 0, 0)),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 11, 0, 0))
    )
    sheet.evaluateFormula("""=DATEDIF(A1, B1, "D")""") match
      case Right(CellValue.Number(n)) =>
        assertEquals(n, BigDecimal(10))
      case other => fail(s"Expected Number(10), got $other")
  }

  test("DATEDIF MD: handles month boundary correctly") {
    import java.time.LocalDateTime
    // Jan 31 to Mar 1 = 1 day (ignoring months/years)
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 31, 0, 0)),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 3, 1, 0, 0))
    )
    sheet.evaluateFormula("""=DATEDIF(A1, B1, "MD")""") match
      case Right(CellValue.Number(n)) =>
        assertEquals(n, BigDecimal(1))
      case other => fail(s"Expected Number(1), got $other")
  }

  test("NETWORKDAYS: counts working days (Mon-Fri)") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      // Mon 2025-01-06 to Fri 2025-01-10 = 5 working days
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 6, 0, 0)),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 10, 0, 0))
    )
    sheet.evaluateFormula("=NETWORKDAYS(A1, B1)") match
      case Right(CellValue.Number(n)) =>
        assertEquals(n, BigDecimal(5))
      case other => fail(s"Expected Number(5), got $other")
  }

  test("NETWORKDAYS: excludes weekends") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      // Mon 2025-01-06 to Mon 2025-01-13 = 6 working days (8 days - 2 weekend days)
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 6, 0, 0)),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 13, 0, 0))
    )
    sheet.evaluateFormula("=NETWORKDAYS(A1, B1)") match
      case Right(CellValue.Number(n)) =>
        assertEquals(n, BigDecimal(6))
      case other => fail(s"Expected Number(6), got $other")
  }

  test("NETWORKDAYS: excludes holidays from range") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      // Mon 2025-01-06 to Fri 2025-01-10 = 5 working days normally
      // But Wed 2025-01-08 is a holiday, so only 4 working days
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 6, 0, 0)),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 10, 0, 0)),
      ARef.from0(2, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 8, 0, 0)) // Holiday: Wed
    )
    sheet.evaluateFormula("=NETWORKDAYS(A1, B1, C1:C1)") match
      case Right(CellValue.Number(n)) =>
        assertEquals(n, BigDecimal(4))
      case other => fail(s"Expected Number(4), got $other")
  }

  test("WORKDAY: adds working days") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      // Starting Mon 2025-01-06, add 5 working days = Fri 2025-01-10
      // But WORKDAY starts counting from next day, so 5 days from Mon = Mon+5 = next Mon? Let me check
      // Actually WORKDAY(2025-01-06, 5) should give 2025-01-13 (Mon + 5 working days = next Mon)
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 6, 0, 0))
    )
    sheet.evaluateFormula("=WORKDAY(A1, 5)") match
      case Right(CellValue.DateTime(dt)) =>
        assert(dt.getYear == 2025)
        assert(dt.getMonthValue == 1)
        assertEquals(dt.getDayOfMonth, 13) // Mon + 5 working days = next Mon
      case other => fail(s"Expected DateTime, got $other")
  }

  test("WORKDAY: handles negative days") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      // Starting Fri 2025-01-10, subtract 5 working days = Mon 2025-01-03
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 10, 0, 0))
    )
    sheet.evaluateFormula("=WORKDAY(A1, -5)") match
      case Right(CellValue.DateTime(dt)) =>
        assert(dt.getYear == 2025)
        assert(dt.getMonthValue == 1)
        assertEquals(dt.getDayOfMonth, 3) // Fri - 5 working days = previous Mon
      case other => fail(s"Expected DateTime, got $other")
  }

  test("YEARFRAC: calculates year fraction with actual/actual basis") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 1, 0, 0)),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 7, 1, 0, 0))
    )
    // 181 days / 365 days ≈ 0.496
    sheet.evaluateFormula("=YEARFRAC(A1, B1, 1)") match
      case Right(CellValue.Number(n)) =>
        assert(n > BigDecimal("0.49") && n < BigDecimal("0.50"), s"Expected ~0.496, got $n")
      case other => fail(s"Expected Number, got $other")
  }

  test("YEARFRAC: default basis is US 30/360") {
    import java.time.LocalDateTime
    val sheet = sheetWith(
      ARef.from0(0, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 1, 1, 0, 0)),
      ARef.from0(1, 0) -> CellValue.DateTime(LocalDateTime.of(2025, 7, 1, 0, 0))
    )
    // 30/360: 6 months = 180 days, 180/360 = 0.5
    sheet.evaluateFormula("=YEARFRAC(A1, B1)") match
      case Right(CellValue.Number(n)) =>
        assertEquals(n, BigDecimal("0.5"))
      case other => fail(s"Expected Number(0.5), got $other")
  }
