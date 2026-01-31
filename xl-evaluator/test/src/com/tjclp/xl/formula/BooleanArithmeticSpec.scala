package com.tjclp.xl.formula

import munit.FunSuite
import com.tjclp.xl.*
import com.tjclp.xl.unsafe.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet

/**
 * Tests for GH-196: Boolean arithmetic coercion.
 *
 * Excel treats booleans as numbers in arithmetic contexts:
 *   - TRUE → 1
 *   - FALSE → 0
 *
 * This enables expressions like `(5>3)*10` to evaluate to `10`.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class BooleanArithmeticSpec extends FunSuite:

  // Helper to create a sheet with data
  private def sheetWith(entries: (ARef, Any)*): Sheet =
    entries.foldLeft(Sheet("Test")) { case (s, (ref, value)) =>
      val cv = value match
        case cv: CellValue => cv
        case str: String   => CellValue.Text(str)
        case n: Int        => CellValue.Number(BigDecimal(n))
        case n: Long       => CellValue.Number(BigDecimal(n))
        case n: Double     => CellValue.Number(BigDecimal(n))
        case n: BigDecimal => CellValue.Number(n)
        case b: Boolean    => CellValue.Bool(b)
        case _             => CellValue.Text(value.toString)
      s.put(ref, cv).unsafe
    }

  // Helper to parse and evaluate a formula returning BigDecimal
  private def evalFormula(
    formula: String,
    sheet: Sheet = Sheet("Test")
  ): Either[String, BigDecimal] =
    FormulaParser.parse(formula) match
      case Right(expr) =>
        Evaluator.eval(expr, sheet) match
          case Right(value: BigDecimal) => Right(value)
          case Right(other)             => Left(s"Expected BigDecimal, got: $other")
          case Left(err)                => Left(s"Eval error: $err")
      case Left(err) => Left(s"Parse error: $err")

  // Helper for expected results
  private def assertEval(formula: String, expected: BigDecimal)(implicit
    loc: munit.Location
  ): Unit =
    evalFormula(formula) match
      case Right(value) => assertEquals(value, expected)
      case Left(err)    => fail(s"Formula evaluation failed: $err")

  private def assertEval(formula: String, expected: Int)(implicit
    loc: munit.Location
  ): Unit =
    assertEval(formula, BigDecimal(expected))

  private def assertEvalWithSheet(formula: String, sheet: Sheet, expected: BigDecimal)(implicit
    loc: munit.Location
  ): Unit =
    evalFormula(formula, sheet) match
      case Right(value) => assertEquals(value, expected)
      case Left(err)    => fail(s"Formula evaluation failed: $err")

  // ===== GH-196: Boolean Arithmetic Coercion =====

  test("(5>3)*10 = 10 (TRUE coerced to 1)") {
    assertEval("=(5>3)*10", 10)
  }

  test("(3>5)*10 = 0 (FALSE coerced to 0)") {
    assertEval("=(3>5)*10", 0)
  }

  test("TRUE+TRUE = 2") {
    assertEval("=TRUE+TRUE", 2)
  }

  test("FALSE+1 = 1") {
    assertEval("=FALSE+1", 1)
  }

  test("TRUE-FALSE = 1") {
    assertEval("=TRUE-FALSE", 1)
  }

  test("(1=1)+(2=2) = 2") {
    assertEval("=(1=1)+(2=2)", 2)
  }

  test("(1<>1)+(2=2) = 1") {
    assertEval("=(1<>1)+(2=2)", 1)
  }

  test("5/(1=1) = 5 (division by TRUE)") {
    assertEval("=5/(1=1)", 5)
  }

  test("5-(3>2) = 4") {
    assertEval("=5-(3>2)", 4)
  }

  test("nested boolean arithmetic: ((5>3)+(2<1))*10 = 10") {
    assertEval("=((5>3)+(2<1))*10", 10)
  }

  test("TRUE*TRUE*TRUE = 1") {
    assertEval("=TRUE*TRUE*TRUE", 1)
  }

  test("FALSE*100 = 0") {
    assertEval("=FALSE*100", 0)
  }

  test("(10>5)*(20>10)*(30>20) = 1 (all TRUE)") {
    assertEval("=(10>5)*(20>10)*(30>20)", 1)
  }

  test("(10>5)*(20<10)*(30>20) = 0 (one FALSE)") {
    assertEval("=(10>5)*(20<10)*(30>20)", 0)
  }

  // ===== Boolean Cell Reference Tests =====

  test("boolean cell reference in arithmetic: A1*10 where A1=TRUE") {
    val sheet = sheetWith(ref"A1" -> true)
    assertEvalWithSheet("=A1*10", sheet, BigDecimal(10))
  }

  test("boolean cell reference in arithmetic: A1*10 where A1=FALSE") {
    val sheet = sheetWith(ref"A1" -> false)
    assertEvalWithSheet("=A1*10", sheet, BigDecimal(0))
  }

  test("boolean cell reference addition: A1+A2 where A1=TRUE, A2=TRUE") {
    val sheet = sheetWith(ref"A1" -> true, ref"A2" -> true)
    assertEvalWithSheet("=A1+A2", sheet, BigDecimal(2))
  }

  test("mixed boolean and numeric: A1+B1 where A1=TRUE, B1=5") {
    val sheet = sheetWith(ref"A1" -> true, ref"B1" -> 5)
    assertEvalWithSheet("=A1+B1", sheet, BigDecimal(6))
  }

  // ===== Comparison Result in Arithmetic =====

  test("ROW comparison in arithmetic context") {
    // Simulating ROW(G2) > ROW($G$1) pattern from GH-196
    val sheet = sheetWith(ref"A1" -> 1, ref"A2" -> 2)
    // (A2>A1)*5 should be (2>1)*5 = TRUE*5 = 5
    assertEvalWithSheet("=(A2>A1)*5", sheet, BigDecimal(5))
  }

  test("complex conditional sum pattern") {
    // Pattern: (condition)*value - common in SUMPRODUCT-like formulas
    val sheet = sheetWith(
      ref"A1" -> 10,
      ref"A2" -> 20,
      ref"B1" -> 5,
      ref"B2" -> 3
    )
    // (B1>4)*A1 + (B2>4)*A2 = (5>4)*10 + (3>4)*20 = 1*10 + 0*20 = 10
    assertEvalWithSheet("=(B1>4)*A1+(B2>4)*A2", sheet, BigDecimal(10))
  }
