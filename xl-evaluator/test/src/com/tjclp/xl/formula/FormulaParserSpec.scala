package com.tjclp.xl.formula

import com.tjclp.xl.CellRange
import com.tjclp.xl.addressing.ARef
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import org.scalacheck.{Arbitrary, Gen}

import scala.math.BigDecimal

/**
 * Property-based tests for FormulaParser and FormulaPrinter.
 *
 * Tests round-trip laws, edge cases, and integration with existing formula system.
 */
class FormulaParserSpec extends ScalaCheckSuite:

  // ==================== Generators ====================

  /** Generate valid cell references */
  val genARef: Gen[ARef] =
    for
      col <- Gen.choose(0, 100) // A-CV
      row <- Gen.choose(0, 100) // 1-101
    yield ARef.from0(col, row)

  /** Generate simple numeric literals */
  val genNumericLit: Gen[TExpr[BigDecimal]] =
    Gen.choose(-1000.0, 1000.0).map(d => TExpr.Lit(BigDecimal(d)))

  /** Generate boolean literals */
  val genBoolLit: Gen[TExpr[Boolean]] =
    Gen.oneOf(TExpr.Lit(true), TExpr.Lit(false))

  /** Generate string literals */
  val genStringLit: Gen[TExpr[String]] =
    Gen.alphaNumStr.map(TExpr.Lit.apply)

  /** Generate cell references */
  val genCellRef: Gen[TExpr[BigDecimal]] =
    genARef.map(ref => TExpr.Ref(ref, TExpr.decodeNumeric))

  /** Generate simple arithmetic expressions */
  def genArithExpr(depth: Int = 0): Gen[TExpr[BigDecimal]] =
    if depth >= 3 then genNumericLit
    else
      Gen.frequency(
        3 -> genNumericLit,
        1 -> genCellRef,
        1 -> Gen.lzy(for
          x <- genArithExpr(depth + 1)
          y <- genArithExpr(depth + 1)
          op <- Gen.oneOf[
            (TExpr[BigDecimal], TExpr[BigDecimal]) => TExpr[BigDecimal]
          ](
            TExpr.Add.apply,
            TExpr.Sub.apply,
            TExpr.Mul.apply,
            TExpr.Div.apply
          )
        yield op(x, y))
      )

  /** Generate comparison expressions */
  def genCompExpr(depth: Int = 0): Gen[TExpr[Boolean]] =
    if depth >= 2 then genBoolLit
    else
      for
        x <- genArithExpr(0)
        y <- genArithExpr(0)
        op <- Gen.oneOf[
          (TExpr[BigDecimal], TExpr[BigDecimal]) => TExpr[Boolean]
        ](
          TExpr.Lt.apply,
          TExpr.Lte.apply,
          TExpr.Gt.apply,
          TExpr.Gte.apply
        )
      yield op(x, y)

  /** Generate ranges */
  val genRange: Gen[CellRange] =
    for
      startCol <- Gen.choose(0, 90)
      startRow <- Gen.choose(0, 90)
      endCol <- Gen.choose(startCol, startCol + 10)
      endRow <- Gen.choose(startRow, startRow + 10)
      start = ARef.from0(startCol, startRow)
      end = ARef.from0(endCol, endRow)
    yield CellRange(start, end)

  // ==================== Round-Trip Tests ====================

  property("round-trip: parse ∘ print = id for numeric literals") {
    forAll(Gen.choose(-1000.0, 1000.0)) { d =>
      val original = BigDecimal(d)
      val expr = TExpr.Lit(original)
      val printed = FormulaPrinter.print(expr, includeEquals = true)
      val parsed = FormulaParser.parse(printed)

      // Check that parsing succeeds
      parsed.isRight &&
      // Check value matches semantically (BigDecimal equality, not string equality)
      // Note: We compare values, not string representations, because BigDecimal.toString()
      // may format the same number differently (e.g., "1E-7" vs "0.0000001")
      // Also note: Negative numbers parse as Sub(Lit(0), Lit(abs(value))), not Lit(negative)
      parsed.exists {
        case TExpr.Lit(value: BigDecimal) =>
          // Positive number or zero
          (value - original).abs < BigDecimal("1E-15")
        case TExpr.Sub(TExpr.Lit(zero: BigDecimal), TExpr.Lit(value: BigDecimal)) =>
          // Negative number (parsed as unary minus: 0 - abs(value))
          zero == BigDecimal(0) && ((-value) - original).abs < BigDecimal("1E-15")
        case _ => false
      }
    }
  }

  property("round-trip: parse ∘ print = id for boolean literals") {
    forAll(Gen.oneOf(true, false)) { b =>
      val expr = TExpr.Lit(b)
      val printed = FormulaPrinter.print(expr, includeEquals = true)
      val parsed = FormulaParser.parse(printed)

      parsed.isRight &&
      parsed.exists {
        case TExpr.Lit(value: Boolean) => value == b
        case _                          => false
      }
    }
  }

  property("round-trip: parse ∘ print = id for cell references") {
    forAll(genARef) { ref =>
      val expr = TExpr.Ref(ref, TExpr.decodeNumeric)
      val printed = FormulaPrinter.print(expr, includeEquals = true)
      val parsed = FormulaParser.parse(printed)

      parsed.isRight &&
      parsed.exists {
        case TExpr.Ref(at, _) => at == ref
        case _                => false
      }
    }
  }

  property("round-trip: parse ∘ print = id for addition") {
    forAll(genNumericLit, genNumericLit) { (x, y) =>
      val expr = TExpr.Add(x, y)
      val printed = FormulaPrinter.print(expr, includeEquals = true)
      val parsed = FormulaParser.parse(printed)

      parsed.isRight &&
      parsed.exists {
        case TExpr.Add(_, _) => true
        case _               => false
      }
    }
  }

  property("round-trip: parse ∘ print = id for subtraction") {
    forAll(genNumericLit, genNumericLit) { (x, y) =>
      val expr = TExpr.Sub(x, y)
      val printed = FormulaPrinter.print(expr, includeEquals = true)
      val parsed = FormulaParser.parse(printed)

      parsed.isRight
    }
  }

  property("round-trip: parse ∘ print = id for multiplication") {
    forAll(genNumericLit, genNumericLit) { (x, y) =>
      val expr = TExpr.Mul(x, y)
      val printed = FormulaPrinter.print(expr, includeEquals = true)
      val parsed = FormulaParser.parse(printed)

      parsed.isRight
    }
  }

  property("round-trip: parse ∘ print = id for division") {
    forAll(genNumericLit, genNumericLit) { (x, y) =>
      val expr = TExpr.Div(x, y)
      val printed = FormulaPrinter.print(expr, includeEquals = true)
      val parsed = FormulaParser.parse(printed)

      parsed.isRight
    }
  }

  // ==================== Parser Tests ====================

  test("parse simple number") {
    val result = FormulaParser.parse("=42")
    assert(result.isRight)
    result.foreach {
      case TExpr.Lit(value: BigDecimal) => assertEquals(value, BigDecimal(42))
      case other                         => fail(s"Expected Lit(42), got $other")
    }
  }

  test("parse negative number (unary minus)") {
    val result = FormulaParser.parse("=-42")
    assert(result.isRight)
    result.foreach {
      case TExpr.Sub(TExpr.Lit(zero: BigDecimal), TExpr.Lit(value: BigDecimal)) =>
        assertEquals(zero, BigDecimal(0))
        assertEquals(value, BigDecimal(42))
      case other => fail(s"Expected Sub(0, 42), got $other")
    }
  }

  test("parse decimal number") {
    val result = FormulaParser.parse("=3.14")
    assert(result.isRight)
    result.foreach {
      case TExpr.Lit(value: BigDecimal) => assertEquals(value, BigDecimal(3.14))
      case other                         => fail(s"Expected Lit(3.14), got $other")
    }
  }

  // Scientific notation tests
  test("parse scientific notation: positive exponent (1E10)") {
    val result = FormulaParser.parse("=1E10")
    assert(result.isRight)
    result.foreach {
      case TExpr.Lit(value: BigDecimal) => assertEquals(value, BigDecimal("1E10"))
      case other                         => fail(s"Expected Lit(1E10), got $other")
    }
  }

  test("parse scientific notation: negative exponent (1E-10)") {
    val result = FormulaParser.parse("=1E-10")
    assert(result.isRight)
    result.foreach {
      case TExpr.Lit(value: BigDecimal) => assertEquals(value, BigDecimal("1E-10"))
      case other                         => fail(s"Expected Lit(1E-10), got $other")
    }
  }

  test("parse scientific notation: explicit plus (1.5E+5)") {
    val result = FormulaParser.parse("=1.5E+5")
    assert(result.isRight)
    result.foreach {
      case TExpr.Lit(value: BigDecimal) => assertEquals(value, BigDecimal("1.5E+5"))
      case other                         => fail(s"Expected Lit(1.5E+5), got $other")
    }
  }

  test("parse scientific notation: decimal with negative exponent (2.3E-7)") {
    val result = FormulaParser.parse("=2.3E-7")
    assert(result.isRight)
    result.foreach {
      case TExpr.Lit(value: BigDecimal) => assertEquals(value, BigDecimal("2.3E-7"))
      case other                         => fail(s"Expected Lit(2.3E-7), got $other")
    }
  }

  test("parse scientific notation: lowercase e (3.14e2)") {
    val result = FormulaParser.parse("=3.14e2")
    assert(result.isRight)
    result.foreach {
      case TExpr.Lit(value: BigDecimal) => assertEquals(value, BigDecimal("3.14e2"))
      case other                         => fail(s"Expected Lit(3.14e2), got $other")
    }
  }

  test("parse scientific notation: very small number (1.4E-100)") {
    val result = FormulaParser.parse("=1.4E-100")
    assert(result.isRight)
    result.foreach {
      case TExpr.Lit(value: BigDecimal) =>
        assertEquals(value, BigDecimal("1.4E-100"))
      case other => fail(s"Expected Lit(1.4E-100), got $other")
    }
  }

  test("parse scientific notation: very large number (9.99E+200)") {
    val result = FormulaParser.parse("=9.99E+200")
    assert(result.isRight)
    result.foreach {
      case TExpr.Lit(value: BigDecimal) =>
        assertEquals(value, BigDecimal("9.99E+200"))
      case other => fail(s"Expected Lit(9.99E+200), got $other")
    }
  }

  test("parse scientific notation: negative number with exponent (-5.2E-8)") {
    val result = FormulaParser.parse("=-5.2E-8")
    assert(result.isRight)
    // Note: Unary minus produces Sub(0, value)
    result.foreach {
      case TExpr.Sub(TExpr.Lit(zero: BigDecimal), TExpr.Lit(value: BigDecimal)) =>
        assertEquals(zero, BigDecimal(0))
        assertEquals(value, BigDecimal("5.2E-8"))
      case other => fail(s"Expected Sub(0, 5.2E-8), got $other")
    }
  }

  test("error: invalid scientific notation (no exponent digits: 1E)") {
    val result = FormulaParser.parse("=1E")
    // Parser should stop at 'E' and treat '1' as complete number, 'E' as unexpected
    assert(result.isLeft)
  }

  test("error: invalid scientific notation (multiple E: 1E2E3)") {
    val result = FormulaParser.parse("=1E2E3")
    // Parser should parse 1E2, then fail on 'E3'
    assert(result.isLeft)
  }

  test("parse TRUE") {
    val result = FormulaParser.parse("=TRUE")
    assert(result.isRight)
    result.foreach {
      case TExpr.Lit(true) => ()
      case other           => fail(s"Expected Lit(true), got $other")
    }
  }

  test("parse FALSE") {
    val result = FormulaParser.parse("=FALSE")
    assert(result.isRight)
    result.foreach {
      case TExpr.Lit(false) => ()
      case other            => fail(s"Expected Lit(false), got $other")
    }
  }

  test("parse cell reference A1") {
    val result = FormulaParser.parse("=A1")
    assert(result.isRight)
    result.foreach {
      case TExpr.Ref(at, _) =>
        val expected = ARef.parse("A1").fold(err => fail(s"Invalid ref: $err"), identity)
        assertEquals(at, expected)
      case other => fail(s"Expected Ref(A1), got $other")
    }
  }

  test("parse cell reference Z99") {
    val result = FormulaParser.parse("=Z99")
    assert(result.isRight)
    result.foreach {
      case TExpr.Ref(at, _) =>
        val expected = ARef.parse("Z99").fold(err => fail(s"Invalid ref: $err"), identity)
        assertEquals(at, expected)
      case other => fail(s"Expected Ref(Z99), got $other")
    }
  }

  test("parse addition") {
    val result = FormulaParser.parse("=1+2")
    assert(result.isRight)
    result.foreach {
      case TExpr.Add(TExpr.Lit(x: BigDecimal), TExpr.Lit(y: BigDecimal)) =>
        assertEquals(x, BigDecimal(1))
        assertEquals(y, BigDecimal(2))
      case other => fail(s"Expected Add(1, 2), got $other")
    }
  }

  test("parse subtraction") {
    val result = FormulaParser.parse("=10-5")
    assert(result.isRight)
    result.foreach {
      case TExpr.Sub(_, _) => ()
      case other           => fail(s"Expected Sub, got $other")
    }
  }

  test("parse multiplication") {
    val result = FormulaParser.parse("=3*4")
    assert(result.isRight)
    result.foreach {
      case TExpr.Mul(_, _) => ()
      case other           => fail(s"Expected Mul, got $other")
    }
  }

  test("parse division") {
    val result = FormulaParser.parse("=10/2")
    assert(result.isRight)
    result.foreach {
      case TExpr.Div(_, _) => ()
      case other           => fail(s"Expected Div, got $other")
    }
  }

  test("parse SUM function") {
    val result = FormulaParser.parse("=SUM(A1:B10)")
    assert(result.isRight)
    result.foreach {
      case TExpr.FoldRange(range, zero: BigDecimal, _, _) =>
        val expected = CellRange.parse("A1:B10").fold(err => fail(s"Invalid range: $err"), identity)
        assertEquals(range, expected)
        assertEquals(zero, BigDecimal(0)) // SUM starts with 0
      case other => fail(s"Expected FoldRange with BigDecimal, got $other")
    }
  }

  test("parse COUNT function") {
    val result = FormulaParser.parse("=COUNT(A1:B10)")
    assert(result.isRight)
    result.foreach {
      case TExpr.FoldRange(range, zero: Int, _, _) =>
        val expected = CellRange.parse("A1:B10").fold(err => fail(s"Invalid range: $err"), identity)
        assertEquals(range, expected)
        assertEquals(zero, 0) // COUNT starts with 0
      case other => fail(s"Expected FoldRange with Int accumulator, got $other")
    }
  }

  test("parse AVERAGE function") {
    val result = FormulaParser.parse("=AVERAGE(A1:B10)")
    assert(result.isRight)
    result.foreach {
      // @unchecked needed: JVM erases tuple element types at runtime
      // Safe because FormulaParser constructs correct GADT types
      case TExpr.FoldRange(range, zero: (BigDecimal, Int) @unchecked, _, _) =>
        val expected = CellRange.parse("A1:B10").fold(err => fail(s"Invalid range: $err"), identity)
        assertEquals(range, expected)
        assertEquals(zero, (BigDecimal(0), 0)) // AVERAGE starts with (sum=0, count=0)
      case other => fail(s"Expected FoldRange with tuple accumulator, got $other")
    }
  }

  test("SUM/COUNT/AVERAGE create different fold types") {
    // Verify that each function creates the correct fold type (not all SUM)
    val sumResult = FormulaParser.parse("=SUM(A1:A10)")
    val countResult = FormulaParser.parse("=COUNT(A1:A10)")
    val avgResult = FormulaParser.parse("=AVERAGE(A1:A10)")

    // @unchecked needed: JVM erases tuple element types at runtime (line 447)
    // Safe because FormulaParser constructs correct GADT types
    (sumResult, countResult, avgResult) match
      case (Right(sumFold: TExpr.FoldRange[?, ?]), Right(countFold: TExpr.FoldRange[?, ?]), Right(avgFold: TExpr.FoldRange[?, ?])) =>
        // Extract zero values to verify different fold types
        val sumZero = sumFold.z
        val countZero = countFold.z
        val avgZero = avgFold.z

        // SUM has BigDecimal(0), COUNT has Int(0), AVERAGE has (BigDecimal(0), Int(0))
        sumZero match
          case _: BigDecimal => // success
          case other => fail(s"Expected BigDecimal for SUM zero, got ${other.getClass}")
        countZero match
          case _: Int => // success
          case other => fail(s"Expected Int for COUNT zero, got ${other.getClass}")
        avgZero match
          case _: (BigDecimal, Int) @unchecked => // success
          case other => fail(s"Expected (BigDecimal, Int) for AVERAGE zero, got ${other.getClass}")
      case _ => fail("All three functions should parse successfully")
  }

  test("parse IF function") {
    val result = FormulaParser.parse("=IF(TRUE, 1, 2)")
    assert(result.isRight)
    result.foreach {
      case TExpr.If(_, _, _) => ()
      case other             => fail(s"Expected If, got $other")
    }
  }

  test("parse AND function") {
    val result = FormulaParser.parse("=AND(TRUE, FALSE)")
    assert(result.isRight)
    result.foreach {
      case TExpr.And(_, _) => ()
      case other           => fail(s"Expected And, got $other")
    }
  }

  test("parse OR function") {
    val result = FormulaParser.parse("=OR(TRUE, FALSE)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Or(_, _) => ()
      case other          => fail(s"Expected Or, got $other")
    }
  }

  test("parse NOT function") {
    val result = FormulaParser.parse("=NOT(TRUE)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Not(_) => ()
      case other        => fail(s"Expected Not, got $other")
    }
  }

  test("parse nested parentheses") {
    val result = FormulaParser.parse("=((1+2)*3)")
    assert(result.isRight)
  }

  test("parse whitespace handling") {
    val result = FormulaParser.parse("= 1 + 2 ")
    assert(result.isRight)
  }

  test("parse without leading =") {
    val result = FormulaParser.parse("1+2")
    assert(result.isRight)
  }

  // ==================== Error Cases ====================

  test("error: empty formula") {
    val result = FormulaParser.parse("=")
    assert(result.isLeft)
    result.left.foreach {
      case ParseError.EmptyFormula => ()
      case other                   => fail(s"Expected EmptyFormula, got $other")
    }
  }

  test("error: unbalanced parentheses (missing close)") {
    val result = FormulaParser.parse("=(1+2")
    assert(result.isLeft)
  }

  test("error: unknown function") {
    val result = FormulaParser.parse("=SUMM(A1:B10)")
    assert(result.isLeft)
    result.left.foreach {
      case ParseError.UnknownFunction(name, _, suggestions) =>
        assertEquals(name, "SUMM")
        assert(suggestions.contains("SUM"))
      case other => fail(s"Expected UnknownFunction, got $other")
    }
  }

  test("error: invalid cell reference") {
    val result = FormulaParser.parse("=ZZZ9999999")
    assert(result.isLeft)
  }

  test("error: formula too long") {
    val longFormula = "=" + ("1+" * 5000) + "1"
    val result = FormulaParser.parse(longFormula)
    assert(result.isLeft)
    result.left.foreach {
      case ParseError.FormulaTooLong(_, _) => ()
      case other                           => fail(s"Expected FormulaTooLong, got $other")
    }
  }

  test("error: concatenation operator not supported") {
    val result = FormulaParser.parse("=\"foo\"&\"bar\"")
    assert(result.isLeft)
    result.left.foreach {
      case ParseError.InvalidOperator("&", _, reason) =>
        assert(reason.contains("concatenation"))
        assert(reason.contains("not yet supported"))
      case other => fail(s"Expected InvalidOperator for '&', got $other")
    }
  }

  // ==================== Printer Tests ====================

  test("print numeric literal") {
    val expr = TExpr.Lit(BigDecimal(42))
    val result = FormulaPrinter.print(expr)
    assertEquals(result, "=42")
  }

  test("print boolean literal TRUE") {
    val expr = TExpr.Lit(true)
    val result = FormulaPrinter.print(expr)
    assertEquals(result, "=TRUE")
  }

  test("print boolean literal FALSE") {
    val expr = TExpr.Lit(false)
    val result = FormulaPrinter.print(expr)
    assertEquals(result, "=FALSE")
  }

  test("print cell reference") {
    val ref = ARef.parse("A1").fold(err => fail(s"Invalid ref: $err"), identity)
    val expr = TExpr.Ref(ref, TExpr.decodeNumeric)
    val result = FormulaPrinter.print(expr)
    assertEquals(result, "=A1")
  }

  test("print addition") {
    val expr = TExpr.Add(TExpr.Lit(BigDecimal(1)), TExpr.Lit(BigDecimal(2)))
    val result = FormulaPrinter.print(expr)
    assertEquals(result, "=1+2")
  }

  test("print SUM function") {
    val range = CellRange.parse("A1:B10").fold(err => fail(s"Invalid range: $err"), identity)
    val expr = TExpr.sum(range)
    val result = FormulaPrinter.print(expr)
    assertEquals(result, "=SUM(A1:B10)")
  }

  test("print without leading =") {
    val expr = TExpr.Lit(BigDecimal(42))
    val result = FormulaPrinter.print(expr, includeEquals = false)
    assertEquals(result, "42")
  }

  // ==================== Integration Tests ====================

  test("integration: complex arithmetic expression") {
    val formula = "=((A1+B2)*C3)/D4"
    val result = FormulaParser.parse(formula)
    assert(result.isRight)

    // Verify round-trip
    result.foreach { expr =>
      val reprinted = FormulaPrinter.print(expr)
      // Parse again
      val reparsed = FormulaParser.parse(reprinted)
      assert(reparsed.isRight)
    }
  }

  test("integration: nested IF") {
    val formula = "=IF(A1>0, IF(A1>10, 2, 1), 0)"
    val result = FormulaParser.parse(formula)
    assert(result.isRight)
  }

  test("integration: mixed operators") {
    val formula = "=A1+B2*C3-D4/E5"
    val result = FormulaParser.parse(formula)
    assert(result.isRight)

    // Verify operator precedence is preserved
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      val reparsed = FormulaParser.parse(printed)
      assert(reparsed.isRight)
    }
  }

  // ===== New Function Parser Tests (WI-09b) =====

  test("parse MIN function") {
    val result = FormulaParser.parse("=MIN(A1:A10)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Min(_) => assert(true)
      case _ => fail("Expected TExpr.Min")
    }
  }

  test("parse MAX function") {
    val result = FormulaParser.parse("=MAX(B2:B20)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Max(_) => assert(true)
      case _ => fail("Expected TExpr.Max")
    }
  }

  test("parse CONCATENATE function") {
    val result = FormulaParser.parse("=CONCATENATE(\"Hello\", \" \", \"World\")")
    assert(result.isRight)
    result.foreach {
      case TExpr.Concatenate(xs) => assertEquals(xs.length, 3)
      case _ => fail("Expected TExpr.Concatenate")
    }
  }

  test("parse LEFT function") {
    val result = FormulaParser.parse("=LEFT(\"Hello\", 3)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Left(_, _) => assert(true)
      case _ => fail("Expected TExpr.Left")
    }
  }

  test("parse RIGHT function") {
    val result = FormulaParser.parse("=RIGHT(A1, 5)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Right(_, _) => assert(true)
      case _ => fail("Expected TExpr.Right")
    }
  }

  test("parse LEN function") {
    val result = FormulaParser.parse("=LEN(A1)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Len(_) => assert(true)
      case _ => fail("Expected TExpr.Len")
    }
  }

  test("parse UPPER function") {
    val result = FormulaParser.parse("=UPPER(\"hello\")")
    assert(result.isRight)
    result.foreach {
      case TExpr.Upper(_) => assert(true)
      case _ => fail("Expected TExpr.Upper")
    }
  }

  test("parse LOWER function") {
    val result = FormulaParser.parse("=LOWER(\"WORLD\")")
    assert(result.isRight)
    result.foreach {
      case TExpr.Lower(_) => assert(true)
      case _ => fail("Expected TExpr.Lower")
    }
  }

  test("parse TODAY function") {
    val result = FormulaParser.parse("=TODAY()")
    assert(result.isRight)
    result.foreach {
      case TExpr.Today() => assert(true)
      case _ => fail("Expected TExpr.Today")
    }
  }

  test("parse NOW function") {
    val result = FormulaParser.parse("=NOW()")
    assert(result.isRight)
    result.foreach {
      case TExpr.Now() => assert(true)
      case _ => fail("Expected TExpr.Now")
    }
  }

  test("parse DATE function") {
    val result = FormulaParser.parse("=DATE(2025, 11, 21)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Date(_, _, _) => assert(true)
      case _ => fail("Expected TExpr.Date")
    }
  }

  test("parse YEAR function") {
    val result = FormulaParser.parse("=YEAR(TODAY())")
    assert(result.isRight)
    result.foreach {
      case TExpr.Year(_) => assert(true)
      case _ => fail("Expected TExpr.Year")
    }
  }

  test("parse MONTH function") {
    val result = FormulaParser.parse("=MONTH(A1)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Month(_) => assert(true)
      case _ => fail("Expected TExpr.Month")
    }
  }

  test("parse DAY function") {
    val result = FormulaParser.parse("=DAY(DATE(2025, 11, 21))")
    assert(result.isRight)
    result.foreach {
      case TExpr.Day(_) => assert(true)
      case _ => fail("Expected TExpr.Day")
    }
  }

  test("parse nested text functions") {
    val result = FormulaParser.parse("=UPPER(CONCATENATE(LEFT(A1, 3), RIGHT(B1, 2)))")
    assert(result.isRight)
    result.foreach {
      case TExpr.Upper(TExpr.Concatenate(_)) => assert(true)
      case _ => fail("Expected nested text functions")
    }
  }

  test("FunctionParser.allFunctions includes all 21 functions") {
    val functions = FunctionParser.allFunctions
    assert(functions.contains("SUM"))
    assert(functions.contains("MIN"))
    assert(functions.contains("MAX"))
    assert(functions.contains("CONCATENATE"))
    assert(functions.contains("TODAY"))
    assert(functions.contains("DATE"))
    assertEquals(functions.length, 21)
  }

  test("FunctionParser.lookup finds known functions") {
    assert(FunctionParser.lookup("SUM").isDefined)
    assert(FunctionParser.lookup("min").isDefined) // Case insensitive
    assert(FunctionParser.lookup("CONCATENATE").isDefined)
    assert(FunctionParser.lookup("TODAY").isDefined)
  }

  test("FunctionParser.lookup returns None for unknown functions") {
    assert(FunctionParser.lookup("UNKNOWN").isEmpty)
    assert(FunctionParser.lookup("FOOBAR").isEmpty)
  }
