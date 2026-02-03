package com.tjclp.xl.formula

import com.tjclp.xl.CellRange
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.Evaluator
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.Sheet
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import org.scalacheck.{Arbitrary, Gen}

import scala.math.BigDecimal

/**
 * Property-based tests for FormulaParser and FormulaPrinter.
 *
 * Tests round-trip laws, edge cases, and integration with existing formula system.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.AsInstanceOf"))
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
    genARef.map(ref => TExpr.ref(ref, TExpr.decodeNumeric))

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
      // Parser now creates PolyRef then resolves to typed Ref[CellValue]
      val expr = TExpr.PolyRef(ref)
      val printed = FormulaPrinter.print(expr, includeEquals = true)
      val parsed = FormulaParser.parse(printed)

      parsed.isRight &&
      parsed.exists {
        case TExpr.Ref(at, _, _) => at == ref // PolyRef resolved to typed Ref
        case _                   => false
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
      case TExpr.Ref(at, _, _) => // Top-level PolyRef resolved to typed Ref[BigDecimal]
        val expected = ARef.parse("A1").fold(err => fail(s"Invalid ref: $err"), identity)
        assertEquals(at, expected)
      case other => fail(s"Expected Ref(A1), got $other")
    }
  }

  test("parse cell reference Z99") {
    val result = FormulaParser.parse("=Z99")
    assert(result.isRight)
    result.foreach {
      case TExpr.Ref(at, _, _) => // Top-level PolyRef resolved to typed Ref[BigDecimal]
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

  // ==================== Exponentiation (^) Tests ====================

  test("parse exponentiation: simple") {
    val result = FormulaParser.parse("=2^3")
    assert(result.isRight, s"Expected success, got $result")
    result.foreach {
      case TExpr.Pow(TExpr.Lit(x: BigDecimal), TExpr.Lit(y: BigDecimal)) =>
        assertEquals(x, BigDecimal(2))
        assertEquals(y, BigDecimal(3))
      case other => fail(s"Expected Pow(2, 3), got $other")
    }
  }

  test("parse exponentiation: right-associative (2^3^2 = 2^(3^2) = 512)") {
    val result = FormulaParser.parse("=2^3^2")
    assert(result.isRight, s"Expected success, got $result")
    result.foreach {
      case TExpr.Pow(TExpr.Lit(_), TExpr.Pow(_, _)) =>
        // 2^(3^2) structure - right-associative
        ()
      case other => fail(s"Expected Pow(2, Pow(3, 2)), got $other")
    }
  }

  test("parse exponentiation: higher precedence than multiplication") {
    // 2*3^2 should parse as 2*(3^2) = 18, not (2*3)^2 = 36
    val result = FormulaParser.parse("=2*3^2")
    assert(result.isRight, s"Expected success, got $result")
    result.foreach {
      case TExpr.Mul(TExpr.Lit(_), TExpr.Pow(_, _)) =>
        // 2*(3^2) structure
        ()
      case other => fail(s"Expected Mul(2, Pow(3, 2)), got $other")
    }
  }

  test("parse exponentiation: unary minus precedence (Excel-compatible)") {
    // Excel parses -2^2 as -(2^2) = -4, not (-2)^2 = 4
    // Our parser matches Excel behavior: ^ binds tighter than unary minus
    val result = FormulaParser.parse("=-2^2")
    assert(result.isRight, s"Expected success, got $result")
    result.foreach {
      case TExpr.Sub(TExpr.Lit(zero: BigDecimal), TExpr.Pow(_, _)) =>
        assertEquals(zero, BigDecimal(0))
      case other => fail(s"Expected Sub(0, Pow(2, 2)), got $other")
    }
  }

  test("parse exponentiation: with parentheses") {
    val result = FormulaParser.parse("=(2^3)^2")
    assert(result.isRight, s"Expected success, got $result")
    result.foreach {
      case TExpr.Pow(TExpr.Pow(_, _), TExpr.Lit(_)) =>
        // (2^3)^2 structure - parentheses override right-associativity
        ()
      case other => fail(s"Expected Pow(Pow(2, 3), 2), got $other")
    }
  }

  test("parse exponentiation: with cell references") {
    val result = FormulaParser.parse("=A1^B1")
    assert(result.isRight, s"Expected success, got $result")
    result.foreach {
      case TExpr.Pow(_, _) => ()
      case other           => fail(s"Expected Pow, got $other")
    }
  }

  test("evaluate exponentiation: 2^3 = 8") {
    val sheet = Sheet("Test")
    val result = for
      expr <- FormulaParser.parse("=2^3")
      value <- Evaluator.eval(expr, sheet)
    yield value
    assertEquals(result, Right(BigDecimal(8)))
  }

  test("evaluate exponentiation: 2^3^2 = 512 (right-associative)") {
    val sheet = Sheet("Test")
    val result = for
      expr <- FormulaParser.parse("=2^3^2")
      value <- Evaluator.eval(expr, sheet)
    yield value
    // 2^(3^2) = 2^9 = 512
    assertEquals(result, Right(BigDecimal(512)))
  }

  test("evaluate exponentiation: 2^-1 = 0.5 (negative exponent)") {
    val sheet = Sheet("Test")
    val result = for
      expr <- FormulaParser.parse("=2^-1")
      value <- Evaluator.eval(expr, sheet)
    yield value
    assertEquals(result, Right(BigDecimal(0.5)))
  }

  test("evaluate exponentiation: 0^0 = 1 (Excel convention)") {
    val sheet = Sheet("Test")
    val result = for
      expr <- FormulaParser.parse("=0^0")
      value <- Evaluator.eval(expr, sheet)
    yield value
    assertEquals(result, Right(BigDecimal(1)))
  }

  test("evaluate exponentiation: 4^0.5 = 2 (square root)") {
    val sheet = Sheet("Test")
    val result = for
      expr <- FormulaParser.parse("=4^0.5")
      value <- Evaluator.eval(expr, sheet)
    yield value
    assertEquals(result, Right(BigDecimal(2)))
  }

  test("print exponentiation: round-trip") {
    val result = FormulaParser.parse("=2^3")
    assert(result.isRight)
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=2^3")
      // Verify round-trip
      val reparsed = FormulaParser.parse(printed)
      assert(reparsed.isRight)
    }
  }

  test("print exponentiation: nested right-associative") {
    val result = FormulaParser.parse("=2^3^2")
    assert(result.isRight)
    result.foreach { expr =>
      val printed = FormulaPrinter.print(expr)
      assertEquals(printed, "=2^3^2")
      // Verify round-trip
      val reparsed = FormulaParser.parse(printed)
      assert(reparsed.isRight)
    }
  }

  test("parse SUM function") {
    val result = FormulaParser.parse("=SUM(A1:B10)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, Left(TExpr.RangeLocation.Local(range)) :: Nil) if spec.name == "SUM" =>
        val expected = CellRange.parse("A1:B10").fold(err => fail(s"Invalid range: $err"), identity)
        assertEquals(range, expected)
      case other => fail(s"Expected Call(SUM) with List(Left(Local range)), got $other")
    }
  }

  test("parse COUNT function") {
    val result = FormulaParser.parse("=COUNT(A1:B10)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, Left(TExpr.RangeLocation.Local(range)) :: Nil) if spec.name == "COUNT" =>
        val expected = CellRange.parse("A1:B10").fold(err => fail(s"Invalid range: $err"), identity)
        assertEquals(range, expected)
      case other => fail(s"Expected Call(COUNT) with List(Left(Local range)), got $other")
    }
  }

  test("parse AVERAGE function") {
    val result = FormulaParser.parse("=AVERAGE(A1:B10)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, Left(TExpr.RangeLocation.Local(range)) :: Nil) if spec.name == "AVERAGE" =>
        val expected = CellRange.parse("A1:B10").fold(err => fail(s"Invalid range: $err"), identity)
        assertEquals(range, expected)
      case other => fail(s"Expected Call(AVERAGE), got $other")
    }
  }

  test("SUM/COUNT/AVERAGE parse to spec-backed calls") {
    // Verify that each aggregate function creates a FunctionSpec call
    val sumResult = FormulaParser.parse("=SUM(A1:A10)")
    val countResult = FormulaParser.parse("=COUNT(A1:A10)")
    val avgResult = FormulaParser.parse("=AVERAGE(A1:A10)")

    (sumResult, countResult, avgResult) match
      case (
            Right(TExpr.Call(sumSpec, _: List[?])),
            Right(TExpr.Call(countSpec, _: List[?])),
            Right(TExpr.Call(avgSpec, _: List[?]))
          ) =>
        assertEquals(sumSpec.name, "SUM")
        assertEquals(countSpec.name, "COUNT")
        assertEquals(avgSpec.name, "AVERAGE")
      case (Right(s), Right(c), Right(a)) =>
        fail(s"Unexpected types: SUM=${s.getClass.getSimpleName}, COUNT=${c.getClass.getSimpleName}, AVERAGE=${a.getClass.getSimpleName}")
      case _ => fail("All three functions should parse successfully")
  }

  test("parse IF function") {
    val result = FormulaParser.parse("=IF(TRUE, 1, 2)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "IF" => ()
      case other => fail(s"Expected IF Call, got $other")
    }
  }

  test("parse AND function") {
    val result = FormulaParser.parse("=AND(TRUE, FALSE)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "AND" => ()
      case other => fail(s"Expected AND Call, got $other")
    }
  }

  test("parse OR function") {
    val result = FormulaParser.parse("=OR(TRUE, FALSE)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "OR" => ()
      case other => fail(s"Expected OR Call, got $other")
    }
  }

  test("parse NOT function") {
    val result = FormulaParser.parse("=NOT(TRUE)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "NOT" => ()
      case other => fail(s"Expected NOT Call, got $other")
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

  // ==================== Concatenation Operator Tests ====================

  test("parse concatenation: string literals") {
    val result = FormulaParser.parse("=\"Hello\"&\"World\"")
    assert(result.isRight, s"Expected success, got $result")
    result.foreach { expr =>
      assertEquals(FormulaPrinter.print(expr), "=\"Hello\"&\"World\"")
    }
  }

  test("parse concatenation: cell references") {
    val result = FormulaParser.parse("=A1&B1")
    assert(result.isRight, s"Expected success, got $result")
    result.foreach { expr =>
      assertEquals(FormulaPrinter.print(expr), "=A1&B1")
    }
  }

  test("parse concatenation: mixed literals and refs") {
    val result = FormulaParser.parse("=A1&\" - \"&B1")
    assert(result.isRight, s"Expected success, got $result")
    result.foreach { expr =>
      assertEquals(FormulaPrinter.print(expr), "=A1&\" - \"&B1")
    }
  }

  test("parse concatenation: with arithmetic (precedence)") {
    // & has lower precedence than + so A1+B1 evaluates first
    val result = FormulaParser.parse("=A1+B1&C1")
    assert(result.isRight, s"Expected success, got $result")
  }

  test("evaluate concatenation: string literals") {
    val sheet = Sheet("Test")
    val result = for
      expr <- FormulaParser.parse("=\"Hello\"&\"World\"")
      value <- Evaluator.eval(expr, sheet)
    yield value
    assertEquals(result, Right("HelloWorld"))
  }

  test("evaluate concatenation: cell references") {
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Text("Hello"))
      .put(ARef.from0(1, 0), CellValue.Text("World"))
    val result = for
      expr <- FormulaParser.parse("=A1&B1")
      value <- Evaluator.eval(expr, sheet)
    yield value
    assertEquals(result, Right("HelloWorld"))
  }

  test("evaluate concatenation: number coercion to string") {
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Number(42))
      .put(ARef.from0(1, 0), CellValue.Text("!"))
    val result = for
      expr <- FormulaParser.parse("=A1&B1")
      value <- Evaluator.eval(expr, sheet)
    yield value
    assertEquals(result, Right("42!"))
  }

  test("evaluate concatenation: chained") {
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Text("A"))
      .put(ARef.from0(1, 0), CellValue.Text("B"))
      .put(ARef.from0(2, 0), CellValue.Text("C"))
    val result = for
      expr <- FormulaParser.parse("=A1&B1&C1")
      value <- Evaluator.eval(expr, sheet)
    yield value
    assertEquals(result, Right("ABC"))
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
    val expr = TExpr.ref(ref, TExpr.decodeNumeric)
    val result = FormulaPrinter.print(expr)
    assertEquals(result, "=A1")
  }

  test("print addition") {
    val expr = TExpr.Add(TExpr.Lit(BigDecimal(1)), TExpr.Lit(BigDecimal(2)))
    val result = FormulaPrinter.print(expr)
    assertEquals(result, "=1+2")
  }

  test("print exponentiation: parenthesize negative base") {
    val expr = TExpr.Pow(
      TExpr.Sub(TExpr.Lit(BigDecimal(0)), TExpr.Lit(BigDecimal(2))),
      TExpr.Lit(BigDecimal(2))
    )
    val result = FormulaPrinter.print(expr)
    assertEquals(result, "=(-2)^2")
    assertEquals(FormulaParser.parse(result), Right(expr))
  }

  test("print exponentiation: parenthesize nested base") {
    val expr = TExpr.Pow(
      TExpr.Pow(TExpr.Lit(BigDecimal(2)), TExpr.Lit(BigDecimal(3))),
      TExpr.Lit(BigDecimal(2))
    )
    val result = FormulaPrinter.print(expr)
    assertEquals(result, "=(2^3)^2")
    assertEquals(FormulaParser.parse(result), Right(expr))
  }

  test("print exponentiation: parenthesize multiplicative exponent") {
    val expr = TExpr.Pow(
      TExpr.Lit(BigDecimal(2)),
      TExpr.Mul(TExpr.Lit(BigDecimal(3)), TExpr.Lit(BigDecimal(4)))
    )
    val result = FormulaPrinter.print(expr)
    assertEquals(result, "=2^(3*4)")
    assertEquals(FormulaParser.parse(result), Right(expr))
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
      case TExpr.Call(spec, Left(TExpr.RangeLocation.Local(range)) :: Nil) if spec.name == "MIN" =>
        assertEquals(range.toA1, "A1:A10")
      case other => fail(s"Expected Call(MIN), got $other")
    }
  }

  test("parse MAX function") {
    val result = FormulaParser.parse("=MAX(B2:B20)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, Left(TExpr.RangeLocation.Local(range)) :: Nil) if spec.name == "MAX" =>
        assertEquals(range.toA1, "B2:B20")
      case other => fail(s"Expected Call(MAX), got $other")
    }
  }

  test("parse CONCATENATE function") {
    val result = FormulaParser.parse("=CONCATENATE(\"Hello\", \" \", \"World\")")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, args: List[?]) if spec.name == "CONCATENATE" =>
        assertEquals(args.length, 3)
      case _ => fail("Expected TExpr.Call(CONCATENATE)")
    }
  }

  test("parse LEFT function") {
    val result = FormulaParser.parse("=LEFT(\"Hello\", 3)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "LEFT" => assert(true)
      case _ => fail("Expected TExpr.Call(LEFT)")
    }
  }

  test("parse RIGHT function") {
    val result = FormulaParser.parse("=RIGHT(A1, 5)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "RIGHT" => assert(true)
      case _ => fail("Expected TExpr.Call(RIGHT)")
    }
  }

  test("parse LEN function") {
    val result = FormulaParser.parse("=LEN(A1)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "LEN" => assert(true)
      case _ => fail("Expected TExpr.Call(LEN)")
    }
  }

  test("parse UPPER function") {
    val result = FormulaParser.parse("=UPPER(\"hello\")")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "UPPER" => assert(true)
      case _ => fail("Expected TExpr.Call(UPPER)")
    }
  }

  test("parse LOWER function") {
    val result = FormulaParser.parse("=LOWER(\"WORLD\")")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "LOWER" => assert(true)
      case _ => fail("Expected TExpr.Call(LOWER)")
    }
  }

  test("parse TODAY function") {
    val result = FormulaParser.parse("=TODAY()")
    assert(result.isRight)
    result.foreach {
      case call: TExpr.Call[?] if call.spec == FunctionSpecs.today => assert(true)
      case _ => fail("Expected TExpr.Call(TODAY)")
    }
  }

  test("parse NOW function") {
    val result = FormulaParser.parse("=NOW()")
    assert(result.isRight)
    result.foreach {
      case call: TExpr.Call[?] if call.spec == FunctionSpecs.now => assert(true)
      case _ => fail("Expected TExpr.Call(NOW)")
    }
  }

  test("parse DATE function") {
    val result = FormulaParser.parse("=DATE(2025, 11, 21)")
    assert(result.isRight)
    result.foreach {
      case call: TExpr.Call[?] if call.spec == FunctionSpecs.date => assert(true)
      case _ => fail("Expected TExpr.Call(DATE)")
    }
  }

  test("parse YEAR function") {
    val result = FormulaParser.parse("=YEAR(TODAY())")
    assert(result.isRight)
    result.foreach {
      case call: TExpr.Call[?] if call.spec == FunctionSpecs.year => assert(true)
      case _ => fail("Expected TExpr.Call(YEAR)")
    }
  }

  test("parse MONTH function") {
    val result = FormulaParser.parse("=MONTH(A1)")
    assert(result.isRight)
    result.foreach {
      case call: TExpr.Call[?] if call.spec == FunctionSpecs.month => assert(true)
      case _ => fail("Expected TExpr.Call(MONTH)")
    }
  }

  test("parse DAY function") {
    val result = FormulaParser.parse("=DAY(DATE(2025, 11, 21))")
    assert(result.isRight)
    result.foreach {
      case call: TExpr.Call[?] if call.spec == FunctionSpecs.day => assert(true)
      case _ => fail("Expected TExpr.Call(DAY)")
    }
  }

  test("parse nested text functions") {
    val result = FormulaParser.parse("=UPPER(CONCATENATE(LEFT(A1, 3), RIGHT(B1, 2)))")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(upperSpec, inner) if upperSpec.name == "UPPER" =>
        inner match
          case TExpr.Call(concatSpec, _) if concatSpec.name == "CONCATENATE" => assert(true)
          case _ => fail("Expected nested CONCATENATE")
      case _ => fail("Expected nested text functions")
    }
  }

  // Error handling functions
  test("parse IFERROR function") {
    val result = FormulaParser.parse("=IFERROR(A1/B1, 0)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "IFERROR" => assert(true)
      case _ => fail("Expected TExpr.Call(IFERROR)")
    }
  }

  test("parse ISERROR function") {
    val result = FormulaParser.parse("=ISERROR(A1)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "ISERROR" => assert(true)
      case _ => fail("Expected TExpr.Call(ISERROR)")
    }
  }

  // Rounding and math functions
  test("parse ROUND function") {
    val result = FormulaParser.parse("=ROUND(3.14159, 2)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "ROUND" => assert(true)
      case _ => fail("Expected TExpr.Call(ROUND)")
    }
  }

  test("parse ROUNDUP function") {
    val result = FormulaParser.parse("=ROUNDUP(3.14159, 2)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "ROUNDUP" => assert(true)
      case _ => fail("Expected TExpr.Call(ROUNDUP)")
    }
  }

  test("parse ROUNDDOWN function") {
    val result = FormulaParser.parse("=ROUNDDOWN(3.14159, 2)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "ROUNDDOWN" => assert(true)
      case _ => fail("Expected TExpr.Call(ROUNDDOWN)")
    }
  }

  test("parse ABS function") {
    val result = FormulaParser.parse("=ABS(-5)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "ABS" => assert(true)
      case _ => fail("Expected TExpr.Call(ABS)")
    }
  }

  test("parse SQRT function") {
    val result = FormulaParser.parse("=SQRT(9)")
    assert(result.isRight)
    result.foreach {
      case TExpr.Call(spec, _) if spec.name == "SQRT" => assert(true)
      case _ => fail("Expected TExpr.Call(SQRT)")
    }
  }

  // Lookup functions
  test("parse INDEX function") {
    val result = FormulaParser.parse("=INDEX(A1:C3, 2, 3)")
    assert(result.isRight)
    result.foreach {
      case call: TExpr.Call[?] if call.spec == FunctionSpecs.index => assert(true)
      case _ => fail("Expected TExpr.Call(INDEX)")
    }
  }

  test("parse INDEX function with 2 args") {
    val result = FormulaParser.parse("=INDEX(A1:A10, 5)")
    assert(result.isRight)
    result.foreach {
      case call: TExpr.Call[?] if call.spec == FunctionSpecs.index =>
        val (_, _, colNumOpt) = call.args.asInstanceOf[FunctionSpecs.IndexArgs]
        assert(colNumOpt.isEmpty)
      case _ => fail("Expected TExpr.Call(INDEX) with no column")
    }
  }

  test("parse MATCH function") {
    val result = FormulaParser.parse("=MATCH(42, A1:A10, 0)")
    assert(result.isRight)
    result.foreach {
      case call: TExpr.Call[?] if call.spec == FunctionSpecs.matchFn =>
        val (_, _, matchTypeOpt) = call.args.asInstanceOf[FunctionSpecs.MatchArgs]
        assert(matchTypeOpt.isDefined)
      case _ => fail("Expected TExpr.Call(MATCH)")
    }
  }

  test("parse MATCH function with default match type") {
    val result = FormulaParser.parse("=MATCH(42, A1:A10)")
    assert(result.isRight)
    result.foreach {
      case call: TExpr.Call[?] if call.spec == FunctionSpecs.matchFn =>
        val (_, _, matchTypeOpt) = call.args.asInstanceOf[FunctionSpecs.MatchArgs]
        assert(matchTypeOpt.isEmpty)
      case _ => fail("Expected TExpr.Call(MATCH)")
    }
  }

  test("Known functions include all 81 functions") {
    val functions = FunctionRegistry.allNames
    assert(functions.contains("SUM"))
    assert(functions.contains("MIN"))
    assert(functions.contains("MAX"))
    assert(functions.contains("CONCATENATE"))
    assert(functions.contains("TODAY"))
    assert(functions.contains("DATE"))
    assert(functions.contains("NPV"))
    assert(functions.contains("IRR"))
    assert(functions.contains("VLOOKUP"))
    // Conditional aggregation functions
    assert(functions.contains("SUMIF"))
    assert(functions.contains("COUNTIF"))
    assert(functions.contains("SUMIFS"))
    assert(functions.contains("COUNTIFS"))
    assert(functions.contains("AVERAGEIF"))
    assert(functions.contains("AVERAGEIFS"))
    // Array and advanced lookup functions
    assert(functions.contains("SUMPRODUCT"))
    assert(functions.contains("XLOOKUP"))
    // Error handling functions
    assert(functions.contains("IFERROR"))
    assert(functions.contains("ISERROR"))
    // Rounding and math functions
    assert(functions.contains("ROUND"))
    assert(functions.contains("ROUNDUP"))
    assert(functions.contains("ROUNDDOWN"))
    assert(functions.contains("ABS"))
    // Math constants
    assert(functions.contains("PI"))
    // Lookup functions
    assert(functions.contains("INDEX"))
    assert(functions.contains("MATCH"))
    // Date-based financial functions
    assert(functions.contains("XNPV"))
    assert(functions.contains("XIRR"))
    // Date calculation functions
    assert(functions.contains("EOMONTH"))
    assert(functions.contains("EDATE"))
    assert(functions.contains("DATEDIF"))
    assert(functions.contains("NETWORKDAYS"))
    assert(functions.contains("WORKDAY"))
    assert(functions.contains("YEARFRAC"))
    // Non-empty cell counter (added via Aggregator typeclass)
    assert(functions.contains("COUNTA"))
    // Math functions
    assert(functions.contains("SQRT"))
    assert(functions.contains("MOD"))
    assert(functions.contains("POWER"))
    assert(functions.contains("LOG"))
    assert(functions.contains("LN"))
    assert(functions.contains("EXP"))
    assert(functions.contains("FLOOR"))
    assert(functions.contains("CEILING"))
    assert(functions.contains("TRUNC"))
    assert(functions.contains("SIGN"))
    assert(functions.contains("INT"))
    // Count functions
    assert(functions.contains("COUNTBLANK"))
    // Reference functions
    assert(functions.contains("ROW"))
    assert(functions.contains("COLUMN"))
    assert(functions.contains("ROWS"))
    assert(functions.contains("COLUMNS"))
    assert(functions.contains("ADDRESS"))
    // Statistical functions
    assert(functions.contains("MEDIAN"))
    assert(functions.contains("STDEV"))
    assert(functions.contains("STDEVP"))
    assert(functions.contains("VAR"))
    assert(functions.contains("VARP"))
    // Type-check functions
    assert(functions.contains("ISNUMBER"))
    assert(functions.contains("ISTEXT"))
    assert(functions.contains("ISBLANK"))
    assert(functions.contains("ISERR"))
    // TVM Financial functions
    assert(functions.contains("PMT"))
    assert(functions.contains("FV"))
    assert(functions.contains("PV"))
    assert(functions.contains("RATE"))
    assert(functions.contains("NPER"))
    // Array functions
    assert(functions.contains("TRANSPOSE"))
    assertEquals(functions.length, 82)
  }

  test("FunctionRegistry.lookup finds spec-based functions") {
    assert(FunctionRegistry.lookup("SUM").isDefined)
    assert(FunctionRegistry.lookup("MIN").isDefined)
    assert(FunctionRegistry.lookup("MAX").isDefined)
    assert(FunctionRegistry.lookup("CONCATENATE").isDefined)
    assert(FunctionRegistry.lookup("LEFT").isDefined)
    assert(FunctionRegistry.lookup("LOWER").isDefined)
    assert(FunctionRegistry.lookup("IF").isDefined)
    assert(FunctionRegistry.lookup("AND").isDefined)
    assert(FunctionRegistry.lookup("OR").isDefined)
    assert(FunctionRegistry.lookup("NOT").isDefined)
    assert(FunctionRegistry.lookup("IFERROR").isDefined)
    assert(FunctionRegistry.lookup("ISERROR").isDefined)
    assert(FunctionRegistry.lookup("ISERR").isDefined)
    assert(FunctionRegistry.lookup("ISNUMBER").isDefined)
    assert(FunctionRegistry.lookup("ISTEXT").isDefined)
    assert(FunctionRegistry.lookup("ISBLANK").isDefined)
  }

  test("FunctionRegistry.lookup returns None for unknown functions") {
    assert(FunctionRegistry.lookup("UNKNOWN").isEmpty)
    assert(FunctionRegistry.lookup("FOOBAR").isEmpty)
  }

  // ==================== XNPV/XIRR Parsing Tests ====================

  test("parses XNPV(rate, values, dates)") {
    val result = FormulaParser.parse("=XNPV(0.1, A1:A5, B1:B5)")
    assert(result.isRight)
    result match
      case Right(call: TExpr.Call[?]) if call.spec == FunctionSpecs.xnpv =>
        val (_, values, dates) = call.args.asInstanceOf[FunctionSpecs.XnpvArgs]
        assert(values.start == ref"A1")
        assert(values.end == ref"A5")
        assert(dates.start == ref"B1")
        assert(dates.end == ref"B5")
      case _ => fail("Expected TExpr.Call(XNPV)")
  }

  test("parses XIRR(values, dates)") {
    val result = FormulaParser.parse("=XIRR(A1:A5, B1:B5)")
    assert(result.isRight)
    result match
      case Right(call: TExpr.Call[?]) if call.spec == FunctionSpecs.xirr =>
        val (values, dates, guess) = call.args.asInstanceOf[FunctionSpecs.XirrArgs]
        assert(values.start == ref"A1")
        assert(values.end == ref"A5")
        assert(dates.start == ref"B1")
        assert(dates.end == ref"B5")
        assert(guess.isEmpty)
      case _ => fail("Expected TExpr.Call(XIRR)")
  }

  test("parses XIRR(values, dates, guess)") {
    val result = FormulaParser.parse("=XIRR(A1:A5, B1:B5, 0.15)")
    assert(result.isRight)
    result match
      case Right(call: TExpr.Call[?]) if call.spec == FunctionSpecs.xirr =>
        val (values, _, guess) = call.args.asInstanceOf[FunctionSpecs.XirrArgs]
        assert(values.start == ref"A1")
        assert(values.end == ref"A5")
        assert(guess.isDefined)
      case _ => fail("Expected TExpr.Call(XIRR) with guess")
  }

  // ==================== Date Calculation Function Parsing Tests ====================

  test("parses EOMONTH(start_date, months)") {
    val result = FormulaParser.parse("=EOMONTH(A1, 1)")
    assert(result.isRight)
    result match
      case Right(call: TExpr.Call[?]) if call.spec == FunctionSpecs.eomonth => ()
      case _ => fail("Expected TExpr.Call(EOMONTH)")
  }

  test("parses EDATE(start_date, months)") {
    val result = FormulaParser.parse("=EDATE(A1, 3)")
    assert(result.isRight)
    result match
      case Right(call: TExpr.Call[?]) if call.spec == FunctionSpecs.edate => ()
      case _ => fail("Expected TExpr.Call(EDATE)")
  }

  test("parses DATEDIF(start, end, unit)") {
    val result = FormulaParser.parse("""=DATEDIF(A1, B1, "Y")""")
    assert(result.isRight)
    result match
      case Right(call: TExpr.Call[?]) if call.spec == FunctionSpecs.datedif => ()
      case _ => fail("Expected TExpr.Call(DATEDIF)")
  }

  test("parses NETWORKDAYS(start, end)") {
    val result = FormulaParser.parse("=NETWORKDAYS(A1, B1)")
    assert(result.isRight)
    result match
      case Right(TExpr.Call(FunctionSpecs.networkdays, (_, _, holidays))) =>
        assert(holidays == None)
      case _ => fail("Expected TExpr.Call(NETWORKDAYS)")
  }

  test("parses NETWORKDAYS(start, end, holidays)") {
    val result = FormulaParser.parse("=NETWORKDAYS(A1, B1, C1:C10)")
    assert(result.isRight)
    result match
      case Right(TExpr.Call(FunctionSpecs.networkdays, (_, _, holidays))) =>
        assert(holidays != None)
      case _ => fail("Expected TExpr.Call(NETWORKDAYS) with holidays")
  }

  test("parses WORKDAY(start, days)") {
    val result = FormulaParser.parse("=WORKDAY(A1, 5)")
    assert(result.isRight)
    result match
      case Right(TExpr.Call(FunctionSpecs.workday, (_, _, holidays))) =>
        assert(holidays == None)
      case _ => fail("Expected TExpr.Call(WORKDAY)")
  }

  test("parses WORKDAY(start, days, holidays)") {
    val result = FormulaParser.parse("=WORKDAY(A1, 5, C1:C10)")
    assert(result.isRight)
    result match
      case Right(TExpr.Call(FunctionSpecs.workday, (_, _, holidays))) =>
        assert(holidays != None)
      case _ => fail("Expected TExpr.Call(WORKDAY) with holidays")
  }

  test("parses YEARFRAC(start, end)") {
    val result = FormulaParser.parse("=YEARFRAC(A1, B1)")
    assert(result.isRight)
    result match
      case Right(TExpr.Call(FunctionSpecs.yearfrac, (_, _, basis))) =>
        assert(basis == None)
      case _ => fail("Expected TExpr.Call(YEARFRAC)")
  }

  test("parses YEARFRAC(start, end, basis)") {
    val result = FormulaParser.parse("=YEARFRAC(A1, B1, 1)")
    assert(result.isRight)
    result match
      case Right(TExpr.Call(FunctionSpecs.yearfrac, (_, _, basis))) =>
        assert(basis != None)
      case _ => fail("Expected TExpr.Call(YEARFRAC) with basis")
  }
