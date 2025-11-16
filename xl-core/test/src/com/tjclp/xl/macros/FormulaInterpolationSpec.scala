package com.tjclp.xl.macros

import com.tjclp.xl.cell.{CellValue, FormulaParser}
import com.tjclp.xl.error.XLError
import munit.FunSuite

class FormulaInterpolationSpec extends FunSuite:

  // ===== Backward Compatibility (Compile-Time Literals) =====

  test("Compile-time literal: fx\"=SUM(A1:A10)\" returns CellValue directly") {
    val f = fx"=SUM(A1:A10)"
    f match
      case formula: CellValue.Formula =>
        assertEquals(formula.expression, "=SUM(A1:A10)")
      case other => fail(s"Expected CellValue.Formula, got $other")
  }

  // ===== Runtime Interpolation (New Functionality) =====

  test("Runtime interpolation: simple SUM formula") {
    val formulaStr = "=SUM(A1:A10)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=SUM(A1:A10)")
      case other =>
        fail(s"Expected Right(Formula), got $other")
  }

  test("Runtime interpolation: IF formula") {
    val formulaStr = "=IF(A1>0,B1,C1)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=IF(A1>0,B1,C1)")
      case other =>
        fail(s"Expected Right(Formula), got $other")
  }

  test("Runtime interpolation: nested parentheses") {
    val formulaStr = "=IF(A1>0,SUM(B1:B10),0)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=IF(A1>0,SUM(B1:B10),0)")
      case other =>
        fail(s"Should parse: $other")
  }

  test("Runtime interpolation: complex formula with multiple functions") {
    val formulaStr = "=AVERAGE(IF(A1:A10>0,B1:B10))"
    fx"$formulaStr" match
      case Right(_) => () // Expected
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime interpolation: formula without = prefix") {
    val formulaStr = "SUM(A1:A10)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "SUM(A1:A10)")  // We don't require =
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }

  // ===== Error Cases =====

  test("Runtime interpolation: empty formula returns Left") {
    val emptyStr = ""
    fx"$emptyStr" match
      case Left(XLError.FormulaError(input, msg)) =>
        assertEquals(input, "")
        assert(msg.contains("empty"))
      case other =>
        fail(s"Expected Left(FormulaError), got $other")
  }

  test("Runtime interpolation: unbalanced opening paren returns Left") {
    val invalidStr = "=SUM(A1:A10"
    fx"$invalidStr" match
      case Left(XLError.FormulaError(input, msg)) =>
        assertEquals(input, "=SUM(A1:A10")
        assert(msg.contains("Unbalanced"))
      case other =>
        fail(s"Expected Left(FormulaError), got $other")
  }

  test("Runtime interpolation: unbalanced closing paren returns Left") {
    val invalidStr = "=SUM(A1:A10))"
    fx"$invalidStr" match
      case Left(XLError.FormulaError(input, msg)) =>
        assertEquals(input, "=SUM(A1:A10))")
        assert(msg.contains("Unbalanced"))
      case other =>
        fail(s"Expected Left(FormulaError), got $other")
  }

  test("Runtime interpolation: mismatched parens returns Left") {
    val invalidStr = "=SUM(A1:A10()"
    fx"$invalidStr" match
      case Left(XLError.FormulaError(_, msg)) =>
        assert(msg.contains("Unbalanced"))
      case other =>
        fail(s"Expected Left(FormulaError), got $other")
  }

  // ===== Mixed Compile-Time and Runtime =====

  test("Mixed interpolation: prefix + variable") {
    val range = "A1:A10"
    fx"=SUM($range)" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=SUM(A1:A10)")
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }

  test("Mixed interpolation: variable function name") {
    val func = "SUM"
    fx"=$func(A1:A10)" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=SUM(A1:A10)")
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }

  test("Mixed interpolation: multiple variables") {
    val func = "IF"
    val cond = "A1>0"
    val thenVal = "B1"
    val elseVal = "C1"
    fx"=$func($cond,$thenVal,$elseVal)" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=IF(A1>0,B1,C1)")
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }

  // ===== String Literal Edge Cases =====

  test("String literal: formula with ) inside string") {
    val formulaStr = """=IF(A1=")", "yes", "no")"""
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, """=IF(A1=")", "yes", "no")""")
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }

  test("String literal: formula with ( inside string") {
    val formulaStr = """=IF(A1="(", "left", "right")"""
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, """=IF(A1="(", "left", "right")""")
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }

  test("String literal: escaped quotes inside string") {
    val formulaStr = "=CONCATENATE(\"Say \"\"hello\"\"\", A1)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=CONCATENATE(\"Say \"\"hello\"\"\", A1)")
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }

  test("String literal: multiple strings with parens") {
    val formulaStr = """=IF(A1=")", B1, "(other)")"""
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, """=IF(A1=")", B1, "(other)")""")
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }

  test("String literal: rejects unclosed string") {
    val formulaStr = """=IF(A1="unclosed, B1, C1)"""
    fx"$formulaStr" match
      case Left(XLError.FormulaError(_, msg)) =>
        assert(msg.contains("Unbalanced"))
      case other => fail(s"Should fail for unclosed string, got $other")
  }

  test("String literal: nested parens with strings") {
    val formulaStr = """=IF(IF(A1=">", SUM(B1:B10), 0), "result", "none")"""
    fx"$formulaStr" match
      case Right(_) => () // Should parse
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }

  // ===== Edge Cases =====

  test("Edge: formula with whitespace") {
    val formulaStr = " =SUM( A1:A10 ) "
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, " =SUM( A1:A10 ) ")  // Preserve whitespace
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }

  test("Edge: formula with many nested parens") {
    val formulaStr = "=IF(IF(A1>0,IF(B1>0,1,0),0),SUM(C1:C10),0)"
    fx"$formulaStr" match
      case Right(_) => () // Should parse
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }

  test("Edge: formula with array syntax {1,2,3}") {
    val formulaStr = "=SUM({1,2,3})"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=SUM({1,2,3})")
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }

  // ===== Integration =====

  test("Integration: for-comprehension with Either") {
    val func = "SUM"
    val range = "A1:A10"

    val result = for
      formula <- fx"=$func($range)"
    yield formula

    result match
      case Right(CellValue.Formula(expr)) =>
        assertEquals(expr, "=SUM(A1:A10)")
      case Left(err) => fail(s"Should parse: $err")
      case other => fail(s"Unexpected result: $other")
  }
