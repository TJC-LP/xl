package com.tjclp.xl.macros

import com.tjclp.xl.cells.{CellValue, FormulaParser}
import com.tjclp.xl.error.XLError
import scala.annotation.unchecked
import munit.FunSuite

/**
 * The `fx` literal: compile-time validation of literals, `Either` for runtime interpolation, and
 * (GH-479) the canonical model shape on every path — the stored `expression` is the BARE text:
 * surrounding whitespace trimmed, the display form's single leading '=' removed, trimmed again (a
 * leading '+' stays, an interior '=' and interior spaces are untouched).
 */
class FormulaInterpolationSpec extends FunSuite:

  // ===== Compile-time rejection (the CHANGELOG's `fx"="` fails to compile, pinned) =====

  test("GH-479: fx\"=\", fx\"\" and a whitespace-only literal fail to compile as empty") {
    val loneEq = compileErrors("""fx"=" """)
    assert(loneEq.contains("Formula literal cannot be empty"), loneEq)
    val empty = compileErrors("""fx"" """)
    assert(empty.contains("Formula literal cannot be empty"), empty)
    val blank = compileErrors("""fx"   " """)
    assert(blank.contains("Formula literal cannot be empty"), blank)
    val blankEq = compileErrors("""fx" = " """)
    assert(blankEq.contains("Formula literal cannot be empty"), blankEq)
  }

  test("GH-479: a literal with unbalanced parentheses fails to compile") {
    val errors = compileErrors("""fx"=SUM(A1:A10" """)
    assert(errors.contains("unbalanced parentheses"), errors)
  }

  // ===== Compile-Time Literals =====

  test("Compile-time literal: fx\"=SUM(A1:A10)\" returns CellValue directly, stored bare") {
    fx"=SUM(A1:A10)" match
      case CellValue.Formula(expr, _, _) => assertEquals(expr, "SUM(A1:A10)")
      case other => fail(s"Expected Formula, got $other")
  }

  test("GH-271: compile-time literal accepts leading unary plus (=+SUM(A1:B2)); the '+' stays") {
    // The macro accepted '=+' before the full parser did; this pins the two staying in agreement
    fx"=+SUM(A1:B2)" match
      case CellValue.Formula(expr, _, _) => assertEquals(expr, "+SUM(A1:B2)")
      case other => fail(s"Expected Formula, got $other")
  }

  test("GH-479: the literal with and without the leading '=' is the same value (both shapes)") {
    assertEquals(fx"SUM(A1:A2)", fx"=SUM(A1:A2)")
    assertEquals(fx"=SUM(A1:A2)", CellValue.Formula("SUM(A1:A2)"))
  }

  test("GH-479: exactly one leading '=' is stripped; interior and doubled '=' survive") {
    fx"=IF(A1=1,2,3)" match
      case CellValue.Formula(expr, _, _) => assertEquals(expr, "IF(A1=1,2,3)")
      case other => fail(s"Expected Formula, got $other")
    fx"==A1" match
      case CellValue.Formula(expr, _, _) => assertEquals(expr, "=A1")
      case other => fail(s"Expected Formula, got $other")
  }

  test("GH-479: all-literal interpolation (compile-time optimized path) stores bare") {
    fx"=SUM(A1:A${10})" match
      case CellValue.Formula(expr, _, _) => assertEquals(expr, "SUM(A1:A10)")
      case other => fail(s"Expected Formula, got $other")
  }

  // ===== Runtime Interpolation =====

  test("Runtime interpolation: simple SUM formula") {
    val formulaStr = "=SUM(A1:A10)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, "SUM(A1:A10)")
      case Left(err) =>
        fail(s"Expected Right(Formula), got Left($err)")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  test("Runtime interpolation: IF formula") {
    val formulaStr = "=IF(A1>0,B1,C1)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, "IF(A1>0,B1,C1)")
      case Left(err) =>
        fail(s"Expected Right(Formula), got Left($err)")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  test("Runtime interpolation: nested parentheses") {
    val formulaStr = "=IF(A1>0,SUM(B1:B10),0)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, "IF(A1>0,SUM(B1:B10),0)")
      case Left(err) =>
        fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
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
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, "SUM(A1:A10)") // We don't require =
      case Left(err) => fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  test("GH-479: runtime interpolation with and without the '=' yields the same value") {
    val withEq = "=SUM(A1:A10)"
    val bare = "SUM(A1:A10)"
    assertEquals(fx"$withEq", fx"$bare")
    assertEquals(fx"$withEq", Right(CellValue.Formula("SUM(A1:A10)")))
    val doubled = "==A1"
    assertEquals(fx"$doubled", Right(CellValue.Formula("=A1")))
  }

  // ===== Error Cases =====

  test("Runtime interpolation: empty formula returns Left") {
    val emptyStr = ""
    fx"$emptyStr" match
      case Left(XLError.FormulaError(input, msg)) =>
        assertEquals(input, "")
        assert(msg.contains("empty"))
      case Right(value) =>
        fail(s"Expected Left(FormulaError), got Right($value)")
      case Left(other) =>
        fail(s"Expected FormulaError, got $other")
  }

  test("GH-479: a lone '=' is an empty formula (Left reporting the ORIGINAL text)") {
    val loneEq = "="
    fx"$loneEq" match
      case Left(XLError.FormulaError(input, msg)) =>
        assertEquals(input, "=")
        assert(msg.contains("empty"))
      case Right(value) =>
        fail(s"Expected Left(FormulaError), got Right($value)")
      case Left(other) =>
        fail(s"Expected FormulaError, got $other")
  }

  test("Runtime interpolation: unbalanced opening paren returns Left") {
    val invalidStr = "=SUM(A1:A10"
    fx"$invalidStr" match
      case Left(XLError.FormulaError(input, msg)) =>
        assertEquals(input, "=SUM(A1:A10")
        assert(msg.contains("Unbalanced"))
      case Right(value) =>
        fail(s"Expected Left(FormulaError), got Right($value)")
      case Left(other) =>
        fail(s"Expected FormulaError, got $other")
  }

  test("Runtime interpolation: unbalanced closing paren returns Left") {
    val invalidStr = "=SUM(A1:A10))"
    fx"$invalidStr" match
      case Left(XLError.FormulaError(input, msg)) =>
        assertEquals(input, "=SUM(A1:A10))")
        assert(msg.contains("Unbalanced"))
      case Right(value) =>
        fail(s"Expected Left(FormulaError), got Right($value)")
      case Left(other) =>
        fail(s"Expected FormulaError, got $other")
  }

  test("Runtime interpolation: mismatched parens returns Left") {
    val invalidStr = "=SUM(A1:A10()"
    fx"$invalidStr" match
      case Left(XLError.FormulaError(_, msg)) =>
        assert(msg.contains("Unbalanced"))
      case Right(value) =>
        fail(s"Expected Left(FormulaError), got Right($value)")
      case Left(other) =>
        fail(s"Expected FormulaError, got $other")
  }

  // ===== Mixed Compile-Time and Runtime =====

  test("Mixed interpolation: prefix + variable") {
    val range = "A1:A10"
    fx"=SUM($range)" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, "SUM(A1:A10)")
      case Left(err) => fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  test("Mixed interpolation: variable function name") {
    val func = "SUM"
    fx"=$func(A1:A10)" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, "SUM(A1:A10)")
      case Left(err) => fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  test("Mixed interpolation: multiple variables") {
    val func = "IF"
    val cond = "A1>0"
    val thenVal = "B1"
    val elseVal = "C1"
    fx"=$func($cond,$thenVal,$elseVal)" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, "IF(A1>0,B1,C1)")
      case Left(err) => fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  // ===== String Literal Edge Cases =====

  test("String literal: formula with ) inside string") {
    val formulaStr = """=IF(A1=")", "yes", "no")"""
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, """IF(A1=")", "yes", "no")""")
      case Left(err) => fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  test("String literal: formula with ( inside string") {
    val formulaStr = """=IF(A1="(", "left", "right")"""
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, """IF(A1="(", "left", "right")""")
      case Left(err) => fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  test("String literal: escaped quotes inside string") {
    val formulaStr = "=CONCATENATE(\"Say \"\"hello\"\"\", A1)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, "CONCATENATE(\"Say \"\"hello\"\"\", A1)")
      case Left(err) => fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  test("String literal: double escaped quotes regression") {
    val formulaStr = "=IF(A1=\"\"test\"\", B1, C1)"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, "IF(A1=\"\"test\"\", B1, C1)")
      case Left(err) => fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  test("String literal: multiple strings with parens") {
    val formulaStr = """=IF(A1=")", B1, "(other)")"""
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, """IF(A1=")", B1, "(other)")""")
      case Left(err) => fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  test("String literal: rejects unclosed string") {
    val formulaStr = """=IF(A1="unclosed, B1, C1)"""
    fx"$formulaStr" match
      case Left(XLError.FormulaError(_, msg)) =>
        assert(msg.contains("Unbalanced"))
      case Right(value) => fail(s"Should fail for unclosed string, got Right($value)")
      case Left(other) =>
        fail(s"Expected FormulaError, got $other")
  }

  test("String literal: nested parens with strings") {
    val formulaStr = """=IF(IF(A1=">", SUM(B1:B10), 0), "result", "none")"""
    fx"$formulaStr" match
      case Right(_) => () // Should parse
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Edge Cases =====

  test("Edge: surrounding whitespace is trimmed, interior spaces are kept (GH-479, one rule)") {
    val formulaStr = " =SUM( A1:A10 ) "
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, "SUM( A1:A10 )") // trim, one leading '=' off, trim
      case Left(err) => fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  test("GH-479: every core entry canonicalises a padded formula to the same value") {
    // The runtime fx path, FormulaParser.parse and CellValue.formula share one rule with the edit
    // interpreter (EditInterpreterSpec pins `" = SUM(B2:B4) "` -> "SUM(B2:B4)") and the CLI's putf.
    val padded = " = SUM( A1:A10 ) "
    val canonical: Either[XLError, CellValue] = Right(CellValue.Formula("SUM( A1:A10 )"))
    assertEquals(fx"$padded", canonical)
    assertEquals(FormulaParser.parse(padded), canonical)
    assertEquals(CellValue.formula(padded).map(identity[CellValue]), canonical)
    assertEquals(fx"$padded", Right(fx"=SUM( A1:A10 )"): Either[XLError, CellValue])
    // and the literal path agrees with the runtime path on the same text
    assertEquals(fx"$padded", Right(fx" = SUM( A1:A10 ) "): Either[XLError, CellValue])
  }

  test("Edge: formula with many nested parens") {
    val formulaStr = "=IF(IF(A1>0,IF(B1>0,1,0),0),SUM(C1:C10),0)"
    fx"$formulaStr" match
      case Right(_) => () // Should parse
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Edge: formula with array syntax {1,2,3}") {
    val formulaStr = "=SUM({1,2,3})"
    fx"$formulaStr" match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, "SUM({1,2,3})")
      case Left(err) => fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }

  // ===== Cell Limit Validation =====

  test("Cell limit: rejects formula exceeding Excel limit") {
    val longFormula = "=" + ("A" * 33000) // Exceeds 32,767
    fx"$longFormula" match
      case Left(XLError.FormulaError(input, msg)) =>
        assert(msg.contains("Excel cell limit"))
        assert(input.length < 100) // Truncated in error message
      case Right(value) => fail(s"Should fail for too-long formula, got Right($value)")
      case Left(other) =>
        fail(s"Expected FormulaError, got $other")
  }

  // ===== Integration =====

  test("Integration: for-comprehension with Either") {
    val func = "SUM"
    val range = "A1:A10"

    val result =
      for formula <- fx"=$func($range)"
      yield formula

    result match
      case Right(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr, "SUM(A1:A10)")
      case Left(err) => fail(s"Should parse: $err")
      case Right(other) =>
        fail(s"Expected Formula, got $other")
  }
