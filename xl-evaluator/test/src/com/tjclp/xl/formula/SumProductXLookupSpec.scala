package com.tjclp.xl.formula

import munit.FunSuite
import com.tjclp.xl.*
import com.tjclp.xl.unsafe.*
import com.tjclp.xl.cells.{CellValue, CellError}
import com.tjclp.xl.sheets.Sheet

/**
 * Comprehensive tests for SUMPRODUCT and XLOOKUP functions.
 *
 * Tests cover:
 *   - Basic functionality
 *   - Type coercion (SUMPRODUCT: TRUE->1, FALSE->0, text->0)
 *   - Match modes (XLOOKUP: exact, wildcard, next smaller, next larger)
 *   - Search modes (XLOOKUP: first-to-last, last-to-first)
 *   - Error handling (dimension mismatch, not found)
 *   - Round-trip parsing
 *   - Dependency graph extraction
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class SumProductXLookupSpec extends FunSuite:

  // Helper to create a sheet with data
  private def sheetWith(entries: (ARef, Any)*): Sheet =
    entries.foldLeft(Sheet("Test")) { case (s, (ref, value)) =>
      val cv = value match
        case cv: CellValue => cv
        case s: String => CellValue.Text(s)
        case n: Int => CellValue.Number(BigDecimal(n))
        case n: Long => CellValue.Number(BigDecimal(n))
        case n: Double => CellValue.Number(BigDecimal(n))
        case n: BigDecimal => CellValue.Number(n)
        case b: Boolean => CellValue.Bool(b)
        case _ => CellValue.Text(value.toString)
      s.put(ref, cv).unsafe
    }

  // Helper to parse and evaluate a formula returning BigDecimal
  private def evalFormula(formula: String, sheet: Sheet): Either[String, BigDecimal] =
    FormulaParser.parse(formula) match
      case Right(expr) =>
        Evaluator.eval(expr, sheet) match
          case Right(value: BigDecimal) => Right(value)
          case Right(other) => Left(s"Expected BigDecimal, got: $other")
          case Left(err) => Left(s"Eval error: $err")
      case Left(err) => Left(s"Parse error: $err")

  // Helper to parse and evaluate a formula returning CellValue
  private def evalFormulaCellValue(formula: String, sheet: Sheet): Either[String, CellValue] =
    FormulaParser.parse(formula) match
      case Right(expr) =>
        Evaluator.eval(expr, sheet) match
          case Right(value: CellValue) => Right(value)
          case Right(other) => Left(s"Expected CellValue, got: $other")
          case Left(err) => Left(s"Eval error: $err")
      case Left(err) => Left(s"Parse error: $err")

  // Helper for expected results (BigDecimal)
  private def assertEval(formula: String, sheet: Sheet, expected: BigDecimal)(implicit
    loc: munit.Location
  ): Unit =
    evalFormula(formula, sheet) match
      case Right(value) => assertEquals(value, expected)
      case Left(err) => fail(s"Formula evaluation failed: $err")

  private def assertEval(formula: String, sheet: Sheet, expected: Int)(implicit
    loc: munit.Location
  ): Unit =
    assertEval(formula, sheet, BigDecimal(expected))

  // Helper for expected CellValue results
  private def assertEvalCellValue(formula: String, sheet: Sheet, expected: CellValue)(implicit
    loc: munit.Location
  ): Unit =
    evalFormulaCellValue(formula, sheet) match
      case Right(value) => assertEquals(value, expected)
      case Left(err) => fail(s"Formula evaluation failed: $err")

  // ===== SUMPRODUCT Basic Tests =====

  test("SUMPRODUCT: basic two-array multiplication") {
    val sheet = sheetWith(
      ref"A1" -> 1, ref"B1" -> 4,
      ref"A2" -> 2, ref"B2" -> 5,
      ref"A3" -> 3, ref"B3" -> 6
    )
    // 1*4 + 2*5 + 3*6 = 4 + 10 + 18 = 32
    assertEval("=SUMPRODUCT(A1:A3, B1:B3)", sheet, 32)
  }

  test("SUMPRODUCT: single array (sum of values)") {
    val sheet = sheetWith(
      ref"A1" -> 10,
      ref"A2" -> 20,
      ref"A3" -> 30
    )
    // SUMPRODUCT with single array = sum * 1 for each
    assertEval("=SUMPRODUCT(A1:A3)", sheet, 60)
  }

  test("SUMPRODUCT: three arrays") {
    val sheet = sheetWith(
      ref"A1" -> 1, ref"B1" -> 2, ref"C1" -> 3,
      ref"A2" -> 4, ref"B2" -> 5, ref"C2" -> 6
    )
    // 1*2*3 + 4*5*6 = 6 + 120 = 126
    assertEval("=SUMPRODUCT(A1:A2, B1:B2, C1:C2)", sheet, 126)
  }

  test("SUMPRODUCT: column vectors") {
    val sheet = sheetWith(
      ref"A1" -> 10, ref"B1" -> 1,
      ref"A2" -> 20, ref"B2" -> 2,
      ref"A3" -> 30, ref"B3" -> 3
    )
    // 10*1 + 20*2 + 30*3 = 10 + 40 + 90 = 140
    assertEval("=SUMPRODUCT(A1:A3, B1:B3)", sheet, 140)
  }

  test("SUMPRODUCT: 2D arrays (2x2)") {
    val sheet = sheetWith(
      ref"A1" -> 1, ref"A2" -> 2,
      ref"B1" -> 3, ref"B2" -> 4,
      ref"C1" -> 5, ref"C2" -> 6,
      ref"D1" -> 7, ref"D2" -> 8
    )
    // Array 1: A1:B2 (1,2,3,4), Array 2: C1:D2 (5,6,7,8)
    // 1*5 + 2*6 + 3*7 + 4*8 = 5 + 12 + 21 + 32 = 70
    assertEval("=SUMPRODUCT(A1:B2, C1:D2)", sheet, 70)
  }

  // ===== SUMPRODUCT Type Coercion Tests =====

  test("SUMPRODUCT: TRUE coerced to 1") {
    val sheet = sheetWith(
      ref"A1" -> true, ref"B1" -> 10,
      ref"A2" -> false, ref"B2" -> 20
    )
    // TRUE*10 + FALSE*20 = 1*10 + 0*20 = 10
    assertEval("=SUMPRODUCT(A1:A2, B1:B2)", sheet, 10)
  }

  test("SUMPRODUCT: FALSE coerced to 0") {
    val sheet = sheetWith(
      ref"A1" -> 5, ref"B1" -> false,
      ref"A2" -> 10, ref"B2" -> true
    )
    // 5*0 + 10*1 = 10
    assertEval("=SUMPRODUCT(A1:A2, B1:B2)", sheet, 10)
  }

  test("SUMPRODUCT: text coerced to 0") {
    val sheet = sheetWith(
      ref"A1" -> 5, ref"B1" -> "text",
      ref"A2" -> 10, ref"B2" -> 2
    )
    // 5*0 + 10*2 = 20
    assertEval("=SUMPRODUCT(A1:A2, B1:B2)", sheet, 20)
  }

  test("SUMPRODUCT: empty cells coerced to 0") {
    val sheet = sheetWith(
      ref"A1" -> 5, ref"B1" -> 2
      // A2, B2 are empty
    )
    // With empty cells at A2, B2: 5*2 + 0*0 = 10
    assertEval("=SUMPRODUCT(A1:A2, B1:B2)", sheet, 10)
  }

  // ===== SUMPRODUCT Error Cases =====

  test("SUMPRODUCT: dimension mismatch error") {
    val sheet = sheetWith(
      ref"A1" -> 1, ref"A2" -> 2, ref"A3" -> 3,
      ref"B1" -> 4, ref"B2" -> 5
    )
    val result = evalFormula("=SUMPRODUCT(A1:A3, B1:B2)", sheet)
    assert(result.isLeft)
    assert(result.left.exists(_.contains("dimension")))
  }

  test("SUMPRODUCT: single cell arrays work") {
    val sheet = sheetWith(
      ref"A1" -> 5,
      ref"B1" -> 3
    )
    assertEval("=SUMPRODUCT(A1:A1, B1:B1)", sheet, 15)
  }

  test("SUMPRODUCT: all zeros") {
    val sheet = sheetWith(
      ref"A1" -> 0, ref"B1" -> 0,
      ref"A2" -> 0, ref"B2" -> 0
    )
    assertEval("=SUMPRODUCT(A1:A2, B1:B2)", sheet, 0)
  }

  // ===== SUMPRODUCT Round-trip Tests =====

  test("SUMPRODUCT: parse -> print -> parse roundtrip (two arrays)") {
    val formula = "=SUMPRODUCT(A1:A3, B1:B3)"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight)
    val printed = FormulaPrinter.print(parsed.toOption.get)
    assertEquals(printed, formula)
  }

  test("SUMPRODUCT: parse -> print -> parse roundtrip (single array)") {
    val formula = "=SUMPRODUCT(A1:A10)"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight)
    val printed = FormulaPrinter.print(parsed.toOption.get)
    assertEquals(printed, formula)
  }

  // ===== GH-197: SUMPRODUCT Array Expression Tests =====

  test("SUMPRODUCT: array expression with numeric comparison") {
    val sheet = sheetWith(
      ref"A1" -> 10, ref"B1" -> 1,
      ref"A2" -> 20, ref"B2" -> 2,
      ref"A3" -> 30, ref"B3" -> 3
    )
    // (A1:A3>15)*B1:B3 = [0,1,1]*[1,2,3] = [0,2,3], sum = 5
    assertEval("=SUMPRODUCT((A1:A3>15)*B1:B3)", sheet, 5)
  }

  test("SUMPRODUCT: single array multiplication expression") {
    val sheet = sheetWith(
      ref"A1" -> 10, ref"B1" -> 2,
      ref"A2" -> 20, ref"B2" -> 3,
      ref"A3" -> 30, ref"B3" -> 4
    )
    // SUMPRODUCT(A1:A3*B1:B3) = sum of [20,60,120] = 200
    assertEval("=SUMPRODUCT(A1:A3*B1:B3)", sheet, 200)
  }

  test("SUMPRODUCT: multiple boolean criteria") {
    val sheet = sheetWith(
      ref"A1" -> 100, ref"B1" -> 50, ref"C1" -> 10,
      ref"A2" -> 200, ref"B2" -> 30, ref"C2" -> 20,
      ref"A3" -> 300, ref"B3" -> 70, ref"C3" -> 30
    )
    // (A1:A3>150)*(B1:B3>40)*C1:C3
    // = [0,1,1]*[1,0,1]*[10,20,30] = [0,0,30]
    // Sum = 30
    assertEval("=SUMPRODUCT((A1:A3>150)*(B1:B3>40)*C1:C3)", sheet, 30)
  }

  test("SUMPRODUCT: mixed ranges and expressions") {
    val sheet = sheetWith(
      ref"A1" -> 1, ref"B1" -> 2, ref"C1" -> true,
      ref"A2" -> 3, ref"B2" -> 4, ref"C2" -> false,
      ref"A3" -> 5, ref"B3" -> 6, ref"C3" -> true
    )
    // SUMPRODUCT(A1:A3, (C1:C3)*B1:B3)
    // C1:C3 = [1,0,1], B1:B3 = [2,4,6]
    // (C1:C3)*B1:B3 = [2,0,6]
    // A1:A3 * [2,0,6] = [1*2,3*0,5*6] = [2,0,30]
    // Sum = 32
    assertEval("=SUMPRODUCT(A1:A3, (C1:C3)*B1:B3)", sheet, 32)
  }

  test("SUMPRODUCT: all-true boolean array") {
    val sheet = sheetWith(
      ref"A1" -> 1, ref"B1" -> 10,
      ref"A2" -> 2, ref"B2" -> 20,
      ref"A3" -> 3, ref"B3" -> 30
    )
    // All A values are >0, so (A1:A3>0) = [1,1,1]
    // SUMPRODUCT((A1:A3>0)*B1:B3) = [1,1,1]*[10,20,30] = 60
    assertEval("=SUMPRODUCT((A1:A3>0)*B1:B3)", sheet, 60)
  }

  test("SUMPRODUCT: all-false boolean array") {
    val sheet = sheetWith(
      ref"A1" -> 1, ref"B1" -> 10,
      ref"A2" -> 2, ref"B2" -> 20,
      ref"A3" -> 3, ref"B3" -> 30
    )
    // All A values are <0 is false, so (A1:A3<0) = [0,0,0]
    // SUMPRODUCT((A1:A3<0)*B1:B3) = [0,0,0]*[10,20,30] = 0
    assertEval("=SUMPRODUCT((A1:A3<0)*B1:B3)", sheet, 0)
  }

  test("SUMPRODUCT: nested arithmetic in expression") {
    val sheet = sheetWith(
      ref"A1" -> 1, ref"B1" -> 2,
      ref"A2" -> 3, ref"B2" -> 4
    )
    // SUMPRODUCT((A1:A2+B1:B2)*2) = sum of ([3,7]*2) = sum of [6,14] = 20
    assertEval("=SUMPRODUCT((A1:A2+B1:B2)*2)", sheet, 20)
  }

  test("SUMPRODUCT: expression dimension mismatch error") {
    val sheet = sheetWith(
      ref"A1" -> 1, ref"A2" -> 2, ref"A3" -> 3,
      ref"B1" -> 1, ref"B2" -> 2
    )
    // A1:A3 is 3x1, B1:B2 is 2x1 - should fail with dimension error
    val result = evalFormula("=SUMPRODUCT((A1:A3>0), B1:B2)", sheet)
    assert(result.isLeft)
    assert(result.left.exists(_.contains("dimension")))
  }

  test("SUMPRODUCT: scalar expression treated as 1x1") {
    val sheet = sheetWith(
      ref"A1" -> 10
    )
    // SUMPRODUCT(A1:A1, 2) - scalar 2 treated as 1x1
    // = [10] * [2] = 20
    assertEval("=SUMPRODUCT(A1:A1, 2)", sheet, 20)
  }

  test("SUMPRODUCT: GH-197 full-column ranges in expressions are bounded") {
    // GH-197: Full-column ranges in array expressions should be bounded to shared usedRange
    // Without bounding, A:A and B:B would create 1M+ row arrays
    val sheet = sheetWith(
      ref"A1" -> 10, ref"B1" -> 1,
      ref"A2" -> 20, ref"B2" -> 2,
      ref"A3" -> 30, ref"B3" -> 3
    )
    // Full-column ranges: (A:A>15)*B:B should be bounded to rows 1-3
    // = [0,1,1]*[1,2,3] = [0,2,3], sum = 5
    assertEval("=SUMPRODUCT((A:A>15)*B:B)", sheet, 5)
  }

  test("SUMPRODUCT: GH-197 mixed full-column and bounded ranges") {
    val sheet = sheetWith(
      ref"A1" -> 5, ref"B1" -> 10,
      ref"A2" -> 15, ref"B2" -> 20,
      ref"A3" -> 25, ref"B3" -> 30
    )
    // Mix of full-column expression and bounded range
    // (A:A>10)*B1:B3 should bound A:A to rows 1-3
    // = [0,1,1]*[10,20,30] = [0,20,30], sum = 50
    assertEval("=SUMPRODUCT((A:A>10)*B1:B3)", sheet, 50)
  }

  test("SUMPRODUCT: GH-197 equality comparison with full-column ranges") {
    val sheet = sheetWith(
      ref"A1" -> "Yes", ref"B1" -> 10,
      ref"A2" -> "No", ref"B2" -> 20,
      ref"A3" -> "Yes", ref"B3" -> 30
    )
    // (A:A="Yes")*B:B with full-column ranges
    // = [1,0,1]*[10,20,30] = [10,0,30], sum = 40
    assertEval("=SUMPRODUCT((A:A=\"Yes\")*B:B)", sheet, 40)
  }

  // ===== XLOOKUP Basic Tests =====

  test("XLOOKUP: exact match - text lookup") {
    val sheet = sheetWith(
      ref"A1" -> "Apple", ref"B1" -> 100,
      ref"A2" -> "Banana", ref"B2" -> 200,
      ref"A3" -> "Cherry", ref"B3" -> 300
    )
    assertEvalCellValue(
      "=XLOOKUP(\"Banana\", A1:A3, B1:B3)",
      sheet,
      CellValue.Number(200)
    )
  }

  test("XLOOKUP: exact match - numeric lookup") {
    val sheet = sheetWith(
      ref"A1" -> 1, ref"B1" -> "One",
      ref"A2" -> 2, ref"B2" -> "Two",
      ref"A3" -> 3, ref"B3" -> "Three"
    )
    assertEvalCellValue(
      "=XLOOKUP(2, A1:A3, B1:B3)",
      sheet,
      CellValue.Text("Two")
    )
  }

  test("XLOOKUP: returns text value") {
    val sheet = sheetWith(
      ref"A1" -> 100, ref"B1" -> "Small",
      ref"A2" -> 200, ref"B2" -> "Medium",
      ref"A3" -> 300, ref"B3" -> "Large"
    )
    assertEvalCellValue(
      "=XLOOKUP(200, A1:A3, B1:B3)",
      sheet,
      CellValue.Text("Medium")
    )
  }

  test("XLOOKUP: returns numeric value") {
    val sheet = sheetWith(
      ref"A1" -> "X", ref"B1" -> 10,
      ref"A2" -> "Y", ref"B2" -> 20,
      ref"A3" -> "Z", ref"B3" -> 30
    )
    assertEvalCellValue(
      "=XLOOKUP(\"Y\", A1:A3, B1:B3)",
      sheet,
      CellValue.Number(20)
    )
  }

  test("XLOOKUP: case-insensitive text match") {
    val sheet = sheetWith(
      ref"A1" -> "Apple", ref"B1" -> 100,
      ref"A2" -> "Banana", ref"B2" -> 200
    )
    assertEvalCellValue(
      "=XLOOKUP(\"APPLE\", A1:A2, B1:B2)",
      sheet,
      CellValue.Number(100)
    )
  }

  test("XLOOKUP: horizontal arrays") {
    val sheet = sheetWith(
      ref"A1" -> "A", ref"B1" -> "B", ref"C1" -> "C",
      ref"A2" -> 1, ref"B2" -> 2, ref"C2" -> 3
    )
    assertEvalCellValue(
      "=XLOOKUP(\"B\", A1:C1, A2:C2)",
      sheet,
      CellValue.Number(2)
    )
  }

  // ===== XLOOKUP Match Modes =====

  test("XLOOKUP: match_mode 0 (exact) - not found returns #N/A") {
    val sheet = sheetWith(
      ref"A1" -> "Apple", ref"B1" -> 100,
      ref"A2" -> "Banana", ref"B2" -> 200
    )
    assertEvalCellValue(
      "=XLOOKUP(\"Cherry\", A1:A2, B1:B2)",
      sheet,
      CellValue.Error(CellError.NA)
    )
  }

  test("XLOOKUP: match_mode -1 (next smaller)") {
    val sheet = sheetWith(
      ref"A1" -> 10, ref"B1" -> "Ten",
      ref"A2" -> 20, ref"B2" -> "Twenty",
      ref"A3" -> 30, ref"B3" -> "Thirty"
    )
    // Looking for 25, should find 20 (next smaller)
    assertEvalCellValue(
      "=XLOOKUP(25, A1:A3, B1:B3, \"\", -1)",
      sheet,
      CellValue.Text("Twenty")
    )
  }

  test("XLOOKUP: match_mode 1 (next larger)") {
    val sheet = sheetWith(
      ref"A1" -> 10, ref"B1" -> "Ten",
      ref"A2" -> 20, ref"B2" -> "Twenty",
      ref"A3" -> 30, ref"B3" -> "Thirty"
    )
    // Looking for 25, should find 30 (next larger)
    assertEvalCellValue(
      "=XLOOKUP(25, A1:A3, B1:B3, \"\", 1)",
      sheet,
      CellValue.Text("Thirty")
    )
  }

  test("XLOOKUP: match_mode 2 (wildcard with *)") {
    val sheet = sheetWith(
      ref"A1" -> "Apple Pie", ref"B1" -> 100,
      ref"A2" -> "Banana Split", ref"B2" -> 200,
      ref"A3" -> "Cherry Tart", ref"B3" -> 300
    )
    assertEvalCellValue(
      "=XLOOKUP(\"Banana*\", A1:A3, B1:B3, \"\", 2)",
      sheet,
      CellValue.Number(200)
    )
  }

  test("XLOOKUP: match_mode 2 (wildcard with ?)") {
    val sheet = sheetWith(
      ref"A1" -> "Cat", ref"B1" -> 100,
      ref"A2" -> "Cut", ref"B2" -> 200,
      ref"A3" -> "Cot", ref"B3" -> 300
    )
    assertEvalCellValue(
      "=XLOOKUP(\"C?t\", A1:A3, B1:B3, \"\", 2)",
      sheet,
      CellValue.Number(100) // First match
    )
  }

  // ===== XLOOKUP Search Modes =====

  test("XLOOKUP: search_mode 1 (first-to-last, finds first match)") {
    val sheet = sheetWith(
      ref"A1" -> "X", ref"B1" -> 100,
      ref"A2" -> "X", ref"B2" -> 200,
      ref"A3" -> "X", ref"B3" -> 300
    )
    assertEvalCellValue(
      "=XLOOKUP(\"X\", A1:A3, B1:B3, \"\", 0, 1)",
      sheet,
      CellValue.Number(100) // First match
    )
  }

  test("XLOOKUP: search_mode -1 (last-to-first, finds last match)") {
    val sheet = sheetWith(
      ref"A1" -> "X", ref"B1" -> 100,
      ref"A2" -> "X", ref"B2" -> 200,
      ref"A3" -> "X", ref"B3" -> 300
    )
    assertEvalCellValue(
      "=XLOOKUP(\"X\", A1:A3, B1:B3, \"\", 0, -1)",
      sheet,
      CellValue.Number(300) // Last match
    )
  }

  // ===== XLOOKUP if_not_found =====

  test("XLOOKUP: if_not_found returns custom value") {
    val sheet = sheetWith(
      ref"A1" -> "Apple", ref"B1" -> 100
    )
    assertEvalCellValue(
      "=XLOOKUP(\"Cherry\", A1:A1, B1:B1, \"Not Found\")",
      sheet,
      CellValue.Text("Not Found")
    )
  }

  test("XLOOKUP: if_not_found with numeric value") {
    val sheet = sheetWith(
      ref"A1" -> "Apple", ref"B1" -> 100
    )
    assertEvalCellValue(
      "=XLOOKUP(\"Cherry\", A1:A1, B1:B1, -1)",
      sheet,
      CellValue.Number(-1)
    )
  }

  // ===== XLOOKUP Error Cases =====

  test("XLOOKUP: dimension mismatch error") {
    val sheet = sheetWith(
      ref"A1" -> "A", ref"A2" -> "B", ref"A3" -> "C",
      ref"B1" -> 1, ref"B2" -> 2
    )
    val result = evalFormulaCellValue("=XLOOKUP(\"A\", A1:A3, B1:B2)", sheet)
    assert(result.isLeft)
    assert(result.left.exists(_.contains("dimension")))
  }

  test("XLOOKUP: no match returns #N/A (default)") {
    val sheet = sheetWith(
      ref"A1" -> "Apple", ref"B1" -> 100
    )
    assertEvalCellValue(
      "=XLOOKUP(\"Banana\", A1:A1, B1:B1)",
      sheet,
      CellValue.Error(CellError.NA)
    )
  }

  test("XLOOKUP: single cell arrays") {
    val sheet = sheetWith(
      ref"A1" -> "X", ref"B1" -> 42
    )
    assertEvalCellValue(
      "=XLOOKUP(\"X\", A1:A1, B1:B1)",
      sheet,
      CellValue.Number(42)
    )
  }

  // ===== XLOOKUP Round-trip Tests =====

  test("XLOOKUP: parse -> print (3 args)") {
    val formula = "=XLOOKUP(\"A\", A1:A10, B1:B10)"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight)
    val printed = FormulaPrinter.print(parsed.toOption.get)
    assertEquals(printed, formula)
  }

  test("XLOOKUP: parse -> print (4 args with if_not_found)") {
    val formula = "=XLOOKUP(\"A\", A1:A10, B1:B10, \"Not Found\")"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight)
    val printed = FormulaPrinter.print(parsed.toOption.get)
    assertEquals(printed, formula)
  }

  test("XLOOKUP: parse -> print (6 args full)") {
    val formula = "=XLOOKUP(\"A\", A1:A10, B1:B10, \"Not Found\", 2, -1)"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight)
    val printed = FormulaPrinter.print(parsed.toOption.get)
    assertEquals(printed, formula)
  }

  // ===== Dependency Graph Tests =====

  test("DependencyGraph: SUMPRODUCT extracts all array dependencies") {
    val formula = "=SUMPRODUCT(A1:A3, B1:B3)"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight)
    val deps = DependencyGraph.extractDependencies(parsed.toOption.get)
    assertEquals(deps.size, 6) // A1, A2, A3, B1, B2, B3
    assert(deps.contains(ref"A1"))
    assert(deps.contains(ref"B3"))
  }

  test("DependencyGraph: SUMPRODUCT multiple arrays") {
    val formula = "=SUMPRODUCT(A1:A2, B1:B2, C1:C2)"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight)
    val deps = DependencyGraph.extractDependencies(parsed.toOption.get)
    assertEquals(deps.size, 6) // A1, A2, B1, B2, C1, C2
  }

  test("DependencyGraph: XLOOKUP extracts all range dependencies") {
    val formula = "=XLOOKUP(\"X\", A1:A5, B1:B5)"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight)
    val deps = DependencyGraph.extractDependencies(parsed.toOption.get)
    assertEquals(deps.size, 10) // A1-A5, B1-B5
  }

  test("DependencyGraph: XLOOKUP with cell reference lookup value") {
    val formula = "=XLOOKUP(C1, A1:A3, B1:B3)"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight)
    val deps = DependencyGraph.extractDependencies(parsed.toOption.get)
    assertEquals(deps.size, 7) // C1, A1-A3, B1-B3
    assert(deps.contains(ref"C1"))
  }
