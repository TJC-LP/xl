package com.tjclp.xl.formula

import munit.FunSuite
import com.tjclp.xl.*
import com.tjclp.xl.unsafe.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet

/**
 * Comprehensive tests for SUMIF, COUNTIF, SUMIFS, COUNTIFS conditional aggregation functions.
 *
 * Tests cover:
 *   - Basic functionality for each function
 *   - Wildcard matching (*, ?)
 *   - Numeric comparisons (>, >=, <, <=, <>)
 *   - Multiple criteria (AND logic for SUMIFS/COUNTIFS)
 *   - Round-trip parsing
 *   - Error cases
 *   - Edge cases
 */
class ConditionalFunctionsSpec extends FunSuite:

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

  // Helper to parse and evaluate a formula
  private def evalFormula(formula: String, sheet: Sheet): Either[String, BigDecimal] =
    FormulaParser.parse(formula) match
      case Right(expr) =>
        Evaluator.eval(expr, sheet) match
          case Right(value: BigDecimal) => Right(value)
          case Right(other) => Left(s"Expected BigDecimal, got: $other")
          case Left(err) => Left(s"Eval error: $err")
      case Left(err) => Left(s"Parse error: $err")

  // Helper for expected results
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

  // ===== SUMIF Basic Tests =====

  test("SUMIF: sum values where text criteria matches exactly") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"B1" -> 10,
      ref"A2" -> "Banana",
      ref"B2" -> 20,
      ref"A3" -> "Apple",
      ref"B3" -> 30
    )
    assertEval("=SUMIF(A1:A3, \"Apple\", B1:B3)", sheet, 40)
  }

  test("SUMIF: sum with same range (2-arg form)") {
    val sheet = sheetWith(
      ref"A1" -> 10,
      ref"A2" -> 20,
      ref"A3" -> 10,
      ref"A4" -> 30
    )
    assertEval("=SUMIF(A1:A4, 10)", sheet, 20)
  }

  test("SUMIF: case-insensitive text matching") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"B1" -> 10,
      ref"A2" -> "APPLE",
      ref"B2" -> 20,
      ref"A3" -> "apple",
      ref"B3" -> 30
    )
    assertEval("=SUMIF(A1:A3, \"apple\", B1:B3)", sheet, 60)
  }

  test("SUMIF: no matches returns 0") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"B1" -> 10,
      ref"A2" -> "Banana",
      ref"B2" -> 20
    )
    assertEval("=SUMIF(A1:A2, \"Cherry\", B1:B2)", sheet, 0)
  }

  test("SUMIF: skip non-numeric values in sum range") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"B1" -> 10,
      ref"A2" -> "Apple",
      ref"B2" -> "text",
      ref"A3" -> "Apple",
      ref"B3" -> 30
    )
    assertEval("=SUMIF(A1:A3, \"Apple\", B1:B3)", sheet, 40)
  }

  // ===== SUMIF Wildcard Tests =====

  test("SUMIF: wildcard * matches any characters") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"B1" -> 10,
      ref"A2" -> "Apricot",
      ref"B2" -> 20,
      ref"A3" -> "Banana",
      ref"B3" -> 30
    )
    assertEval("=SUMIF(A1:A3, \"A*\", B1:B3)", sheet, 30)
  }

  test("SUMIF: wildcard * matches suffix") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"B1" -> 10,
      ref"A2" -> "Pineapple",
      ref"B2" -> 20,
      ref"A3" -> "Banana",
      ref"B3" -> 30
    )
    assertEval("=SUMIF(A1:A3, \"*apple\", B1:B3)", sheet, 30)
  }

  test("SUMIF: wildcard ? matches single character") {
    val sheet = sheetWith(
      ref"A1" -> "Cat",
      ref"B1" -> 10,
      ref"A2" -> "Cut",
      ref"B2" -> 20,
      ref"A3" -> "Cart",
      ref"B3" -> 30
    )
    assertEval("=SUMIF(A1:A3, \"C?t\", B1:B3)", sheet, 30)
  }

  test("SUMIF: escaped wildcard ~* matches literal asterisk") {
    val sheet = sheetWith(
      ref"A1" -> "test*value",
      ref"B1" -> 10,
      ref"A2" -> "testvalue",
      ref"B2" -> 20
    )
    assertEval("=SUMIF(A1:A2, \"test~*value\", B1:B2)", sheet, 10)
  }

  // ===== SUMIF Numeric Comparison Tests =====

  test("SUMIF: greater than comparison") {
    val sheet = sheetWith(
      ref"A1" -> 50,
      ref"A2" -> 150,
      ref"A3" -> 200
    )
    assertEval("=SUMIF(A1:A3, \">100\")", sheet, 350)
  }

  test("SUMIF: greater than or equal comparison") {
    val sheet = sheetWith(
      ref"A1" -> 50,
      ref"A2" -> 100,
      ref"A3" -> 150
    )
    assertEval("=SUMIF(A1:A3, \">=100\")", sheet, 250)
  }

  test("SUMIF: less than comparison") {
    val sheet = sheetWith(
      ref"A1" -> 50,
      ref"A2" -> 100,
      ref"A3" -> 150
    )
    assertEval("=SUMIF(A1:A3, \"<100\")", sheet, 50)
  }

  test("SUMIF: less than or equal comparison") {
    val sheet = sheetWith(
      ref"A1" -> 50,
      ref"A2" -> 100,
      ref"A3" -> 150
    )
    assertEval("=SUMIF(A1:A3, \"<=100\")", sheet, 150)
  }

  test("SUMIF: not equal comparison") {
    val sheet = sheetWith(
      ref"A1" -> 0,
      ref"A2" -> 100,
      ref"A3" -> 0,
      ref"A4" -> 200
    )
    assertEval("=SUMIF(A1:A4, \"<>0\")", sheet, 300)
  }

  test("SUMIF: negative number comparison") {
    val sheet = sheetWith(
      ref"A1" -> -50,
      ref"A2" -> 0,
      ref"A3" -> 50
    )
    assertEval("=SUMIF(A1:A3, \">-25\")", sheet, 50)
  }

  // ===== COUNTIF Tests =====

  test("COUNTIF: count matching text") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"A2" -> "Banana",
      ref"A3" -> "Apple",
      ref"A4" -> "Cherry"
    )
    assertEval("=COUNTIF(A1:A4, \"Apple\")", sheet, 2)
  }

  test("COUNTIF: count with wildcard") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"A2" -> "Apricot",
      ref"A3" -> "Banana",
      ref"A4" -> "Avocado"
    )
    assertEval("=COUNTIF(A1:A4, \"A*\")", sheet, 3)
  }

  test("COUNTIF: count with numeric comparison") {
    val sheet = sheetWith(
      ref"A1" -> 50,
      ref"A2" -> 150,
      ref"A3" -> 200,
      ref"A4" -> 75
    )
    assertEval("=COUNTIF(A1:A4, \">100\")", sheet, 2)
  }

  test("COUNTIF: count exact number match") {
    val sheet = sheetWith(
      ref"A1" -> 100,
      ref"A2" -> 200,
      ref"A3" -> 100,
      ref"A4" -> 300
    )
    assertEval("=COUNTIF(A1:A4, 100)", sheet, 2)
  }

  test("COUNTIF: no matches returns 0") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"A2" -> "Banana"
    )
    assertEval("=COUNTIF(A1:A2, \"Cherry\")", sheet, 0)
  }

  test("COUNTIF: count boolean values") {
    val sheet = sheetWith(
      ref"A1" -> true,
      ref"A2" -> false,
      ref"A3" -> true,
      ref"A4" -> true
    )
    assertEval("=COUNTIF(A1:A4, TRUE)", sheet, 3)
  }

  // ===== SUMIFS Tests (Multiple Criteria) =====

  test("SUMIFS: two criteria AND logic") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"B1" -> "Red",
      ref"C1" -> 10,
      ref"A2" -> "Apple",
      ref"B2" -> "Green",
      ref"C2" -> 20,
      ref"A3" -> "Banana",
      ref"B3" -> "Yellow",
      ref"C3" -> 30
    )
    assertEval("=SUMIFS(C1:C3, A1:A3, \"Apple\", B1:B3, \"Red\")", sheet, 10)
  }

  test("SUMIFS: mixed text and numeric criteria") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"B1" -> 100,
      ref"C1" -> 10,
      ref"A2" -> "Apple",
      ref"B2" -> 50,
      ref"C2" -> 20,
      ref"A3" -> "Banana",
      ref"B3" -> 150,
      ref"C3" -> 30
    )
    assertEval("=SUMIFS(C1:C3, A1:A3, \"Apple\", B1:B3, \">75\")", sheet, 10)
  }

  test("SUMIFS: all criteria must match") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"B1" -> "Red",
      ref"C1" -> 10,
      ref"A2" -> "Apple",
      ref"B2" -> "Green",
      ref"C2" -> 20
    )
    // Apple + Red matches only row 1
    assertEval("=SUMIFS(C1:C2, A1:A2, \"Apple\", B1:B2, \"Red\")", sheet, 10)
    // Apple alone would match both, but with Blue nothing matches
    assertEval("=SUMIFS(C1:C2, A1:A2, \"Apple\", B1:B2, \"Blue\")", sheet, 0)
  }

  test("SUMIFS: wildcard in multiple criteria") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"B1" -> "Red Apple",
      ref"C1" -> 10,
      ref"A2" -> "Apple",
      ref"B2" -> "Green Apple",
      ref"C2" -> 20,
      ref"A3" -> "Banana",
      ref"B3" -> "Yellow Banana",
      ref"C3" -> 30
    )
    assertEval("=SUMIFS(C1:C3, A1:A3, \"A*\", B1:B3, \"*Apple\")", sheet, 30)
  }

  // ===== COUNTIFS Tests (Multiple Criteria) =====

  test("COUNTIFS: two criteria") {
    val sheet = sheetWith(
      ref"A1" -> "Apple",
      ref"B1" -> "Red",
      ref"A2" -> "Apple",
      ref"B2" -> "Green",
      ref"A3" -> "Banana",
      ref"B3" -> "Yellow"
    )
    assertEval("=COUNTIFS(A1:A3, \"Apple\", B1:B3, \"Red\")", sheet, 1)
  }

  test("COUNTIFS: numeric comparisons") {
    val sheet = sheetWith(
      ref"A1" -> 100,
      ref"B1" -> 50,
      ref"A2" -> 200,
      ref"B2" -> 150,
      ref"A3" -> 300,
      ref"B3" -> 75
    )
    // Count where A > 100 AND B < 100
    assertEval("=COUNTIFS(A1:A3, \">100\", B1:B3, \"<100\")", sheet, 1)
  }

  test("COUNTIFS: all criteria must match") {
    val sheet = sheetWith(
      ref"A1" -> "Yes",
      ref"B1" -> "Yes",
      ref"A2" -> "Yes",
      ref"B2" -> "No",
      ref"A3" -> "No",
      ref"B3" -> "Yes"
    )
    assertEval("=COUNTIFS(A1:A3, \"Yes\", B1:B3, \"Yes\")", sheet, 1)
  }

  // ===== Round-trip Tests =====

  test("SUMIF: parse -> print -> parse roundtrip") {
    val formula = "=SUMIF(A1:A10, \"Apple\", B1:B10)"
    FormulaParser.parse(formula) match
      case Right(expr) =>
        val printed = FormulaPrinter.print(expr)
        assertEquals(printed, formula)
      case Left(err) => fail(s"Parse error: $err")
  }

  test("SUMIF: 2-arg form roundtrip") {
    val formula = "=SUMIF(A1:A10, \">100\")"
    FormulaParser.parse(formula) match
      case Right(expr) =>
        val printed = FormulaPrinter.print(expr)
        assertEquals(printed, formula)
      case Left(err) => fail(s"Parse error: $err")
  }

  test("COUNTIF: parse -> print -> parse roundtrip") {
    val formula = "=COUNTIF(A1:A10, \"Apple\")"
    FormulaParser.parse(formula) match
      case Right(expr) =>
        val printed = FormulaPrinter.print(expr)
        assertEquals(printed, formula)
      case Left(err) => fail(s"Parse error: $err")
  }

  test("SUMIFS: parse -> print -> parse roundtrip") {
    val formula = "=SUMIFS(C1:C10, A1:A10, \"Apple\", B1:B10, \">100\")"
    FormulaParser.parse(formula) match
      case Right(expr) =>
        val printed = FormulaPrinter.print(expr)
        assertEquals(printed, formula)
      case Left(err) => fail(s"Parse error: $err")
  }

  test("COUNTIFS: parse -> print -> parse roundtrip") {
    val formula = "=COUNTIFS(A1:A10, \"Apple\", B1:B10, \">100\")"
    FormulaParser.parse(formula) match
      case Right(expr) =>
        val printed = FormulaPrinter.print(expr)
        assertEquals(printed, formula)
      case Left(err) => fail(s"Parse error: $err")
  }

  // ===== Edge Cases =====

  test("SUMIF: empty range returns 0") {
    val sheet = Sheet("Test")
    assertEval("=SUMIF(A1:A3, \"Apple\")", sheet, 0)
  }

  test("COUNTIF: empty range returns 0") {
    val sheet = Sheet("Test")
    assertEval("=COUNTIF(A1:A3, \"Apple\")", sheet, 0)
  }

  test("SUMIF: formula cells with cached values") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=1+1", Some(CellValue.Number(BigDecimal(100)))),
      ref"B1" -> 10,
      ref"A2" -> CellValue.Number(BigDecimal(100)),
      ref"B2" -> 20
    )
    // Should match both: formula with cached 100, and literal 100
    assertEval("=SUMIF(A1:A2, 100, B1:B2)", sheet, 30)
  }

  test("SUMIF: single cell range") {
    val sheet = sheetWith(ref"A1" -> "Apple", ref"B1" -> 100)
    assertEval("=SUMIF(A1:A1, \"Apple\", B1:B1)", sheet, 100)
  }

  test("COUNTIF: single cell match") {
    val sheet = sheetWith(ref"A1" -> "Apple")
    assertEval("=COUNTIF(A1:A1, \"Apple\")", sheet, 1)
  }

  test("SUMIF: decimal number criteria") {
    val sheet = sheetWith(
      ref"A1" -> BigDecimal("3.14"),
      ref"B1" -> 10,
      ref"A2" -> BigDecimal("3.14"),
      ref"B2" -> 20
    )
    assertEval("=SUMIF(A1:A2, 3.14, B1:B2)", sheet, 30)
  }

  // ===== Dependency Graph Tests =====

  test("DependencyGraph: SUMIF extracts all range dependencies") {
    val formula = "=SUMIF(A1:A3, \"Apple\", B1:B3)"
    FormulaParser.parse(formula) match
      case Right(expr) =>
        val deps = DependencyGraph.extractDependencies(expr)
        // Should include A1, A2, A3 (criteria range) and B1, B2, B3 (sum range)
        assert(deps.contains(ref"A1"))
        assert(deps.contains(ref"A2"))
        assert(deps.contains(ref"A3"))
        assert(deps.contains(ref"B1"))
        assert(deps.contains(ref"B2"))
        assert(deps.contains(ref"B3"))
        assertEquals(deps.size, 6)
      case Left(err) => fail(s"Parse error: $err")
  }

  test("DependencyGraph: COUNTIF extracts range dependencies") {
    val formula = "=COUNTIF(A1:A3, \"Apple\")"
    FormulaParser.parse(formula) match
      case Right(expr) =>
        val deps = DependencyGraph.extractDependencies(expr)
        assert(deps.contains(ref"A1"))
        assert(deps.contains(ref"A2"))
        assert(deps.contains(ref"A3"))
        assertEquals(deps.size, 3)
      case Left(err) => fail(s"Parse error: $err")
  }

  test("DependencyGraph: SUMIFS extracts all ranges") {
    val formula = "=SUMIFS(C1:C2, A1:A2, \"X\", B1:B2, \"Y\")"
    FormulaParser.parse(formula) match
      case Right(expr) =>
        val deps = DependencyGraph.extractDependencies(expr)
        // Should include C1, C2, A1, A2, B1, B2
        assertEquals(deps.size, 6)
      case Left(err) => fail(s"Parse error: $err")
  }
