package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.addressing.SheetName
// conversions.given and SheetEvaluator extension methods now available from com.tjclp.xl.{*, given}
import java.time.{LocalDate, LocalDateTime}
import munit.FunSuite

/**
 * Regression tests for P1 bug fix: Cell reference type coercion.
 *
 * Previously, all cell references were decoded as numeric (TExpr.decodeNumeric), causing text/date
 * functions to fail with TypeMismatch when referencing non-numeric cells.
 *
 * Fix: Parser creates PolyRef for cell references; function parsers convert to typed Ref with
 * appropriate coercing decoder matching Excel semantics.
 *
 * Tests verify:
 *   - Text functions work with cell references (not just literals)
 *   - Date functions work with cell references
 *   - Automatic type coercion matches Excel behavior (42 → "42" in text functions)
 *   - Round-trip parsing preserves PolyRef
 */
class TypeCoercionSpec extends FunSuite:

  val emptySheet = new Sheet(name = SheetName.unsafe("Test"))

  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  // ============================================================================
  // Text Function Coercion Tests (9 tests)
  // ============================================================================

  test("LEFT coerces numeric cell to string") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(12345)))
    val result = sheet.evaluateFormula("=LEFT(A1, 3)")
    assertEquals(result, Right(CellValue.Text("123")))
  }

  test("RIGHT coerces numeric cell to string") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(12345)))
    val result = sheet.evaluateFormula("=RIGHT(A1, 2)")
    assertEquals(result, Right(CellValue.Text("45")))
  }

  test("LEN coerces numeric cell to string") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(12345)))
    val result = sheet.evaluateFormula("=LEN(A1)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(5))))
  }

  test("UPPER with text cell reference") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("hello"))
    val result = sheet.evaluateFormula("=UPPER(A1)")
    assertEquals(result, Right(CellValue.Text("HELLO")))
  }

  test("LOWER with text cell reference") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("WORLD"))
    val result = sheet.evaluateFormula("=LOWER(A1)")
    assertEquals(result, Right(CellValue.Text("world")))
  }

  test("CONCATENATE with mixed cell types") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(42)),
      ref"B1" -> CellValue.Text("text")
    )
    val result = sheet.evaluateFormula("=CONCATENATE(A1, B1)")
    // Excel coerces 42 to "42", concatenates with "text"
    assertEquals(result, Right(CellValue.Text("42text")))
  }

  test("LEFT with empty cell returns empty string") {
    val sheet = sheetWith(ref"A1" -> CellValue.Empty)
    val result = sheet.evaluateFormula("=LEFT(A1, 5)")
    assertEquals(result, Right(CellValue.Text("")))
  }

  test("LEN with empty cell returns 0") {
    val sheet = sheetWith(ref"A1" -> CellValue.Empty)
    val result = sheet.evaluateFormula("=LEN(A1)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(0))))
  }

  test("CONCATENATE with literal and cell reference") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(100)))
    val result = sheet.evaluateFormula("=CONCATENATE(\"Value: \", A1)")
    assertEquals(result, Right(CellValue.Text("Value: 100")))
  }

  // ============================================================================
  // Date Function Coercion Tests (6 tests)
  // ============================================================================

  test("YEAR with date cell reference") {
    val sheet = sheetWith(ref"A1" -> CellValue.DateTime(LocalDateTime.of(2025, 11, 21, 10, 30)))
    val result = sheet.evaluateFormula("=YEAR(A1)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(2025))))
  }

  test("MONTH with date cell reference") {
    val sheet = sheetWith(ref"A1" -> CellValue.DateTime(LocalDateTime.of(2025, 11, 21, 10, 30)))
    val result = sheet.evaluateFormula("=MONTH(A1)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(11))))
  }

  test("DAY with date cell reference") {
    val sheet = sheetWith(ref"A1" -> CellValue.DateTime(LocalDateTime.of(2025, 11, 21, 10, 30)))
    val result = sheet.evaluateFormula("=DAY(A1)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(21))))
  }

  test("Date extraction: All three functions with same cell") {
    val originalDate = LocalDateTime.of(2025, 11, 21, 0, 0)
    val sheet = sheetWith(ref"A1" -> CellValue.DateTime(originalDate))

    // Verify all three extraction functions work with cell references
    val year = sheet.evaluateFormula("=YEAR(A1)")
    val month = sheet.evaluateFormula("=MONTH(A1)")
    val day = sheet.evaluateFormula("=DAY(A1)")

    assertEquals(year, Right(CellValue.Number(BigDecimal(2025))))
    assertEquals(month, Right(CellValue.Number(BigDecimal(11))))
    assertEquals(day, Right(CellValue.Number(BigDecimal(21))))
  }

  test("YEAR with TODAY function") {
    val clock = Clock.fixedDate(LocalDate.of(2025, 11, 21))
    val result = emptySheet.evaluateFormula("=YEAR(TODAY())", clock)
    assertEquals(result, Right(CellValue.Number(BigDecimal(2025))))
  }

  test("YEAR with non-date cell returns error") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("not a date"))
    val result = sheet.evaluateFormula("=YEAR(A1)")
    assert(result.isLeft, "Expected error for text cell in YEAR function")
  }

  // ============================================================================
  // Logical Function Coercion Tests (4 tests)
  // ============================================================================

  test("IF with cell reference in condition") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(10)))
    val result = sheet.evaluateFormula("=IF(A1>0, \"Positive\", \"Negative\")")
    assertEquals(result, Right(CellValue.Text("Positive")))
  }

  test("AND with boolean cell references") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Bool(true),
      ref"B1" -> CellValue.Bool(true)
    )
    val result = sheet.evaluateFormula("=AND(A1, B1)")
    assertEquals(result, Right(CellValue.Bool(true)))
  }

  test("OR with boolean cell references") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Bool(false),
      ref"B1" -> CellValue.Bool(true)
    )
    val result = sheet.evaluateFormula("=OR(A1, B1)")
    assertEquals(result, Right(CellValue.Bool(true)))
  }

  test("NOT with boolean cell reference") {
    val sheet = sheetWith(ref"A1" -> CellValue.Bool(false))
    val result = sheet.evaluateFormula("=NOT(A1)")
    assertEquals(result, Right(CellValue.Bool(true)))
  }

  // ============================================================================
  // Arithmetic Function Coercion Tests (3 tests)
  // ============================================================================

  test("Addition with cell references (regression for P1 bug)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(100)),
      ref"A2" -> CellValue.Number(BigDecimal(200))
    )
    val result = sheet.evaluateFormula("=A1+A2")
    assertEquals(result, Right(CellValue.Number(BigDecimal(300))))
  }

  test("MIN with cell range") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(5)),
      ref"A3" -> CellValue.Number(BigDecimal(20))
    )
    val result = sheet.evaluateFormula("=MIN(A1:A3)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(5))))
  }

  test("MAX with cell range") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(5)),
      ref"A3" -> CellValue.Number(BigDecimal(20))
    )
    val result = sheet.evaluateFormula("=MAX(A1:A3)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(20))))
  }

  // ============================================================================
  // Round-Trip Parser Tests (3 tests)
  // ============================================================================

  test("Round-trip: LEFT with cell reference") {
    val original = "=LEFT(A1, 3)"
    FormulaParser.parse(original) match
      case Right(expr) =>
        val printed = FormulaPrinter.print(expr, includeEquals = true)
        assertEquals(printed, original)

        // Verify re-parsing produces equivalent AST
        val reparsed = FormulaParser.parse(printed)
        assert(reparsed.isRight, "Round-trip parse failed")

      case Left(err) => fail(s"Failed to parse: $err")
  }

  test("Round-trip: YEAR with cell reference") {
    val original = "=YEAR(B5)"
    FormulaParser.parse(original) match
      case Right(expr) =>
        val printed = FormulaPrinter.print(expr, includeEquals = true)
        assertEquals(printed, original)

        val reparsed = FormulaParser.parse(printed)
        assert(reparsed.isRight, "Round-trip parse failed")

      case Left(err) => fail(s"Failed to parse: $err")
  }

  test("Round-trip: IF with cell references") {
    val original = "=IF(A1, B1, C1)"
    FormulaParser.parse(original) match
      case Right(expr) =>
        val printed = FormulaPrinter.print(expr, includeEquals = true)
        assertEquals(printed, original)

        val reparsed = FormulaParser.parse(printed)
        assert(reparsed.isRight, "Round-trip parse failed")

      case Left(err) => fail(s"Failed to parse: $err")
  }

  // ============================================================================
  // Complex Integration Tests (2 tests)
  // ============================================================================

  test("Complex: Text function chain with cell references") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Text("hello world"),
      ref"B1" -> CellValue.Number(BigDecimal(5))
    )
    // LEFT(UPPER(A1), B1) → "HELLO"
    val result = sheet.evaluateFormula("=LEFT(UPPER(A1), B1)")
    assertEquals(result, Right(CellValue.Text("HELLO")))
  }

  test("Complex: Date extraction with arithmetic") {
    val sheet = sheetWith(ref"A1" -> CellValue.DateTime(LocalDateTime.of(2025, 11, 21, 0, 0)))
    // YEAR(A1) + MONTH(A1) → 2025 + 11 = 2036 (tests BigDecimal return type)
    val result = sheet.evaluateFormula("=YEAR(A1) + MONTH(A1)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(2036))))
  }
