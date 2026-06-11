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

  // ============================================================================
  // GH-306: cross-type CALL-RESULT coercion in typed argument positions
  // (the wave-3 let-rand reviewer's probe list + totality sweep)
  // ============================================================================

  test("GH-306: UPPER(SUM(A1:A2)) renders the numeric call result as text") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(20))
    )
    assertEquals(sheet.evaluateFormula("=UPPER(SUM(A1:A2))"), Right(CellValue.Text("30")))
  }

  test("GH-306: LEN(SUM(A1:A2)) measures the rendered number") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(20))
    )
    assertEquals(sheet.evaluateFormula("=LEN(SUM(A1:A2))"), Right(CellValue.Number(BigDecimal(2))))
  }

  test("GH-306: IF(SUM(A1:A2), 1, 2) uses Excel truthiness (non-zero = TRUE)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(20))
    )
    assertEquals(sheet.evaluateFormula("=IF(SUM(A1:A2), 1, 2)"), Right(CellValue.Number(BigDecimal(1))))
  }

  test("GH-306: IF(SUM(empty), 1, 2) — zero is FALSE") {
    val sheet = sheetWith(ref"Z9" -> CellValue.Text("unrelated"))
    assertEquals(sheet.evaluateFormula("=IF(SUM(A1:A2), 1, 2)"), Right(CellValue.Number(BigDecimal(2))))
  }

  test("GH-306: LEFT(text, SUM(...)) — numeric call result in int position (probe seed)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(1)),
      ref"A2" -> CellValue.Number(BigDecimal(2))
    )
    assertEquals(sheet.evaluateFormula("=LEFT(\"hello\", SUM(A1:A2))"), Right(CellValue.Text("hel")))
  }

  test("GH-306/GH-307: fractional call result in int position TRUNCATES like Excel") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal("1.5")),
      ref"A2" -> CellValue.Number(BigDecimal("1.2"))
    )
    // SUM = 2.7 → LEFT("hello", 2.7) → Excel truncates → "he"
    assertEquals(sheet.evaluateFormula("=LEFT(\"hello\", SUM(A1:A2))"), Right(CellValue.Text("he")))
  }

  test("GH-306: numeric TEXT call result coerces in numeric position (SQRT(LEFT(\"16ab\",2)))") {
    assertEquals(
      emptySheet.evaluateFormula("=SQRT(LEFT(\"16ab\", 2))"),
      Right(CellValue.Number(BigDecimal(4)))
    )
  }

  test("GH-306: boolean call results in boolean positions are unchanged (regression guard)") {
    assertEquals(
      emptySheet.evaluateFormula("=IF(AND(TRUE, TRUE), 1, 2)"),
      Right(CellValue.Number(BigDecimal(1)))
    )
    assertEquals(
      emptySheet.evaluateFormula("=IF(NOT(FALSE), 1, 2)"),
      Right(CellValue.Number(BigDecimal(1)))
    )
  }

  test("GH-306: boolean call result coerces to text (UPPER(AND(TRUE,TRUE)) = \"TRUE\")") {
    assertEquals(
      emptySheet.evaluateFormula("=UPPER(AND(TRUE, TRUE))"),
      Right(CellValue.Text("TRUE"))
    )
  }

  test("GH-306: IF branch value flows through a text position (LOWER(IF(...)))") {
    assertEquals(
      emptySheet.evaluateFormula("=LOWER(IF(TRUE, \"ABC\", 5))"),
      Right(CellValue.Text("abc"))
    )
    // the numeric branch renders as text
    assertEquals(
      emptySheet.evaluateFormula("=LOWER(IF(FALSE, \"ABC\", 5))"),
      Right(CellValue.Text("5"))
    )
  }

  test("GH-306: time call result in a date position (YEAR(NOW()))") {
    val clock = Clock.fixed(LocalDate.of(2026, 6, 10), LocalDateTime.of(2026, 6, 10, 12, 0))
    assertEquals(
      emptySheet.evaluateFormula("=YEAR(NOW())", clock),
      Right(CellValue.Number(BigDecimal(2026)))
    )
  }

  test("GH-306: serial arithmetic in a date position (MONTH(TODAY()+40))") {
    val clock = Clock.fixed(LocalDate.of(2026, 6, 10), LocalDateTime.of(2026, 6, 10, 12, 0))
    // 2026-06-10 + 40 days = 2026-07-20 → month 7
    assertEquals(
      emptySheet.evaluateFormula("=MONTH(TODAY()+40)", clock),
      Right(CellValue.Number(BigDecimal(7)))
    )
  }

  test("GH-306: uncoercible call results are clean per-cell errors, never thrown") {
    // text where a number is needed
    val r1 = emptySheet.evaluateFormula("=ABS(UPPER(\"xy\"))")
    assert(r1.isLeft, s"expected clean Left for text in numeric position, got $r1")
    // text where a boolean is needed
    val r2 = emptySheet.evaluateFormula("=IF(\"a\"&\"b\", 1, 2)")
    assert(r2.isLeft, s"expected clean Left for text in boolean position, got $r2")
    // text where an int is needed
    val r3 = emptySheet.evaluateFormula("=LEFT(\"hello\", UPPER(\"xy\"))")
    assert(r3.isLeft, s"expected clean Left for text in int position, got $r3")
    // text where a date is needed
    val r4 = emptySheet.evaluateFormula("=YEAR(UPPER(\"xy\"))")
    assert(r4.isLeft, s"expected clean Left for text in date position, got $r4")
  }

  test("GH-306: totality sweep — cross-type compositions never throw") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(20)),
      ref"B1" -> CellValue.Text("text")
    )
    val probes = List(
      "=UPPER(SUM(A1:A2))",
      "=IF(SUM(A1:A2), 1, 2)",
      "=LEFT(B1, SUM(A1:A2))",
      "=SQRT(CONCATENATE(\"1\", \"6\"))",
      "=IF(CONCATENATE(\"a\", \"b\"), 1, 2)",
      "=ABS(IF(TRUE, B1, B1))",
      "=YEAR(SUM(A1:A2))",
      "=LOWER(MAX(A1:A2))",
      "=NOT(SUM(A1:A2))",
      "=LEN(AVERAGE(A1:A2))"
    )
    probes.foreach { f =>
      val result = sheet.evaluateFormula(f) // must not throw — Left is acceptable
      assert(result.isLeft || result.isRight, s"unreachable, evaluated: $f")
    }
  }

  test("GH-306: round-trip — coercion wrappers print transparently") {
    val formulas = List(
      "=UPPER(SUM(A1:A2))",
      "=IF(SUM(A1:A2), 1, 2)",
      "=LEFT(\"hello\", SUM(A1:A2))",
      "=YEAR(NOW())",
      "=SQRT(LEFT(\"16ab\", 2))"
    )
    formulas.foreach { f =>
      FormulaParser.parse(f) match
        case Right(expr) =>
          assertEquals(FormulaPrinter.print(expr), f)
          assertEquals(FormulaParser.parse(FormulaPrinter.print(expr)), Right(expr))
        case Left(err) => fail(s"parse failed for $f: $err")
    }
  }
