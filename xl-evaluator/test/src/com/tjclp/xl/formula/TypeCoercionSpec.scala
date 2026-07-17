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
  // GH-385: blank cells coerce to 0 in scalar numeric positions
  // ============================================================================

  test("GH-385: =A1+1 with explicitly Empty A1 is 1 (blank-as-zero)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Empty)
    assertEquals(sheet.evaluateFormula("=A1+1"), Right(CellValue.Number(BigDecimal(1))))
  }

  test("GH-385: =A1+1 with absent A1 is 1 (blank-as-zero)") {
    val sheet = sheetWith(ref"Z9" -> CellValue.Text("unrelated"))
    assertEquals(sheet.evaluateFormula("=A1+1"), Right(CellValue.Number(BigDecimal(1))))
  }

  test("GH-385: blank cells are 0 across the scalar operators (*, %, unary -, ^)") {
    val sheet = sheetWith(ref"Z9" -> CellValue.Text("unrelated"))
    assertEquals(sheet.evaluateFormula("=A1*1"), Right(CellValue.Number(BigDecimal(0))))
    assertEquals(sheet.evaluateFormula("=A1%"), Right(CellValue.Number(BigDecimal(0))))
    assertEquals(sheet.evaluateFormula("=-A1"), Right(CellValue.Number(BigDecimal(0))))
    assertEquals(sheet.evaluateFormula("=A1^2"), Right(CellValue.Number(BigDecimal(0))))
  }

  test("GH-385: blank ref as a scalar function argument is 0 (=ABS(A1))") {
    val sheet = sheetWith(ref"Z9" -> CellValue.Text("unrelated"))
    assertEquals(sheet.evaluateFormula("=ABS(A1)"), Right(CellValue.Number(BigDecimal(0))))
  }

  test("GH-385: SUM args mixing blank direct refs no longer reject (issue repro)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(10)))
    // A2 is blank: Excel evaluates SUM(A1, A2) = 10 rather than a strict-typing error
    assertEquals(sheet.evaluateFormula("=SUM(A1, A2)"), Right(CellValue.Number(BigDecimal(10))))
  }

  test("GH-385: an error-valued cell still propagates through =A1+1 (not masked to 1)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(com.tjclp.xl.cells.CellError.Div0))
    // GH-344: propagation now carries the Excel error VALUE with its code (was a loud Left)
    assertEquals(
      sheet.evaluateFormula("=A1+1"),
      Right(CellValue.Error(com.tjclp.xl.cells.CellError.Div0))
    )
  }

  test("GH-385: text cell in numeric position is still a clean error (not 0)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("abc"))
    val result = sheet.evaluateFormula("=A1+1")
    assert(result.isLeft, s"expected error for text cell in numeric position, got $result")
  }

  test("GH-385: range aggregates still SKIP blanks while scalar refs coerce to 0") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      // A2 deliberately blank
      ref"A3" -> CellValue.Number(BigDecimal(20))
    )
    // Both must hold simultaneously: blank-skip in range folds, blank-as-zero at scalar refs
    assertEquals(sheet.evaluateFormula("=SUM(A1:A3)"), Right(CellValue.Number(BigDecimal(30))))
    assertEquals(sheet.evaluateFormula("=COUNT(A1:A3)"), Right(CellValue.Number(BigDecimal(2))))
    assertEquals(sheet.evaluateFormula("=MIN(A1:A3)"), Right(CellValue.Number(BigDecimal(10))))
    assertEquals(sheet.evaluateFormula("=AVERAGE(A1:A3)"), Right(CellValue.Number(BigDecimal(15))))
    assertEquals(sheet.evaluateFormula("=A2+1"), Right(CellValue.Number(BigDecimal(1))))
  }

  // ============================================================================
  // GH-385: date functions coerce raw Excel serial Numbers in date positions
  // ============================================================================

  private val serialDate = LocalDateTime.of(2024, 3, 15, 0, 0)
  private val dateSerial = BigDecimal(CellValue.dateTimeToExcelSerial(serialDate))

  test("GH-385: YEAR/MONTH/DAY accept a raw serial Number cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(dateSerial))
    assertEquals(sheet.evaluateFormula("=YEAR(A1)"), Right(CellValue.Number(BigDecimal(2024))))
    assertEquals(sheet.evaluateFormula("=MONTH(A1)"), Right(CellValue.Number(BigDecimal(3))))
    assertEquals(sheet.evaluateFormula("=DAY(A1)"), Right(CellValue.Number(BigDecimal(15))))
  }

  test("GH-385: EOMONTH accepts a raw serial Number cell (issue repro)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(dateSerial))
    sheet.evaluateFormula("=EOMONTH(A1, 0)") match
      case Right(CellValue.DateTime(dt)) =>
        assertEquals(dt.toLocalDate, java.time.LocalDate.of(2024, 3, 31))
      case other => fail(s"Expected DateTime, got $other")
  }

  test("GH-385: EDATE accepts a raw serial Number cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(dateSerial))
    sheet.evaluateFormula("=EDATE(A1, 1)") match
      case Right(CellValue.DateTime(dt)) =>
        assertEquals(dt.toLocalDate, java.time.LocalDate.of(2024, 4, 15))
      case other => fail(s"Expected DateTime, got $other")
  }

  test("GH-385: YEARFRAC accepts raw serial Number cells") {
    val start = BigDecimal(CellValue.dateTimeToExcelSerial(LocalDateTime.of(2024, 1, 1, 0, 0)))
    val end = BigDecimal(CellValue.dateTimeToExcelSerial(LocalDateTime.of(2025, 1, 1, 0, 0)))
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(start),
      ref"B1" -> CellValue.Number(end)
    )
    assertEquals(sheet.evaluateFormula("=YEARFRAC(A1, B1)"), Right(CellValue.Number(BigDecimal(1))))
  }

  test("GH-385: fractional serial (date + time) still lands on the right date") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(dateSerial + BigDecimal("0.5")))
    assertEquals(sheet.evaluateFormula("=YEAR(A1)"), Right(CellValue.Number(BigDecimal(2024))))
    assertEquals(sheet.evaluateFormula("=DAY(A1)"), Right(CellValue.Number(BigDecimal(15))))
  }

  test("GH-385: formula cell with cached serial Number coerces in date position") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=B1*1", Some(CellValue.Number(dateSerial)))
    )
    assertEquals(sheet.evaluateFormula("=YEAR(A1)"), Right(CellValue.Number(BigDecimal(2024))))
  }

  test("GH-385: formula cell with cached DateTime coerces in date position") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=DATE(2024,3,15)", Some(CellValue.DateTime(serialDate)))
    )
    assertEquals(sheet.evaluateFormula("=YEAR(A1)"), Right(CellValue.Number(BigDecimal(2024))))
  }

  test("GH-385: negative serial in a date position is a clean error") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(-1)))
    val result = sheet.evaluateFormula("=YEAR(A1)")
    assert(result.isLeft, s"expected error for negative serial in date position, got $result")
  }

  test("GH-385: oversized serial (> 9999-12-31) in a date position is a clean error") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(2958466)))
    val result = sheet.evaluateFormula("=YEAR(A1)")
    assert(result.isLeft, s"expected error for oversized serial in date position, got $result")
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
    assertEquals(
      sheet.evaluateFormula("=IF(SUM(A1:A2), 1, 2)"),
      Right(CellValue.Number(BigDecimal(1)))
    )
  }

  test("GH-306: IF(SUM(empty), 1, 2) — zero is FALSE") {
    val sheet = sheetWith(ref"Z9" -> CellValue.Text("unrelated"))
    assertEquals(
      sheet.evaluateFormula("=IF(SUM(A1:A2), 1, 2)"),
      Right(CellValue.Number(BigDecimal(2)))
    )
  }

  test("GH-306: LEFT(text, SUM(...)) — numeric call result in int position (probe seed)") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(1)),
      ref"A2" -> CellValue.Number(BigDecimal(2))
    )
    assertEquals(
      sheet.evaluateFormula("=LEFT(\"hello\", SUM(A1:A2))"),
      Right(CellValue.Text("hel"))
    )
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

  // ============================================================================
  // GH-396: remaining coercion edges — the decoder tables mirror ScalarCoercion
  // exactly (int positions, blank/Bool cells in date positions)
  // ============================================================================

  test("GH-396: =LEFT(\"hello\", A1) with blank A1 is \"\" (Empty is 0 in int positions)") {
    // Issue repro: Excel evaluates LEFT("hello", <blank>) as LEFT("hello", 0) = ""
    val sheet = sheetWith(ref"Z9" -> CellValue.Text("unrelated"))
    assertEquals(sheet.evaluateFormula("=LEFT(\"hello\", A1)"), Right(CellValue.Text("")))
  }

  test("GH-396: numeric-text cell in int position parses (coerceInteger mirror)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("3"))
    assertEquals(sheet.evaluateFormula("=LEFT(\"hello\", A1)"), Right(CellValue.Text("hel")))
  }

  test("GH-396: cached-formula Number cell in int position unwraps") {
    val sheet =
      sheetWith(ref"A1" -> CellValue.Formula("=1+2", Some(CellValue.Number(BigDecimal(3)))))
    assertEquals(sheet.evaluateFormula("=LEFT(\"hello\", A1)"), Right(CellValue.Text("hel")))
  }

  test("GH-396: cached-formula Text cell in int position parses numeric text") {
    val sheet =
      sheetWith(ref"A1" -> CellValue.Formula("=\"3\"", Some(CellValue.Text("3"))))
    assertEquals(sheet.evaluateFormula("=LEFT(\"hello\", A1)"), Right(CellValue.Text("hel")))
  }

  test("GH-396: fractional Number cell in int position TRUNCATES like Excel") {
    // The cached-call boundary already truncated (GH-306/307 pin above); the direct-cell
    // boundary must agree — LEFT("hello", 2.7) is "he" whether 2.7 is a literal, a call
    // result, or a cell value
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal("2.7")))
    assertEquals(sheet.evaluateFormula("=LEFT(\"hello\", A1)"), Right(CellValue.Text("he")))
  }

  test("GH-396: non-numeric text cell in int position is a clean error (not 0)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("abc"))
    val result = sheet.evaluateFormula("=LEFT(\"hello\", A1)")
    assert(result.isLeft, s"expected error for non-numeric text in int position, got $result")
  }

  test("GH-396: YEAR/MONTH of a blank cell match Excel's serial-0 rendering (1900/1)") {
    // Excel renders a blank date argument as serial 0 = the phantom "January 0, 1900":
    // =YEAR(blank) is 1900, =MONTH(blank) is 1, =DAY(blank) is 0. LocalDate cannot represent
    // a day-0 date, so the Empty arm maps to 1900-01-01 — YEAR and MONTH are Excel-exact,
    // DAY reads 1 instead of 0 (documented divergence, the phantom day is unrepresentable).
    val sheet = sheetWith(ref"Z9" -> CellValue.Text("unrelated"))
    assertEquals(sheet.evaluateFormula("=YEAR(A1)"), Right(CellValue.Number(BigDecimal(1900))))
    assertEquals(sheet.evaluateFormula("=MONTH(A1)"), Right(CellValue.Number(BigDecimal(1))))
    assertEquals(sheet.evaluateFormula("=DAY(A1)"), Right(CellValue.Number(BigDecimal(1))))
  }

  test("GH-396: blank call result in a date position (coerceDate Empty arm, table mirror)") {
    // INDEX returns the raw cell value, so a blank hit routes CellValue.Empty through the
    // Coerced(Date) path — both tables (decodeAsDate, coerceDate) must agree on 1900-01-01
    val sheet = sheetWith(ref"A2" -> CellValue.Number(BigDecimal(99)))
    assertEquals(
      sheet.evaluateFormula("=YEAR(INDEX(A1:A2, 1))"),
      Right(CellValue.Number(BigDecimal(1900)))
    )
  }

  test("GH-396: Bool CELL in a date position matches the =YEAR(TRUE) literal (GH-307)") {
    // Pre-existing asymmetry: the literal coerced via ScalarCoercion (TRUE = serial 1 =
    // 1900-01-01) while the cell errored via decodeAsDate
    val sheet = sheetWith(ref"A1" -> CellValue.Bool(true))
    assertEquals(sheet.evaluateFormula("=YEAR(A1)"), Right(CellValue.Number(BigDecimal(1900))))
    assertEquals(sheet.evaluateFormula("=YEAR(TRUE)"), Right(CellValue.Number(BigDecimal(1900))))
  }

  test("GH-396: cached-formula Bool cell in a date position unwraps") {
    val sheet = sheetWith(ref"A1" -> CellValue.Formula("=1=1", Some(CellValue.Bool(true))))
    assertEquals(sheet.evaluateFormula("=YEAR(A1)"), Right(CellValue.Number(BigDecimal(1900))))
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
