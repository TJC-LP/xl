package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.addressing.SheetName
import munit.FunSuite
import java.time.{LocalDate, LocalDateTime}

/**
 * GH-449: DateTime CELLS are their Excel serial number in arithmetic positions.
 *
 * `=C28-C27` over two date cells used to fail with "Expected Numeric, got DateTime" — the
 * direct-cell numeric decoder refused DateTime while `ScalarCoercion.coerceNumeric` (literals, call
 * results, LET bindings) already converted dates to serials, and `normalizeForCompare` already did
 * so for comparisons. These tests pin the closed parity gap: every arithmetic position now converts
 * via [[CellValue.dateTimeToExcelSerial]], the same conversion the writer serializes with.
 *
 * Result-type note: arithmetic over dates yields a serial `CellValue.Number`, not a `DateTime` —
 * the established convention pinned by [[DateArithmeticSpec]] for `=TODAY()+30`, and Excel's own
 * model (a date IS a number; date-ness is a number format, not a value type).
 */
class DateTimeCellArithmeticSpec extends FunSuite:

  private val emptySheet = new Sheet(name = SheetName.unsafe("Test"))

  private def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) => s.put(ref, value) }

  private def serialOf(dt: LocalDateTime): BigDecimal =
    BigDecimal(CellValue.dateTimeToExcelSerial(dt))

  private def dateCell(y: Int, m: Int, d: Int): CellValue =
    CellValue.DateTime(LocalDate.of(y, m, d).atStartOfDay())

  private def assertNumber(result: XLResult[CellValue], expected: BigDecimal): Unit =
    result match
      case Right(CellValue.Number(n)) => assertEquals(n.toDouble, expected.toDouble, 1e-9)
      case other => fail(s"Expected Number($expected), got $other")

  // ============================================================================
  // The literal GH-449 repro: two date cells, a third subtracting them
  // ============================================================================

  test("GH-449 repro: =C28-C27 over two date cells evaluates to the day count") {
    val sheet = sheetWith(
      ref"C27" -> dateCell(2024, 1, 1),
      ref"C28" -> dateCell(2024, 3, 31),
      ref"C29" -> CellValue.Formula("=C28-C27")
    )
    assertEquals(sheet.evaluateCell(ref"C29"), Right(CellValue.Number(BigDecimal(90))))
  }

  test("GH-449 repro: the same subtraction as an ad-hoc formula") {
    val sheet = sheetWith(
      ref"C27" -> dateCell(2024, 1, 1),
      ref"C28" -> dateCell(2024, 3, 31)
    )
    assertNumber(sheet.evaluateFormula("=C28-C27"), BigDecimal(90))
  }

  test("date - date is negative when the operands are reversed") {
    val sheet = sheetWith(
      ref"A1" -> dateCell(2024, 1, 1),
      ref"A2" -> dateCell(2024, 3, 31)
    )
    assertNumber(sheet.evaluateFormula("=A1-A2"), BigDecimal(-90))
  }

  test("date - date carries the time component as a fractional day") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.DateTime(LocalDateTime.of(2024, 1, 1, 0, 0)),
      ref"A2" -> CellValue.DateTime(LocalDateTime.of(2024, 1, 2, 12, 0))
    )
    assertNumber(sheet.evaluateFormula("=A2-A1"), BigDecimal(1.5))
  }

  // ============================================================================
  // date offsets: date +/- number (Excel serial semantics)
  // ============================================================================

  test("date + number offsets the serial") {
    val sheet = sheetWith(ref"A1" -> dateCell(2024, 1, 1))
    assertNumber(
      sheet.evaluateFormula("=A1+30"),
      serialOf(LocalDate.of(2024, 1, 31).atStartOfDay())
    )
  }

  test("number + date offsets the serial (operands reversed)") {
    val sheet = sheetWith(ref"A1" -> dateCell(2024, 1, 1))
    assertNumber(
      sheet.evaluateFormula("=30+A1"),
      serialOf(LocalDate.of(2024, 1, 31).atStartOfDay())
    )
  }

  test("date - number offsets the serial backwards") {
    val sheet = sheetWith(ref"A1" -> dateCell(2024, 3, 31))
    assertNumber(
      sheet.evaluateFormula("=A1-30"),
      serialOf(LocalDate.of(2024, 3, 1).atStartOfDay())
    )
  }

  test("date + fractional number adds hours") {
    val sheet = sheetWith(ref"A1" -> CellValue.DateTime(LocalDateTime.of(2024, 1, 1, 0, 0)))
    assertNumber(sheet.evaluateFormula("=A1+0.5"), serialOf(LocalDateTime.of(2024, 1, 1, 12, 0)))
  }

  // ============================================================================
  // Other arithmetic operators coerce to the serial
  // ============================================================================

  test("date * number multiplies the serial") {
    val sheet = sheetWith(ref"A1" -> dateCell(2024, 1, 1))
    assertNumber(
      sheet.evaluateFormula("=A1*2"),
      serialOf(LocalDate.of(2024, 1, 1).atStartOfDay()) * 2
    )
  }

  test("date / number divides the serial") {
    val sheet = sheetWith(ref"A1" -> dateCell(2024, 1, 1))
    assertNumber(
      sheet.evaluateFormula("=A1/2"),
      serialOf(LocalDate.of(2024, 1, 1).atStartOfDay()) / 2
    )
  }

  test("unary minus on a date negates the serial") {
    val sheet = sheetWith(ref"A1" -> dateCell(2024, 1, 1))
    assertNumber(sheet.evaluateFormula("=-A1"), -serialOf(LocalDate.of(2024, 1, 1).atStartOfDay()))
  }

  test("percent on a date scales the serial") {
    val sheet = sheetWith(ref"A1" -> dateCell(2024, 1, 1))
    assertNumber(
      sheet.evaluateFormula("=A1%"),
      serialOf(LocalDate.of(2024, 1, 1).atStartOfDay()) / 100
    )
  }

  test("array broadcast: a date range minus a date scalar is a day-count array") {
    val sheet = sheetWith(
      ref"A1" -> dateCell(2024, 1, 1),
      ref"A2" -> dateCell(2024, 3, 31),
      ref"B1" -> dateCell(2024, 1, 1)
    )
    assertNumber(sheet.evaluateFormula("=SUM(A1:A2-B1)"), BigDecimal(90))
  }

  // ============================================================================
  // Aggregates over date ranges (dates are numbers; booleans stay skipped)
  // ============================================================================

  test("SUM over a range of date cells sums the serials") {
    val sheet = sheetWith(
      ref"A1" -> dateCell(2024, 1, 1),
      ref"A2" -> dateCell(2024, 1, 2)
    )
    assertNumber(
      sheet.evaluateFormula("=SUM(A1:A2)"),
      serialOf(LocalDate.of(2024, 1, 1).atStartOfDay()) +
        serialOf(LocalDate.of(2024, 1, 2).atStartOfDay())
    )
  }

  test("MIN/MAX over a date column are the earliest and latest serials") {
    val sheet = sheetWith(
      ref"A1" -> dateCell(2024, 3, 31),
      ref"A2" -> dateCell(2024, 1, 1),
      ref"A3" -> dateCell(2024, 2, 15)
    )
    assertNumber(
      sheet.evaluateFormula("=MIN(A1:A3)"),
      serialOf(LocalDate.of(2024, 1, 1).atStartOfDay())
    )
    assertNumber(
      sheet.evaluateFormula("=MAX(A1:A3)"),
      serialOf(LocalDate.of(2024, 3, 31).atStartOfDay())
    )
  }

  test("COUNT over a date column counts the dates") {
    val sheet = sheetWith(
      ref"A1" -> dateCell(2024, 1, 1),
      ref"A2" -> dateCell(2024, 3, 31),
      ref"A3" -> CellValue.Text("not a date")
    )
    assertNumber(sheet.evaluateFormula("=COUNT(A1:A3)"), BigDecimal(2))
  }

  test("MAX(dates) - MIN(dates) is the span in days") {
    val sheet = sheetWith(
      ref"A1" -> dateCell(2024, 3, 31),
      ref"A2" -> dateCell(2024, 1, 1),
      ref"A3" -> dateCell(2024, 2, 15)
    )
    assertNumber(sheet.evaluateFormula("=MAX(A1:A3)-MIN(A1:A3)"), BigDecimal(90))
  }

  // ============================================================================
  // Comparisons (pre-existing normalizeForCompare behavior — pinned here)
  // ============================================================================

  test("date < date compares by serial") {
    val sheet = sheetWith(
      ref"A1" -> dateCell(2024, 1, 1),
      ref"A2" -> dateCell(2024, 3, 31)
    )
    assertEquals(sheet.evaluateFormula("=A1<A2"), Right(CellValue.Bool(true)))
    assertEquals(sheet.evaluateFormula("=A2<A1"), Right(CellValue.Bool(false)))
  }

  test("date = date compares by serial") {
    val sheet = sheetWith(
      ref"A1" -> dateCell(2024, 1, 1),
      ref"A2" -> dateCell(2024, 1, 1),
      ref"A3" -> dateCell(2024, 3, 31)
    )
    assertEquals(sheet.evaluateFormula("=A1=A2"), Right(CellValue.Bool(true)))
    assertEquals(sheet.evaluateFormula("=A1=A3"), Right(CellValue.Bool(false)))
  }

  test("date compares against a raw serial number cell") {
    val sheet = sheetWith(
      ref"A1" -> dateCell(2024, 1, 1),
      ref"A2" -> CellValue.Number(serialOf(LocalDate.of(2024, 1, 1).atStartOfDay()))
    )
    assertEquals(sheet.evaluateFormula("=A1=A2"), Right(CellValue.Bool(true)))
    assertEquals(sheet.evaluateFormula("=A1>=A2"), Right(CellValue.Bool(true)))
  }

  test("date >= against a DATE() literal") {
    val sheet = sheetWith(ref"A1" -> dateCell(2024, 3, 31))
    assertEquals(sheet.evaluateFormula("=A1>DATE(2024,1,1)"), Right(CellValue.Bool(true)))
  }

  // ============================================================================
  // Finance-flavored: hold period and day-count fraction
  // ============================================================================

  test("GH-449: hold-period day count / 365 over acquisition and exit date cells") {
    val sheet = sheetWith(
      ref"B2" -> dateCell(2021, 6, 30), // acquisition
      ref"B3" -> dateCell(2026, 6, 30) // exit
    )
    assertNumber(sheet.evaluateFormula("=B3-B2"), BigDecimal(1826))
    assertNumber(sheet.evaluateFormula("=(B3-B2)/365"), BigDecimal(1826) / 365)
  }

  test("GH-449: accrual day count feeding a rate calculation") {
    val sheet = sheetWith(
      ref"B2" -> dateCell(2024, 1, 1),
      ref"B3" -> dateCell(2024, 12, 31),
      ref"B4" -> CellValue.Number(BigDecimal(1000000)),
      ref"B5" -> CellValue.Number(BigDecimal("0.085"))
    )
    // principal * rate * (days / 365)
    assertNumber(
      sheet.evaluateFormula("=B4*B5*(B3-B2)/365"),
      BigDecimal(1000000) * BigDecimal("0.085") * BigDecimal(365) / 365
    )
  }

  test("date arithmetic composes with date functions") {
    val sheet = sheetWith(
      ref"A1" -> dateCell(2024, 1, 1),
      ref"A2" -> dateCell(2024, 3, 31)
    )
    assertNumber(sheet.evaluateFormula("=YEARFRAC(A1,A2,3)"), BigDecimal(90) / 365)
    assertNumber(sheet.evaluateFormula("=(A2-A1)/365"), BigDecimal(90) / 365)
  }

  // ============================================================================
  // Constructors unchanged: date functions still return DateTime
  // ============================================================================

  test("DATE() still returns a DateTime, not a serial") {
    assertEquals(
      emptySheet.evaluateFormula("=DATE(2024,3,31)"),
      Right(CellValue.DateTime(LocalDate.of(2024, 3, 31).atStartOfDay()))
    )
  }

  test("EDATE() over a date cell still returns a DateTime") {
    val sheet = sheetWith(ref"A1" -> dateCell(2024, 1, 31))
    assertEquals(
      sheet.evaluateFormula("=EDATE(A1,1)"),
      Right(CellValue.DateTime(LocalDate.of(2024, 2, 29).atStartOfDay()))
    )
  }

  test("EOMONTH() over a date cell still returns a DateTime") {
    val sheet = sheetWith(ref"A1" -> dateCell(2024, 2, 15))
    assertEquals(
      sheet.evaluateFormula("=EOMONTH(A1,0)"),
      Right(CellValue.DateTime(LocalDate.of(2024, 2, 29).atStartOfDay()))
    )
  }

  test("DATEDIF over date cells is unchanged") {
    val sheet = sheetWith(
      ref"A1" -> dateCell(2024, 1, 1),
      ref"A2" -> dateCell(2024, 3, 31)
    )
    assertNumber(sheet.evaluateFormula("=DATEDIF(A1,A2,\"D\")"), BigDecimal(90))
  }
