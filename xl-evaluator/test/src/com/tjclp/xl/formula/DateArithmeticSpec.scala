package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite
import java.time.{LocalDate, LocalDateTime}

/**
 * Tests for date arithmetic functionality in formula evaluation.
 *
 * Excel stores dates as serial numbers (days since 1899-12-30) and allows arithmetic like:
 *   - TODAY() + 30 → date 30 days in future
 *   - NOW() + 0.5 → 12 hours later (0.5 days)
 *   - DATE(2025,1,15) - 7 → January 8th
 *
 * These tests verify that the TExpr.DateToSerial wrapper correctly converts date functions
 * to serial numbers for arithmetic operations.
 */
class DateArithmeticSpec extends FunSuite:

  val emptySheet: Sheet = Sheet("Test")

  // ============================================================================
  // TODAY() Arithmetic Tests
  // ============================================================================

  test("TODAY() + 30 returns date 30 days in future as serial") {
    val clock = Clock.fixedDate(LocalDate.of(2025, 1, 1))
    val result = emptySheet.evaluateFormula("=TODAY()+30", clock)
    result match
      case Right(CellValue.Number(serial)) =>
        val expectedSerial = CellValue.dateTimeToExcelSerial(
          LocalDate.of(2025, 1, 31).atStartOfDay()
        )
        assertEquals(serial.toDouble, expectedSerial, 0.0001)
      case other => fail(s"Expected Number, got $other")
  }

  test("TODAY() - 7 returns date 7 days ago") {
    val clock = Clock.fixedDate(LocalDate.of(2025, 1, 15))
    val result = emptySheet.evaluateFormula("=TODAY()-7", clock)
    result match
      case Right(CellValue.Number(serial)) =>
        val expectedSerial = CellValue.dateTimeToExcelSerial(
          LocalDate.of(2025, 1, 8).atStartOfDay()
        )
        assertEquals(serial.toDouble, expectedSerial, 0.0001)
      case other => fail(s"Expected Number, got $other")
  }

  test("TODAY() + cell reference for days offset") {
    val sheet = emptySheet.put(ref"A1", CellValue.Number(BigDecimal(30)))
    val clock = Clock.fixedDate(LocalDate.of(2025, 1, 1))
    val result = sheet.evaluateFormula("=TODAY()+A1", clock)
    result match
      case Right(CellValue.Number(serial)) =>
        val expectedSerial = CellValue.dateTimeToExcelSerial(
          LocalDate.of(2025, 1, 31).atStartOfDay()
        )
        assertEquals(serial.toDouble, expectedSerial, 0.0001)
      case other => fail(s"Expected Number, got $other")
  }

  test("TODAY() * 1 returns same serial (identity)") {
    val clock = Clock.fixedDate(LocalDate.of(2025, 1, 1))
    val result = emptySheet.evaluateFormula("=TODAY()*1", clock)
    result match
      case Right(CellValue.Number(serial)) =>
        val expectedSerial = CellValue.dateTimeToExcelSerial(
          LocalDate.of(2025, 1, 1).atStartOfDay()
        )
        assertEquals(serial.toDouble, expectedSerial, 0.0001)
      case other => fail(s"Expected Number, got $other")
  }

  // ============================================================================
  // NOW() Arithmetic Tests
  // ============================================================================

  test("NOW() + 0.5 adds 12 hours") {
    val fixedNow = LocalDateTime.of(2025, 1, 1, 0, 0)
    val clock = Clock.fixed(fixedNow.toLocalDate, fixedNow)
    val result = emptySheet.evaluateFormula("=NOW()+0.5", clock)
    result match
      case Right(CellValue.Number(serial)) =>
        val expectedSerial = CellValue.dateTimeToExcelSerial(
          LocalDateTime.of(2025, 1, 1, 12, 0)
        )
        assertEquals(serial.toDouble, expectedSerial, 0.0001)
      case other => fail(s"Expected Number, got $other")
  }

  test("NOW() + 1 adds 1 day") {
    val fixedNow = LocalDateTime.of(2025, 1, 1, 10, 30)
    val clock = Clock.fixed(fixedNow.toLocalDate, fixedNow)
    val result = emptySheet.evaluateFormula("=NOW()+1", clock)
    result match
      case Right(CellValue.Number(serial)) =>
        val expectedSerial = CellValue.dateTimeToExcelSerial(
          LocalDateTime.of(2025, 1, 2, 10, 30)
        )
        assertEquals(serial.toDouble, expectedSerial, 0.0001)
      case other => fail(s"Expected Number, got $other")
  }

  test("NOW() - 0.25 subtracts 6 hours") {
    val fixedNow = LocalDateTime.of(2025, 1, 1, 12, 0)
    val clock = Clock.fixed(fixedNow.toLocalDate, fixedNow)
    val result = emptySheet.evaluateFormula("=NOW()-0.25", clock)
    result match
      case Right(CellValue.Number(serial)) =>
        val expectedSerial = CellValue.dateTimeToExcelSerial(
          LocalDateTime.of(2025, 1, 1, 6, 0)
        )
        assertEquals(serial.toDouble, expectedSerial, 0.0001)
      case other => fail(s"Expected Number, got $other")
  }

  // ============================================================================
  // Complex Expressions
  // ============================================================================

  test("TODAY() + 30 - 7 chain arithmetic") {
    val clock = Clock.fixedDate(LocalDate.of(2025, 1, 1))
    val result = emptySheet.evaluateFormula("=TODAY()+30-7", clock)
    result match
      case Right(CellValue.Number(serial)) =>
        // Jan 1 + 30 = Jan 31, - 7 = Jan 24
        val expectedSerial = CellValue.dateTimeToExcelSerial(
          LocalDate.of(2025, 1, 24).atStartOfDay()
        )
        assertEquals(serial.toDouble, expectedSerial, 0.0001)
      case other => fail(s"Expected Number, got $other")
  }

  test("Date comparison: TODAY() > literal date") {
    val clock = Clock.fixedDate(LocalDate.of(2025, 6, 15))
    // Compare TODAY() serial (2025-06-15) with DATE(2025,1,1) serial
    val result = emptySheet.evaluateFormula("=TODAY()>DATE(2025,1,1)", clock)
    assertEquals(result, Right(CellValue.Bool(true)))
  }

  test("Date in past: TODAY() - 365 gives previous year") {
    val clock = Clock.fixedDate(LocalDate.of(2025, 6, 15))
    val result = emptySheet.evaluateFormula("=TODAY()-365", clock)
    result match
      case Right(CellValue.Number(serial)) =>
        val expectedSerial = CellValue.dateTimeToExcelSerial(
          LocalDate.of(2024, 6, 15).atStartOfDay()
        )
        assertEquals(serial.toDouble, expectedSerial, 0.0001)
      case other => fail(s"Expected Number, got $other")
  }
