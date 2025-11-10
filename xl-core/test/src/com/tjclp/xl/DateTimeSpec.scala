package com.tjclp.xl

import munit.FunSuite
import java.time.LocalDateTime

/** Tests for DateTime serialization to Excel serial numbers */
class DateTimeSpec extends FunSuite:

  test("dateTimeToExcelSerial for known dates") {
    // January 1, 2000 at midnight should be serial 36526
    val dt = LocalDateTime.of(2000, 1, 1, 0, 0, 0)
    val serial = CellValue.dateTimeToExcelSerial(dt)
    assertEquals(serial, 36526.0, 0.001)
  }

  test("dateTimeToExcelSerial for date with time") {
    // January 1, 2000 at noon (12:00:00) should be 36526.5
    val dt = LocalDateTime.of(2000, 1, 1, 12, 0, 0)
    val serial = CellValue.dateTimeToExcelSerial(dt)
    assertEquals(serial, 36526.5, 0.001)
  }

  test("dateTimeToExcelSerial for current date") {
    // November 10, 2025 at midnight should be serial 45971
    val dt = LocalDateTime.of(2025, 11, 10, 0, 0, 0)
    val serial = CellValue.dateTimeToExcelSerial(dt)
    assertEquals(serial, 45971.0, 0.001)
  }

  test("dateTimeToExcelSerial for date with fractional time") {
    // January 1, 2000 at 06:00:00 (quarter day) should be 36526.25
    val dt = LocalDateTime.of(2000, 1, 1, 6, 0, 0)
    val serial = CellValue.dateTimeToExcelSerial(dt)
    assertEquals(serial, 36526.25, 0.001)
  }

  test("dateTimeToExcelSerial for date with seconds") {
    // January 1, 2000 at 00:00:01 should be 36526 + 1/86400
    val dt = LocalDateTime.of(2000, 1, 1, 0, 0, 1)
    val serial = CellValue.dateTimeToExcelSerial(dt)
    val expected = 36526.0 + (1.0 / 86400.0)
    assertEquals(serial, expected, 0.00001)
  }

  test("dateTimeToExcelSerial epoch is 1899-12-30") {
    // Excel epoch (December 30, 1899) should be serial 0
    val epoch = LocalDateTime.of(1899, 12, 30, 0, 0, 0)
    val serial = CellValue.dateTimeToExcelSerial(epoch)
    assertEquals(serial, 0.0, 0.001)
  }

  test("dateTimeToExcelSerial for next day after epoch") {
    // December 31, 1899 should be serial 1
    val dt = LocalDateTime.of(1899, 12, 31, 0, 0, 0)
    val serial = CellValue.dateTimeToExcelSerial(dt)
    assertEquals(serial, 1.0, 0.001)
  }

  test("dateTimeToExcelSerial for year 2020") {
    // January 1, 2020 at midnight should be serial 43831
    val dt = LocalDateTime.of(2020, 1, 1, 0, 0, 0)
    val serial = CellValue.dateTimeToExcelSerial(dt)
    assertEquals(serial, 43831.0, 0.001)
  }
