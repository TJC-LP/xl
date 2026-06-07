package com.tjclp.xl

import com.tjclp.xl.cells.CellValue
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

  // GH-239: Excel 1900 leap-year parity. (Excel has no serials below 1; dates before 1900-01-01
  // extend the number line and are outside Excel's representable range.)
  test("dateTimeToExcelSerial: 1900-01-01 is serial 1 (Excel 1900 leap-year parity)") {
    val dt = LocalDateTime.of(1900, 1, 1, 0, 0, 0)
    assertEquals(CellValue.dateTimeToExcelSerial(dt), 1.0, 0.001)
  }

  test("dateTimeToExcelSerial: 1900-02-28 is serial 59 (day before the phantom leap day)") {
    val dt = LocalDateTime.of(1900, 2, 28, 0, 0, 0)
    assertEquals(CellValue.dateTimeToExcelSerial(dt), 59.0, 0.001)
  }

  test("dateTimeToExcelSerial for year 2020") {
    // January 1, 2020 at midnight should be serial 43831
    val dt = LocalDateTime.of(2020, 1, 1, 0, 0, 0)
    val serial = CellValue.dateTimeToExcelSerial(dt)
    assertEquals(serial, 43831.0, 0.001)
  }

  test("dateTimeToExcelSerial: 1900-03-01 is serial 61 (after the phantom leap day)") {
    val dt = LocalDateTime.of(1900, 3, 1, 0, 0, 0)
    assertEquals(CellValue.dateTimeToExcelSerial(dt), 61.0, 0.001)
  }

  test("excelSerialToDateTime: serial 60 (Excel's phantom 1900-02-29) maps to 1900-02-28") {
    assertEquals(
      CellValue.excelSerialToDateTime(60.0).toLocalDate,
      java.time.LocalDate.of(1900, 2, 28)
    )
  }

  test("date serial round-trips across the 1900 leap-year boundary and for modern dates") {
    val dates = List(
      LocalDateTime.of(1900, 1, 1, 0, 0, 0),
      LocalDateTime.of(1900, 1, 15, 0, 0, 0),
      LocalDateTime.of(1900, 2, 28, 0, 0, 0),
      LocalDateTime.of(1900, 3, 1, 0, 0, 0),
      LocalDateTime.of(2000, 1, 1, 0, 0, 0),
      LocalDateTime.of(2025, 11, 10, 0, 0, 0)
    )
    dates.foreach { dt =>
      val serial = CellValue.dateTimeToExcelSerial(dt)
      assertEquals(
        CellValue.excelSerialToDateTime(serial),
        dt,
        s"round-trip failed for $dt (serial $serial)"
      )
    }
  }
