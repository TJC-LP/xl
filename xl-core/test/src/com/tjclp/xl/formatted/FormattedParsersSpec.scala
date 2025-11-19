package com.tjclp.xl.formatted

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.errors.XLError
import com.tjclp.xl.styles.numfmt.NumFmt
import munit.FunSuite
import java.time.LocalDate

class FormattedParsersSpec extends FunSuite:

  // ===== Money Parser Tests (8 tests) =====

  test("parseMoney: $1,234.56 with comma separator") {
    FormattedParsers.parseMoney("$1,234.56") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1234.56))
        assertEquals(f.numFmt, NumFmt.Currency)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseMoney: 1234.56 without dollar sign") {
    FormattedParsers.parseMoney("1234.56") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1234.56))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseMoney: $1234.56 without commas") {
    FormattedParsers.parseMoney("$1234.56") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1234.56))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseMoney: $ 1,234.56 with spaces") {
    FormattedParsers.parseMoney("$ 1,234.56") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1234.56))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseMoney: rejects non-numeric $ABC") {
    FormattedParsers.parseMoney("$ABC") match
      case Left(XLError.MoneyFormatError(value, _)) =>
        assertEquals(value, "$ABC")
      case other => fail(s"Should fail, got $other")
  }

  test("parseMoney: rejects empty string") {
    FormattedParsers.parseMoney("") match
      case Left(XLError.MoneyFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseMoney: rejects multiple decimals $1.2.3") {
    FormattedParsers.parseMoney("$1.2.3") match
      case Left(XLError.MoneyFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseMoney: large number $1,000,000.00") {
    FormattedParsers.parseMoney("$1,000,000.00") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1000000.00))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseMoney: preserves high-precision BigDecimal") {
    FormattedParsers.parseMoney("$1234567890.123456789") match
      case Right(f) =>
        f.value match
          case CellValue.Number(bd) =>
            assertEquals(bd, BigDecimal("1234567890.123456789"))
          case _ => fail("Expected Number")
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseMoney: rejects input exceeding Excel cell limit") {
    val tooLong = "$" + ("1" * 33000) // Exceeds 32,767
    FormattedParsers.parseMoney(tooLong) match
      case Left(XLError.MoneyFormatError(input, msg)) =>
        assert(msg.contains("Excel cell limit"))
        assert(input.length < 100) // Truncated
      case other => fail(s"Should fail for too-long input, got $other")
  }

  // ===== Percent Parser Tests (8 tests) =====

  test("parsePercent: 45.5% with percent sign") {
    FormattedParsers.parsePercent("45.5%") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.455))
        assertEquals(f.numFmt, NumFmt.Percent)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parsePercent: 50 without percent sign") {
    FormattedParsers.parsePercent("50") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.50))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parsePercent: 100% full percent") {
    FormattedParsers.parsePercent("100%") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1.0))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parsePercent: 0% zero percent") {
    FormattedParsers.parsePercent("0%") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.0))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parsePercent: rejects non-numeric ABC%") {
    FormattedParsers.parsePercent("ABC%") match
      case Left(XLError.PercentFormatError(value, _)) =>
        assertEquals(value, "ABC%")
      case other => fail(s"Should fail, got $other")
  }

  test("parsePercent: rejects multiple percent signs 50%%") {
    FormattedParsers.parsePercent("50%%") match
      case Left(XLError.PercentFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parsePercent: rejects empty string") {
    FormattedParsers.parsePercent("") match
      case Left(XLError.PercentFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parsePercent: handles > 100%") {
    FormattedParsers.parsePercent("150%") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1.50))
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Date Parser Tests (8 tests) =====

  test("parseDate: 2025-11-10 ISO format") {
    FormattedParsers.parseDate("2025-11-10") match
      case Right(f) =>
        f.value match
          case CellValue.DateTime(dt) =>
            assertEquals(dt.toLocalDate.toString, "2025-11-10")
          case _ => fail("Expected DateTime")
        assertEquals(f.numFmt, NumFmt.Date)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseDate: 2000-01-01 Y2K date") {
    FormattedParsers.parseDate("2000-01-01") match
      case Right(f) =>
        f.value match
          case CellValue.DateTime(dt) =>
            assertEquals(dt.toLocalDate, LocalDate.of(2000, 1, 1))
          case _ => fail("Expected DateTime")
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseDate: 2025-12-31 end of year") {
    FormattedParsers.parseDate("2025-12-31") match
      case Right(_) => () // Should parse
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseDate: rejects US format 11/10/2025") {
    FormattedParsers.parseDate("11/10/2025") match
      case Left(XLError.DateFormatError(value, reason)) =>
        assertEquals(value, "11/10/2025")
        assert(reason.contains("ISO format"))
      case other => fail(s"Should fail, got $other")
  }

  test("parseDate: rejects invalid date 2025-02-30") {
    FormattedParsers.parseDate("2025-02-30") match
      case Left(XLError.DateFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseDate: rejects invalid month 2025-13-01") {
    FormattedParsers.parseDate("2025-13-01") match
      case Left(XLError.DateFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseDate: rejects non-date string") {
    FormattedParsers.parseDate("not-a-date") match
      case Left(XLError.DateFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseDate: rejects empty string") {
    FormattedParsers.parseDate("") match
      case Left(XLError.DateFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  // ===== Accounting Parser Tests (8 tests) =====

  test("parseAccounting: $123.45 positive amount") {
    FormattedParsers.parseAccounting("$123.45") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(123.45))
        assertEquals(f.numFmt, NumFmt.Currency)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseAccounting: ($123.45) negative with parens") {
    FormattedParsers.parseAccounting("($123.45)") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(-123.45))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseAccounting: ($1,234.56) negative with commas") {
    FormattedParsers.parseAccounting("($1,234.56)") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(-1234.56))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseAccounting: $0.00 zero amount") {
    FormattedParsers.parseAccounting("$0.00") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.00))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("parseAccounting: rejects non-numeric $ABC") {
    FormattedParsers.parseAccounting("$ABC") match
      case Left(XLError.AccountingFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseAccounting: rejects empty parens $()") {
    FormattedParsers.parseAccounting("$()") match
      case Left(XLError.AccountingFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseAccounting: rejects empty string") {
    FormattedParsers.parseAccounting("") match
      case Left(XLError.AccountingFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("parseAccounting: large negative ($1,000,000.00)") {
    FormattedParsers.parseAccounting("($1,000,000.00)") match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(-1000000.00))
      case Left(err) => fail(s"Should parse: $err")
  }
