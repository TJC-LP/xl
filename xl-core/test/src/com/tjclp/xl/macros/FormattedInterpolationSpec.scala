package com.tjclp.xl.macros

import com.tjclp.xl.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError
import com.tjclp.xl.formatted.Formatted
import com.tjclp.xl.styles.numfmt.NumFmt
import munit.{FunSuite, ScalaCheckSuite}
import org.scalacheck.Prop.*
import java.time.LocalDate

class FormattedInterpolationSpec extends ScalaCheckSuite:

  // ===== Backward Compatibility (Compile-Time Literals) =====

  test("Compile-time: money\"$$1,234.56\" returns Formatted directly") {
    val formatted: Formatted = money"$$1,234.56"
    assertEquals(formatted.numFmt, NumFmt.Currency)
    assertEquals(formatted.value, CellValue.Number(1234.56))
  }

  test("Compile-time: percent\"45.5%\" returns Formatted directly") {
    val formatted: Formatted = percent"45.5%"
    assertEquals(formatted.numFmt, NumFmt.Percent)
  }

  test("Compile-time: date\"2025-11-10\" returns Formatted directly") {
    val formatted: Formatted = date"2025-11-10"
    assertEquals(formatted.numFmt, NumFmt.Date)
  }

  test("Compile-time: accounting\"$$123.45\" returns Formatted directly") {
    val formatted: Formatted = accounting"$$123.45"
    assertEquals(formatted.numFmt, NumFmt.Currency)
  }

  // ===== Money Runtime Interpolation (7 tests) =====

  test("Runtime money: variable with dollar and commas") {
    val priceStr = "$1,234.56"
    money"$priceStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1234.56))
        assertEquals(f.numFmt, NumFmt.Currency)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime money: variable without dollar sign") {
    val priceStr = "999.99"
    money"$priceStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(999.99))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime money: mixed prefix and variable") {
    val amount = "1234.56"
    money"$$$amount" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1234.56))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime money: invalid variable returns Left") {
    val invalidStr = "$ABC"
    money"$invalidStr" match
      case Left(XLError.MoneyFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("Runtime money: empty string returns Left") {
    val emptyStr = ""
    money"$emptyStr" match
      case Left(_) => () // Expected
      case Right(_) => fail("Empty string should fail")
  }

  test("Runtime money: for-comprehension") {
    val str1 = "$100.00"
    val str2 = "$200.00"

    val result = for
      f1 <- money"$str1"
      f2 <- money"$str2"
    yield (f1.value, f2.value)

    result match
      case Right((v1, v2)) =>
        assertEquals(v1, CellValue.Number(100.00))
        assertEquals(v2, CellValue.Number(200.00))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime money: large amount") {
    val largeStr = "$1,000,000.00"
    money"$largeStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1000000.00))
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Percent Runtime Interpolation (7 tests) =====

  test("Runtime percent: variable with percent sign") {
    val pctStr = "45.5%"
    percent"$pctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.455))
        assertEquals(f.numFmt, NumFmt.Percent)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime percent: variable without percent sign") {
    val pctStr = "50"
    percent"$pctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.50))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime percent: mixed variable and suffix") {
    val value = "75"
    percent"$value%" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.75))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime percent: invalid variable returns Left") {
    val invalidStr = "ABC%"
    percent"$invalidStr" match
      case Left(XLError.PercentFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("Runtime percent: empty string returns Left") {
    val emptyStr = ""
    percent"$emptyStr" match
      case Left(_) => () // Expected
      case Right(_) => fail("Empty string should fail")
  }

  test("Runtime percent: > 100%") {
    val largeStr = "150%"
    percent"$largeStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(1.50))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime percent: 0%") {
    val zeroStr = "0%"
    percent"$zeroStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.0))
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Date Runtime Interpolation (7 tests) =====

  test("Runtime date: ISO format variable") {
    val dateStr = "2025-11-10"
    date"$dateStr" match
      case Right(f) =>
        f.value match
          case CellValue.DateTime(dt) =>
            assertEquals(dt.toLocalDate.toString, "2025-11-10")
          case _ => fail("Expected DateTime")
        assertEquals(f.numFmt, NumFmt.Date)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime date: Y2K date") {
    val dateStr = "2000-01-01"
    date"$dateStr" match
      case Right(f) =>
        f.value match
          case CellValue.DateTime(dt) =>
            assertEquals(dt.toLocalDate, LocalDate.of(2000, 1, 1))
          case _ => fail("Expected DateTime")
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime date: end of year") {
    val dateStr = "2025-12-31"
    date"$dateStr" match
      case Right(_) => () // Should parse
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime date: invalid format returns Left") {
    val invalidStr = "11/10/2025"  // US format not supported
    date"$invalidStr" match
      case Left(XLError.DateFormatError(value, reason)) =>
        assertEquals(value, "11/10/2025")
        assert(reason.contains("ISO"))
      case other => fail(s"Should fail, got $other")
  }

  test("Runtime date: invalid date returns Left") {
    val invalidStr = "2025-02-30"  // Feb 30 doesn't exist
    date"$invalidStr" match
      case Left(XLError.DateFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("Runtime date: non-date string returns Left") {
    val invalidStr = "not-a-date"
    date"$invalidStr" match
      case Left(XLError.DateFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("Runtime date: empty string returns Left") {
    val emptyStr = ""
    date"$emptyStr" match
      case Left(_) => () // Expected
      case Right(_) => fail("Empty string should fail")
  }

  // ===== Accounting Runtime Interpolation (7 tests) =====

  test("Runtime accounting: positive amount") {
    val acctStr = "$123.45"
    accounting"$acctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(123.45))
        assertEquals(f.numFmt, NumFmt.Currency)
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime accounting: negative with parens") {
    val acctStr = "($123.45)"
    accounting"$acctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(-123.45))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime accounting: negative with commas") {
    val acctStr = "($1,234.56)"
    accounting"$acctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(-1234.56))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime accounting: zero amount") {
    val acctStr = "$0.00"
    accounting"$acctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(0.00))
      case Left(err) => fail(s"Should parse: $err")
  }

  test("Runtime accounting: invalid variable returns Left") {
    val invalidStr = "$ABC"
    accounting"$invalidStr" match
      case Left(XLError.AccountingFormatError(_, _)) => () // Expected
      case other => fail(s"Should fail, got $other")
  }

  test("Runtime accounting: empty string returns Left") {
    val emptyStr = ""
    accounting"$emptyStr" match
      case Left(_) => () // Expected
      case Right(_) => fail("Empty string should fail")
  }

  test("Runtime accounting: large negative") {
    val acctStr = "($1,000,000.00)"
    accounting"$acctStr" match
      case Right(f) =>
        assertEquals(f.value, CellValue.Number(-1000000.00))
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Property-Based Tests =====

  property("Money: valid strings always parse") {
    forAll(Generators.genMoneyString) { str =>
      money"$str" match
        case Right(_) => () // Expected
        case Left(err) => fail(s"Valid money '$str' should parse: $err")
    }
  }

  property("Percent: valid strings always parse") {
    forAll(Generators.genPercentString) { str =>
      percent"$str" match
        case Right(_) => () // Expected
        case Left(err) => fail(s"Valid percent '$str' should parse: $err")
    }
  }

  property("Date: valid strings always parse") {
    forAll(Generators.genDateString) { str =>
      date"$str" match
        case Right(_) => () // Expected
        case Left(err) => fail(s"Valid date '$str' should parse: $err")
    }
  }

  property("Money: invalid strings always return Left") {
    forAll(Generators.genInvalidMoney) { str =>
      money"$str" match
        case Left(_) => () // Expected
        case Right(_) => fail(s"Invalid money '$str' should fail")
    }
  }
