package com.tjclp.xl.errors

import munit.FunSuite

class XLErrorSpec extends FunSuite:

  test("MoneyFormatError.message includes value and reason") {
    val err = XLError.MoneyFormatError("$ABC", "non-numeric characters")
    val msg = err.message
    assert(msg.contains("$ABC"))
    assert(msg.contains("non-numeric"))
    assert(msg.contains("money format"))
  }

  test("PercentFormatError.message includes value and reason") {
    val err = XLError.PercentFormatError("ABC%", "invalid number")
    val msg = err.message
    assert(msg.contains("ABC%"))
    assert(msg.contains("invalid"))
    assert(msg.contains("percent format"))
  }

  test("DateFormatError.message includes value and reason") {
    val err = XLError.DateFormatError("not-a-date", "unparseable")
    val msg = err.message
    assert(msg.contains("not-a-date"))
    assert(msg.contains("unparseable"))
    assert(msg.contains("date format"))
  }

  test("AccountingFormatError.message includes value and reason") {
    val err = XLError.AccountingFormatError("$ABC", "invalid format")
    val msg = err.message
    assert(msg.contains("$ABC"))
    assert(msg.contains("invalid"))
    assert(msg.contains("accounting format"))
  }

  test("MoneyFormatError pattern matches correctly") {
    val err: XLError = XLError.MoneyFormatError("$123", "test")
    err match
      case XLError.MoneyFormatError(value, reason) =>
        assertEquals(value, "$123")
        assertEquals(reason, "test")
      case _ => fail("Should match MoneyFormatError")
  }

  test("PercentFormatError pattern matches correctly") {
    val err: XLError = XLError.PercentFormatError("50%", "test")
    err match
      case XLError.PercentFormatError(value, reason) =>
        assertEquals(value, "50%")
        assertEquals(reason, "test")
      case _ => fail("Should match PercentFormatError")
  }

  test("DateFormatError pattern matches correctly") {
    val err: XLError = XLError.DateFormatError("2025-11-10", "test")
    err match
      case XLError.DateFormatError(value, reason) =>
        assertEquals(value, "2025-11-10")
        assertEquals(reason, "test")
      case _ => fail("Should match DateFormatError")
  }

  test("AccountingFormatError pattern matches correctly") {
    val err: XLError = XLError.AccountingFormatError("($123)", "test")
    err match
      case XLError.AccountingFormatError(value, reason) =>
        assertEquals(value, "($123)")
        assertEquals(reason, "test")
      case _ => fail("Should match AccountingFormatError")
  }
