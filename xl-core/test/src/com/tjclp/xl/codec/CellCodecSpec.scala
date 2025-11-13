package com.tjclp.xl.codec

import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import com.tjclp.xl.api.*
import com.tjclp.xl.cell.{Cell, CellValue}
import com.tjclp.xl.codec.CodecError
import com.tjclp.xl.macros.ref
import com.tjclp.xl.style.numfmt.NumFmt

import java.time.{LocalDate, LocalDateTime}

/** Property-based tests for CellCodec instances */
class CellCodecSpec extends ScalaCheckSuite:

  // ========== String Codec ==========

  test("String codec: read text ref") {
    val cell = Cell(ref"A1", CellValue.Text("Hello"))
    assertEquals(CellCodec[String].read(cell), Right(Some("Hello")))
  }

  test("String codec: empty cell returns None") {
    val cell = Cell(ref"A1", CellValue.Empty)
    assertEquals(CellCodec[String].read(cell), Right(None))
  }

  test("String codec: convert number to string") {
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal("123.45")))
    assertEquals(CellCodec[String].read(cell), Right(Some("123.45")))
  }

  test("String codec: convert boolean to string") {
    val cell = Cell(ref"A1", CellValue.Bool(true))
    assertEquals(CellCodec[String].read(cell), Right(Some("true")))
  }

  test("String codec: write produces text value with no style") {
    val (value, style) = CellCodec[String].write("Test")
    assertEquals(value, CellValue.Text("Test"))
    assertEquals(style, None)
  }

  // ========== Int Codec ==========

  test("Int codec: read number ref") {
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal(42)))
    assertEquals(CellCodec[Int].read(cell), Right(Some(42)))
  }

  test("Int codec: empty cell returns None") {
    val cell = Cell(ref"A1", CellValue.Empty)
    assertEquals(CellCodec[Int].read(cell), Right(None))
  }

  test("Int codec: decimal number fails") {
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal("123.45")))
    assert(CellCodec[Int].read(cell).isLeft)
  }

  test("Int codec: text cell fails") {
    val cell = Cell(ref"A1", CellValue.Text("not a number"))
    assert(CellCodec[Int].read(cell).isLeft)
  }

  test("Int codec: write produces number with general format") {
    val (value, styleOpt) = CellCodec[Int].write(42)
    assertEquals(value, CellValue.Number(BigDecimal(42)))
    assert(styleOpt.exists(_.numFmt == NumFmt.General))
  }

  // ========== Long Codec ==========

  test("Long codec: read large number") {
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal(9876543210L)))
    assertEquals(CellCodec[Long].read(cell), Right(Some(9876543210L)))
  }

  test("Long codec: decimal number fails") {
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal("9876543210.5")))
    assert(CellCodec[Long].read(cell).isLeft)
  }

  // ========== Double Codec ==========

  test("Double codec: read number ref") {
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal("123.45")))
    assertEquals(CellCodec[Double].read(cell), Right(Some(123.45)))
  }

  test("Double codec: write produces number") {
    val (value, styleOpt) = CellCodec[Double].write(123.45)
    assertEquals(value, CellValue.Number(BigDecimal(123.45)))
  }

  // ========== BigDecimal Codec ==========

  test("BigDecimal codec: read number ref") {
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal("123.456789")))
    assertEquals(CellCodec[BigDecimal].read(cell), Right(Some(BigDecimal("123.456789"))))
  }

  test("BigDecimal codec: write produces number with decimal format") {
    val (value, styleOpt) = CellCodec[BigDecimal].write(BigDecimal("123.45"))
    assertEquals(value, CellValue.Number(BigDecimal("123.45")))
    assert(styleOpt.exists(_.numFmt == NumFmt.Decimal))
  }

  test("BigDecimal codec: auto-infers decimal format") {
    val (_, styleOpt) = CellCodec[BigDecimal].write(BigDecimal("123.45"))
    assert(styleOpt.isDefined, "Style should be defined for BigDecimal")
    assert(styleOpt.get.numFmt == NumFmt.Decimal, "Should have Decimal format")
  }

  // ========== Boolean Codec ==========

  test("Boolean codec: read boolean ref") {
    val cell = Cell(ref"A1", CellValue.Bool(true))
    assertEquals(CellCodec[Boolean].read(cell), Right(Some(true)))
  }

  test("Boolean codec: read 0 as false") {
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal(0)))
    assertEquals(CellCodec[Boolean].read(cell), Right(Some(false)))
  }

  test("Boolean codec: read 1 as true") {
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal(1)))
    assertEquals(CellCodec[Boolean].read(cell), Right(Some(true)))
  }

  test("Boolean codec: other numbers fail") {
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal(42)))
    assert(CellCodec[Boolean].read(cell).isLeft)
  }

  test("Boolean codec: write produces boolean value") {
    val (value, style) = CellCodec[Boolean].write(true)
    assertEquals(value, CellValue.Bool(true))
    assertEquals(style, None)
  }

  // ========== LocalDate Codec ==========

  test("LocalDate codec: read datetime ref") {
    val dt = LocalDateTime.of(2025, 11, 10, 14, 30)
    val cell = Cell(ref"A1", CellValue.DateTime(dt))
    assertEquals(CellCodec[LocalDate].read(cell), Right(Some(LocalDate.of(2025, 11, 10))))
  }

  test("LocalDate codec: read Excel serial number") {
    // November 10, 2025 = 45971 days since Dec 30, 1899
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal(45971)))
    val result = CellCodec[LocalDate].read(cell)
    assert(result.isRight)
    result.foreach(opt => assert(opt.exists(_.getYear == 2025)))
  }

  test("LocalDate codec: write produces datetime with date format") {
    val date = LocalDate.of(2025, 11, 10)
    val (value, styleOpt) = CellCodec[LocalDate].write(date)
    value match
      case CellValue.DateTime(dt) =>
        assertEquals(dt.toLocalDate, date)
        assert(styleOpt.exists(_.numFmt == NumFmt.Date))
      case other => fail(s"Expected DateTime, got $other")
  }

  test("LocalDate codec: auto-infers date format") {
    val (_, styleOpt) = CellCodec[LocalDate].write(LocalDate.of(2025, 11, 10))
    assert(styleOpt.isDefined, "Style should be defined for LocalDate")
    assert(styleOpt.get.numFmt == NumFmt.Date, "Should have Date format")
  }

  // ========== LocalDateTime Codec ==========

  test("LocalDateTime codec: read datetime ref") {
    val dt = LocalDateTime.of(2025, 11, 10, 14, 30, 45)
    val cell = Cell(ref"A1", CellValue.DateTime(dt))
    assertEquals(CellCodec[LocalDateTime].read(cell), Right(Some(dt)))
  }

  test("LocalDateTime codec: read Excel serial number") {
    // November 10, 2025 14:30:45 ≈ 45971.604 days
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal("45971.604166667")))
    val result = CellCodec[LocalDateTime].read(cell)
    assert(result.isRight)
    result.foreach { opt =>
      assert(opt.exists { dt =>
        dt.getYear == 2025 && dt.getMonthValue == 11 && dt.getDayOfMonth == 10
      })
    }
  }

  test("LocalDateTime codec: write produces datetime with datetime format") {
    val dt = LocalDateTime.of(2025, 11, 10, 14, 30)
    val (value, styleOpt) = CellCodec[LocalDateTime].write(dt)
    assertEquals(value, CellValue.DateTime(dt))
    assert(styleOpt.exists(_.numFmt == NumFmt.DateTime))
  }

  test("LocalDateTime codec: auto-infers datetime format") {
    val (_, styleOpt) = CellCodec[LocalDateTime].write(LocalDateTime.of(2025, 11, 10, 14, 30))
    assert(styleOpt.isDefined, "Style should be defined for LocalDateTime")
    assert(styleOpt.get.numFmt == NumFmt.DateTime, "Should have DateTime format")
  }

  // ========== Identity Laws ==========

  test("Identity law: String round-trip") {
    val original = "Hello, World!"
    val (cellValue, _) = CellCodec[String].write(original)
    val cell = Cell(ref"A1", cellValue)
    assertEquals(CellCodec[String].read(cell), Right(Some(original)))
  }

  test("Identity law: Int round-trip") {
    val original = 42
    val (cellValue, _) = CellCodec[Int].write(original)
    val cell = Cell(ref"A1", cellValue)
    assertEquals(CellCodec[Int].read(cell), Right(Some(original)))
  }

  test("Identity law: Long round-trip") {
    val original = 9876543210L
    val (cellValue, _) = CellCodec[Long].write(original)
    val cell = Cell(ref"A1", cellValue)
    assertEquals(CellCodec[Long].read(cell), Right(Some(original)))
  }

  test("Identity law: Double round-trip") {
    val original = 123.45
    val (cellValue, _) = CellCodec[Double].write(original)
    val cell = Cell(ref"A1", cellValue)
    val result = CellCodec[Double].read(cell)
    assert(result.isRight)
    result.foreach(opt => assert(opt.exists(v => math.abs(v - original) < 0.0001)))
  }

  test("Identity law: BigDecimal round-trip") {
    val original = BigDecimal("123.456789")
    val (cellValue, _) = CellCodec[BigDecimal].write(original)
    val cell = Cell(ref"A1", cellValue)
    assertEquals(CellCodec[BigDecimal].read(cell), Right(Some(original)))
  }

  test("Identity law: Boolean round-trip") {
    val original = true
    val (cellValue, _) = CellCodec[Boolean].write(original)
    val cell = Cell(ref"A1", cellValue)
    assertEquals(CellCodec[Boolean].read(cell), Right(Some(original)))
  }

  test("Identity law: LocalDate round-trip") {
    val original = LocalDate.of(2025, 11, 10)
    val (cellValue, _) = CellCodec[LocalDate].write(original)
    val cell = Cell(ref"A1", cellValue)
    assertEquals(CellCodec[LocalDate].read(cell), Right(Some(original)))
  }

  test("Identity law: LocalDateTime round-trip") {
    val original = LocalDateTime.of(2025, 11, 10, 14, 30, 45)
    val (cellValue, _) = CellCodec[LocalDateTime].write(original)
    val cell = Cell(ref"A1", cellValue)
    val result = CellCodec[LocalDateTime].read(cell)
    assert(result.isRight)
    // DateTime may lose subsecond precision in Excel serial format
    result.foreach(opt => assert(opt.exists { dt =>
      dt.getYear == original.getYear &&
      dt.getMonthValue == original.getMonthValue &&
      dt.getDayOfMonth == original.getDayOfMonth &&
      dt.getHour == original.getHour &&
      dt.getMinute == original.getMinute &&
      math.abs(dt.getSecond - original.getSecond) <= 1 // Allow 1 second tolerance
    }))
  }

  // ========== Error Cases ==========

  test("Type mismatch: Int from text") {
    val cell = Cell(ref"A1", CellValue.Text("not a number"))
    CellCodec[Int].read(cell) match
      case Left(CodecError.TypeMismatch(expected, actual)) =>
        assertEquals(expected, "Int")
      case other => fail(s"Expected TypeMismatch error, got $other")
  }

  test("Parse error: Int from decimal") {
    val cell = Cell(ref"A1", CellValue.Number(BigDecimal("123.45")))
    CellCodec[Int].read(cell) match
      case Left(CodecError.ParseError(value, targetType, detail)) =>
        assertEquals(targetType, "Int")
        assert(detail.contains("decimal"))
      case other => fail(s"Expected ParseError, got $other")
  }

  // ========== CodecError to XLError Conversion ==========

  test("CodecError.toXLError: TypeMismatch conversion") {
    val cellError = CodecError.TypeMismatch("Int", CellValue.Text("abc"))
    val xlError = cellError.toXLError(ref"A1")
    xlError match
      case XLError.TypeMismatch(exp, act, r) =>
        assertEquals(exp, "Int")
        assertEquals(r, "A1")
      case other => fail(s"Expected XLError.TypeMismatch, got $other")
  }

  test("CodecError.toXLError: ParseError conversion") {
    val cellError = CodecError.ParseError("123.45", "Int", "has decimal")
    val xlError = cellError.toXLError(ref"B2")
    xlError match
      case XLError.ParseError(loc, rsn) =>
        assertEquals(loc, "B2")
        assert(rsn.contains("123.45"))
        assert(rsn.contains("Int"))
      case other => fail(s"Expected XLError.ParseError, got $other")
  }
