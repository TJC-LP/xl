package com.tjclp.xl.codec

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.codec.syntax.*
import com.tjclp.xl.richtext.RichText.{*, given}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.error.XLException
import com.tjclp.xl.macros.ref
// Removed: BatchPutMacro is dead code (shadowed by Sheet.put member)
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.unsafe.*

import java.time.{LocalDate, LocalDateTime}

/** Tests for batch update extension methods */
class BatchUpdateSpec extends FunSuite:

  test("put: single string value") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Hello")
      .unsafe

    assertEquals(sheet(ref"A1").value, CellValue.Text("Hello"))
    assertEquals(sheet.getCellStyle(ref"A1"), None) // Strings don't get styling
  }

  test("put: single int value with auto-style") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> 42)
      .unsafe

    assertEquals(sheet(ref"A1").value, CellValue.Number(BigDecimal(42)))
    val style = sheet.getCellStyle(ref"A1")
    assert(style.exists(_.numFmt == NumFmt.General), "Should have General format")
  }

  test("put: single BigDecimal with decimal format") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> BigDecimal("123.45"))
      .unsafe

    assertEquals(sheet(ref"A1").value, CellValue.Number(BigDecimal("123.45")))
    val style = sheet.getCellStyle(ref"A1")
    assert(style.exists(_.numFmt == NumFmt.Decimal), "Should have Decimal format")
  }

  test("put: single LocalDate with date format") {
    val date = LocalDate.of(2025, 11, 10)
    val sheet = Sheet("Test")
      .put(ref"A1" -> date)
      .unsafe

    sheet(ref"A1").value match
      case CellValue.DateTime(datetime) => assertEquals(datetime.toLocalDate, date)
      case other => fail(s"Expected DateTime, got $other")

    val style = sheet.getCellStyle(ref"A1")
    assert(style.exists(_.numFmt == NumFmt.Date), "Should have Date format")
  }

  test("put: single LocalDateTime with datetime format") {
    val dt = LocalDateTime.of(2025, 11, 10, 14, 30)
    val sheet = Sheet("Test")
      .put(ref"A1" -> dt)
      .unsafe

    assertEquals(sheet(ref"A1").value, CellValue.DateTime(dt))
    val style = sheet.getCellStyle(ref"A1")
    assert(style.exists(_.numFmt == NumFmt.DateTime), "Should have DateTime format")
  }

  test("put: multiple values with mixed types") {
    val sheet = Sheet("Test")
      .put(
        ref"A1" -> "Title",
        ref"B1" -> 42,
        ref"C1" -> BigDecimal("123.45"),
        ref"D1" -> LocalDate.of(2025, 11, 10),
        ref"E1" -> true
      )
      .unsafe

    // Verify values
    assertEquals(sheet(ref"A1").value, CellValue.Text("Title"))
    assertEquals(sheet(ref"B1").value, CellValue.Number(BigDecimal(42)))
    assertEquals(sheet(ref"C1").value, CellValue.Number(BigDecimal("123.45")))
    assert(sheet(ref"D1").value match { case _: CellValue.DateTime => true; case _ => false }, "D1 should contain DateTime")
    assertEquals(sheet(ref"E1").value, CellValue.Bool(true))

    // Verify styles applied correctly
    assert(sheet.getCellStyle(ref"A1").isEmpty, "String should have no style")
    assert(sheet.getCellStyle(ref"B1").exists(_.numFmt == NumFmt.General), "Int should have General format")
    assert(sheet.getCellStyle(ref"C1").exists(_.numFmt == NumFmt.Decimal), "BigDecimal should have Decimal format")
    assert(sheet.getCellStyle(ref"D1").exists(_.numFmt == NumFmt.Date), "LocalDate should have Date format")
    assert(sheet.getCellStyle(ref"E1").isEmpty, "Boolean should have no style")
  }

  test("put: batch semantics equivalent to individual puts") {
    val sheet = Sheet("Test")

    // Manual individual puts
    val manual = sheet
      .put(ref"A1", CellValue.Text("Hello"))
      .put(ref"B1", CellValue.Number(BigDecimal(42)))

    // Using batch put
    val batched = sheet.put(
      ref"A1" -> "Hello",
      ref"B1" -> 42
    )
      .unsafe

    // Values should be the same (styles might differ due to auto-inference)
    assertEquals(batched(ref"A1").value, manual(ref"A1").value)
    assertEquals(batched(ref"B1").value, manual(ref"B1").value)
  }

  test("put: overwrites existing cells") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Old")
      .unsafe
      .put(ref"A1" -> "New")
      .unsafe

    assertEquals(sheet(ref"A1").value, CellValue.Text("New"))
  }

  // ========== Type Safety Tests ==========
  // Note: Unsupported types now cause COMPILE-TIME errors, not runtime errors.
  // The following tests document the type-safe behavior introduced by CellWriter[CellWritable]:
  //
  // This code would NOT compile:
  //   case class UnsupportedType(value: String)
  //   sheet.put(ref"A1" -> UnsupportedType("Invalid"))
  //   // Error: No given instance of type CellWriter[UnsupportedType]
  //
  // To support custom types, users define a CellWriter instance:
  //   given CellWriter[MyType] with
  //     def write(value: MyType) = (CellValue.Text(value.toString), None)

  test("put: type-safe API accepts CellWritable types") {
    // All CellWritable types compile and work correctly
    val sheet = Sheet("Test")
      .put(
        ref"A1" -> "String",
        ref"B1" -> 42,
        ref"C1" -> 3.14,
        ref"D1" -> true,
        ref"E1" -> BigDecimal("99.99")
      )
      .unsafe

    assertEquals(sheet(ref"A1").value, CellValue.Text("String"))
    assertEquals(sheet(ref"B1").value, CellValue.Number(BigDecimal(42)))
    assertEquals(sheet(ref"C1").value, CellValue.Number(BigDecimal(3.14)))
    assertEquals(sheet(ref"D1").value, CellValue.Bool(true))
    assertEquals(sheet(ref"E1").value, CellValue.Number(BigDecimal("99.99")))
  }

  test("put: writes single typed value with auto-inferred style") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> BigDecimal("123.45"))
      .unsafe

    assertEquals(sheet(ref"A1").value, CellValue.Number(BigDecimal("123.45")))
    assert(sheet.getCellStyle(ref"A1").exists(_.numFmt == NumFmt.Decimal), "Should have Decimal format")
  }

  test("readTyped: reads existing ref") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Number(BigDecimal(42)))

    val result = sheet.readTyped[Int](ref"A1")
    assertEquals(result, Right(Some(42)))
  }

  test("readTyped: returns None for empty ref") {
    val sheet = Sheet("Test")

    val result = sheet.readTyped[Int](ref"A1")
    assertEquals(result, Right(None))
  }

  test("readTyped: returns error for type mismatch") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("not a number"))

    val result = sheet.readTyped[Int](ref"A1")
    assert(result.isLeft, "Should return error for type mismatch")
  }

  test("putMixed + readTyped: round-trip") {
    val sheet = Sheet("Test")
      .put(
        ref"A1" -> 42,
        ref"B1" -> BigDecimal("123.45"),
        ref"C1" -> "Text"
      )
      .unsafe

    assertEquals(sheet.readTyped[Int](ref"A1"), Right(Some(42)))
    assertEquals(sheet.readTyped[BigDecimal](ref"B1"), Right(Some(BigDecimal("123.45"))))
    assertEquals(sheet.readTyped[String](ref"C1"), Right(Some("Text")))
  }

  test("put: StyleRegistry size increases correctly") {
    val sheet = Sheet("Test")
      .put(
        ref"A1" -> "No style", // Strings don't add styles
        ref"B1" -> BigDecimal("123.45"), // Adds decimal style
        ref"C1" -> LocalDate.of(2025, 11, 10) // Adds date style
      )
      .unsafe

    // Should have: default + decimal + date = 3 styles
    assert(sheet.styleRegistry.size >= 3, s"Expected at least 3 styles, got ${sheet.styleRegistry.size}")
  }

  test("put: deduplicates identical styles") {
    val sheet = Sheet("Test")
      .put(
        ref"A1" -> BigDecimal("123.45"),
        ref"A2" -> BigDecimal("678.90") // Same type, should reuse style
      )
      .unsafe

    // Both cells should have the same styleId
    assertEquals(
      sheet(ref"A1").styleId,
      sheet(ref"A2").styleId,
      "Cells with same type should have same style"
    )
  }

  // ========== RichText Codec Integration ==========

  test("put: RichText value") {
    val richText = "Bold".bold + " normal"
    val sheet = Sheet("Test")
      .put(ref"A1" -> richText)
      .unsafe

    sheet(ref"A1").value match
      case CellValue.RichText(rt) =>
        assertEquals(rt.toPlainText, "Bold normal")
        assert(rt.runs(0).font.exists(_.bold), "First run should be bold")
      case other => fail(s"Expected RichText, got $other")
  }

  test("put: mix RichText with other types") {
    val richText = "Error: ".red.bold + "File not found"
    val sheet = Sheet("Test")
      .put(
        ref"A1" -> "Plain text",
        ref"A2" -> richText,
        ref"A3" -> 42
      )
      .unsafe

    assertEquals(sheet(ref"A1").value, CellValue.Text("Plain text"))
    assert(sheet(ref"A2").value match { case _: CellValue.RichText => true; case _ => false }, "A2 should be RichText")
    assertEquals(sheet(ref"A3").value, CellValue.Number(BigDecimal(42)))
  }

  test("readTyped: read RichText from ref") {
    val richText = "Bold".bold
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.RichText(richText))

    val result = sheet.readTyped[RichText](ref"A1")
    result match
      case Right(Some(rt)) => assertEquals(rt.toPlainText, "Bold")
      case other => fail(s"Expected Right(Some(RichText)), got $other")
  }

  test("readTyped: convert plain Text to RichText") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Hello"))

    val result = sheet.readTyped[RichText](ref"A1")
    result match
      case Right(Some(rt)) =>
        assertEquals(rt.toPlainText, "Hello")
        assert(rt.isPlainText, "Converted from plain text should be plain")
      case other => fail(s"Expected Right(Some(RichText)), got $other")
  }

  test("RichText: identity law") {
    val original = "Bold".bold.red + " normal"
    val (cellValue, _) = CellCodec[RichText].write(original)
    val cell = Cell(ref"A1", cellValue)

    CellCodec[RichText].read(cell) match
      case Right(Some(rt)) =>
        assertEquals(rt.toPlainText, original.toPlainText)
        assertEquals(rt.runs.size, original.runs.size)
      case other => fail(s"Expected Right(Some(RichText)), got $other")
  }
