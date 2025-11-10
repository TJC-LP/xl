package com.tjclp.xl.codec

import munit.FunSuite
import com.tjclp.xl.*
import com.tjclp.xl.codec.{*, given}
import com.tjclp.xl.macros.{cell, range}
import java.time.{LocalDate, LocalDateTime}

/** Tests for batch update extension methods */
class BatchUpdateSpec extends FunSuite:

  test("putMixed: single string value") {
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .putMixed(cell"A1" -> "Hello")

    assertEquals(sheet(cell"A1").value, CellValue.Text("Hello"))
    assertEquals(sheet.getCellStyle(cell"A1"), None) // Strings don't get styling
  }

  test("putMixed: single int value with auto-style") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .putMixed(cell"A1" -> 42)

    assertEquals(sheet(cell"A1").value, CellValue.Number(BigDecimal(42)))
    val style = sheet.getCellStyle(cell"A1")
    assert(style.exists(_.numFmt == NumFmt.General), "Should have General format")
  }

  test("putMixed: single BigDecimal with decimal format") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .putMixed(cell"A1" -> BigDecimal("123.45"))

    assertEquals(sheet(cell"A1").value, CellValue.Number(BigDecimal("123.45")))
    val style = sheet.getCellStyle(cell"A1")
    assert(style.exists(_.numFmt == NumFmt.Decimal), "Should have Decimal format")
  }

  test("putMixed: single LocalDate with date format") {
    val date = LocalDate.of(2025, 11, 10)
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .putMixed(cell"A1" -> date)

    sheet(cell"A1").value match
      case CellValue.DateTime(datetime) => assertEquals(datetime.toLocalDate, date)
      case other => fail(s"Expected DateTime, got $other")

    val style = sheet.getCellStyle(cell"A1")
    assert(style.exists(_.numFmt == NumFmt.Date), "Should have Date format")
  }

  test("putMixed: single LocalDateTime with datetime format") {
    val dt = LocalDateTime.of(2025, 11, 10, 14, 30)
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .putMixed(cell"A1" -> dt)

    assertEquals(sheet(cell"A1").value, CellValue.DateTime(dt))
    val style = sheet.getCellStyle(cell"A1")
    assert(style.exists(_.numFmt == NumFmt.DateTime), "Should have DateTime format")
  }

  test("putMixed: multiple values with mixed types") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .putMixed(
        cell"A1" -> "Title",
        cell"B1" -> 42,
        cell"C1" -> BigDecimal("123.45"),
        cell"D1" -> LocalDate.of(2025, 11, 10),
        cell"E1" -> true
      )

    // Verify values
    assertEquals(sheet(cell"A1").value, CellValue.Text("Title"))
    assertEquals(sheet(cell"B1").value, CellValue.Number(BigDecimal(42)))
    assertEquals(sheet(cell"C1").value, CellValue.Number(BigDecimal("123.45")))
    assert(sheet(cell"D1").value.isInstanceOf[CellValue.DateTime], "D1 should contain DateTime")
    assertEquals(sheet(cell"E1").value, CellValue.Bool(true))

    // Verify styles applied correctly
    assert(sheet.getCellStyle(cell"A1").isEmpty, "String should have no style")
    assert(sheet.getCellStyle(cell"B1").exists(_.numFmt == NumFmt.General), "Int should have General format")
    assert(sheet.getCellStyle(cell"C1").exists(_.numFmt == NumFmt.Decimal), "BigDecimal should have Decimal format")
    assert(sheet.getCellStyle(cell"D1").exists(_.numFmt == NumFmt.Date), "LocalDate should have Date format")
    assert(sheet.getCellStyle(cell"E1").isEmpty, "Boolean should have no style")
  }

  test("putMixed: builds on existing putAll") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))

    // Manual putAll
    val cells = Seq(
      Cell(cell"A1", CellValue.Text("Hello")),
      Cell(cell"B1", CellValue.Number(BigDecimal(42)))
    )
    val manual = sheet.putAll(cells)

    // Using putMixed
    val batched = sheet.putMixed(
      cell"A1" -> "Hello",
      cell"B1" -> 42
    )

    // Values should be the same (styles might differ)
    assertEquals(batched(cell"A1").value, manual(cell"A1").value)
    assertEquals(batched(cell"B1").value, manual(cell"B1").value)
  }

  test("putMixed: overwrites existing cells") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .putMixed(cell"A1" -> "Old")
      .putMixed(cell"A1" -> "New")

    assertEquals(sheet(cell"A1").value, CellValue.Text("New"))
  }

  test("putMixed: unsupported types are ignored") {
    // This is a manual test - unsupported types should be silently ignored
    case class UnsupportedType(value: String)

    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .putMixed(
        cell"A1" -> "Valid",
        cell"B1" -> UnsupportedType("Invalid") // Should be ignored
      )

    assertEquals(sheet(cell"A1").value, CellValue.Text("Valid"))
    assert(!sheet.contains(cell"B1")) // Unsupported type should not create a cell
  }

  test("putTyped: writes single typed value") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .putTyped(cell"A1", BigDecimal("123.45"))

    assertEquals(sheet(cell"A1").value, CellValue.Number(BigDecimal("123.45")))
    assert(sheet.getCellStyle(cell"A1").exists(_.numFmt == NumFmt.Decimal), "Should have Decimal format")
  }

  test("readTyped: reads existing cell") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(cell"A1", CellValue.Number(BigDecimal(42)))

    val result = sheet.readTyped[Int](cell"A1")
    assertEquals(result, Right(Some(42)))
  }

  test("readTyped: returns None for empty cell") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))

    val result = sheet.readTyped[Int](cell"A1")
    assertEquals(result, Right(None))
  }

  test("readTyped: returns error for type mismatch") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(cell"A1", CellValue.Text("not a number"))

    val result = sheet.readTyped[Int](cell"A1")
    assert(result.isLeft, "Should return error for type mismatch")
  }

  test("putMixed + readTyped: round-trip") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .putMixed(
        cell"A1" -> 42,
        cell"B1" -> BigDecimal("123.45"),
        cell"C1" -> "Text"
      )

    assertEquals(sheet.readTyped[Int](cell"A1"), Right(Some(42)))
    assertEquals(sheet.readTyped[BigDecimal](cell"B1"), Right(Some(BigDecimal("123.45"))))
    assertEquals(sheet.readTyped[String](cell"C1"), Right(Some("Text")))
  }

  test("putMixed: StyleRegistry size increases correctly") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .putMixed(
        cell"A1" -> "No style", // Strings don't add styles
        cell"B1" -> BigDecimal("123.45"), // Adds decimal style
        cell"C1" -> LocalDate.of(2025, 11, 10) // Adds date style
      )

    // Should have: default + decimal + date = 3 styles
    assert(sheet.styleRegistry.size >= 3, s"Expected at least 3 styles, got ${sheet.styleRegistry.size}")
  }

  test("putMixed: deduplicates identical styles") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .putMixed(
        cell"A1" -> BigDecimal("123.45"),
        cell"A2" -> BigDecimal("678.90") // Same type, should reuse style
      )

    // Both cells should have the same styleId
    assertEquals(
      sheet(cell"A1").styleId,
      sheet(cell"A2").styleId,
      "Cells with same type should have same style"
    )
  }
