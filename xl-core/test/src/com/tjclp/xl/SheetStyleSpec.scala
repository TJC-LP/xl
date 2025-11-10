package com.tjclp.xl

import munit.FunSuite
import com.tjclp.xl.macros.{cell, range}
import com.tjclp.xl.conversions.given

/** Tests for Sheet style extension methods */
class SheetStyleSpec extends FunSuite:

  test("withCellStyle registers and applies style") {
    val boldStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .put(cell"A1", CellValue.Text("Header"))
      .withCellStyle(cell"A1", boldStyle)

    // Registry should have 2 styles (default + bold)
    assertEquals(sheet.styleRegistry.size, 2)

    // Cell should have styleId
    assert(sheet(cell"A1").styleId.isDefined)

    // Can retrieve style
    assertEquals(sheet.getCellStyle(cell"A1"), Some(boldStyle))
  }

  test("withCellStyle deduplicates same style applied twice") {
    val redStyle = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .put(cell"A1", "Text1")
      .put(cell"A2", "Text2")
      .withCellStyle(cell"A1", redStyle)
      .withCellStyle(cell"A2", redStyle)

    // Should only have 2 styles (default + red)
    assertEquals(sheet.styleRegistry.size, 2)

    // Both cells should have same styleId
    assertEquals(
      sheet(cell"A1").styleId,
      sheet(cell"A2").styleId
    )
  }

  test("withRangeStyle applies to all cells in range") {
    val headerStyle = CellStyle.default.withFont(Font("Arial", 12.0, bold = true))
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .put(cell"A1", "Col1")
      .put(cell"B1", "Col2")
      .put(cell"C1", "Col3")
      .withRangeStyle(range"A1:C1", headerStyle)

    // All cells should have styleId
    assert(sheet(cell"A1").styleId.isDefined)
    assert(sheet(cell"B1").styleId.isDefined)
    assert(sheet(cell"C1").styleId.isDefined)

    // All should be same styleId
    val styleId = sheet(cell"A1").styleId
    assertEquals(sheet(cell"B1").styleId, styleId)
    assertEquals(sheet(cell"C1").styleId, styleId)

    // All should retrieve same style
    assert(sheet.getCellStyle(cell"A1") == Some(headerStyle))
    assert(sheet.getCellStyle(cell"B1") == Some(headerStyle))
    assert(sheet.getCellStyle(cell"C1") == Some(headerStyle))
  }

  test("getCellStyle returns None for unstyled cell") {
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .put(cell"A1", CellValue.Text("Text"))

    assert(sheet.getCellStyle(cell"A1") == None)
  }

  test("getCellStyle returns None for empty cell") {
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))

    assert(sheet.getCellStyle(cell"A1") == None)
  }

  test("multiple different styles accumulate in registry") {
    val bold = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
    val red = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))
    val bordered = CellStyle.default.withBorder(Border.all(BorderStyle.Thin))

    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .put(cell"A1", CellValue.Text("Bold"))
      .put(cell"A2", CellValue.Text("Red"))
      .put(cell"A3", CellValue.Text("Bordered"))
      .withCellStyle(cell"A1", bold)
      .withCellStyle(cell"A2", red)
      .withCellStyle(cell"A3", bordered)

    // Should have 4 styles (default + 3 custom)
    assertEquals(sheet.styleRegistry.size, 4)

    // Each cell should have different styleId
    assert(sheet(cell"A1").styleId != sheet(cell"A2").styleId, "A1 and A2 should have different styles")
    assert(sheet(cell"A2").styleId != sheet(cell"A3").styleId, "A2 and A3 should have different styles")
  }

  test("withCellStyle creates cell if it doesn't exist") {
    val boldStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .withCellStyle(cell"A1", boldStyle) // No put() first

    // Cell should exist with style
    assert(sheet.contains(cell"A1"))
    assert(sheet(cell"A1").styleId.isDefined)
    assertEquals(sheet.getCellStyle(cell"A1"), Some(boldStyle))
  }

  test("complex styling workflow") {
    val headerStyle = CellStyle.default
      .withFont(Font("Arial", 14.0, bold = true))
      .withFill(Fill.Solid(Color.Rgb(0xFFCCCCCC)))
      .withAlign(Align(HAlign.Center, VAlign.Middle))

    val dataStyle = CellStyle.default
      .withFont(Font("Calibri", 11.0))
      .withNumFmt(NumFmt.Decimal)

    val sheet = Sheet("Report").getOrElse(fail("Failed to create sheet"))
      // Header row
      .put(cell"A1", CellValue.Text("Name"))
      .put(cell"B1", CellValue.Text("Amount"))
      .withRangeStyle(range"A1:B1", headerStyle)
      // Data rows
      .put(cell"A2", CellValue.Text("Item 1"))
      .put(cell"B2", CellValue.Number(BigDecimal("123.45")))
      .withCellStyle(cell"B2", dataStyle)

    // Should have 3 styles (default + header + data)
    assertEquals(sheet.styleRegistry.size, 3)

    // Verify styles applied correctly
    assert(sheet.getCellStyle(cell"A1") == Some(headerStyle))
    assert(sheet.getCellStyle(cell"B1") == Some(headerStyle))
    assert(sheet.getCellStyle(cell"B2") == Some(dataStyle))
  }
