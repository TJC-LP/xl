package com.tjclp.xl

import com.tjclp.xl.api.*
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.conversions.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheet.syntax.*
import com.tjclp.xl.style.CellStyle
import com.tjclp.xl.style.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.style.border.{Border, BorderStyle}
import com.tjclp.xl.style.color.Color
import com.tjclp.xl.style.fill.Fill
import com.tjclp.xl.style.font.Font
import com.tjclp.xl.style.numfmt.NumFmt
import munit.FunSuite

/** Tests for Sheet style extension methods */
class SheetStyleSpec extends FunSuite:

  test("withCellStyle registers and applies style") {
    val boldStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .put(ref"A1", CellValue.Text("Header"))
      .withCellStyle(ref"A1", boldStyle)

    // Registry should have 2 styles (default + bold)
    assertEquals(sheet.styleRegistry.size, 2)

    // Cell should have styleId
    assert(sheet(ref"A1").styleId.isDefined)

    // Can retrieve style
    assertEquals(sheet.getCellStyle(ref"A1"), Some(boldStyle))
  }

  test("withCellStyle deduplicates same style applied twice") {
    val redStyle = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .put(ref"A1", "Text1")
      .put(ref"A2", "Text2")
      .withCellStyle(ref"A1", redStyle)
      .withCellStyle(ref"A2", redStyle)

    // Should only have 2 styles (default + red)
    assertEquals(sheet.styleRegistry.size, 2)

    // Both cells should have same styleId
    assertEquals(
      sheet(ref"A1").styleId,
      sheet(ref"A2").styleId
    )
  }

  test("withRangeStyle applies to all cells in ref") {
    val headerStyle = CellStyle.default.withFont(Font("Arial", 12.0, bold = true))
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .put(ref"A1", "Col1")
      .put(ref"B1", "Col2")
      .put(ref"C1", "Col3")
      .withRangeStyle(ref"A1:C1", headerStyle)

    // All cells should have styleId
    assert(sheet(ref"A1").styleId.isDefined)
    assert(sheet(ref"B1").styleId.isDefined)
    assert(sheet(ref"C1").styleId.isDefined)

    // All should be same styleId
    val styleId = sheet(ref"A1").styleId
    assertEquals(sheet(ref"B1").styleId, styleId)
    assertEquals(sheet(ref"C1").styleId, styleId)

    // All should retrieve same style
    assert(sheet.getCellStyle(ref"A1") == Some(headerStyle))
    assert(sheet.getCellStyle(ref"B1") == Some(headerStyle))
    assert(sheet.getCellStyle(ref"C1") == Some(headerStyle))
  }

  test("getCellStyle returns None for unstyled ref") {
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .put(ref"A1", CellValue.Text("Text"))

    assert(sheet.getCellStyle(ref"A1") == None)
  }

  test("getCellStyle returns None for empty ref") {
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))

    assert(sheet.getCellStyle(ref"A1") == None)
  }

  test("multiple different styles accumulate in registry") {
    val bold = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
    val red = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))
    val bordered = CellStyle.default.withBorder(Border.all(BorderStyle.Thin))

    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .put(ref"A1", CellValue.Text("Bold"))
      .put(ref"A2", CellValue.Text("Red"))
      .put(ref"A3", CellValue.Text("Bordered"))
      .withCellStyle(ref"A1", bold)
      .withCellStyle(ref"A2", red)
      .withCellStyle(ref"A3", bordered)

    // Should have 4 styles (default + 3 custom)
    assertEquals(sheet.styleRegistry.size, 4)

    // Each cell should have different styleId
    assert(sheet(ref"A1").styleId != sheet(ref"A2").styleId, "A1 and A2 should have different styles")
    assert(sheet(ref"A2").styleId != sheet(ref"A3").styleId, "A2 and A3 should have different styles")
  }

  test("withCellStyle creates cell if it doesn't exist") {
    val boldStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
    val sheet = Sheet("Test").getOrElse(fail("Failed to create sheet"))
      .withCellStyle(ref"A1", boldStyle) // No put() first

    // Cell should exist with style
    assert(sheet.contains(ref"A1"))
    assert(sheet(ref"A1").styleId.isDefined)
    assertEquals(sheet.getCellStyle(ref"A1"), Some(boldStyle))
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
      .put(ref"A1", CellValue.Text("Name"))
      .put(ref"B1", CellValue.Text("Amount"))
      .withRangeStyle(ref"A1:B1", headerStyle)
      // Data rows
      .put(ref"A2", CellValue.Text("Item 1"))
      .put(ref"B2", CellValue.Number(BigDecimal("123.45")))
      .withCellStyle(ref"B2", dataStyle)

    // Should have 3 styles (default + header + data)
    assertEquals(sheet.styleRegistry.size, 3)

    // Verify styles applied correctly
    assert(sheet.getCellStyle(ref"A1") == Some(headerStyle))
    assert(sheet.getCellStyle(ref"B1") == Some(headerStyle))
    assert(sheet.getCellStyle(ref"B2") == Some(dataStyle))
  }
