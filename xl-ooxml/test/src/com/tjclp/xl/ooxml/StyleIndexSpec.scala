package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.{CellStyle, Font, Fill, Color}
import com.tjclp.xl.macros.ref

/** Tests for StyleIndex.fromWorkbook with style remapping */
@SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
class StyleIndexSpec extends FunSuite:

  test("fromWorkbook with single sheet extracts styles from registry") {
    val boldStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
    val redStyle = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))

    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Bold"))
      .withCellStyle(ref"A1", boldStyle)
      .put(ref"A2", CellValue.Text("Red"))
      .withCellStyle(ref"A2", redStyle)

    val wb = Workbook(Vector(sheet))
    val (index, remappings) = StyleIndex.fromWorkbook(wb)

    // Should have 3 styles (default + bold + red)
    assertEquals(index.cellStyles.size, 3)

    // Sheet 0 should have remapping for all 3 local indices
    val sheetRemapping = remappings.getOrElse(0, Map.empty)
    assertEquals(sheetRemapping.size, 3)

    // Default should map to 0
    assertEquals(sheetRemapping.get(0), Some(0))

    // Bold and red should map to 1 and 2
    assert(sheetRemapping.contains(1))
    assert(sheetRemapping.contains(2))
  }

  test("fromWorkbook with multiple sheets shares identical styles") {
    val boldStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))

    val sheet1 = Sheet("Sheet1")
      .put(ref"A1", CellValue.Text("Bold1"))
      .withCellStyle(ref"A1", boldStyle)

    val sheet2 = Sheet("Sheet2")
      .put(ref"B1", CellValue.Text("Bold2"))
      .withCellStyle(ref"B1", boldStyle)

    val wb = Workbook(Vector(sheet1, sheet2))
    val (index, remappings) = StyleIndex.fromWorkbook(wb)

    // Should have 2 styles (default + bold), NOT 3
    assertEquals(index.cellStyles.size, 2, "Bold should be deduplicated")

    // Both sheets' local index 1 should map to same global index
    val sheet1Remapping = remappings.getOrElse(0, Map.empty)
    val sheet2Remapping = remappings.getOrElse(1, Map.empty)

    assertEquals(sheet1Remapping.get(1), sheet2Remapping.get(1), "Same style should map to same global index")
  }

  test("fromWorkbook with conflicting local indices remaps correctly") {
    val boldStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
    val redStyle = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))

    // Sheet1: localId=1 is bold
    val sheet1 = Sheet("Sheet1")
      .put(ref"A1", CellValue.Text("Bold"))
      .withCellStyle(ref"A1", boldStyle)

    // Sheet2: localId=1 is red (different style)
    val sheet2 = Sheet("Sheet2")
      .put(ref"A1", CellValue.Text("Red"))
      .withCellStyle(ref"A1", redStyle)

    val wb = Workbook(Vector(sheet1, sheet2))
    val (index, remappings) = StyleIndex.fromWorkbook(wb)

    // Should have 3 styles (default + bold + red)
    assertEquals(index.cellStyles.size, 3)

    val sheet1Remapping = remappings.getOrElse(0, Map.empty)
    val sheet2Remapping = remappings.getOrElse(1, Map.empty)

    // Local index 1 should map to DIFFERENT global indices
    assertNotEquals(
      sheet1Remapping.get(1),
      sheet2Remapping.get(1),
      "Different styles should have different global indices"
    )
  }

  test("fromWorkbook with empty styleRegistry uses default only") {
    val sheet = Sheet("Plain")
      .put(ref"A1", CellValue.Text("No Style"))
    // Don't apply any styles - registry is default

    val wb = Workbook(Vector(sheet))
    val (index, remappings) = StyleIndex.fromWorkbook(wb)

    // Should only have default style
    assertEquals(index.cellStyles.size, 1)
    assertEquals(index.cellStyles.head, CellStyle.default)

    // Remapping should only have default
    val sheetRemapping = remappings.getOrElse(0, Map.empty)
    assertEquals(sheetRemapping.size, 1)
    assertEquals(sheetRemapping.get(0), Some(0))
  }

  test("fromWorkbook deduplicates fonts, fills, and borders") {
    val style1 = CellStyle.default.withFont(Font("Arial", 12.0, bold = true))
    val style2 = CellStyle.default.withFont(Font("Arial", 12.0, bold = true)).withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))
    // Both use same font

    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Style1"))
      .withCellStyle(ref"A1", style1)
      .put(ref"A2", CellValue.Text("Style2"))
      .withCellStyle(ref"A2", style2)

    val wb = Workbook(Vector(sheet))
    val (index, _) = StyleIndex.fromWorkbook(wb)

    // Should have 3 CellStyles: default, style1, style2
    assertEquals(index.cellStyles.size, 3)

    // But fonts should be deduplicated (default font + Arial bold)
    assert(index.fonts.size < index.cellStyles.size, "Fonts should be deduplicated across styles")
  }

  test("fromWorkbook preserves canonical key equality across sheets") {
    // Two structurally equal styles in different sheets
    val boldStyle = CellStyle.default.withFont(Font("Arial", 12.0, bold = true))

    val sheet1 = Sheet("Sheet1")
      .put(ref"A1", CellValue.Text("Bold"))
      .withCellStyle(ref"A1", boldStyle)

    val sheet2 = Sheet("Sheet2")
      .put(ref"A1", CellValue.Text("Also Bold"))
      .withCellStyle(ref"A1", boldStyle)

    val wb = Workbook(Vector(sheet1, sheet2))
    val (index, remappings) = StyleIndex.fromWorkbook(wb)

    // Should have 2 styles (default + bold), deduplicated across sheets
    assertEquals(index.cellStyles.size, 2, "Identical styles across sheets should be deduplicated")

    // Both sheets should map their local index 1 to same global index
    val sheet1Remapping = remappings.getOrElse(0, Map.empty)
    val sheet2Remapping = remappings.getOrElse(1, Map.empty)

    assertEquals(sheet1Remapping.get(1), sheet2Remapping.get(1), "Same style should map to same global index")
  }
