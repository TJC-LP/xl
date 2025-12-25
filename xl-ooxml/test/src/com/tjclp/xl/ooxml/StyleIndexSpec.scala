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

  test("fromWorkbookWithSource adds new font/fill components for new styles (regression #style-components-bug)") {
    // Regression test: When adding new styles to existing files, the font/fill/border
    // components weren't being added to the output. New styles referenced non-existent
    // component indices, falling back to defaults (fontId=0, fillId=0).

    import java.nio.file.Files

    // Create initial workbook with just default style
    val initialSheet = Sheet("Test").put(ref"A1", CellValue.Text("Original"))
    val initialWb = Workbook(Vector(initialSheet))

    val path = Files.createTempFile("style-components-test", ".xlsx")
    try
      // Write initial workbook
      XlsxWriter.write(initialWb, path).fold(err => fail(s"Write failed: $err"), identity)

      // Read it back
      val readWb = XlsxReader.read(path).fold(err => fail(s"Read failed: $err"), identity)

      // Add new style with bold font and blue fill
      val boldBlueFill = CellStyle.default
        .withFont(Font("Calibri", 11.0, bold = true, color = Some(Color.Rgb(0xFFFFFFFF))))
        .withFill(Fill.Solid(Color.Rgb(0xFF003366)))

      val modifiedSheet = readWb.sheets.head
        .put(ref"A1", CellValue.Text("Styled"))
        .withCellStyle(ref"A1", boldBlueFill)

      val modifiedWb = readWb.put(modifiedSheet)

      // Write modified workbook
      val outputPath = Files.createTempFile("style-components-output", ".xlsx")
      try
        XlsxWriter.write(modifiedWb, outputPath).fold(err => fail(s"Write failed: $err"), identity)

        // Read the styles.xml directly to verify components were added
        val zip = new java.util.zip.ZipFile(outputPath.toFile)
        val stylesEntry = zip.getEntry("xl/styles.xml")
        val stylesXml = scala.io.Source.fromInputStream(zip.getInputStream(stylesEntry)).mkString
        zip.close()

        // Verify new font with bold was added
        // Note: StAX outputs <b></b> while ScalaXml outputs <b/>
        assert(
          stylesXml.contains("<b/>") || stylesXml.contains("<b>") || stylesXml.contains("<b "),
          s"Bold font element missing from styles.xml. New font components not added."
        )

        // Verify new fill with blue color was added
        assert(
          stylesXml.contains("003366") || stylesXml.contains("003366"),
          s"Blue fill color missing from styles.xml. New fill components not added."
        )

        // Verify cellXf references non-zero fontId and fillId
        // Parse to check the last cellXf has fontId > 0 and fillId > 0
        val fontIdPattern = """fontId="(\d+)"""".r
        val fillIdPattern = """fillId="(\d+)"""".r
        val fontIds = fontIdPattern.findAllMatchIn(stylesXml).map(_.group(1).toInt).toList
        val fillIds = fillIdPattern.findAllMatchIn(stylesXml).map(_.group(1).toInt).toList

        assert(fontIds.exists(_ > 0), "No non-default fontId found in cellXfs")
        assert(fillIds.exists(_ > 1), "No non-default fillId found in cellXfs (0=none, 1=gray125)")
      finally Files.deleteIfExists(outputPath)
    finally Files.deleteIfExists(path)
  }
