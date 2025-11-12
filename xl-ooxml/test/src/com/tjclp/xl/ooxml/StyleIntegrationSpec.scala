package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.macros.cell
import com.tjclp.xl.sheet.syntax.*
import com.tjclp.xl.style.{CellStyle, Font, Fill, Color}
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile
import scala.xml.XML

/** End-to-end integration tests for style application */
class StyleIntegrationSpec extends FunSuite:

  val tempDir: FunFixture[Path] = FunFixture[Path](
    setup = _ => Files.createTempDirectory("xl-styles-"),
    teardown = dir =>
      Files
        .walk(dir)
        .sorted(java.util.Comparator.reverseOrder())
        .forEach(Files.delete)
  )

  tempDir.test("generated XLSX includes styled cells in XML") { dir =>
    val boldStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
    val redStyle = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))

    val sheet = Sheet("Styled").getOrElse(fail("Failed to create sheet"))
      .put(cell"A1", CellValue.Text("Bold Text"))
      .withCellStyle(cell"A1", boldStyle)
      .put(cell"A2", CellValue.Text("Red Background"))
      .withCellStyle(cell"A2", redStyle)

    val wb = Workbook(Vector(sheet))
    val path = dir.resolve("styled.xlsx")

    // Write workbook
    val result = XlsxWriter.write(wb, path)
    assert(result.isRight, s"Failed to write: $result")

    // Verify styles.xml structure
    val zipFile = new ZipFile(path.toFile)
    try {
      val stylesEntry = zipFile.getEntry("xl/styles.xml")
      val stylesXml = XML.load(zipFile.getInputStream(stylesEntry))

      // Should have cellXfs with 3 entries (default, bold, red)
      val cellXfs = (stylesXml \\ "cellXfs" \\ "xf")
      assertEquals(cellXfs.size, 3, "Should have 3 cellXf entries")

      // Verify worksheet has s= attributes
      val sheet1Entry = zipFile.getEntry("xl/worksheets/sheet1.xml")
      val sheetXml = XML.load(zipFile.getInputStream(sheet1Entry))

      val cells = (sheetXml \\ "sheetData" \\ "row" \\ "c")
      val cellA1 = cells.find(c => (c \ "@r").text == "A1")
      val cellA2 = cells.find(c => (c \ "@r").text == "A2")

      assert(cellA1.isDefined, "Cell A1 should exist")
      assert(cellA2.isDefined, "Cell A2 should exist")

      // Both should have s= attributes (not 0, which would be omitted)
      val a1StyleAttr = (cellA1.get \ "@s").text
      val a2StyleAttr = (cellA2.get \ "@s").text

      assert(a1StyleAttr.nonEmpty, "A1 should have style attribute")
      assert(a2StyleAttr.nonEmpty, "A2 should have style attribute")
      assert(a1StyleAttr != a2StyleAttr, "Different styles should have different indices")

      // Indices should be 1 and 2 (not 0 which is default)
      assert(a1StyleAttr.toInt > 0, "Style index should be > 0")
      assert(a2StyleAttr.toInt > 0, "Style index should be > 0")
    } finally {
      zipFile.close()
    }
  }

  tempDir.test("multi-sheet workbook deduplicates shared styles") { dir =>
    val headerStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))

    val sheet1 = Sheet("Sales").getOrElse(fail(""))
      .put(cell"A1", CellValue.Text("Product"))
      .withCellStyle(cell"A1", headerStyle)

    val sheet2 = Sheet("Inventory").getOrElse(fail(""))
      .put(cell"A1", CellValue.Text("Item"))
      .withCellStyle(cell"A1", headerStyle)

    val wb = Workbook(Vector(sheet1, sheet2))
    val path = dir.resolve("multi-styled.xlsx")

    val result = XlsxWriter.write(wb, path)
    assert(result.isRight)

    val zipFile = new ZipFile(path.toFile)
    try {
      val stylesEntry = zipFile.getEntry("xl/styles.xml")
      val stylesXml = XML.load(zipFile.getInputStream(stylesEntry))

      // Should have 2 cellXfs (default + header), NOT 3
      val cellXfs = (stylesXml \\ "cellXfs" \\ "xf")
      assertEquals(cellXfs.size, 2, "Shared style should be deduplicated")

      // Both sheets should reference same style index for their A1 cells
      val sheet1Entry = zipFile.getEntry("xl/worksheets/sheet1.xml")
      val sheet1Xml = XML.load(zipFile.getInputStream(sheet1Entry))
      val sheet1A1 = (sheet1Xml \\ "c").find(c => (c \ "@r").text == "A1").get
      val sheet1Style = (sheet1A1 \ "@s").text

      val sheet2Entry = zipFile.getEntry("xl/worksheets/sheet2.xml")
      val sheet2Xml = XML.load(zipFile.getInputStream(sheet2Entry))
      val sheet2A1 = (sheet2Xml \\ "c").find(c => (c \ "@r").text == "A1").get
      val sheet2Style = (sheet2A1 \ "@s").text

      assertEquals(sheet1Style, sheet2Style, "Both sheets should reference same style index")
    } finally {
      zipFile.close()
    }
  }

  tempDir.test("unstyled cells do not have s= attribute") { dir =>
    val sheet = Sheet("Plain").getOrElse(fail(""))
      .put(cell"A1", CellValue.Text("No Style"))
      .put(cell"A2", CellValue.Text("Also No Style"))
    // No styles applied

    val wb = Workbook(Vector(sheet))
    val path = dir.resolve("plain.xlsx")

    val result = XlsxWriter.write(wb, path)
    assert(result.isRight)

    val zipFile = new ZipFile(path.toFile)
    try {
      val sheet1Entry = zipFile.getEntry("xl/worksheets/sheet1.xml")
      val sheetXml = XML.load(zipFile.getInputStream(sheet1Entry))

      val cells = (sheetXml \\ "c")
      cells.foreach { cell =>
        val styleAttr = (cell \ "@s").text
        assert(styleAttr.isEmpty, s"Unstyled cells should not have s= attribute, got: $styleAttr")
      }
    } finally {
      zipFile.close()
    }
  }

  tempDir.test("round-trip with styles preserves formatting") { dir =>
    val boldStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))

    val sheet = Sheet("Test").getOrElse(fail(""))
      .put(cell"A1", CellValue.Text("Bold"))
      .withCellStyle(cell"A1", boldStyle)

    val wb = Workbook(Vector(sheet))
    val path = dir.resolve("round-trip.xlsx")

    // Write
    val writeResult = XlsxWriter.write(wb, path)
    assert(writeResult.isRight)

    // Read back
    val readResult = XlsxReader.read(path)
    assert(readResult.isRight, s"Read failed: $readResult")

    readResult.foreach { readWb =>
      assertEquals(readWb.sheets.size, 1)
      val readSheet = readWb.sheets(0)

      // Cell should exist with value
      assertEquals(readSheet(cell"A1").value, CellValue.Text("Bold"))

      // Cell should have styleId (TODO: Eventually verify it resolves to bold font)
      assert(readSheet(cell"A1").styleId.isDefined, "Styled cell should have styleId after round-trip")
    }
  }
