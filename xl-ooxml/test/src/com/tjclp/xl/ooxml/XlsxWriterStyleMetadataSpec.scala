package com.tjclp.xl.ooxml

import java.io.ByteArrayInputStream
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
import javax.xml.parsers.DocumentBuilderFactory

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * Regression tests for style and metadata preservation in surgical modification.
 *
 * These tests prevent regression of secondary corruption issues related to:
 *   1. Style ID conflicts between original and new styles
 *   2. Worksheet metadata loss (cols, views, conditionalFormatting, etc.)
 *   3. OOXML schema element ordering violations
 *   4. Differential formats (dxfs) missing
 *
 * Corresponds to commits:
 *   - 8f58f15: Surgical style mode
 *   - 2304cec: Worksheet metadata preservation
 *   - 802e020: Element ordering
 *   - 4998af2: dxfs preservation
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.AsInstanceOf"))
class XlsxWriterStyleMetadataSpec extends FunSuite:

  test("surgical style mode preserves original style IDs for unmodified sheets") {
    // Regression test for commit 8f58f15
    // Bug: Style deduplication caused unmodified sheets to reference non-existent style IDs
    // Solution: StyleIndex.fromWorkbookSurgical preserves all original styles

    val source = createWorkbookWithManyStyles()

    // Modify only sheet 1 (adds new styles), leave sheet 2 untouched
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet = sheet.put(ref"A1" -> "Modified")
    yield wb.put(updatedSheet)

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    // Write back
    val output = Files.createTempFile("style-surgical", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify styles.xml has original style count + new styles
    val outputZip = new ZipFile(output.toFile)
    val stylesXml = readEntryString(outputZip, outputZip.getEntry("xl/styles.xml"))

    val doc = parseXml(stylesXml)
    val cellXfsElems = doc.getElementsByTagName("cellXfs")
    assert(cellXfsElems.getLength > 0, "cellXfs element missing")

    val cellXfs = cellXfsElems.item(0).asInstanceOf[org.w3c.dom.Element]
    val countAttr = cellXfs.getAttribute("count").toInt

    // Should have at least 5 original styles (from source)
    assert(
      countAttr >= 5,
      s"Style count $countAttr too low - original styles may have been dropped"
    )

    // Verify unmodified Sheet2 uses original style IDs (e.g., s="3")
    val sheet2Xml = readEntryString(outputZip, outputZip.getEntry("xl/worksheets/sheet2.xml"))
    assert(
      sheet2Xml.contains("""s="3""""),
      "Sheet2 should reference original style ID 3 (surgical mode preserves IDs)"
    )

    // Verify modified Sheet1 can use new styles (higher IDs)
    val sheet1Xml = readEntryString(outputZip, outputZip.getEntry("xl/worksheets/sheet1.xml"))
    // Sheet1 should have been regenerated successfully
    assert(sheet1Xml.contains("Modified"), "Modified cell content missing")

    // Clean up
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  test("preserves worksheet metadata during sheet regeneration") {
    // Regression test for commit 2304cec
    // Bug: 14 metadata fields lost during regeneration (cols, views, conditionalFormatting, etc.)
    // Impact: Conditional formatting, print settings, frozen panes all lost

    val source = createWorkbookWithRichMetadata()

    // Modify one cell
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet = sheet.put(ref"A1" -> "Modified")
    yield wb.put(updatedSheet)

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    // Write back
    val output = Files.createTempFile("metadata-preserve", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify sheet1.xml preserves metadata
    val outputZip = new ZipFile(output.toFile)
    val sheetXml = readEntryString(outputZip, outputZip.getEntry("xl/worksheets/sheet1.xml"))

    // Critical metadata fields
    assert(sheetXml.contains("<sheetViews>"), "sheetViews metadata missing (frozen panes, zoom)")
    assert(sheetXml.contains("<cols>"), "cols metadata missing (column widths)")
    assert(
      sheetXml.contains("<conditionalFormatting"),
      "conditionalFormatting metadata missing (data bars, color scales)"
    )
    assert(sheetXml.contains("<pageMargins"), "pageMargins metadata missing (print layout)")
    assert(sheetXml.contains("<pageSetup"), "pageSetup metadata missing (page size, orientation)")

    // Verify frozen panes preserved
    assert(
      sheetXml.contains("frozenPane") || sheetXml.contains("<pane"),
      "Frozen pane configuration lost"
    )

    // Verify column width preserved (col element)
    // Note: Attribute order may vary between backends but data is preserved
    assert(sheetXml.contains("min=\"1\""), "Column width definition lost")

    // Clean up
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  test("emits worksheet elements in OOXML schema order (§18.3.1.99)") {
    // Regression test for commit 2304cec
    // Bug: Elements emitted in wrong order, violates OOXML Part 1 §18.3.1.99
    // Impact: Excel may reject file or repair it

    val source = createWorkbookWithAllMetadataTypes()

    // Modify one cell
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet = sheet.put(ref"A1" -> "Modified")
    yield wb.put(updatedSheet)

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    // Write back
    val output = Files.createTempFile("element-order", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify sheet1.xml has correct element order
    val outputZip = new ZipFile(output.toFile)
    val sheetXml = readEntryString(outputZip, outputZip.getEntry("xl/worksheets/sheet1.xml"))

    // Extract positions of key elements
    val sheetPrPos = sheetXml.indexOf("<sheetPr")
    val dimensionPos = sheetXml.indexOf("<dimension")
    val sheetViewsPos = sheetXml.indexOf("<sheetViews")
    val colsPos = sheetXml.indexOf("<cols")
    val sheetDataPos = sheetXml.indexOf("<sheetData")
    val conditionalFormattingPos = sheetXml.indexOf("<conditionalFormatting")
    val pageMarginsPos = sheetXml.indexOf("<pageMargins")

    // OOXML schema order (§18.3.1.99):
    // sheetPr → dimension → sheetViews → sheetFormatPr → cols → sheetData → conditionalFormatting → pageMargins → ...

    // Verify ordering (earlier elements should have lower positions)
    if (sheetPrPos > 0 && dimensionPos > 0) {
      assert(sheetPrPos < dimensionPos, "sheetPr should come before dimension")
    }

    if (dimensionPos > 0 && sheetViewsPos > 0) {
      assert(dimensionPos < sheetViewsPos, "dimension should come before sheetViews")
    }

    if (sheetViewsPos > 0 && colsPos > 0) {
      assert(sheetViewsPos < colsPos, "sheetViews should come before cols")
    }

    if (colsPos > 0 && sheetDataPos > 0) {
      assert(colsPos < sheetDataPos, "cols should come before sheetData")
    }

    if (sheetDataPos > 0 && conditionalFormattingPos > 0) {
      assert(
        sheetDataPos < conditionalFormattingPos,
        "sheetData should come before conditionalFormatting"
      )
    }

    if (conditionalFormattingPos > 0 && pageMarginsPos > 0) {
      assert(
        conditionalFormattingPos < pageMarginsPos,
        "conditionalFormatting should come before pageMargins"
      )
    }

    // Clean up
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  test("preserves differential formats (dxfs) for conditional formatting") {
    // Regression test for commit 4998af2
    // Bug: Missing <dxfs> section caused Excel repair message
    // Impact: Conditional formatting styles lost

    val source = createWorkbookWithDxfs()

    // Modify one cell
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet = sheet.put(ref"A1" -> "Modified")
    yield wb.put(updatedSheet)

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    // Write back
    val output = Files.createTempFile("dxfs-preserve", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify styles.xml has <dxfs> section
    val outputZip = new ZipFile(output.toFile)
    val stylesXml = readEntryString(outputZip, outputZip.getEntry("xl/styles.xml"))

    assert(stylesXml.contains("<dxfs"), "dxfs section missing (conditional formatting styles lost)")

    // Parse and verify dxfs position (should be after cellXfs, before tableStyles)
    val cellXfsPos = stylesXml.indexOf("</cellXfs>")
    val dxfsPos = stylesXml.indexOf("<dxfs")
    val tableStylesPos = stylesXml.indexOf("<tableStyles")

    assert(dxfsPos > 0, "dxfs element not found")
    assert(
      cellXfsPos < dxfsPos,
      "dxfs should come after cellXfs (OOXML schema order)"
    )

    if (tableStylesPos > 0) {
      assert(
        dxfsPos < tableStylesPos,
        "dxfs should come before tableStyles (OOXML schema order)"
      )
    }

    // Verify dxfs has at least one dxf element
    assert(
      stylesXml.contains("<dxf>"),
      "dxfs should contain at least one dxf element (differential format)"
    )

    // Clean up
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ===== Helper Methods =====

  private def readEntryString(zip: ZipFile, entry: ZipEntry): String =
    val is = zip.getInputStream(entry)
    try new String(is.readAllBytes(), "UTF-8")
    finally is.close()

  private def parseXml(xml: String): org.w3c.dom.Document =
    val factory = DocumentBuilderFactory.newInstance()
    factory.setNamespaceAware(true)
    val builder = factory.newDocumentBuilder()
    builder.parse(new ByteArrayInputStream(xml.getBytes("UTF-8")))

  // ===== Test File Creators =====

  private def createWorkbookWithManyStyles(): Path =
    val path = Files.createTempFile("test-many-styles", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))
    out.setLevel(1) // Match production compression level

    try
      writeEntry(
        out,
        "[Content_Types].xml",
        """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""
      )

      writeEntry(
        out,
        "_rels/.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
      )

      writeEntry(
        out,
        "xl/workbook.xml",
        """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"""
      )

      writeEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
      )

      // Styles with 5 different cellXfs
      writeEntry(
        out,
        "xl/styles.xml",
        """<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><i/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
</styleSheet>"""
      )

      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Bold</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

      // Sheet2 uses style ID 3 (should be preserved in output)
      writeEntry(
        out,
        "xl/worksheets/sheet2.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="3" t="inlineStr"><is><t>Yellow Fill</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

    finally out.close()

    path

  private def createWorkbookWithRichMetadata(): Path =
    val path = Files.createTempFile("test-metadata", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))
    out.setLevel(1) // Match production compression level

    try
      writeEntry(
        out,
        "[Content_Types].xml",
        """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
      )

      writeEntry(
        out,
        "_rels/.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
      )

      writeEntry(
        out,
        "xl/workbook.xml",
        """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
      )

      writeEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
      )

      // Worksheet with rich metadata
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
    </sheetView>
  </sheetViews>
  <cols>
    <col min="1" max="1" width="15.5" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Data</t></is></c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="A1:A10">
    <cfRule type="dataBar" priority="1">
      <dataBar>
        <cfvo type="min"/><cfvo type="max"/>
        <color rgb="FF638EC6"/>
      </dataBar>
    </cfRule>
  </conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup paperSize="1" orientation="portrait"/>
</worksheet>"""
      )

    finally out.close()

    path

  private def createWorkbookWithAllMetadataTypes(): Path =
    val path = Files.createTempFile("test-all-metadata", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))
    out.setLevel(1) // Match production compression level

    try
      writeEntry(
        out,
        "[Content_Types].xml",
        """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
      )

      writeEntry(
        out,
        "_rels/.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
      )

      writeEntry(
        out,
        "xl/workbook.xml",
        """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
      )

      writeEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
      )

      // Worksheet with ALL metadata types in OOXML schema order
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><outlinePr summaryBelow="0"/></sheetPr>
  <dimension ref="A1:C10"/>
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="12"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Data</t></is></c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="A1:A10">
    <cfRule type="cellIs" operator="greaterThan" priority="1">
      <formula>10</formula>
    </cfRule>
  </conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup paperSize="1"/>
</worksheet>"""
      )

    finally out.close()

    path

  private def createWorkbookWithDxfs(): Path =
    val path = Files.createTempFile("test-dxfs", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))
    out.setLevel(1) // Match production compression level

    try
      writeEntry(
        out,
        "[Content_Types].xml",
        """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""
      )

      writeEntry(
        out,
        "_rels/.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
      )

      writeEntry(
        out,
        "xl/workbook.xml",
        """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
      )

      writeEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
      )

      // Styles with dxfs section (for conditional formatting)
      writeEntry(
        out,
        "xl/styles.xml",
        """<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <dxfs count="2">
    <dxf>
      <font><color rgb="FF9C0006"/></font>
      <fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill>
    </dxf>
    <dxf>
      <font><color rgb="FF006100"/></font>
      <fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill>
    </dxf>
  </dxfs>
</styleSheet>"""
      )

      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Data</t></is></c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="A1:A10">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan">
      <formula>5</formula>
    </cfRule>
    <cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan">
      <formula>10</formula>
    </cfRule>
  </conditionalFormatting>
</worksheet>"""
      )

    finally out.close()

    path

  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    val entry = new ZipEntry(name)
    out.putNextEntry(entry)
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()
