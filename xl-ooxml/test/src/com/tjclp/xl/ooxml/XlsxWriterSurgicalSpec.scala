package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import com.tjclp.xl.api.*
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * Tests for XlsxWriter surgical modification support (Phase 4).
 *
 * Verifies that:
 *   1. Clean workbooks are copied verbatim (fast path)
 *   1. Modified workbooks use hybrid write (preserve + regenerate)
 *   1. Unknown parts (charts, images) are preserved byte-for-byte
 *   1. Modified sheets are regenerated with new content
 *   1. Unmodified sheets are copied from source
 *   1. Workbooks without SourceContext use full regeneration
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class XlsxWriterSurgicalSpec extends FunSuite:

  test("clean workbook write copies source verbatim") {
    // Create minimal workbook with chart
    val source = createMinimalWorkbookWithChart()

    // Read workbook (creates SourceContext)
    val wb = XlsxReader
      .read(source)
      .fold(err => fail(s"Failed to read: $err"), identity)

    assert(wb.sourceContext.isDefined, "Workbook should have SourceContext")
    assert(wb.sourceContext.get.isClean, "Freshly read workbook should be clean")

    // Write back without modifications
    val output = Files.createTempFile("clean-write", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify source and output are identical size (verbatim copy)
    assertEquals(Files.size(output), Files.size(source))

    // Clean up
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  test("modified sheet triggers hybrid write with preservation") {
    // Create workbook with chart
    val source = createMinimalWorkbookWithChart()

    // Read and modify one cell
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet <- sheet.put(ref"A1" -> "Modified")
      updated <- wb.put(updatedSheet)
    yield updated

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    assert(wb.sourceContext.isDefined)
    assert(!wb.sourceContext.get.isClean, "Modified workbook should be dirty")

    // Write back
    val output = Files.createTempFile("modified-write", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify chart is preserved in output
    val outputZip = new ZipFile(output.toFile)
    val chartEntry = outputZip.getEntry("xl/charts/chart1.xml")
    assert(chartEntry != null, "Chart should be preserved in output")

    // Verify chart bytes are identical to source (byte-for-byte preservation)
    val sourceZip = new ZipFile(source.toFile)
    val sourceChartEntry = sourceZip.getEntry("xl/charts/chart1.xml")
    val sourceBytes = readEntryBytes(sourceZip, sourceChartEntry)
    val outputBytes = readEntryBytes(outputZip, chartEntry)

    assertEquals(outputBytes.toSeq, sourceBytes.toSeq, "Chart should be byte-identical")

    // Verify modified cell is in output
    val reloaded = XlsxReader
      .read(output)
      .fold(err => fail(s"Failed to reload: $err"), identity)
    val cellValue = reloaded("Sheet1")
      .flatMap(sheet => Right(sheet.cells.get(ref"A1")))
      .fold(err => fail(s"Failed to get cell: $err"), identity)

    assertEquals(cellValue.map(_.value), Some(CellValue.Text("Modified")))

    // Clean up
    sourceZip.close()
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  test("unmodified sheets are copied, modified sheets regenerated") {
    // Create workbook with 2 sheets
    val source = createWorkbookWith2Sheets()

    // Modify only sheet 1, leave sheet 2 untouched
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet <- sheet.put(ref"A1" -> "Changed")
      updated <- wb.put(updatedSheet)
    yield updated

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    val tracker = wb.sourceContext.get.modificationTracker
    assertEquals(tracker.modifiedSheets, Set(0), "Only sheet 0 should be modified")

    // Write back
    val output = Files.createTempFile("partial-modify", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify sheet2 is byte-identical to source (copied, not regenerated)
    val sourceZip = new ZipFile(source.toFile)
    val outputZip = new ZipFile(output.toFile)

    val sourceSheet2 = readEntryBytes(sourceZip, sourceZip.getEntry("xl/worksheets/sheet2.xml"))
    val outputSheet2 = readEntryBytes(outputZip, outputZip.getEntry("xl/worksheets/sheet2.xml"))

    assertEquals(
      outputSheet2.toSeq,
      sourceSheet2.toSeq,
      "Unmodified sheet2 should be byte-identical"
    )

    // Verify sheet1 has new content
    val reloaded = XlsxReader
      .read(output)
      .fold(err => fail(s"Failed to reload: $err"), identity)
    val cell = reloaded("Sheet1")
      .flatMap(s => Right(s.cells.get(ref"A1")))
      .fold(err => fail(s"Failed to get cell: $err"), identity)

    assertEquals(cell.map(_.value), Some(CellValue.Text("Changed")))

    // Clean up
    sourceZip.close()
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  test("workbook without SourceContext uses full regeneration") {
    // Create workbook programmatically (no SourceContext)
    val result = for
      wb <- Workbook("TestSheet")
      sheet <- wb("TestSheet")
      updatedSheet <- sheet.put(ref"A1" -> "Test")
      updated <- wb.put(updatedSheet)
    yield updated

    val wb = result.fold(err => fail(s"Failed to create workbook: $err"), identity)

    assert(wb.sourceContext.isEmpty, "Programmatic workbook should not have SourceContext")

    // Write (should use regenerateAll, not throw error)
    val output = Files.createTempFile("programmatic", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify output is valid XLSX
    val reloaded = XlsxReader
      .read(output)
      .fold(err => fail(s"Failed to reload: $err"), identity)

    assertEquals(reloaded.sheets.size, 1)
    assertEquals(reloaded.sheets(0).name.value, "TestSheet")

    // Clean up
    Files.deleteIfExists(output)
  }

  test("unknown parts in unparsedParts are preserved") {
    // Create workbook with unknown part
    val source = createWorkbookWithUnknownPart()

    // Read and modify
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet <- sheet.put(ref"B2" -> "Modified")
      updated <- wb.put(updatedSheet)
    yield updated

    val wb = modified.fold(err => fail(s"Failed: $err"), identity)

    // Write back
    val output = Files.createTempFile("unknown-preserved", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify unknown part is preserved
    val outputZip = new ZipFile(output.toFile)
    val unknownEntry = outputZip.getEntry("xl/custom/data.xml")
    assert(unknownEntry != null, "Unknown part should be preserved")

    // Verify content is identical
    val sourceZip = new ZipFile(source.toFile)
    val sourceBytes = readEntryBytes(sourceZip, sourceZip.getEntry("xl/custom/data.xml"))
    val outputBytes = readEntryBytes(outputZip, unknownEntry)

    assertEquals(outputBytes.toSeq, sourceBytes.toSeq)

    // Clean up
    sourceZip.close()
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  test("sharedStrings part is preserved and regenerated sheet uses SST references") {
    val source = createWorkbookWithSharedStrings()

    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet <- sheet.put(ref"A1" -> "Updated Text")
      updated <- wb.put(updatedSheet)
    yield updated

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    val output = Files.createTempFile("sst-preserve", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    val sourceZip = new ZipFile(source.toFile)
    val outputZip = new ZipFile(output.toFile)

    val sharedStringsEntry = "xl/sharedStrings.xml"
    val sourceSst = readEntryBytes(sourceZip, sourceZip.getEntry(sharedStringsEntry))
    val outputSst = readEntryBytes(outputZip, outputZip.getEntry(sharedStringsEntry))
    assertEquals(outputSst.toSeq, sourceSst.toSeq, "sharedStrings.xml should be preserved byte-for-byte")

    val sheetXmlBytes = readEntryBytes(outputZip, outputZip.getEntry("xl/worksheets/sheet1.xml"))
    val sheetXml = new String(sheetXmlBytes, "UTF-8")

    // Modified sheet should use SST references for existing strings, inline for new strings
    assert(sheetXml.contains("""t="s""""), "Existing cells should use SST references (t=\"s\")")
    // New cell "Updated Text" may use inline if not in original SST
    assert(sheetXml.contains("Updated Text"), "Modified cell should be present")

    sourceZip.close()
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // Helper: Read entry bytes from ZIP
  private def readEntryBytes(zip: ZipFile, entry: ZipEntry): Array[Byte] =
    val is = zip.getInputStream(entry)
    try is.readAllBytes()
    finally is.close()

  // Helper: Create minimal workbook with chart (unknown part)
  private def createMinimalWorkbookWithChart(): Path =
    val path = Files.createTempFile("test-chart", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))

    try
      // Content Types
      writeEntry(
        out,
        "[Content_Types].xml",
        """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>"""
      )

      // Root rels
      writeEntry(
        out,
        "_rels/.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
      )

      // Workbook
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

      // Workbook rels
      writeEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
      )

      // Worksheet
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Test</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

      // Chart (unknown part)
      writeEntry(
        out,
        "xl/charts/chart1.xml",
        """<?xml version="1.0"?>
<chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <chart>
    <title><tx><rich><p><r><t>Sales Chart</t></r></p></rich></tx></title>
    <plotArea/>
  </chart>
</chartSpace>"""
      )

    finally out.close()

    path

  // Helper: Create workbook with 2 sheets
  private def createWorkbookWith2Sheets(): Path =
    val path = Files.createTempFile("test-2-sheets", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))

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
</Relationships>"""
      )

      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Sheet1 Data</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

      writeEntry(
        out,
        "xl/worksheets/sheet2.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="B1" t="inlineStr"><is><t>Sheet2 Data</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

    finally out.close()

    path

  // Helper: Create workbook with unknown part
  private def createWorkbookWithUnknownPart(): Path =
    val path = Files.createTempFile("test-unknown", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))

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
</worksheet>"""
      )

      // Unknown part (custom data)
      writeEntry(
        out,
        "xl/custom/data.xml",
        """<?xml version="1.0"?>
<customData>
  <metadata>This is custom data that XL doesn't understand</metadata>
</customData>"""
      )

    finally out.close()

    path

  private def createWorkbookWithSharedStrings(): Path =
    val path = Files.createTempFile("test-sst", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))

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
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""
      )

      writeEntry(
        out,
        "xl/sharedStrings.xml",
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Original Value</t></si>
  <si><t>Preserved</t></si>
</sst>"""
      )

      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"""
      )

    finally out.close()

    path

  // Helper: Write ZIP entry
  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    val entry = new ZipEntry(name)
    out.putNextEntry(entry)
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()
