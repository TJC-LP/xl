package com.tjclp.xl.ooxml

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipOutputStream}

import com.tjclp.xl.api.*
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * Tests for XlsxReader surgical modification support (Phase 3).
 *
 * Verifies that:
 *   1. Reading from file creates SourceContext with indexed unknown parts
 *   2. Reading from stream/bytes does NOT create SourceContext (backwards compat)
 *   3. Unknown parts are indexed but NOT loaded into memory
 *   4. PartManifest has accurate metadata
 */
class XlsxReaderPassthroughSpec extends FunSuite:

  test("read from file creates SourceContext") {
    // Create minimal XLSX with one known part and one unknown part
    val path = createMinimalWorkbookWithChart()

    val wb = XlsxReader
      .read(path)
      .fold(err => fail(s"Failed to read workbook: $err"), identity)

    // Should have SourceContext since read from file
    assert(wb.sourceContext.isDefined, "Reading from file should create SourceContext")

    val ctx = wb.sourceContext.get
    assertEquals(ctx.sourcePath, path)
    assert(ctx.isClean, "Freshly read workbook should have clean tracker")

    // PartManifest should contain both known and unknown parts
    val manifest = ctx.partManifest
    assert(manifest.parsedParts.contains("xl/workbook.xml"))
    assert(manifest.unparsedParts.contains("xl/charts/chart1.xml"))

    // Clean up
    Files.deleteIfExists(path)
  }

  test("read from bytes does NOT create SourceContext") {
    val bytes = createMinimalWorkbookBytes()

    val wb = XlsxReader
      .readFromBytes(bytes)
      .fold(err => fail(s"Failed to read workbook: $err"), identity)

    // Should NOT have SourceContext (no file path available)
    assert(wb.sourceContext.isEmpty, "Reading from bytes should not create SourceContext")
  }

  test("unknown parts are indexed but not loaded") {
    // Create workbook with a large unknown part (chart XML)
    val largeChart = "<chart>" + ("x" * 1024 * 1024) + "</chart>" // 1MB chart
    val path = createWorkbookWithLargeChart(largeChart)

    val wb = XlsxReader
      .read(path)
      .fold(err => fail(s"Failed to read workbook: $err"), identity)

    // Verify SourceContext exists
    assert(wb.sourceContext.isDefined)

    val manifest = wb.sourceContext.get.partManifest

    // Chart should be in unparsed parts (indexed but not loaded)
    assert(manifest.unparsedParts.contains("xl/charts/chart1.xml"))

    // Verify metadata captured
    val chartEntry = manifest.entries.get("xl/charts/chart1.xml")
    assert(chartEntry.isDefined, "Chart entry should be in manifest")
    assert(!chartEntry.get.parsed, "Chart should NOT be marked as parsed")

    // Note: size may not be available from ZipInputStream (returns -1 for DEFLATED entries)
    // This is expected behavior - size is optional metadata

    // Clean up
    Files.deleteIfExists(path)
  }

  test("PartManifest distinguishes parsed vs unparsed parts") {
    val path = createMinimalWorkbookWithChart()

    val wb = XlsxReader
      .read(path)
      .fold(err => fail(s"Failed to read workbook: $err"), identity)

    val manifest = wb.sourceContext.get.partManifest

    // Parsed parts (known to XL)
    val expectedParsed = Set(
      "xl/workbook.xml",
      "xl/worksheets/sheet1.xml",
      "xl/_rels/workbook.xml.rels",
      "_rels/.rels",
      "[Content_Types].xml"
    )
    assert(expectedParsed.subsetOf(manifest.parsedParts))

    // Unparsed parts (unknown to XL - should be preserved)
    assert(manifest.unparsedParts.contains("xl/charts/chart1.xml"))

    // Clean up
    Files.deleteIfExists(path)
  }

  test("backwards compatibility: existing tests using readFromBytes work unchanged") {
    // This test verifies that existing code that uses readFromBytes continues to work
    val bytes = createMinimalWorkbookBytes()

    val result = XlsxReader.readFromBytes(bytes)

    assert(result.isRight, "readFromBytes should succeed")
    val wb = result.getOrElse(fail("Expected Right"))
    assert(wb.sheets.nonEmpty, "Workbook should have sheets")
    assert(wb.sourceContext.isEmpty, "Bytes-based read should not create SourceContext")
  }

  // Helper: Create minimal XLSX with one worksheet and one chart (unknown part)
  private def createMinimalWorkbookWithChart(): Path =
    val path = Files.createTempFile("test-workbook-chart", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))

    try
      // Content Types
      writeEntry(out, "[Content_Types].xml", """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>""")

      // Root rels
      writeEntry(out, "_rels/.rels", """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""")

      // Workbook
      writeEntry(out, "xl/workbook.xml", """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>""")

      // Workbook rels
      writeEntry(out, "xl/_rels/workbook.xml.rels", """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""")

      // Worksheet
      writeEntry(out, "xl/worksheets/sheet1.xml", """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Test</t></is></c>
    </row>
  </sheetData>
</worksheet>""")

      // Chart (unknown part - should be preserved)
      writeEntry(out, "xl/charts/chart1.xml", """<?xml version="1.0"?>
<chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <chart>
    <plotArea/>
  </chart>
</chartSpace>""")

    finally out.close()

    path

  // Helper: Create minimal XLSX bytes without unknown parts
  private def createMinimalWorkbookBytes(): Array[Byte] =
    val baos = new ByteArrayOutputStream()
    val out = new ZipOutputStream(baos)

    try
      // Content Types
      writeEntry(out, "[Content_Types].xml", """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>""")

      // Root rels
      writeEntry(out, "_rels/.rels", """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""")

      // Workbook
      writeEntry(out, "xl/workbook.xml", """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>""")

      // Workbook rels
      writeEntry(out, "xl/_rels/workbook.xml.rels", """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""")

      // Worksheet
      writeEntry(out, "xl/worksheets/sheet1.xml", """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Test</t></is></c>
    </row>
  </sheetData>
</worksheet>""")

    finally out.close()

    baos.toByteArray

  // Helper: Create workbook with large chart
  private def createWorkbookWithLargeChart(chartContent: String): Path =
    val path = Files.createTempFile("test-large-chart", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))

    try
      // Same structure as createMinimalWorkbookWithChart but with custom chart content
      writeEntry(out, "[Content_Types].xml", """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>""")

      writeEntry(out, "_rels/.rels", """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""")

      writeEntry(out, "xl/workbook.xml", """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>""")

      writeEntry(out, "xl/_rels/workbook.xml.rels", """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""")

      writeEntry(out, "xl/worksheets/sheet1.xml", """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Test</t></is></c>
    </row>
  </sheetData>
</worksheet>""")

      // Large chart (unknown part)
      writeEntry(out, "xl/charts/chart1.xml", chartContent)

    finally out.close()

    path

  // Helper: Write ZIP entry
  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    val entry = new ZipEntry(name)
    out.putNextEntry(entry)
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()
