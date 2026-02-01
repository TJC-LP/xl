package com.tjclp.xl.ooxml

import java.io.ByteArrayInputStream
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
import javax.xml.parsers.DocumentBuilderFactory

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.styleSyntax.withCellStyle
import com.tjclp.xl.styles.{CellStyle, Color, Fill}
import munit.FunSuite

/**
 * Regression tests for surgical modification corruption fixes.
 *
 * These tests prevent regression of 26 critical fixes made during surgical
 * modification implementation (feat/surgical-mod branch).
 *
 * **Primary Corruption Triggers (P0)**:
 *   1. Missing mc:Ignorable attribute → Excel corruption warning
 *   2. Theme colors converted to black RGB → Visual corruption
 *   3. Row attributes lost → 6% data loss
 *   4. SST binary garbage → String corruption
 *
 * Corresponds to commits:
 *   - 4bfb9f4: Namespace preservation
 *   - a40fd02: Theme color preservation
 *   - e1f36fe: Row attribute preservation
 *   - 802e020: SST handling (9 fixes)
 */
// Test code uses .get/.head for brevity in assertions
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class XlsxWriterCorruptionRegressionSpec extends FunSuite:

  test("preserves mc:Ignorable namespace attribute on workbook root (PRIMARY corruption fix)") {
    // Regression test for commit 4bfb9f4
    // Bug: Missing mc:Ignorable caused PRIMARY Excel corruption warning
    // Root cause: Workbook.toXml didn't preserve namespace metadata from source

    val source = createWorkbookWithNamespaces()

    // Read and modify
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet = sheet.put(ref"A1" -> "Modified")
    yield wb.put(updatedSheet)

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    // Write back
    val output = Files.createTempFile("namespace-preserve", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify workbook.xml has mc:Ignorable attribute
    val outputZip = new ZipFile(output.toFile)
    val workbookXml = readEntryString(outputZip, outputZip.getEntry("xl/workbook.xml"))

    // Parse XML to check root attributes
    val doc = parseXml(workbookXml)
    val rootElem = doc.getDocumentElement

    // CRITICAL: mc:Ignorable must be present (tells Excel to ignore unknown namespaces)
    val mcIgnorable = rootElem.getAttributeNS("http://schemas.openxmlformats.org/markup-compatibility/2006", "Ignorable")
    assert(
      mcIgnorable != null && mcIgnorable.nonEmpty,
      s"mc:Ignorable attribute missing - PRIMARY corruption trigger! Found: '$mcIgnorable'"
    )

    // Should contain typical values like "x15 xr xr6 xr10 xr2"
    assert(
      mcIgnorable.contains("x15") || mcIgnorable.contains("xr"),
      s"mc:Ignorable should reference forward-compatibility namespaces, found: '$mcIgnorable'"
    )

    // Verify namespace pollution is prevented (child elements shouldn't re-declare namespaces)
    val sheetsCount = workbookXml.split("<sheets").length - 1
    val xmlnsCount = workbookXml.split("xmlns=").length - 1

    // Should have xmlns only on root (or very few declarations), not on every child
    assert(
      xmlnsCount <= 12,
      s"Namespace pollution detected: $xmlnsCount xmlns declarations (should be ~8-12 on root only)"
    )

    // Clean up
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  test("preserves theme colors in fills and borders (prevents black RGB corruption)") {
    // Regression test for commit a40fd02
    // Bug: Theme colors (theme="0" tint="-0.05") converted to rgb="00000000" (black)
    // Impact: MAJOR visual corruption - cells showed black fills instead of white/light backgrounds

    val source = createWorkbookWithThemeColors()

    // Read and modify
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet = sheet.put(ref"A1" -> "Modified")
    yield wb.put(updatedSheet)

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    // Write back
    val output = Files.createTempFile("theme-preserve", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify styles.xml preserves theme colors
    val outputZip = new ZipFile(output.toFile)
    val stylesXml = readEntryString(outputZip, outputZip.getEntry("xl/styles.xml"))

    // CRITICAL: Should have theme="0" attributes, NOT rgb="00000000"
    assert(
      stylesXml.contains("theme=\"0\"") || stylesXml.contains("theme=\"1\""),
      "Theme color attributes missing - causes visual corruption (black fills/borders)"
    )

    // Should NOT have black RGB as replacement for theme colors
    val blackRgbCount = stylesXml.split("rgb=\"00000000\"").length - 1
    assert(
      blackRgbCount == 0,
      s"Found $blackRgbCount instances of rgb=\"00000000\" - indicates theme colors were incorrectly converted to black"
    )

    // Should have tint attributes for theme variations
    assert(
      stylesXml.contains("tint="),
      "Tint attributes missing - theme color variations not preserved"
    )

    // Clean up
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  test("preserves all 10 row attributes during regeneration (prevents 6% data loss)") {
    // Regression test for commit e1f36fe
    // Bug: Row attributes lost during regeneration (spans, style, height, etc.)
    // Impact: 2,170 bytes data loss, 6% of row metadata

    val source = createWorkbookWithRichRowAttributes()

    // Read and modify one cell
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet = sheet.put(ref"A3" -> "Modified")
    yield wb.put(updatedSheet)

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    // Write back
    val output = Files.createTempFile("row-attrs", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify sheet1.xml preserves row attributes
    val outputZip = new ZipFile(output.toFile)
    val sheetXml = readEntryString(outputZip, outputZip.getEntry("xl/worksheets/sheet1.xml"))

    // Row 1: Empty row with formatting (was getting dropped)
    // Note: Attribute order may vary between backends but row is preserved
    assert(sheetXml.contains("""r="1"""") && sheetXml.contains("<row"), "Empty row 1 should be preserved")

    // Row 2: All 10 attributes
    assert(sheetXml.contains("""spans="1:5""""), "Row 'spans' attribute missing")
    assert(sheetXml.contains("""s="1""""), "Row 's' (style) attribute missing")
    assert(sheetXml.contains("""customHeight="1""""), "Row 'customHeight' attribute missing")
    assert(sheetXml.contains("""ht=""""), "Row 'ht' (height) attribute missing")

    // Namespaced attributes (these were completely lost before fix)
    assert(
      sheetXml.contains("dyDescent="),
      "Row 'x14ac:dyDescent' namespaced attribute missing (affects vertical alignment)"
    )

    // Row 3: thickBot attribute
    assert(sheetXml.contains("""thickBot="1""""), "Row 'thickBot' attribute missing")

    // Verify modified cell is present
    assert(sheetXml.contains("Modified"), "Modified cell content missing")

    // Clean up
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  test("SST count vs uniqueCount correctness and RichText SST references") {
    // Regression test for commit 802e020 (9 fixes)
    // Bugs:
    //   - count=uniqueCount (incorrect, should be total refs vs unique strings)
    //   - RichText inlined instead of using SST references
    //   - Binary garbage from system sheets in SST

    val source = createWorkbookWithRichTextSST()

    // Read and modify
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet = sheet.put(ref"A1" -> "New String")
    yield wb.put(updatedSheet)

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    // Write back
    val output = Files.createTempFile("sst-correct", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify sharedStrings.xml has correct counts
    val outputZip = new ZipFile(output.toFile)
    val sstXml = readEntryString(outputZip, outputZip.getEntry("xl/sharedStrings.xml"))

    // Parse SST root element
    val doc = parseXml(sstXml)
    val sstElem = doc.getDocumentElement

    val countAttr = sstElem.getAttribute("count")
    val uniqueCountAttr = sstElem.getAttribute("uniqueCount")

    assert(countAttr.nonEmpty, "SST 'count' attribute missing")
    assert(uniqueCountAttr.nonEmpty, "SST 'uniqueCount' attribute missing")

    val count = countAttr.toInt
    val uniqueCount = uniqueCountAttr.toInt

    // Count (total references) should be >= uniqueCount (unique strings)
    assert(
      count >= uniqueCount,
      s"SST count=$count should be >= uniqueCount=$uniqueCount (count = total refs, uniqueCount = unique strings)"
    )

    // Verify RichText in SST (not inlined)
    val sheetXml = readEntryString(outputZip, outputZip.getEntry("xl/worksheets/sheet1.xml"))

    // Cell B1 (RichText) should use t="s" (SST reference), not t="inlineStr"
    val b1Cell = sheetXml.split("""<c r="B1"""").drop(1).headOption.getOrElse("")
    assert(
      b1Cell.contains("""t="s""""),
      "RichText cell B1 should use SST reference (t=\"s\"), not inline (prevents corruption)"
    )

    // Should NOT have inlineStr for RichText cells
    assert(
      !sstXml.contains("inlineStr"),
      "SST should not contain 'inlineStr' (RichText should be in SST, not inlined)"
    )

    // Verify no binary garbage in SST (check for reasonable string content)
    val siCount = sstXml.split("<si>").length - 1
    assert(
      siCount == uniqueCount,
      s"SST has $siCount <si> entries but uniqueCount=$uniqueCount (mismatch indicates corruption)"
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

  private def createWorkbookWithNamespaces(): Path =
    val path = Files.createTempFile("test-namespaces", ".xlsx")
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

      // Workbook with full namespace declarations including mc:Ignorable
      writeEntry(
        out,
        "xl/workbook.xml",
        """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" mc:Ignorable="x15 xr xr6 xr10">
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
      <c r="A1" t="inlineStr"><is><t>Test</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

    finally out.close()

    path

  private def createWorkbookWithThemeColors(): Path =
    val path = Files.createTempFile("test-theme", ".xlsx")
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

      // Styles with theme colors (theme="0" tint="-0.05" for light background)
      writeEntry(
        out,
        "xl/styles.xml",
        """<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
    <font><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor theme="0" tint="-0.04999" /><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color theme="1" tint="0.5"/></left><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="1" applyFill="1" applyBorder="1"/>
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
      <c r="A1" s="1" t="inlineStr"><is><t>Themed Cell</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

    finally out.close()

    path

  private def createWorkbookWithRichRowAttributes(): Path =
    val path = Files.createTempFile("test-row-attrs", ".xlsx")
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

      writeEntry(
        out,
        "xl/styles.xml",
        """<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="0" fillId="1" borderId="0" applyFill="1"/>
  </cellXfs>
</styleSheet>"""
      )

      // Worksheet with rich row attributes
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetData>
    <row r="1" spans="1:5" s="1" customFormat="1" ht="20.5" customHeight="1" x14ac:dyDescent="0.25"/>
    <row r="2" spans="1:5" s="1" customFormat="1" ht="15.0" customHeight="1" x14ac:dyDescent="0.3">
      <c r="A2" t="inlineStr"><is><t>Row 2 Data</t></is></c>
    </row>
    <row r="3" thickBot="1" x14ac:dyDescent="0.2">
      <c r="A3" t="inlineStr"><is><t>Row 3 Data</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

    finally out.close()

    path

  private def createWorkbookWithRichTextSST(): Path =
    val path = Files.createTempFile("test-richtext-sst", ".xlsx")
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

      // SST with 3 unique strings, 5 total references (count != uniqueCount)
      writeEntry(
        out,
        "xl/sharedStrings.xml",
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="3">
  <si><t>Plain Text</t></si>
  <si>
    <r><rPr><b/><color rgb="FFFF0000"/></rPr><t>Bold</t></r>
    <r><t xml:space="preserve"> </t></r>
    <r><rPr><i/></rPr><t>Italic</t></r>
  </si>
  <si><t>Repeated String</t></si>
</sst>"""
      )

      // Worksheet with SST references (count > uniqueCount due to repeats)
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"""
      )

    finally out.close()

    path

  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    val entry = new ZipEntry(name)
    out.putNextEntry(entry)
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()

  // ===== New Regression Tests for v0.1.5 Corruption Fixes =====

  test("styles.xml has no negative fontId/fillId/borderIdx (prevents OOXML schema violation)") {
    // Regression test for: fontMap.getOrElse(style.font, -1) -> 0
    // Bug: Styles with unrecognized fonts wrote fontId="-1" which is invalid OOXML
    // Impact: Excel shows repair dialog, styles render incorrectly

    // Create a workbook with custom styles
    import com.tjclp.xl.styles.{CellStyle, Font, Fill}
    import com.tjclp.xl.styles.color.Color

    val boldRedStyle = CellStyle.default
      .withFont(Font("Arial", 12.0, bold = true))
      .withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))

    val sheet = Sheet("Test")
      .put(ref"A1" -> "Styled")
      .withCellStyle(ref"A1", boldRedStyle)

    val wb = Workbook(Vector(sheet))

    val output = Files.createTempFile("styles-no-negative", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify styles.xml has no negative indices
    val outputZip = new ZipFile(output.toFile)
    val stylesXml = readEntryString(outputZip, outputZip.getEntry("xl/styles.xml"))

    // CRITICAL: Must not have negative indices
    assert(
      !stylesXml.contains("fontId=\"-"),
      "fontId contains negative value - OOXML violation!"
    )
    assert(
      !stylesXml.contains("fillId=\"-"),
      "fillId contains negative value - OOXML violation!"
    )
    assert(
      !stylesXml.contains("borderId=\"-"),
      "borderId contains negative value - OOXML violation!"
    )

    // Clean up
    outputZip.close()
    Files.deleteIfExists(output)
  }

  test("styles.xml includes required cellStyleXfs and cellStyles elements (ECMA-376 compliance)") {
    // Regression test for missing required OOXML elements
    // Bug: Generated styles.xml omitted <cellStyleXfs> and <cellStyles>
    // Impact: Excel repair dialog, strict OOXML validators fail

    val sheet = Sheet("Test").put(ref"A1" -> "Data")
    val wb = Workbook(Vector(sheet))

    val output = Files.createTempFile("styles-required-elements", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify styles.xml has required elements
    val outputZip = new ZipFile(output.toFile)
    val stylesXml = readEntryString(outputZip, outputZip.getEntry("xl/styles.xml"))

    // REQUIRED: cellStyleXfs (master cell formatting records)
    assert(
      stylesXml.contains("<cellStyleXfs"),
      "Missing <cellStyleXfs> element - required per ECMA-376 §18.8.9"
    )

    // REQUIRED: cellStyles (named styles, at least "Normal")
    assert(
      stylesXml.contains("<cellStyles"),
      "Missing <cellStyles> element - required per ECMA-376 §18.8.8"
    )
    assert(
      stylesXml.contains("name=\"Normal\""),
      "Missing default 'Normal' cellStyle"
    )

    // Verify correct OOXML element order: cellStyleXfs before cellXfs, cellStyles after cellXfs
    val cellStyleXfsPos = stylesXml.indexOf("<cellStyleXfs")
    val cellXfsPos = stylesXml.indexOf("<cellXfs")
    val cellStylesPos = stylesXml.indexOf("<cellStyles")

    assert(
      cellStyleXfsPos < cellXfsPos,
      "OOXML order violation: cellStyleXfs must come before cellXfs"
    )
    assert(
      cellXfsPos < cellStylesPos,
      "OOXML order violation: cellXfs must come before cellStyles"
    )

    // Clean up
    outputZip.close()
    Files.deleteIfExists(output)
  }

  test("theme1.xml is preserved from source during surgical modification") {
    // Regression test for missing theme file
    // Bug: xl/theme/theme1.xml was parsed but never written back
    // Impact: Excel shows corruption warning, theme colors unresolvable

    val source = createWorkbookWithThemeFile()

    // Read and modify
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet = sheet.put(ref"A2" -> "New Data")
    yield wb.put(updatedSheet)

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    // Write back
    val output = Files.createTempFile("theme-preserved", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // CRITICAL: theme1.xml must be preserved
    val outputZip = new ZipFile(output.toFile)
    val themeEntry = outputZip.getEntry("xl/theme/theme1.xml")

    assert(
      themeEntry != null,
      "xl/theme/theme1.xml missing - causes Excel corruption warning!"
    )

    // Verify theme content is valid XML
    val themeXml = readEntryString(outputZip, themeEntry)
    assert(themeXml.contains("<a:clrScheme"), "Theme file missing color scheme")

    // Clean up
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  test("dimension element reflects actual cell range after put operations") {
    // Regression test for dimension not updated after cell modifications
    // Bug: preserved.dimension used directly without recalculating
    // Impact: Excel may not recognize all data rows

    val source = createWorkbookWithDimension()

    // Read and add cells OUTSIDE original dimension
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      // Original dimension is A1:C3, add cell at E5
      updatedSheet = sheet.put(ref"E5" -> "Extended")
    yield wb.put(updatedSheet)

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    // Write back
    val output = Files.createTempFile("dimension-updated", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify dimension was updated to include new cell
    val outputZip = new ZipFile(output.toFile)
    val sheetXml = readEntryString(outputZip, outputZip.getEntry("xl/worksheets/sheet1.xml"))

    // Extract dimension ref attribute
    val dimMatch = """<dimension ref="([^"]+)"""".r.findFirstMatchIn(sheetXml)
    assert(dimMatch.isDefined, "dimension element missing")

    val dimRef = dimMatch.get.group(1)
    // Should now extend to at least E5
    assert(
      dimRef.contains("E") && dimRef.contains("5"),
      s"Dimension '$dimRef' should include E5 after put operation"
    )

    // Verify the actual cell E5 exists
    assert(sheetXml.contains("""r="E5""""), "Cell E5 missing from sheetData")
    assert(sheetXml.contains("Extended"), "Cell E5 content missing")

    // Clean up
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ===== Additional Test File Creators =====

  private def createWorkbookWithThemeFile(): Path =
    val path = Files.createTempFile("test-with-theme", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))
    out.setLevel(1)

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
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"""
      )

      // Minimal theme file
      writeEntry(
        out,
        "xl/theme/theme1.xml",
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>"""
      )

      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Original Data</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

    finally out.close()

    path

  private def createWorkbookWithDimension(): Path =
    val path = Files.createTempFile("test-with-dimension", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))
    out.setLevel(1)

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

      // Worksheet with explicit dimension A1:C3
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C3"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>A1</t></is></c>
      <c r="B1" t="inlineStr"><is><t>B1</t></is></c>
      <c r="C1" t="inlineStr"><is><t>C1</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>A2</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>A3</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

    finally out.close()

    path

  // ===== Regression Test for SharedStrings Metadata Bug (Jan 2026) =====

  test("sharedStrings.xml is properly referenced in Content_Types and workbook.xml.rels when generated") {
    // Regression test for: SharedStrings metadata missing when SST generated from preserved source
    // Bug: When batch operations added enough strings to trigger SST generation (SstPolicy.Auto),
    //      the preserved Content_Types.xml and workbook.xml.rels from source didn't include
    //      sharedStrings references, even though sharedStrings.xml was written to the ZIP.
    // Impact: Excel shows "We found a problem with some content" corruption warning

    // Create a source workbook WITHOUT sharedStrings (uses inline strings)
    val source = createWorkbookWithoutSST()

    // Read the workbook
    val wb1 = XlsxReader.read(source).fold(err => fail(s"Failed to read: $err"), identity)

    // Add strings with duplicates to trigger SST generation
    // SST heuristic: totalCells > uniqueCount && totalCells > 10
    // So we need more than 10 cells with some duplicates
    val sheet = wb1.sheets.head
    val updatedSheet = sheet
      // 5 unique strings repeated across 15 cells = will trigger SST
      .put(ref"A1" -> "Header A", ref"A2" -> "Header B", ref"A3" -> "Header C")
      .put(ref"A4" -> "Data Value", ref"A5" -> "Data Value", ref"A6" -> "Data Value")
      .put(ref"A7" -> "Header A", ref"A8" -> "Header B", ref"A9" -> "Header C")
      .put(ref"A10" -> "Data Value", ref"A11" -> "Data Value", ref"A12" -> "Total")
      .put(ref"A13" -> "Header A", ref"A14" -> "Total", ref"A15" -> "Data Value")
    val wb2 = wb1.put(updatedSheet)

    // Write the modified workbook
    val output = Files.createTempFile("sst-metadata-regression", ".xlsx")
    XlsxWriter
      .write(wb2, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // Verify the output has sharedStrings.xml
    val outputZip = new ZipFile(output.toFile)
    val sstEntry = outputZip.getEntry("xl/sharedStrings.xml")
    assert(sstEntry != null, "xl/sharedStrings.xml should exist when SST is generated")

    // CRITICAL: Verify Content_Types.xml references sharedStrings
    val contentTypesXml = readEntryString(outputZip, outputZip.getEntry("[Content_Types].xml"))
    assert(
      contentTypesXml.contains("/xl/sharedStrings.xml"),
      "Content_Types.xml must reference /xl/sharedStrings.xml when SST is present - causes Excel corruption!"
    )
    assert(
      contentTypesXml.contains("sharedStrings+xml"),
      "Content_Types.xml must have sharedStrings content type"
    )

    // CRITICAL: Verify workbook.xml.rels references sharedStrings
    val workbookRelsXml = readEntryString(outputZip, outputZip.getEntry("xl/_rels/workbook.xml.rels"))
    assert(
      workbookRelsXml.contains("sharedStrings.xml"),
      "workbook.xml.rels must reference sharedStrings.xml when SST is present - causes Excel corruption!"
    )
    assert(
      workbookRelsXml.contains("relationships/sharedStrings"),
      "workbook.xml.rels must have sharedStrings relationship type"
    )

    // Clean up
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  private def createWorkbookWithoutSST(): Path =
    // Create a minimal workbook that uses inline strings (no sharedStrings.xml)
    val path = Files.createTempFile("test-no-sst", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))
    out.setLevel(1)

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

      writeEntry(
        out,
        "xl/styles.xml",
        """<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""
      )

      // Worksheet with NO sharedStrings - uses inline strings
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Original</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

    finally out.close()

    path
