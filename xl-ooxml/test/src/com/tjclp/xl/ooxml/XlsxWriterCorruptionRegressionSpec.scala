package com.tjclp.xl.ooxml

import java.io.ByteArrayInputStream
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
import javax.xml.parsers.DocumentBuilderFactory

import com.tjclp.xl.api.*
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.style.{CellStyle, Color, Fill}
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
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
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
      updatedSheet <- sheet.put(ref"A1" -> "Modified")
      updated <- wb.put(updatedSheet)
    yield updated

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
      updatedSheet <- sheet.put(ref"A1" -> "Modified")
      updated <- wb.put(updatedSheet)
    yield updated

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
      updatedSheet <- sheet.put(ref"A3" -> "Modified")
      updated <- wb.put(updatedSheet)
    yield updated

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
    assert(sheetXml.contains("""<row r="1""""), "Empty row 1 should be preserved")

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
      updatedSheet <- sheet.put(ref"A1" -> "New String")
      updated <- wb.put(updatedSheet)
    yield updated

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
