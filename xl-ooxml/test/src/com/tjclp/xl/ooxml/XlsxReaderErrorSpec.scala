package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.macros.ref
import com.tjclp.xl.cell.{CellError, CellValue}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.api.Workbook
import java.io.ByteArrayOutputStream
import java.util.zip.{ZipEntry, ZipOutputStream}
import java.nio.charset.StandardCharsets

/**
 * Regression tests for error handling paths inside XlsxReader.
 *
 * These cover malformed XML, missing parts, and corrupted ZIP archives to
 * ensure we surface precise XLErrors instead of throwing.
 */
class XlsxReaderErrorSpec extends FunSuite:

  private val contentTypesXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""

  private val rootRelationshipsXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

  private val workbookRelationshipsXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""

  private val validWorkbookXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""

  private val validWorksheetXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is><t>Hello</t></is>
      </c>
    </row>
  </sheetData>
</worksheet>"""

  private val sharedStringsXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Only</t></si>
</sst>"""

  private val minimalStylesXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

  private val baseParts: Map[String, String] = Map(
    "[Content_Types].xml" -> contentTypesXml,
    "_rels/.rels" -> rootRelationshipsXml,
    "xl/workbook.xml" -> validWorkbookXml,
    "xl/_rels/workbook.xml.rels" -> workbookRelationshipsXml,
    "xl/worksheets/sheet1.xml" -> validWorksheetXml,
    "xl/styles.xml" -> minimalStylesXml
  )

  test("XlsxReader rejects XLSX missing workbook.xml") {
    val bytes = buildWorkbook(omit = Set("xl/workbook.xml"))
    assertParseError(XlsxReader.readFromBytes(bytes), "xl/workbook.xml", "Missing workbook.xml")
  }

  test("XlsxReader rejects malformed workbook.xml") {
    val malformed = "<workbook><sheets></workbook>" // malformed closing tags
    val bytes = buildWorkbook(overrides = Map("xl/workbook.xml" -> malformed))

    XlsxReader.readFromBytes(bytes) match
      case Left(XLError.ParseError(location, message)) =>
        assertEquals(location, "xl/workbook.xml")
        assert(
          message.toLowerCase.contains("xml parse"),
          s"Expected XML parse error message, got: $message"
        )
      case other => fail(s"Expected ParseError for malformed workbook, got $other")
  }

  test("XlsxReader rejects workbook.xml missing required sheets element") {
    val invalidWorkbook =
      """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
</workbook>"""
    val bytes = buildWorkbook(overrides = Map("xl/workbook.xml" -> invalidWorkbook))
    XlsxReader.readFromBytes(bytes) match
      case Left(XLError.ParseError(location, message)) =>
        assertEquals(location, "xl/workbook.xml")
        assert(
          message.contains("Missing required child element: sheets"),
          s"Expected missing sheets error, got: $message"
        )
      case other => fail(s"Expected ParseError for missing sheets node, got $other")
  }

  test("XlsxReader reports styles parse errors with precise location") {
    val bytes = buildWorkbook(
      overrides = Map("xl/styles.xml" -> "<styleSheet><fonts></styleSheet>")
    )
    XlsxReader.readFromBytes(bytes) match
      case Left(XLError.ParseError(location, message)) =>
        assertEquals(location, "xl/styles.xml")
        assert(
          message.toLowerCase.contains("xml parse"),
          s"Expected XML parse error, got: $message"
        )
      case other => fail(s"Expected ParseError for malformed styles, got $other")
  }

  test("XlsxReader errors when worksheet part referenced by workbook is missing") {
    val bytes = buildWorkbook(omit = Set("xl/worksheets/sheet1.xml"))
    assertParseError(
      XlsxReader.readFromBytes(bytes),
      "xl/worksheets/sheet1.xml",
      "Missing worksheet: xl/worksheets/sheet1.xml"
    )
  }

  test("XlsxReader emits warning when styles.xml missing") {
    val bytes = buildWorkbook(omit = Set("xl/styles.xml"))
    val result = XlsxReader.readFromBytesWithWarnings(bytes).getOrElse(fail("Should read"))
    assertEquals(result.warnings, Vector(XlsxReader.Warning.MissingStylesXml))
  }

  test("XlsxReader errors when workbook relationship points to missing target") {
    val rels =
      """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/ghost.xml"/>
</Relationships>"""
    val bytes = buildWorkbook(
      overrides = Map("xl/_rels/workbook.xml.rels" -> rels)
    )
    assertParseError(
      XlsxReader.readFromBytes(bytes),
      "xl/worksheets/ghost.xml",
      "Missing worksheet: xl/worksheets/ghost.xml"
    )
  }

  test("XlsxReader gracefully handles workbooks without [Content_Types].xml") {
    val bytes = buildWorkbook(omit = Set("[Content_Types].xml"))

    val workbook = XlsxReader.readFromBytes(bytes).getOrElse(fail("Workbook should parse"))
    val sheet = workbook("Sheet1").getOrElse(fail("Expected Sheet1"))
    assertEquals(sheet(ref"A1").value, com.tjclp.xl.cell.CellValue.Text("Hello"))
  }

  test("XlsxReader rejects non-ZIP input instead of silently succeeding") {
    val bytes = "not-a-zip".getBytes(StandardCharsets.UTF_8)
    XlsxReader.readFromBytes(bytes) match
      case Left(XLError.ParseError(location, message)) =>
        assertEquals(location, "xl/workbook.xml")
        assertEquals(message, "Missing workbook.xml")
      case Left(XLError.IOError(_)) =>
        () // acceptable alternate failure mode
      case other => fail(s"Expected failure for non-zip input, got $other")
  }

  test("XlsxReader rejects truncated ZIP archives") {
    val goodBytes = buildWorkbook()
    val truncated = goodBytes.take(goodBytes.length / 2)

    XlsxReader.readFromBytes(truncated) match
      case Left(XLError.IOError(reason)) =>
        assert(
          reason.toLowerCase.contains("failed to read bytes"),
          s"Expected IO failure, got: $reason"
        )
      case other => fail(s"Expected IOError for truncated archive, got $other")
  }

  test("XlsxReader reports CellError when shared string index is invalid") {
    val sheetWithBadSst =
      """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>5</v></c>
    </row>
  </sheetData>
</worksheet>"""

    val bytes = buildWorkbook(
      overrides = Map(
        "xl/worksheets/sheet1.xml" -> sheetWithBadSst,
        "xl/sharedStrings.xml" -> sharedStringsXml
      )
    )

    val workbook = XlsxReader.readFromBytes(bytes).getOrElse(fail("Workbook should parse"))
    val sheet = workbook("Sheet1").getOrElse(fail("Expected Sheet1"))
    assertEquals(sheet(ref"A1").value, CellValue.Error(CellError.Ref))
  }

  private def buildWorkbook(
    overrides: Map[String, String] = Map.empty,
    omit: Set[String] = Set.empty
  ): Array[Byte] =
    val finalParts = (baseParts ++ overrides).filterNot { case (name, _) => omit.contains(name) }
    writeZip(finalParts)

  private def writeZip(entries: Map[String, String]): Array[Byte] =
    val baos = ByteArrayOutputStream()
    val zos = ZipOutputStream(baos)
    entries.foreach { case (name, content) =>
      zos.putNextEntry(ZipEntry(name))
      zos.write(content.getBytes(StandardCharsets.UTF_8))
      zos.closeEntry()
    }
    zos.close()
    baos.toByteArray

  private def assertParseError(
    result: Either[XLError, Workbook],
    expectedLocation: String,
    expectedMessage: String
  ): Unit =
    result match
      case Left(XLError.ParseError(location, message)) =>
        assertEquals(location, expectedLocation)
        assertEquals(message, expectedMessage)
      case other =>
        fail(s"Expected ParseError at $expectedLocation with '$expectedMessage', got $other")
