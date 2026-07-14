package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.macros.ref
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.api.Workbook
import java.io.ByteArrayOutputStream
import java.util.zip.{ZipEntry, ZipOutputStream}
import java.nio.charset.StandardCharsets

/**
 * Regression tests for error handling paths inside XlsxReader.
 *
 * These cover malformed XML, missing parts, and corrupted ZIP archives to ensure we surface precise
 * XLErrors instead of throwing.
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

  test("XlsxReader tolerates a leading DOCTYPE without internal subset (GH-350)") {
    // Third-party producers may emit a benign DOCTYPE; Excel/openpyxl read such files fine.
    val doctypeWorkbook =
      """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE workbook SYSTEM "http://example.com/workbook.dtd">
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
    val bytes = buildWorkbook(overrides = Map("xl/workbook.xml" -> doctypeWorkbook))
    val wb = XlsxReader
      .readFromBytes(bytes)
      .fold(err => fail(s"DOCTYPE-bearing workbook must read: $err"), identity)
    assertEquals(wb.sheets(0)(ref"A1").value, CellValue.Text("Hello"))
  }

  test("XlsxReader tolerates a leading DOCTYPE with bracketed internal subset (GH-350)") {
    // Internal subset with quoted '>' and ']' stresses the conservative prolog scanner. The
    // declared entity is never referenced in the document, so not honoring it is benign.
    val doctypeSheet =
      """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE worksheet [
  <!ELEMENT worksheet ANY>
  <!ENTITY unused "tricky > ] value">
]>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is><t>Hello</t></is>
      </c>
    </row>
  </sheetData>
</worksheet>"""
    val bytes = buildWorkbook(overrides = Map("xl/worksheets/sheet1.xml" -> doctypeSheet))
    val wb = XlsxReader
      .readFromBytes(bytes)
      .fold(err => fail(s"DOCTYPE-bearing worksheet must read: $err"), identity)
    assertEquals(wb.sheets(0)(ref"A1").value, CellValue.Text("Hello"))
  }

  test("XlsxReader tolerates a UTF-8 BOM before a leading DOCTYPE (GH-350)") {
    // Producers that emit DOCTYPEs are exactly the ones that emit BOMs; the strip must see
    // through the BOM (which survives the bytes->String decode as U+FEFF).
    val bomDoctypeWorkbook = "\uFEFF" +
      """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE workbook SYSTEM "http://example.invalid/workbook.dtd">
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
    val bytes = buildWorkbook(overrides = Map("xl/workbook.xml" -> bomDoctypeWorkbook))
    val wb = XlsxReader
      .readFromBytes(bytes)
      .fold(err => fail(s"BOM + DOCTYPE workbook must read: $err"), identity)
    assertEquals(wb.sheets(0)(ref"A1").value, CellValue.Text("Hello"))
  }

  test("GH-350: synthesized field-class workbook reads (doctypes + huge styles + externalLinks)") {
    // The field databook class GH-350 names: DOCTYPE headers on core parts, a styles part with
    // thousands of cellXfs, and an externalLinks part. Every dimension at once must read.
    val doctypeWorkbook =
      """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE workbook SYSTEM "http://example.invalid/workbook.dtd">
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <externalReferences>
    <externalReference r:id="rId2"/>
  </externalReferences>
</workbook>"""
    val relsWithExternalLink =
      """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""
    val contentTypesWithExtras =
      """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>"""
    val externalLinkXml =
      """<?xml version="1.0" encoding="UTF-8"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rId1">
    <sheetNames><sheetName val="Extern"/></sheetNames>
  </externalBook>
</externalLink>"""
    val externalLinkRels =
      """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="extern.xlsx" TargetMode="External"/>
</Relationships>"""
    // BOM + DOCTYPE + ~4000 cellXfs: the huge-styles dimension of the field class. The xfs are
    // VARIED (cycling numFmtId/alignment) so the part stays under the reader's 100:1 zip-bomb
    // compression-ratio guard, like real field styles would.
    val horiz = Vector("general", "left", "center", "right", "fill", "justify")
    val hugeXfs = (1 to 4000).map { i =>
      s"""<xf numFmtId="${i % 50}" fontId="0" fillId="0" borderId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="${horiz(
          i % 6
        )}" indent="${i % 16}" textRotation="${i % 91}"/></xf>"""
    }.mkString
    val bomDoctypeHugeStyles = "\uFEFF" +
      s"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE styleSheet SYSTEM "http://example.invalid/styles.dtd">
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4001"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>$hugeXfs</cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""
    val doctypeSharedStrings =
      """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE sst [ <!-- producers put comments here, don't choke --> ]>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Shared</t></si>
</sst>"""
    val sheetUsingSst =
      """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE worksheet [
  <!ELEMENT worksheet ANY>
  <!-- a ']' in a comment: ] must not truncate the strip -->
]>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" s="3999"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"""
    val bytes = buildWorkbook(overrides =
      Map(
        "[Content_Types].xml" -> contentTypesWithExtras,
        "xl/workbook.xml" -> doctypeWorkbook,
        "xl/_rels/workbook.xml.rels" -> relsWithExternalLink,
        "xl/worksheets/sheet1.xml" -> sheetUsingSst,
        "xl/styles.xml" -> bomDoctypeHugeStyles,
        "xl/sharedStrings.xml" -> doctypeSharedStrings,
        "xl/externalLinks/externalLink1.xml" -> externalLinkXml,
        "xl/externalLinks/_rels/externalLink1.xml.rels" -> externalLinkRels
      )
    )
    val wb = XlsxReader
      .readFromBytes(bytes)
      .fold(err => fail(s"field-class workbook must read: $err"), identity)
    assertEquals(wb.sheets(0)(ref"A1").value, CellValue.Text("Shared"))
    assertEquals(wb.sheets(0)(ref"B1").value, CellValue.Number(BigDecimal(42)))
  }

  test("XlsxReader parse errors name the offending line and column (GH-350)") {
    val malformed =
      """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets></wrong>
</workbook>"""
    val bytes = buildWorkbook(overrides = Map("xl/workbook.xml" -> malformed))
    XlsxReader.readFromBytes(bytes) match
      case Left(XLError.ParseError(location, message)) =>
        assertEquals(location, "xl/workbook.xml")
        assert(message.contains("line 3"), s"Expected line 3 in message, got: $message")
        assert(message.contains("column"), s"Expected column in message, got: $message")
      case other => fail(s"Expected ParseError with position info, got $other")
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

  test("XlsxReader survives non-positive font sz in styles.xml (GH-278)") {
    // Font requires sizePt > 0; a malformed val="-4" must fall back to the
    // default size at the parse site instead of throwing through the require.
    val malformedStyles =
      """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="-4"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"""
    val bytes = buildWorkbook(overrides = Map("xl/styles.xml" -> malformedStyles))
    XlsxReader.readFromBytes(bytes) match
      case Right(wb) =>
        assertEquals(wb.sheets(0)(ref"A1").value, CellValue.Text("Hello"))
      case Left(err) => fail(s"read must not fail on a malformed font size: $err")
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
    assertEquals(sheet(ref"A1").value, com.tjclp.xl.cells.CellValue.Text("Hello"))
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
