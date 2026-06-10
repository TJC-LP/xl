package com.tjclp.xl.ooxml

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}

import java.time.LocalDateTime

import munit.FunSuite

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.api.{Sheet, Workbook}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref

/**
 * GH-243: 1904 date system support.
 *
 * Files written by legacy Mac Excel set `<workbookPr date1904="1"/>` and store date serials
 * counted from the 1904-01-01 epoch (no phantom 1900 leap day). Dropping the flag silently shifts
 * every date by 1462 days (~4 years). These tests pin: the flag is read into
 * `WorkbookMetadata.date1904`, preserved on write (preserve-the-system), and DateTime values are
 * serialized with the 1904 epoch when the flag is set.
 */
class Date1904Spec extends FunSuite:

  // Microsoft's documented example pair: July 5, 1998 is serial 34519 in the 1904
  // date system and serial 35981 (= 34519 + 1462) in the 1900 date system.
  private val july5_1998Serial1904 = 34519.0

  private val contentTypesXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
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

  /** workbook.xml with a configurable workbookPr element (legacy Mac Excel shape). */
  private def workbookXml(workbookPr: String): String =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  $workbookPr
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""

  /** Worksheet with A1 = serial 34519 styled with the built-in date format (numFmtId 14). */
  private val worksheetXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1"><v>34519</v></c>
    </row>
  </sheetData>
</worksheet>"""

  private val stylesXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

  /** Build a minimal xlsx from raw parts, with the given workbookPr element. */
  private def build1904Workbook(workbookPr: String = """<workbookPr date1904="1"/>"""): Array[Byte] =
    val parts = Map(
      "[Content_Types].xml" -> contentTypesXml,
      "_rels/.rels" -> rootRelationshipsXml,
      "xl/workbook.xml" -> workbookXml(workbookPr),
      "xl/_rels/workbook.xml.rels" -> workbookRelationshipsXml,
      "xl/worksheets/sheet1.xml" -> worksheetXml,
      "xl/styles.xml" -> stylesXml
    )
    writeZip(parts)

  private def writeZip(parts: Map[String, String]): Array[Byte] =
    val baos = new ByteArrayOutputStream()
    val zip = new ZipOutputStream(baos)
    parts.foreach { case (name, content) =>
      zip.putNextEntry(new ZipEntry(name))
      zip.write(content.getBytes(StandardCharsets.UTF_8))
      zip.closeEntry()
    }
    zip.close()
    baos.toByteArray

  /** Extract one entry from an xlsx byte array as UTF-8 text. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def zipEntryText(bytes: Array[Byte], entryName: String): Option[String] =
    val zip = new ZipInputStream(new ByteArrayInputStream(bytes))
    try
      var entry = zip.getNextEntry
      var found: Option[String] = None
      while entry != null && found.isEmpty do
        if entry.getName == entryName then
          found = Some(new String(zip.readAllBytes(), StandardCharsets.UTF_8))
        entry = zip.getNextEntry
      found
    finally zip.close()

  // ---------------------------------------------------------------------------
  // Round-trip preservation (preserve-the-system)
  // ---------------------------------------------------------------------------

  test("GH-243: date1904 flag survives read → write → read (full regeneration)") {
    val wb = XlsxReader
      .readFromBytes(build1904Workbook())
      .getOrElse(fail("Failed to read 1904 fixture"))

    val written = XlsxWriter.writeToBytes(wb).getOrElse(fail("Failed to write workbook"))
    val workbookXmlOut = zipEntryText(written, "xl/workbook.xml").getOrElse(
      fail("Written xlsx missing xl/workbook.xml")
    )

    assert(
      workbookXmlOut.contains("date1904=\"1\""),
      s"date1904 flag dropped on write; serials would silently shift by 1462 days. Got: $workbookXmlOut"
    )

    // And the raw serial must ride through unchanged (the system travels with it).
    val reread = XlsxReader.readFromBytes(written).getOrElse(fail("Failed to re-read"))
    val sheet = reread("Sheet1").getOrElse(fail("Missing Sheet1"))
    assertEquals(sheet(ref"A1").value, CellValue.Number(BigDecimal(34519)))
    assert(reread.metadata.date1904, "re-read workbook must keep date1904 = true")
  }

  // ---------------------------------------------------------------------------
  // Reading the flag into WorkbookMetadata
  // ---------------------------------------------------------------------------

  test("GH-243: reader populates WorkbookMetadata.date1904 from workbookPr") {
    val wb = XlsxReader
      .readFromBytes(build1904Workbook())
      .getOrElse(fail("Failed to read 1904 fixture"))

    assert(wb.metadata.date1904, "expected date1904 = true from <workbookPr date1904=\"1\"/>")

    // The serial stays raw on read; interpreting it with the workbook's system yields the
    // documented date (1998-07-05 = serial 34519 in the 1904 system).
    val sheet = wb("Sheet1").getOrElse(fail("Missing Sheet1"))
    sheet(ref"A1").value match
      case CellValue.Number(serial) =>
        assertEquals(
          CellValue.excelSerialToDateTime(serial.toDouble, date1904 = true),
          LocalDateTime.of(1998, 7, 5, 0, 0, 0)
        )
      case other => fail(s"Expected Number cell, got $other")
  }

  test("GH-243: reader accepts the xsd:boolean form date1904=\"true\"") {
    val wb = XlsxReader
      .readFromBytes(build1904Workbook("""<workbookPr date1904="true"/>"""))
      .getOrElse(fail("Failed to read fixture"))
    assert(wb.metadata.date1904)
  }

  test("GH-243: date1904 defaults to false (absent attribute, absent workbookPr, explicit 0)") {
    val absentAttr = XlsxReader
      .readFromBytes(build1904Workbook("<workbookPr/>"))
      .getOrElse(fail("Failed to read fixture"))
    assert(!absentAttr.metadata.date1904)

    val absentElem = XlsxReader
      .readFromBytes(build1904Workbook(""))
      .getOrElse(fail("Failed to read fixture"))
    assert(!absentElem.metadata.date1904)

    val explicitZero = XlsxReader
      .readFromBytes(build1904Workbook("""<workbookPr date1904="0"/>"""))
      .getOrElse(fail("Failed to read fixture"))
    assert(!explicitZero.metadata.date1904)

    val explicitFalse = XlsxReader
      .readFromBytes(build1904Workbook("""<workbookPr date1904="false"/>"""))
      .getOrElse(fail("Failed to read fixture"))
    assert(!explicitFalse.metadata.date1904)
  }

  // ---------------------------------------------------------------------------
  // Writing: DateTime serials use the workbook's date system
  // ---------------------------------------------------------------------------

  test("GH-243: DateTime cells in a date1904 workbook are serialized with the 1904 epoch") {
    val dt = LocalDateTime.of(1998, 7, 5, 0, 0, 0)
    val sheet = Sheet(SheetName.unsafe("Dates")).put(ref"A1", CellValue.DateTime(dt))
    val wb = Workbook(Vector(sheet))
    val wb1904 = wb.copy(metadata = wb.metadata.copy(date1904 = true))

    val written = XlsxWriter.writeToBytes(wb1904).getOrElse(fail("Failed to write workbook"))

    val workbookXmlOut = zipEntryText(written, "xl/workbook.xml").getOrElse(
      fail("Written xlsx missing xl/workbook.xml")
    )
    assert(
      workbookXmlOut.contains("date1904=\"1\""),
      s"workbookPr must declare the 1904 system. Got: $workbookXmlOut"
    )

    val reread = XlsxReader.readFromBytes(written).getOrElse(fail("Failed to re-read"))
    assert(reread.metadata.date1904)
    val rereadSheet = reread("Dates").getOrElse(fail("Missing Dates sheet"))
    rereadSheet(ref"A1").value match
      case CellValue.Number(serial) =>
        // 1904-system serial (34519), NOT the 1900-system serial (35981 = 34519 + 1462).
        assertEquals(serial.toDouble, 34519.0)
        assertEquals(CellValue.excelSerialToDateTime(serial.toDouble, date1904 = true), dt)
      case other => fail(s"Expected Number cell after round-trip, got $other")
  }

  test("GH-243: surgical write keeps the 1904 system when editing a 1904 file") {
    val srcPath = java.nio.file.Files.createTempFile("xl-1904-", ".xlsx")
    val outPath = java.nio.file.Files.createTempFile("xl-1904-out-", ".xlsx")
    try
      java.nio.file.Files.write(srcPath, build1904Workbook())

      // File read attaches a SourceContext → write goes down the surgical path.
      val wb = XlsxReader.read(srcPath).getOrElse(fail("Failed to read 1904 file"))
      assert(wb.metadata.date1904)

      val dt = LocalDateTime.of(2026, 6, 10, 0, 0, 0)
      val edited = wb
        .update(SheetName.unsafe("Sheet1"), _.put(ref"B1", CellValue.DateTime(dt)))
        .getOrElse(fail("Failed to update sheet"))
      XlsxWriter.write(edited, outPath).getOrElse(fail("Failed to write surgically"))

      val reread = XlsxReader.read(outPath).getOrElse(fail("Failed to re-read"))
      assert(reread.metadata.date1904, "surgical write must preserve the 1904 flag")
      val sheet = reread("Sheet1").getOrElse(fail("Missing Sheet1"))
      // Untouched 1904 serial rides through; the new DateTime is written with the 1904 epoch.
      assertEquals(sheet(ref"A1").value, CellValue.Number(BigDecimal(34519)))
      sheet(ref"B1").value match
        case CellValue.Number(serial) =>
          assertEquals(CellValue.excelSerialToDateTime(serial.toDouble, date1904 = true), dt)
          assertEquals(
            serial.toDouble,
            CellValue.dateTimeToExcelSerial(dt, date1904 = true)
          )
        case other => fail(s"Expected Number cell for B1, got $other")
    finally
      java.nio.file.Files.deleteIfExists(srcPath)
      java.nio.file.Files.deleteIfExists(outPath)
  }

  test("GH-243: default workbooks keep the 1900 system (no date1904 attribute, 1900 serials)") {
    val dt = LocalDateTime.of(1998, 7, 5, 0, 0, 0)
    val sheet = Sheet(SheetName.unsafe("Dates")).put(ref"A1", CellValue.DateTime(dt))
    val wb = Workbook(Vector(sheet))

    val written = XlsxWriter.writeToBytes(wb).getOrElse(fail("Failed to write workbook"))

    val workbookXmlOut = zipEntryText(written, "xl/workbook.xml").getOrElse(
      fail("Written xlsx missing xl/workbook.xml")
    )
    assert(
      !workbookXmlOut.contains("date1904"),
      s"1900-system workbooks must not declare date1904. Got: $workbookXmlOut"
    )

    val reread = XlsxReader.readFromBytes(written).getOrElse(fail("Failed to re-read"))
    assert(!reread.metadata.date1904)
    val rereadSheet = reread("Dates").getOrElse(fail("Missing Dates sheet"))
    rereadSheet(ref"A1").value match
      case CellValue.Number(serial) => assertEquals(serial.toDouble, 35981.0)
      case other => fail(s"Expected Number cell after round-trip, got $other")
  }
