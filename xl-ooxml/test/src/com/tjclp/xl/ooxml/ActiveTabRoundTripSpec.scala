package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * GH-294: `Workbook.activeSheetIndex` ⇄ `<bookViews><workbookView activeTab>` serialization.
 *
 * Contract:
 *   - Fresh (model-driven) workbook.xml always ships `<bookViews><workbookView activeTab="N"/>`
 *     (Excel always writes bookViews).
 *   - Preserved bookViews: the model value WINS for activeTab (docProps GH-242 precedent) while
 *     every unmodeled attribute (window geometry, tabRatio, ...) rides through. When the model says
 *     0 and the source omitted the attribute, it stays omitted (0 is the schema default) — keeping
 *     parse→merge an identity for both the Excel (attr omitted) and openpyxl (attr explicit)
 *     spellings.
 *   - activeTab is clamped into [0, sheetCount-1] on read (lenient) and on write.
 */
class ActiveTabRoundTripSpec extends FunSuite:

  private def zipEntryString(path: Path, entry: String): String =
    val zf = new ZipFile(path.toFile)
    try
      val is = zf.getInputStream(zf.getEntry(entry))
      try new String(is.readAllBytes(), StandardCharsets.UTF_8)
      finally is.close()
    finally zf.close()

  private def writeRead(wb: Workbook): (Workbook, Path) =
    val out = Files.createTempFile("activetab", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    (reread, out)

  private def twoSheetWorkbook: Workbook =
    Workbook(
      Sheet("First").put(ref"A1" -> 1),
      Sheet("Second").put(ref"A1" -> 2)
    )

  test("GH-294: activeSheetIndex=1 survives write→read (fresh workbook)") {
    val wb = twoSheetWorkbook
      .setActiveSheet(1)
      .fold(e => fail(s"setActiveSheet failed: $e"), identity)
    val (reread, out) = writeRead(wb)

    val workbookXml = zipEntryString(out, "xl/workbook.xml")
    assert(workbookXml.contains("<bookViews>"), s"bookViews missing: $workbookXml")
    assert(workbookXml.contains("activeTab=\"1\""), s"activeTab missing: $workbookXml")
    assertEquals(reread.activeSheetIndex, 1, "activeSheetIndex lost on round-trip")
    Files.deleteIfExists(out)
  }

  test("GH-294: fresh write always emits bookViews (activeTab 0 explicit)") {
    val (reread, out) = writeRead(twoSheetWorkbook)
    val workbookXml = zipEntryString(out, "xl/workbook.xml")
    assert(workbookXml.contains("<bookViews>"), s"bookViews missing for default: $workbookXml")
    assert(workbookXml.contains("activeTab=\"0\""), s"explicit activeTab=0 missing: $workbookXml")
    assertEquals(reread.activeSheetIndex, 0)
    Files.deleteIfExists(out)
  }

  test("GH-294: bookViews precedes sheets in workbook.xml (schema order)") {
    val (_, out) = writeRead(twoSheetWorkbook)
    val workbookXml = zipEntryString(out, "xl/workbook.xml")
    val bookViews = workbookXml.indexOf("<bookViews>")
    val sheets = workbookXml.indexOf("<sheets>")
    assert(bookViews >= 0 && sheets >= 0, s"elements missing: $workbookXml")
    assert(bookViews < sheets, "bookViews must precede sheets (CT_Workbook sequence)")
    Files.deleteIfExists(out)
  }

  test("GH-294: surgical cell edit keeps foreign bookViews attrs AND the parsed activeTab") {
    val src = rawTwoSheetFixture(
      """<bookViews><workbookView xWindow="120" yWindow="60" windowWidth="15000" windowHeight="9000" activeTab="1"/></bookViews>"""
    )
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    assertEquals(wb.activeSheetIndex, 1, "reader did not parse activeTab")

    val edited = wb("First")
      .map(sheet => wb.put(sheet.put(ref"B1" -> 42)))
      .fold(e => fail(s"edit failed: $e"), identity)
    val out = Files.createTempFile("activetab-surgical", ".xlsx")
    XlsxWriter.write(edited, out).fold(e => fail(s"write failed: $e"), identity)

    val workbookXml = zipEntryString(out, "xl/workbook.xml")
    assert(workbookXml.contains("xWindow=\"120\""), s"window geometry lost: $workbookXml")
    assert(workbookXml.contains("activeTab=\"1\""), s"activeTab lost: $workbookXml")
    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    assertEquals(reread.activeSheetIndex, 1)
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-294: setActiveSheet on a read workbook survives a file→file write (model wins)") {
    val seed = Files.createTempFile("activetab-seed", ".xlsx")
    XlsxWriter.write(twoSheetWorkbook, seed).fold(e => fail(s"seed write failed: $e"), identity)

    val activated = XlsxReader
      .read(seed)
      .flatMap(_.setActiveSheet(1))
      .fold(e => fail(s"activate failed: $e"), identity)
    val out = Files.createTempFile("activetab-activated", ".xlsx")
    XlsxWriter.write(activated, out).fold(e => fail(s"write failed: $e"), identity)

    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    assertEquals(reread.activeSheetIndex, 1, "activeTab-only change was dropped on write")
    Files.deleteIfExists(seed)
    Files.deleteIfExists(out)
  }

  test("GH-294: read clamps out-of-range and negative activeTab (lenient)") {
    val tooBig = rawTwoSheetFixture("""<bookViews><workbookView activeTab="7"/></bookViews>""")
    val wbBig = XlsxReader.read(tooBig).fold(e => fail(s"read failed: $e"), identity)
    assertEquals(wbBig.activeSheetIndex, 1, "out-of-range activeTab should clamp to last sheet")

    val negative = rawTwoSheetFixture("""<bookViews><workbookView activeTab="-3"/></bookViews>""")
    val wbNeg = XlsxReader.read(negative).fold(e => fail(s"read failed: $e"), identity)
    assertEquals(wbNeg.activeSheetIndex, 0, "negative activeTab should clamp to 0")

    Files.deleteIfExists(tooBig)
    Files.deleteIfExists(negative)
  }

  test("GH-294: write clamps a model index beyond the sheet count") {
    val wb = twoSheetWorkbook.copy(activeSheetIndex = 9) // bypasses setActiveSheet validation
    val (reread, out) = writeRead(wb)
    val workbookXml = zipEntryString(out, "xl/workbook.xml")
    assert(workbookXml.contains("activeTab=\"1\""), s"clamped activeTab missing: $workbookXml")
    assertEquals(reread.activeSheetIndex, 1)
    Files.deleteIfExists(out)
  }

  test("GH-294: missing bookViews reads as activeSheetIndex 0") {
    val src = rawTwoSheetFixture("")
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    assertEquals(wb.activeSheetIndex, 0)
    Files.deleteIfExists(src)
  }

  // ========== helpers ==========

  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    out.putNextEntry(new ZipEntry(name))
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()

  /** Minimal Excel-shaped two-sheet fixture with the given raw bookViews XML (may be empty). */
  private def rawTwoSheetFixture(bookViewsXml: String): Path =
    val path = Files.createTempFile("activetab-fixture", ".xlsx")
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
        s"""<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  $bookViewsXml
  <sheets>
    <sheet name="First" sheetId="1" r:id="rId1"/>
    <sheet name="Second" sheetId="2" r:id="rId2"/>
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
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"""
      )
      writeEntry(
        out,
        "xl/worksheets/sheet2.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>2</v></c></row>
  </sheetData>
</worksheet>"""
      )
    finally out.close()
    path
