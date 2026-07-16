package com.tjclp.xl.ooxml

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}

import scala.xml.Elem

import munit.FunSuite

import com.tjclp.xl.api.{CalcPr, Workbook}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref

/**
 * GH-373 (a): `WorkbookMetadata.calcPr` ⇄ `<calcPr>` iterative-calculation attributes.
 *
 * Contract (the GH-243 date1904 / GH-294 activeTab machinery):
 *   - The reader parses `iterate` / `iterateCount` / `iterateDelta` into the friendly-named
 *     [[CalcPr]] model; a calcPr with only unmodeled attributes (calcId, refMode, ...) parses to
 *     None and rides through on write.
 *   - `Workbook.withCalcPr` marks metadata modified so a surgical write regenerates workbook.xml —
 *     without it the untouched fast path copies preserved bytes and silently drops the change.
 *   - Reconcile is model-wins for the three modeled attributes ONLY: unmodeled preserved attributes
 *     survive an overlay; the preserved element rides byte-stable when the model agrees.
 *   - Scratch builds emit `<calcPr>` with just the modeled attributes (no calcId — the LibreOffice
 *     precedent), in CT_Workbook schema position (after sheets/definedNames, before extLst).
 */
class CalcPrRoundTripSpec extends FunSuite:

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

  /** workbook.xml with a configurable raw calcPr element (schema position: after sheets). */
  private def workbookXml(calcPr: String): String =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Model" sheetId="1" r:id="rId1"/>
  </sheets>
  $calcPr
</workbook>"""

  private val worksheetXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"""

  private def buildCalcPrWorkbook(calcPr: String): Array[Byte] =
    writeZip(
      Map(
        "[Content_Types].xml" -> contentTypesXml,
        "_rels/.rels" -> rootRelationshipsXml,
        "xl/workbook.xml" -> workbookXml(calcPr),
        "xl/_rels/workbook.xml.rels" -> workbookRelationshipsXml,
        "xl/worksheets/sheet1.xml" -> worksheetXml
      )
    )

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

  private def readBytes(bytes: Array[Byte]): Workbook =
    XlsxReader.readFromBytes(bytes).fold(e => fail(s"read failed: ${e.message}"), identity)

  private def calcPrElemOf(workbookXml: String): Option[Elem] =
    val root = XmlSecurity
      .parseSafe(workbookXml, "xl/workbook.xml")
      .fold(e => fail(e.message), identity)
    (root \ "calcPr").headOption.collect { case e: Elem => e }

  private def parseElem(xml: String): Elem =
    XmlSecurity.parseSafe(xml, "test").fold(e => fail(e.message), identity)

  private val houseCalcPr =
    CalcPr(
      iterativeCalculation = true,
      maxIterations = Some(100),
      maxChange = Some(BigDecimal("0.001"))
    )

  // ---------------------------------------------------------------------------
  // The wither MUST mark metadata modified (the silent-drop fast path)
  // ---------------------------------------------------------------------------

  test("GH-373: withCalcPr on a read workbook survives a file→file write (model wins)") {
    val src = Files.createTempFile("calcpr-src", ".xlsx")
    val out = Files.createTempFile("calcpr-out", ".xlsx")
    try
      Files.write(src, buildCalcPrWorkbook("""<calcPr calcId="191029"/>"""))

      val wb = XlsxReader.read(src).fold(e => fail(s"read failed: ${e.message}"), identity)
      assertEquals(wb.metadata.calcPr, None, "calcId-only calcPr must parse to None")

      // The ONLY change is calcPr — without markMetadataModified the untouched fast path
      // copies the source bytes verbatim and the authored setting silently vanishes.
      val authored = wb.withCalcPr(houseCalcPr)
      XlsxWriter.write(authored, out).fold(e => fail(s"write failed: ${e.message}"), identity)

      val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: ${e.message}"), identity)
      assertEquals(
        reread.metadata.calcPr,
        Some(houseCalcPr),
        "calcPr-only change was dropped on write"
      )

      val calcPr = calcPrElemOf(
        zipEntryText(Files.readAllBytes(out), "xl/workbook.xml").getOrElse(
          fail("workbook.xml missing")
        )
      ).getOrElse(fail("calcPr element missing"))
      assertEquals(calcPr \@ "iterate", "1")
      assertEquals(calcPr \@ "iterateCount", "100")
      assertEquals(calcPr \@ "iterateDelta", "0.001")
      assertEquals(calcPr \@ "calcId", "191029", "unmodeled calcId must survive the overlay")
    finally
      Files.deleteIfExists(src)
      Files.deleteIfExists(out)
  }

  // ---------------------------------------------------------------------------
  // Scratch authoring (the from-scratch replica case)
  // ---------------------------------------------------------------------------

  test("GH-373: scratch workbook emits authored calcPr and round-trips") {
    val wb = Workbook(com.tjclp.xl.api.Sheet("Model").put(ref"A1" -> 1))
      .withCalcPr(
        CalcPr(
          iterativeCalculation = true,
          maxIterations = Some(100),
          maxChange = Some(BigDecimal("0.0001"))
        )
      )

    val bytes = XlsxWriter.writeToBytes(wb).fold(e => fail(s"write failed: ${e.message}"), identity)
    val workbookOut = zipEntryText(bytes, "xl/workbook.xml").getOrElse(fail("workbook.xml missing"))

    val calcPr = calcPrElemOf(workbookOut).getOrElse(fail(s"calcPr missing: $workbookOut"))
    assertEquals(calcPr \@ "iterate", "1")
    assertEquals(calcPr \@ "iterateCount", "100")
    assertEquals(calcPr \@ "iterateDelta", "0.0001")
    assertEquals(calcPr \@ "calcId", "", "scratch emission needs no calcId (LO precedent)")

    // CT_Workbook schema order: sheets before calcPr
    val sheetsIdx = workbookOut.indexOf("<sheets>")
    val calcPrIdx = workbookOut.indexOf("<calcPr")
    assert(sheetsIdx >= 0 && calcPrIdx >= 0, workbookOut)
    assert(sheetsIdx < calcPrIdx, "calcPr must follow sheets (CT_Workbook sequence)")

    val reread = readBytes(bytes)
    assertEquals(
      reread.metadata.calcPr,
      Some(CalcPr(true, Some(100), Some(BigDecimal("0.0001"))))
    )
  }

  test("GH-373: workbooks without authored calcPr emit none (scratch default unchanged)") {
    val wb = Workbook(com.tjclp.xl.api.Sheet("Plain").put(ref"A1" -> 1))
    val bytes = XlsxWriter.writeToBytes(wb).fold(e => fail(s"write failed: ${e.message}"), identity)
    val workbookOut = zipEntryText(bytes, "xl/workbook.xml").getOrElse(fail("workbook.xml missing"))
    assert(!workbookOut.contains("<calcPr"), s"no calcPr expected: $workbookOut")
    assertEquals(readBytes(bytes).metadata.calcPr, None)
  }

  // ---------------------------------------------------------------------------
  // Reader forms
  // ---------------------------------------------------------------------------

  test("GH-373: reader parses the OOXML attribute spellings into friendly names") {
    // xsd:boolean form
    assertEquals(
      readBytes(
        buildCalcPrWorkbook("""<calcPr iterate="true" iterateCount="50"/>""")
      ).metadata.calcPr,
      Some(CalcPr(iterativeCalculation = true, maxIterations = Some(50), maxChange = None))
    )
    // the LibreOffice fixture shape: iterate explicitly false, count+delta present
    assertEquals(
      readBytes(
        buildCalcPrWorkbook(
          """<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.0001"/>"""
        )
      ).metadata.calcPr,
      Some(
        CalcPr(
          iterativeCalculation = false,
          maxIterations = Some(100),
          maxChange = Some(BigDecimal("0.0001"))
        )
      )
    )
    // unmodeled-only calcPr parses to None (and rides through on write elsewhere)
    assertEquals(
      readBytes(
        buildCalcPrWorkbook("""<calcPr calcId="191029" fullCalcOnLoad="1"/>""")
      ).metadata.calcPr,
      None
    )
    // absent element
    assertEquals(readBytes(buildCalcPrWorkbook("")).metadata.calcPr, None)
  }

  // ---------------------------------------------------------------------------
  // Metadata-modified writes: preserved calcPr rides through when the model agrees
  // ---------------------------------------------------------------------------

  test("GH-373: a metadata write keeps the preserved calcPr when the model agrees") {
    val src = Files.createTempFile("calcpr-agree-src", ".xlsx")
    val out = Files.createTempFile("calcpr-agree-out", ".xlsx")
    try
      Files.write(
        src,
        buildCalcPrWorkbook(
          """<calcPr calcId="191029" iterate="1" iterateCount="50" iterateDelta="0.01" refMode="A1"/>"""
        )
      )
      val wb = XlsxReader.read(src).fold(e => fail(s"read failed: ${e.message}"), identity)
      assertEquals(wb.metadata.calcPr, Some(CalcPr(true, Some(50), Some(BigDecimal("0.01")))))

      // trigger the metadata-modified (fromDomain) path with an unrelated authoring action
      val named = wb.withDefinedName("MyRange", "Model!$A$1")
      XlsxWriter.write(named, out).fold(e => fail(s"write failed: ${e.message}"), identity)

      val calcPr = calcPrElemOf(
        zipEntryText(Files.readAllBytes(out), "xl/workbook.xml").getOrElse(
          fail("workbook.xml missing")
        )
      ).getOrElse(fail("calcPr dropped on metadata write"))
      assertEquals(calcPr \@ "calcId", "191029")
      assertEquals(calcPr \@ "refMode", "A1")
      assertEquals(calcPr \@ "iterate", "1")
      assertEquals(calcPr \@ "iterateCount", "50")
      assertEquals(calcPr \@ "iterateDelta", "0.01")
    finally
      Files.deleteIfExists(src)
      Files.deleteIfExists(out)
  }

  // ---------------------------------------------------------------------------
  // reconcileCalcPr unit laws (model wins on the three modeled attrs only)
  // ---------------------------------------------------------------------------

  test("GH-373: reconcileCalcPr is the identity when the model agrees with the preserved bytes") {
    val pr =
      parseElem("""<calcPr calcId="191029" iterate="1" iterateCount="50" iterateDelta="0.01"/>""")
    val model = OoxmlWorkbook.parseCalcPr(Some(pr))
    assert(
      OoxmlWorkbook.reconcileCalcPr(Some(pr), model).exists(_ eq pr),
      "must return the preserved element untouched"
    )
    assertEquals(OoxmlWorkbook.reconcileCalcPr(None, None), None)
  }

  test("GH-373: reconcileCalcPr overlays the model in place (foreign attrs survive)") {
    val pr =
      parseElem("""<calcPr calcId="191029" iterate="false" iterateCount="100" refMode="A1"/>""")
    val model = Some(
      CalcPr(
        iterativeCalculation = true,
        maxIterations = Some(200),
        maxChange = Some(BigDecimal("0.05"))
      )
    )
    val result =
      OoxmlWorkbook.reconcileCalcPr(Some(pr), model).getOrElse(fail("expected an element"))
    assertEquals(result \@ "calcId", "191029")
    assertEquals(result \@ "refMode", "A1")
    assertEquals(result \@ "iterate", "1")
    assertEquals(result \@ "iterateCount", "200")
    assertEquals(result \@ "iterateDelta", "0.05")
  }

  test("GH-373: reconcileCalcPr with a cleared model strips only the modeled attrs") {
    val pr = parseElem("""<calcPr calcId="191029" iterate="1" iterateCount="100" refMode="A1"/>""")
    val result =
      OoxmlWorkbook.reconcileCalcPr(Some(pr), None).getOrElse(fail("expected an element"))
    assertEquals(result \@ "calcId", "191029")
    assertEquals(result \@ "refMode", "A1")
    assertEquals(result.attribute("iterate"), None)
    assertEquals(result.attribute("iterateCount"), None)
    assertEquals(result.attribute("iterateDelta"), None)
  }

  test("GH-373: reconcileCalcPr emits false-with-bounds without an iterate attribute") {
    // iterativeCalculation=false is the schema default: the attribute is omitted, not spelled out
    val model =
      Some(CalcPr(iterativeCalculation = false, maxIterations = Some(10), maxChange = None))
    val result = OoxmlWorkbook.reconcileCalcPr(None, model).getOrElse(fail("expected an element"))
    assertEquals(result.attribute("iterate"), None)
    assertEquals(result \@ "iterateCount", "10")
    // and it parses back to the same model (round-trip)
    assertEquals(OoxmlWorkbook.parseCalcPr(Some(result)), model)
  }
