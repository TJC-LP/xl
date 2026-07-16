package com.tjclp.xl.ooxml

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}

import scala.xml.Elem

import munit.FunSuite

import com.tjclp.xl.api.{CalcMode, CalcPr, Workbook}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref

/**
 * GH-373 (a) + GH-400: `WorkbookMetadata.calcPr` ⇄ `<calcPr>` modeled attributes.
 *
 * Contract (the GH-243 date1904 / GH-294 activeTab machinery):
 *   - The reader parses the iterate triple (`iterate` / `iterateCount` / `iterateDelta`) plus
 *     `calcMode` / `fullCalcOnLoad` / `calcId` (GH-400) into the friendly-named [[CalcPr]] model; a
 *     calcPr with only still-unmodeled attributes (refMode, concurrentCalc, ...) parses to None and
 *     rides through on write.
 *   - `Workbook.withCalcPr` marks metadata modified so a surgical write regenerates workbook.xml —
 *     without it the untouched fast path copies preserved bytes and silently drops the change.
 *   - Reconcile: the iterate triple is model truth (a None field emits nothing and removes a stale
 *     preserved attribute); the GH-400 facts overlay only when Some (a None fact NEVER strips a
 *     preserved attribute). Still-unmodeled preserved attributes survive every overlay; the
 *     preserved element rides byte-stable when the model agrees.
 *   - Scratch builds emit `<calcPr>` with just the authored attributes (no calcId unless authored —
 *     the LibreOffice precedent), in CT_Workbook schema position (after sheets/definedNames, before
 *     extLst), with deterministically sorted attributes on both XML backends.
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
      assertEquals(
        wb.metadata.calcPr,
        Some(CalcPr(calcId = Some(191029))),
        "calcId-only calcPr must surface calcId (GH-400)"
      )

      // The ONLY change is calcPr — without markMetadataModified the untouched fast path
      // copies the source bytes verbatim and the authored setting silently vanishes.
      val authored = wb.withCalcPr(houseCalcPr)
      XlsxWriter.write(authored, out).fold(e => fail(s"write failed: ${e.message}"), identity)

      val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: ${e.message}"), identity)
      assertEquals(
        reread.metadata.calcPr,
        // the authored model has no opinion on calcId (None) — the preserved 191029 rides through
        // and the reader surfaces it (GH-400)
        Some(houseCalcPr.copy(calcId = Some(191029))),
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

  test("GH-400: scratch workbook writes the exact TJC house calcPr in schema position") {
    // The issue's acceptance, byte-for-byte: authoring the house doctrine from scratch emits
    // <calcPr calcId="191029" calcMode="autoNoTable" iterate="1"/> — sorted attributes, after
    // sheets (CT_Workbook sequence), retiring tjc-modeling's last zip patch.
    val wb = Workbook(com.tjclp.xl.api.Sheet("Model").put(ref"A1" -> 1))
      .withCalcPr(
        CalcPr(
          iterativeCalculation = true,
          calcMode = Some(CalcMode.AutoNoTable),
          calcId = Some(191029)
        )
      )

    val bytes = XlsxWriter.writeToBytes(wb).fold(e => fail(s"write failed: ${e.message}"), identity)
    val workbookOut = zipEntryText(bytes, "xl/workbook.xml").getOrElse(fail("workbook.xml missing"))

    assert(
      workbookOut.contains("""<calcPr calcId="191029" calcMode="autoNoTable" iterate="1"/>"""),
      s"exact house calcPr missing from raw workbook.xml: $workbookOut"
    )
    val sheetsEnd = workbookOut.indexOf("</sheets>")
    val calcPrIdx = workbookOut.indexOf("<calcPr")
    assert(sheetsEnd >= 0 && calcPrIdx > sheetsEnd, "calcPr must follow sheets (CT_Workbook)")

    val reread = readBytes(bytes)
    assertEquals(
      reread.metadata.calcPr,
      Some(
        CalcPr(
          iterativeCalculation = true,
          calcMode = Some(CalcMode.AutoNoTable),
          calcId = Some(191029)
        )
      ),
      "read-back must surface the authored GH-400 fields"
    )
  }

  test("GH-400: scratch calcPr lands between definedNames and nothing else (schema order)") {
    val wb = Workbook(com.tjclp.xl.api.Sheet("Model").put(ref"A1" -> 1))
      .withDefinedName("Rate", "0.08")
      .withCalcPr(CalcPr(calcMode = Some(CalcMode.Manual)))

    val bytes = XlsxWriter.writeToBytes(wb).fold(e => fail(s"write failed: ${e.message}"), identity)
    val workbookOut = zipEntryText(bytes, "xl/workbook.xml").getOrElse(fail("workbook.xml missing"))

    val dnEnd = workbookOut.indexOf("</definedNames>")
    val calcPrIdx = workbookOut.indexOf("<calcPr")
    assert(dnEnd >= 0, s"definedNames missing: $workbookOut")
    assert(calcPrIdx > dnEnd, "calcPr must follow definedNames (CT_Workbook sequence)")
    assert(workbookOut.contains("""<calcPr calcMode="manual"/>"""), workbookOut)
  }

  test("GH-400: all six modeled fields round-trip through a scratch write") {
    val full = CalcPr(
      iterativeCalculation = true,
      maxIterations = Some(50),
      maxChange = Some(BigDecimal("0.005")),
      calcMode = Some(CalcMode.Auto),
      fullCalcOnLoad = Some(true),
      calcId = Some(999999)
    )
    val wb = Workbook(com.tjclp.xl.api.Sheet("Model").put(ref"A1" -> 1)).withCalcPr(full)
    val bytes = XlsxWriter.writeToBytes(wb).fold(e => fail(s"write failed: ${e.message}"), identity)

    val calcPr = calcPrElemOf(
      zipEntryText(bytes, "xl/workbook.xml").getOrElse(fail("workbook.xml missing"))
    ).getOrElse(fail("calcPr missing"))
    assertEquals(calcPr \@ "calcMode", "auto")
    assertEquals(calcPr \@ "fullCalcOnLoad", "1")
    assertEquals(calcPr \@ "calcId", "999999")
    assertEquals(calcPr \@ "iterate", "1")
    assertEquals(calcPr \@ "iterateCount", "50")
    assertEquals(calcPr \@ "iterateDelta", "0.005")

    assertEquals(readBytes(bytes).metadata.calcPr, Some(full), "write→read must be the identity")
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
    // GH-400: calcId/fullCalcOnLoad are modeled now — they surface instead of parsing to None
    assertEquals(
      readBytes(
        buildCalcPrWorkbook("""<calcPr calcId="191029" fullCalcOnLoad="1"/>""")
      ).metadata.calcPr,
      Some(CalcPr(fullCalcOnLoad = Some(true), calcId = Some(191029)))
    )
    // still-unmodeled-only calcPr parses to None (and rides through on write elsewhere)
    assertEquals(
      readBytes(
        buildCalcPrWorkbook("""<calcPr refMode="A1" concurrentCalc="0"/>""")
      ).metadata.calcPr,
      None
    )
    // absent element
    assertEquals(readBytes(buildCalcPrWorkbook("")).metadata.calcPr, None)
  }

  test("GH-400: reader parses calcMode / fullCalcOnLoad / calcId (all spellings, total)") {
    // the TJC house shape
    assertEquals(
      readBytes(
        buildCalcPrWorkbook("""<calcPr calcId="191029" calcMode="autoNoTable" iterate="1"/>""")
      ).metadata.calcPr,
      Some(
        CalcPr(
          iterativeCalculation = true,
          calcMode = Some(CalcMode.AutoNoTable),
          calcId = Some(191029)
        )
      )
    )
    // every ST_CalcMode token
    assertEquals(
      readBytes(buildCalcPrWorkbook("""<calcPr calcMode="manual"/>""")).metadata.calcPr,
      Some(CalcPr(calcMode = Some(CalcMode.Manual)))
    )
    assertEquals(
      readBytes(buildCalcPrWorkbook("""<calcPr calcMode="auto"/>""")).metadata.calcPr,
      Some(CalcPr(calcMode = Some(CalcMode.Auto)))
    )
    // xsd:boolean spellings of fullCalcOnLoad
    assertEquals(
      readBytes(buildCalcPrWorkbook("""<calcPr fullCalcOnLoad="true"/>""")).metadata.calcPr,
      Some(CalcPr(fullCalcOnLoad = Some(true)))
    )
    assertEquals(
      readBytes(buildCalcPrWorkbook("""<calcPr fullCalcOnLoad="0"/>""")).metadata.calcPr,
      Some(CalcPr(fullCalcOnLoad = Some(false)))
    )
    // total parse: unknown/malformed values contribute no field, the read never fails —
    // presence of a modeled attribute name still yields Some (the iterate="garbage" precedent)
    assertEquals(
      readBytes(
        buildCalcPrWorkbook(
          """<calcPr calcMode="bogus" fullCalcOnLoad="maybe" calcId="not-a-number"/>"""
        )
      ).metadata.calcPr,
      Some(CalcPr())
    )
    // calcId beyond Int range (xsd:unsignedInt max) degrades to None, never throws
    assertEquals(
      readBytes(buildCalcPrWorkbook("""<calcPr calcId="4294967295"/>""")).metadata.calcPr,
      Some(CalcPr())
    )
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
      assertEquals(
        wb.metadata.calcPr,
        // GH-400: calcId surfaces alongside the iterate triple
        Some(CalcPr(true, Some(50), Some(BigDecimal("0.01")), calcId = Some(191029)))
      )

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

  test("GH-400: flipping calcMode overlays while unmodeled attrs and extLst ride through") {
    val src = Files.createTempFile("calcpr-flip-src", ".xlsx")
    val out = Files.createTempFile("calcpr-flip-out", ".xlsx")
    try
      Files.write(
        src,
        buildCalcPrWorkbook(
          """<calcPr calcId="171027" calcMode="manual" concurrentCalc="0" refMode="R1C1" iterate="1" iterateCount="50" iterateDelta="0.01"/>
  <extLst><ext uri="{TEST-EXT}"/></extLst>"""
        )
      )
      val wb = XlsxReader.read(src).fold(e => fail(s"read failed: ${e.message}"), identity)
      val parsed = wb.metadata.calcPr.getOrElse(fail("calcPr must parse"))
      assertEquals(parsed.calcMode, Some(CalcMode.Manual))
      assertEquals(parsed.calcId, Some(171027))

      // the modeled edit: flip calcMode only, everything else carried from the parsed model
      val flipped = wb.withCalcPr(parsed.copy(calcMode = Some(CalcMode.AutoNoTable)))
      XlsxWriter.write(flipped, out).fold(e => fail(s"write failed: ${e.message}"), identity)

      val workbookOut = zipEntryText(Files.readAllBytes(out), "xl/workbook.xml")
        .getOrElse(fail("workbook.xml missing"))
      val calcPr = calcPrElemOf(workbookOut).getOrElse(fail("calcPr dropped"))
      assertEquals(calcPr \@ "calcMode", "autoNoTable", "modeled edit must win")
      assertEquals(calcPr \@ "calcId", "171027")
      assertEquals(calcPr \@ "concurrentCalc", "0", "unmodeled concurrentCalc dropped")
      assertEquals(calcPr \@ "refMode", "R1C1", "unmodeled refMode dropped")
      assertEquals(calcPr \@ "iterate", "1")
      assertEquals(calcPr \@ "iterateCount", "50")
      assertEquals(calcPr \@ "iterateDelta", "0.01")
      // the overlay must replace, never duplicate, a preserved modeled attribute
      assertEquals(
        "calcMode=".r.findAllIn(workbookOut).size,
        1,
        s"duplicate calcMode attribute: $workbookOut"
      )
      // CT_Workbook sequence: calcPr still before the preserved extLst
      val calcPrIdx = workbookOut.indexOf("<calcPr")
      val extIdx = workbookOut.indexOf("<extLst")
      assert(extIdx >= 0, s"preserved extLst dropped: $workbookOut")
      assert(calcPrIdx >= 0 && calcPrIdx < extIdx, "calcPr must precede extLst")
    finally
      Files.deleteIfExists(src)
      Files.deleteIfExists(out)
  }

  test("GH-400: a calcMode-only edit on a read workbook survives the surgical write path") {
    val src = Files.createTempFile("calcmode-only-src", ".xlsx")
    val out = Files.createTempFile("calcmode-only-out", ".xlsx")
    try
      Files.write(
        src,
        buildCalcPrWorkbook("""<calcPr calcId="191029" calcMode="autoNoTable" iterate="1"/>""")
      )
      val wb = XlsxReader.read(src).fold(e => fail(s"read failed: ${e.message}"), identity)
      val parsed = wb.metadata.calcPr.getOrElse(fail("calcPr must parse"))

      // The ONLY change is calcMode — without the markMetadataModified trigger the untouched
      // fast path would copy the source bytes verbatim and silently keep autoNoTable.
      val edited = wb.withCalcPr(parsed.copy(calcMode = Some(CalcMode.Manual)))
      XlsxWriter.write(edited, out).fold(e => fail(s"write failed: ${e.message}"), identity)

      val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: ${e.message}"), identity)
      assertEquals(
        reread.metadata.calcPr,
        Some(
          CalcPr(
            iterativeCalculation = true,
            calcMode = Some(CalcMode.Manual),
            calcId = Some(191029)
          )
        ),
        "calcMode-only edit was dropped by the surgical write"
      )
    finally
      Files.deleteIfExists(src)
      Files.deleteIfExists(out)
  }

  // ---------------------------------------------------------------------------
  // reconcileCalcPr unit laws (model wins on the three modeled attrs only)
  // ---------------------------------------------------------------------------

  test("GH-373: reconcileCalcPr is the identity when the model agrees with the preserved bytes") {
    val pr =
      parseElem(
        """<calcPr calcId="191029" calcMode="autoNoTable" fullCalcOnLoad="1" iterate="1" iterateCount="50" iterateDelta="0.01" refMode="A1"/>"""
      )
    val model = OoxmlWorkbook.parseCalcPr(Some(pr))
    assert(
      OoxmlWorkbook.reconcileCalcPr(Some(pr), model).exists(_ eq pr),
      "must return the preserved element untouched"
    )
    assertEquals(OoxmlWorkbook.reconcileCalcPr(None, None), None)
  }

  test("GH-400: a None fact field never strips a preserved attribute") {
    // model authored without opinions on calcMode/fullCalcOnLoad/calcId — the preserved
    // spellings must ride through untouched (only the iterate triple is model truth)
    val pr = parseElem(
      """<calcPr calcId="191029" calcMode="autoNoTable" fullCalcOnLoad="1" refMode="A1" iterate="1" iterateCount="100"/>"""
    )
    val model = Some(CalcPr(iterativeCalculation = true))
    val result =
      OoxmlWorkbook.reconcileCalcPr(Some(pr), model).getOrElse(fail("expected an element"))
    assertEquals(result \@ "calcId", "191029", "None calcId stripped the preserved attr")
    assertEquals(result \@ "calcMode", "autoNoTable", "None calcMode stripped the preserved attr")
    assertEquals(result \@ "fullCalcOnLoad", "1", "None fullCalcOnLoad stripped the preserved attr")
    assertEquals(result \@ "refMode", "A1")
    assertEquals(result \@ "iterate", "1")
    // the iterate triple IS model truth: absent fields remove stale preserved attributes
    assertEquals(result.attribute("iterateCount"), None, "stale iterateCount must not survive")
    assertEquals(result.attribute("iterateDelta"), None)
  }

  test("GH-400: Some fact fields overlay preserved attributes in place (no duplicates)") {
    val pr = parseElem("""<calcPr calcId="150000" calcMode="auto" refMode="A1"/>""")
    val model = Some(
      CalcPr(calcMode = Some(CalcMode.Manual), fullCalcOnLoad = Some(true), calcId = Some(191029))
    )
    val result =
      OoxmlWorkbook.reconcileCalcPr(Some(pr), model).getOrElse(fail("expected an element"))
    assertEquals(result \@ "calcMode", "manual")
    assertEquals(result \@ "fullCalcOnLoad", "1")
    assertEquals(result \@ "calcId", "191029")
    assertEquals(result \@ "refMode", "A1")
    assertEquals(
      "calcMode=".r.findAllIn(result.toString).size,
      1,
      s"duplicate calcMode: ${result.toString}"
    )
    assertEquals(
      "calcId=".r.findAllIn(result.toString).size,
      1,
      s"duplicate calcId: ${result.toString}"
    )
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

  test("GH-373: reconcileCalcPr with a cleared model strips the iterate triple only") {
    // a cleared model (removeCalcPr) has no opinion on the GH-400 facts — they survive, like
    // every still-unmodeled attribute; only the iterate triple is removed
    val pr = parseElem(
      """<calcPr calcId="191029" calcMode="autoNoTable" fullCalcOnLoad="1" iterate="1" iterateCount="100" refMode="A1"/>"""
    )
    val result =
      OoxmlWorkbook.reconcileCalcPr(Some(pr), None).getOrElse(fail("expected an element"))
    assertEquals(result \@ "calcId", "191029")
    assertEquals(result \@ "calcMode", "autoNoTable")
    assertEquals(result \@ "fullCalcOnLoad", "1")
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

  test("GH-400: every CalcMode token and both fullCalcOnLoad values round-trip via reconcile") {
    List(
      CalcMode.Manual -> "manual",
      CalcMode.Auto -> "auto",
      CalcMode.AutoNoTable -> "autoNoTable"
    ).foreach { (mode, token) =>
      val model = Some(CalcPr(calcMode = Some(mode)))
      val result = OoxmlWorkbook.reconcileCalcPr(None, model).getOrElse(fail("expected an element"))
      assertEquals(result \@ "calcMode", token)
      assertEquals(OoxmlWorkbook.parseCalcPr(Some(result)), model, s"round-trip failed for $mode")
    }
    List(true -> "1", false -> "0").foreach { (b, token) =>
      val model = Some(CalcPr(fullCalcOnLoad = Some(b)))
      val result = OoxmlWorkbook.reconcileCalcPr(None, model).getOrElse(fail("expected an element"))
      assertEquals(result \@ "fullCalcOnLoad", token)
      assertEquals(OoxmlWorkbook.parseCalcPr(Some(result)), model, s"round-trip failed for $b")
    }
  }
