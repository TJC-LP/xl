package com.tjclp.xl.ooxml

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets
import java.nio.file.Files
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}

import scala.xml.XML

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

import com.tjclp.xl.Generators
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}

/**
 * GH-430: `<f t="array">` and `<f t="dataTable">` formula records survive read -> model -> DIRTY
 * write. The sibling `put` forces full sheet regeneration, so verbatim passthrough is not a pass.
 * Also pins the reader's lenient-total junk-attr fallbacks and the exact-kind round-trip law.
 */
class FormulaRecordPreservationSpec extends ScalaCheckSuite:

  /** Issue exemplars, verbatim: single-cell CSE array + 2-D data table (incl. cached #NUM!). */
  private val recordsSheetXml =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      |  <sheetData>
      |    <row r="1">
      |      <c r="A1"><v>1</v></c>
      |      <c r="B1"><v>7</v></c>
      |      <c r="C1" t="n"><f t="array" ref="C1:C1">SUM(A1:A3*10)</f><v>60</v></c>
      |    </row>
      |    <row r="2">
      |      <c r="A2"><v>2</v></c>
      |      <c r="F2"><f t="dataTable" ref="F2:G2" dt2D="1" dtr="0" r1="A1" r2="A2"/><v>42</v></c>
      |      <c r="G2" t="e"><f t="dataTable" ref="F2:G2" dt2D="1" dtr="0" r1="A1" r2="A2" ca="1"/><v>#NUM!</v></c>
      |    </row>
      |    <row r="3"><c r="A3"><v>3</v></c></row>
      |  </sheetData>
      |</worksheet>""".stripMargin

  List(
    "ScalaXml" -> WriterConfig(backend = XmlBackend.ScalaXml),
    "SaxStax" -> WriterConfig(backend = XmlBackend.SaxStax)
  ).foreach { case (backendName, config) =>
    test(s"dirty regeneration preserves array + dataTable records byte-exactly ($backendName)") {
      val workbook =
        XlsxReader.readFromBytes(rawXlsx(recordsSheetXml)).fold(err => fail(err.message), identity)
      val sheet = workbook.sheets.headOption.getOrElse(fail("missing Sheet1"))

      // Sibling put forces model regeneration of the whole sheet (the GH-430 failure mode).
      val edited = workbook.put(sheet.put(ref"A5", CellValue.Text("dirty")))
      val outputPath = Files.createTempFile("xl-formula-records-", ".xlsx")
      XlsxWriter.writeWith(edited, outputPath, config).fold(err => fail(err.message), identity)
      val output = Files.readAllBytes(outputPath)
      Files.deleteIfExists(outputPath)
      val sheetXml = zipEntry(output, "xl/worksheets/sheet1.xml")

      // Record byte law: emitted <f> elements equal the source exemplars.
      assert(
        sheetXml.contains("""<f t="array" ref="C1:C1">SUM(A1:A3*10)</f>"""),
        s"array record degraded: $sheetXml"
      )
      assert(
        sheetXml.contains("""<f t="dataTable" ref="F2:G2" dt2D="1" dtr="0" r1="A1" r2="A2"/>"""),
        s"dataTable record degraded/baked: $sheetXml"
      )
      assert(
        sheetXml.contains(
          """<f t="dataTable" ref="F2:G2" dt2D="1" dtr="0" r1="A1" r2="A2" ca="1"/>"""
        ),
        s"dataTable ca record degraded/baked: $sheetXml"
      )
      // Cached <v> seeds intact (incl. the #NUM! Sentinel interior).
      assert(sheetXml.contains("<v>42</v>"), s"dataTable cache lost: $sheetXml")
      assert(sheetXml.contains("<v>#NUM!</v>"), s"cached #NUM! interior lost: $sheetXml")
      assert(sheetXml.contains("<v>60</v>"), s"array cache lost: $sheetXml")

      // Semantic law: the output re-reads with the records still formula-shaped.
      val reread = XlsxReader.readFromBytes(output).fold(err => fail(err.message), identity)
      val rereadSheet = reread.sheets.headOption.getOrElse(fail("missing reread Sheet1"))
      rereadSheet(ref"F2").value match
        case f: CellValue.Formula =>
          assertEquals(f.cachedValue, Some(CellValue.Number(BigDecimal(42))))
        case other => fail(s"F2 baked to constant on re-read: $other")
      rereadSheet(ref"C1").value match
        case f: CellValue.Formula =>
          assertEquals(f.expression, "SUM(A1:A3*10)")
          assertEquals(f.cachedValue, Some(CellValue.Number(BigDecimal(60))))
        case other => fail(s"C1 lost formula-hood on re-read: $other")
    }
  }

  test("formula-records.xlsx: clean read->write copies the worksheet verbatim") {
    val in = TestFixtures.copyToTemp("formula-records.xlsx")
    val wb = XlsxReader.read(in).fold(err => fail(err.message), identity)
    val out = Files.createTempFile("xl-records-clean-", ".xlsx")
    try
      XlsxWriter.write(wb, out).fold(err => fail(err.message), identity)
      assertEquals(
        zipEntry(Files.readAllBytes(out), "xl/worksheets/sheet1.xml"),
        zipEntry(Files.readAllBytes(in), "xl/worksheets/sheet1.xml"),
        "clean write must ride the verbatim-copy loop byte-identically"
      )
    finally Files.deleteIfExists(out)
  }

  test("formula-records.xlsx: dirty regeneration keeps every record and cached <v> seed") {
    val in = TestFixtures.copyToTemp("formula-records.xlsx")
    val wb = XlsxReader.read(in).fold(err => fail(err.message), identity)
    val sheet = wb.sheets.headOption.getOrElse(fail("missing Records sheet"))
    val edited = wb.put(sheet.put(ref"A20", CellValue.Text("dirty")))
    val out = Files.createTempFile("xl-records-dirty-", ".xlsx")
    try
      XlsxWriter.write(edited, out).fold(err => fail(err.message), identity)
      val sheetXml = zipEntry(Files.readAllBytes(out), "xl/worksheets/sheet1.xml")
      List(
        """<f t="dataTable" ref="F2:G3" dt2D="1" dtr="0" r1="A1" r2="A2"/>""",
        """<f t="dataTable" ref="F2:G3" dt2D="1" dtr="0" r1="A1" r2="A2" ca="1"/>""",
        """<f t="dataTable" ref="H2:H4" dt2D="0" dtr="0" r1="G1"/>""",
        """<f t="dataTable" ref="B6:D6" dt2D="0" dtr="1" r1="A6"/>""",
        """<f t="array" ref="C8:C9">A8:A9*10</f>""",
        """<f t="array" ref="E10:E10">SUM(A8:A9*10)</f>""",
        "<v>#NUM!</v>"
      ).foreach { needle =>
        assert(sheetXml.contains(needle), s"record lost on dirty regen: $needle\n$sheetXml")
      }
      // The non-anchor array member stays a cached constant (per-cell faithfulness).
      assert(sheetXml.contains("""<c r="C9" t="n"><v>20</v></c>"""), sheetXml)
    finally Files.deleteIfExists(out)
  }

  test("DOM reader models record kinds per cell (issue exemplars)") {
    val worksheet = XML.loadString(
      recordsSheetXml.linesIterator.filterNot(_.startsWith("<?xml")).mkString("\n")
    )
    val parsed = OoxmlWorksheet.fromXml(worksheet).fold(error => fail(error), identity)

    val c1 = CellRange.parse("C1:C1").fold(fail(_), identity)
    val f2g2 = CellRange.parse("F2:G2").fold(fail(_), identity)
    val a1 = cellRef("A1")
    val a2 = cellRef("A2")

    assertEquals(
      valueAt(parsed, "C1"),
      CellValue.Formula(
        "SUM(A1:A3*10)",
        Some(CellValue.Number(BigDecimal(60))),
        FormulaKind.ArrayFormula(c1)
      )
    )
    assertEquals(
      valueAt(parsed, "F2"),
      CellValue.Formula(
        "TABLE(A1,A2)",
        Some(CellValue.Number(BigDecimal(42))),
        FormulaKind.DataTable(f2g2, dt2D = true, dtr = false, r1 = Some(a1), r2 = Some(a2))
      )
    )
    assertEquals(
      valueAt(parsed, "G2"),
      CellValue.Formula(
        "TABLE(A1,A2)",
        Some(CellValue.Error(CellError.Num)),
        FormulaKind
          .DataTable(f2g2, dt2D = true, dtr = false, r1 = Some(a1), r2 = Some(a2), ca = true)
      )
    )
  }

  test("junk record attrs degrade visibly and never throw (lenient totality)") {
    val worksheet = XML.loadString(
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1">
        |      <c r="A1"><f t="array">A9+1</f><v>1</v></c>
        |      <c r="B1"><f t="array" ref="NOT_A_RANGE">B9+1</f><v>2</v></c>
        |      <c r="C1"><f t="array" ref="C1:C1"/><v>3</v></c>
        |      <c r="D1"><f t="dataTable" dt2D="1" dtr="0" r1="A1" r2="A2"/><v>4</v></c>
        |      <c r="E1"><f t="dataTable" ref="E1:E2" dt2D="banana" dtr="true" r1="NOPE" r2="A2" ca="true"/><v>5</v></c>
        |      <c r="F1"><f t="normal">F9+1</f><v>6</v></c>
        |    </row>
        |  </sheetData>
        |</worksheet>""".stripMargin
    )
    val parsed = OoxmlWorksheet.fromXml(worksheet).fold(error => fail(error), identity)

    // Array with no/corrupt ref: degrade to Normal, keep the text (never invent a range).
    assertEquals(
      valueAt(parsed, "A1"),
      CellValue.Formula("A9+1", Some(CellValue.Number(BigDecimal(1))))
    )
    assertEquals(
      valueAt(parsed, "B1"),
      CellValue.Formula("B9+1", Some(CellValue.Number(BigDecimal(2))))
    )
    // Array with empty text: nothing to preserve — cached constant (today's behavior).
    assertEquals(valueAt(parsed, "C1"), CellValue.Number(BigDecimal(3)))
    // DataTable without a parsable ref: cached constant (today's exact behavior).
    assertEquals(valueAt(parsed, "D1"), CellValue.Number(BigDecimal(4)))
    // LibreOffice boolean lexicals parse; junk booleans ("banana") and junk inputs ("NOPE")
    // fall to schema defaults (false / None). The 1-D display slot renders empty for a
    // missing input: TABLE(,).
    val e1e2 = CellRange.parse("E1:E2").fold(fail(_), identity)
    assertEquals(
      valueAt(parsed, "E1"),
      CellValue.Formula(
        "TABLE(,)",
        Some(CellValue.Number(BigDecimal(5))),
        FormulaKind.DataTable(
          e1e2,
          dt2D = false,
          dtr = true,
          r1 = None,
          r2 = Some(cellRef("A2")),
          ca = true
        )
      )
    )
    // Explicit t="normal" is a plain formula.
    assertEquals(
      valueAt(parsed, "F1"),
      CellValue.Formula("F9+1", Some(CellValue.Number(BigDecimal(6))))
    )
  }

  // Exact-kind round-trip law: unlike WorkbookEquivalence (text-only formula compare), this
  // asserts full CellValue equality including the record kind, on BOTH writer backends.
  private val genRecordFormula: Gen[CellValue] =
    for
      kind <- Generators.genFormulaKind
      cached <- Gen.option(Gen.choose(-1000000, 1000000).map(n => CellValue.Number(BigDecimal(n))))
    yield kind match
      case dt: FormulaKind.DataTable => CellValue.dataTable(dt, cached)
      case other => CellValue.Formula("SUM(A1:B2)", cached, other)

  List(
    "ScalaXml" -> WriterConfig(backend = XmlBackend.ScalaXml),
    "SaxStax" -> WriterConfig(backend = XmlBackend.SaxStax)
  ).foreach { case (backendName, config) =>
    property(s"forAll genFormulaKind: write -> read is identity on the kind ($backendName)") {
      forAll(genRecordFormula) { (value: CellValue) =>
        val wb = Workbook(Sheet(SheetName.unsafe("K")).put(ref"B2", value))
        val outputPath = Files.createTempFile("xl-record-prop-", ".xlsx")
        try
          XlsxWriter
            .writeWith(wb, outputPath, config)
            .fold(err => fail(s"write failed: ${err.message}"), identity)
          val reread = XlsxReader.read(outputPath).fold(err => fail(err.message), identity)
          val sheet = reread.sheets.headOption.getOrElse(fail("missing sheet"))
          assertEquals(sheet(ref"B2").value, value)
          true
        finally Files.deleteIfExists(outputPath)
      }
    }
  }

  private def valueAt(worksheet: OoxmlWorksheet, address: String): CellValue =
    worksheet.rows
      .flatMap(_.cells)
      .find(_.ref == cellRef(address))
      .map(_.value)
      .getOrElse(fail(s"missing $address"))

  private def cellRef(address: String): ARef =
    ARef.parse(address).fold(error => fail(error), identity)

  private val workbookXml =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      |          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      |  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
      |</workbook>""".stripMargin

  private val workbookRelsXml =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      |  <Relationship Id="rId1"
      |    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
      |    Target="worksheets/sheet1.xml"/>
      |</Relationships>""".stripMargin

  private val rootRelsXml =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      |  <Relationship Id="rId1"
      |    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
      |    Target="xl/workbook.xml"/>
      |</Relationships>""".stripMargin

  private val contentTypesXml =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      |  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      |  <Default Extension="xml" ContentType="application/xml"/>
      |  <Override PartName="/xl/workbook.xml"
      |    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      |  <Override PartName="/xl/worksheets/sheet1.xml"
      |    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      |</Types>""".stripMargin

  private def rawXlsx(sheetXml: String): Array[Byte] =
    val output = new ByteArrayOutputStream()
    val zip = new ZipOutputStream(output)
    try
      List(
        "[Content_Types].xml" -> contentTypesXml,
        "_rels/.rels" -> rootRelsXml,
        "xl/workbook.xml" -> workbookXml,
        "xl/_rels/workbook.xml.rels" -> workbookRelsXml,
        "xl/worksheets/sheet1.xml" -> sheetXml
      ).foreach { case (name, content) =>
        zip.putNextEntry(new ZipEntry(name))
        zip.write(content.getBytes(StandardCharsets.UTF_8))
        zip.closeEntry()
      }
    finally zip.close()
    output.toByteArray

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def zipEntry(bytes: Array[Byte], name: String): String =
    val zip = new ZipInputStream(new ByteArrayInputStream(bytes))
    var current = zip.getNextEntry
    var result: Option[String] = None
    try
      while current != null && result.isEmpty do
        if current.getName == name then
          result = Some(new String(zip.readAllBytes(), StandardCharsets.UTF_8))
        else
          zip.closeEntry()
          current = zip.getNextEntry
    finally zip.close()
    result.getOrElse(fail(s"missing ZIP entry $name"))
