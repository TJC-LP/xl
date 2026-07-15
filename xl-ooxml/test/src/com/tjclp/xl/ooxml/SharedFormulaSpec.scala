package com.tjclp.xl.ooxml

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets
import java.nio.file.Files
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}

import scala.xml.{Elem, XML}

import munit.FunSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}

/** GH-370: shared formula dependents remain formulas through read and sheet regeneration. */
class SharedFormulaSpec extends FunSuite:

  test("shared formula translation is anchor-aware and lexical") {
    val master = SharedFormula.Master(
      cellRef("B2"),
      """A1+$A1+A$1+$A$1+SUM(A1:B2)+'Other 1'!C1+"A1"+LOG10(A1)+""" +
        """'[Book.xlsx]S 1'!D2+[1]S1!A1+T1[A1]+A:A+$A:B+1:2+$1:2+S1:S2!A1+""" +
        """Rate?A1+€A1+A١"""
    )

    assertEquals(
      SharedFormula.translate(master, cellRef("C3")),
      """B2+$A2+B$1+$A$1+SUM(B2:C3)+'Other 1'!D2+"A1"+LOG10(B2)+""" +
        """'[Book.xlsx]S 1'!E3+[1]S1!B2+T1[A1]+B:B+$A:C+2:3+$1:3+S1:S2!B2+""" +
        """Rate?A1+€A1+A١"""
    )
  }

  test("shared formula indexes normalize the xsd:unsignedInt lexical space") {
    assertEquals(SharedFormula.parseIndex(" 0001 "), Some(1L))
    assertEquals(SharedFormula.parseIndex("4294967295"), Some(4_294_967_295L))
    assertEquals(SharedFormula.parseIndex("-1"), None)
    assertEquals(SharedFormula.parseIndex("4294967296"), None)
  }

  test("shared formula translation renders out-of-bounds references as #REF!") {
    val master = SharedFormula.Master(cellRef("B2"), "A1+B2+$A1+A$1+SUM(A1:B2)")
    assertEquals(
      SharedFormula.translate(master, cellRef("A1")),
      "#REF!+A1+#REF!+#REF!+SUM(#REF!)"
    )
  }

  test("DOM reader resolves a dependent even when its master appears later") {
    val worksheet = XML.loadString(
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1"><c r="B1"><f t="shared" si="1">Z99</f><v>2</v></c></row>
        |    <row r="2"><c r="B2"><f t="shared" si="01" ref="B1:B2">A2*2</f><v>4</v></c></row>
        |  </sheetData>
        |</worksheet>""".stripMargin
    )

    val parsed = OoxmlWorksheet.fromXml(worksheet).fold(error => fail(error), identity)
    assertEquals(valueAt(parsed, "B1"), CellValue.Formula("A1*2", Some(CellValue.Number(2))))
    assertEquals(valueAt(parsed, "B2"), CellValue.Formula("A2*2", Some(CellValue.Number(4))))
  }

  test("malformed shared groups stay formula-shaped and duplicate masters keep the first") {
    val worksheet = XML.loadString(
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1">
        |      <c r="A1"><f t="shared"/><v>1</v></c>
        |      <c r="B1"><f t="shared" si="4" ref="B1:B2">A1</f><v>1</v></c>
        |      <c r="C1"><f t="shared" si="4" ref="C1:C2">Z1</f><v>1</v></c>
        |      <c r="D1"><f t="shared">D9+1</f><v>2</v></c>
        |    </row>
        |    <row r="2">
        |      <c r="B2"><f t="shared" si="4"/><v>2</v></c>
        |      <c r="E2"><f t="shared" si="99"/><v>9</v></c>
        |      <c r="F2"><f t="shared" si="4"/><v>8</v></c>
        |    </row>
        |  </sheetData>
        |</worksheet>""".stripMargin
    )

    val parsed = OoxmlWorksheet.fromXml(worksheet).fold(error => fail(error), identity)
    assertEquals(valueAt(parsed, "A1"), CellValue.Formula("#REF!", Some(CellValue.Number(1))))
    assertEquals(valueAt(parsed, "B2"), CellValue.Formula("A2", Some(CellValue.Number(2))))
    assertEquals(valueAt(parsed, "D1"), CellValue.Formula("D9+1", Some(CellValue.Number(2))))
    assertEquals(valueAt(parsed, "E2"), CellValue.Formula("#REF!", Some(CellValue.Number(9))))
    assertEquals(valueAt(parsed, "F2"), CellValue.Formula("#REF!", Some(CellValue.Number(8))))
  }

  test("shared dependents retain numeric, boolean, string, and error caches") {
    val worksheet = XML.loadString(
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1">
        |      <c r="A1"><f t="shared" si="0" ref="A1:A2">1</f><v>1</v></c>
        |      <c r="B1" t="b"><f t="shared" si="1" ref="B1:B2">TRUE</f><v>1</v></c>
        |      <c r="C1" t="str"><f t="shared" si="2" ref="C1:C2">&quot;x&quot;</f><v>x</v></c>
        |      <c r="D1" t="e"><f t="shared" si="3" ref="D1:D2">1/0</f><v>#DIV/0!</v></c>
        |    </row>
        |    <row r="2">
        |      <c r="A2"><f t="shared" si="0"/><v>2</v></c>
        |      <c r="B2" t="b"><f t="shared" si="1"/><v>0</v></c>
        |      <c r="C2" t="str"><f t="shared" si="2"/><v>y</v></c>
        |      <c r="D2" t="e"><f t="shared" si="3"/><v>#DIV/0!</v></c>
        |    </row>
        |  </sheetData>
        |</worksheet>""".stripMargin
    )

    val parsed = OoxmlWorksheet.fromXml(worksheet).fold(error => fail(error), identity)
    assertEquals(valueAt(parsed, "A2"), CellValue.Formula("1", Some(CellValue.Number(2))))
    assertEquals(valueAt(parsed, "B2"), CellValue.Formula("TRUE", Some(CellValue.Bool(false))))
    assertEquals(valueAt(parsed, "C2"), CellValue.Formula("\"x\"", Some(CellValue.Text("y"))))
    assertEquals(
      valueAt(parsed, "D2"),
      CellValue.Formula("1/0", Some(CellValue.Error(CellError.Div0)))
    )
  }

  List(
    "ScalaXml" -> WriterConfig(backend = XmlBackend.ScalaXml),
    "SaxStax" -> WriterConfig(backend = XmlBackend.SaxStax)
  ).foreach { case (backendName, config) =>
    test(
      s"read-edit-rewrite expands shared dependents instead of converting them to constants ($backendName)"
    ) {
      verifyReadEditRewrite(config)
    }
  }

  private def verifyReadEditRewrite(config: WriterConfig): Unit =
    val input = rawXlsx(
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1">
        |      <c r="A1"><v>1</v></c>
        |      <c r="B1"><f t="shared" si="0" ref="B1:B3">A1*2</f><v>2</v></c>
        |    </row>
        |    <row r="2"><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>
        |    <row r="3"><c r="B3"><f t="shared" si="0"/><v>6</v></c></row>
        |  </sheetData>
        |</worksheet>""".stripMargin
    )

    val workbook = XlsxReader.readFromBytes(input).fold(err => fail(err.message), identity)
    val originalSheet = workbook.sheets.headOption.getOrElse(fail("missing Sheet1"))
    assertEquals(originalSheet(ref"B2").value, formula("A2*2", 4))
    assertEquals(originalSheet(ref"B3").value, formula("A3*2", 6))

    val edited = workbook.put(originalSheet.put(ref"A1", CellValue.Number(10)))
    val outputPath = Files.createTempFile("xl-shared-formula-rewrite-", ".xlsx")
    XlsxWriter.writeWith(edited, outputPath, config).fold(err => fail(err.message), identity)
    val output = Files.readAllBytes(outputPath)
    Files.deleteIfExists(outputPath)
    val outputSheetXml = XML.loadString(zipEntry(output, "xl/worksheets/sheet1.xml"))

    val formulas = (outputSheetXml \\ "c").collect { case cell: Elem =>
      val address = cell \@ "r"
      val expression = (cell \ "f").headOption.map(_.text)
      address -> expression
    }.toMap
    assertEquals(formulas.get("B1"), Some(Some("A1*2")))
    assertEquals(formulas.get("B2"), Some(Some("A2*2")))
    assertEquals(formulas.get("B3"), Some(Some("A3*2")))

    val reread = XlsxReader.readFromBytes(output).fold(err => fail(err.message), identity)
    val rewrittenSheet = reread.sheets.headOption.getOrElse(fail("missing rewritten Sheet1"))
    assertEquals(rewrittenSheet(ref"B2").value, formula("A2*2", 4))
    assertEquals(rewrittenSheet(ref"B3").value, formula("A3*2", 6))

  private def valueAt(worksheet: OoxmlWorksheet, address: String): CellValue =
    worksheet.rows
      .flatMap(_.cells)
      .find(_.ref == cellRef(address))
      .map(_.value)
      .getOrElse(fail(s"missing $address"))

  private def cellRef(address: String): ARef =
    ARef.parse(address).fold(error => fail(error), identity)

  private def formula(expression: String, cached: Int): CellValue =
    CellValue.Formula(expression, Some(CellValue.Number(BigDecimal(cached))))

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
