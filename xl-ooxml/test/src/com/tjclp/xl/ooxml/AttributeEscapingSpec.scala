package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
import scala.xml.Elem

import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.forAll

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import com.tjclp.xl.sheets.{DataValidation, DvKind, DvMessages}
import com.tjclp.xl.workbooks.DefinedName

/**
 * GH-649: TAB, LF and CR inside an attribute value are written as the character references `&#9;`,
 * `&#10;` and `&#13;` by both writers. scala.xml and the JDK StAX writer emit them raw, and XML
 * attribute-value normalization (XML 1.0 §3.3.3) turns a raw one into a space on the next parse, so
 * a data-validation prompt with a line break came back as `l1 l2` — from a typed entry and, worse,
 * from an Excel-authored Preserved entry whose `&#10;` xl had already parsed. Excel writes the
 * references; so does xl now, and the two backends agree byte for byte.
 */
class AttributeEscapingSpec extends ScalaCheckSuite:

  private val decl = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"""

  private def parse(xml: String): Elem =
    XmlSecurity.parseSafe(xml, "test").fold(e => fail(s"parse failed: ${e.message}"), identity)

  /** Legal XML characters incl. every escaping hazard; no surrogates, no illegal controls. */
  private val genAttrText: Gen[String] =
    Gen
      .listOf(
        Gen.frequency(
          6 -> Gen.alphaNumChar,
          2 -> Gen.const(' '),
          1 -> Gen.oneOf('\n', '\r', '\t'),
          1 -> Gen.oneOf('"', '\'', '<', '>', '&', '_', 'ü', '✓')
        )
      )
      .map(_.mkString)

  private def staxBytes(body: StaxSaxWriter => Unit): String =
    val out = new ByteArrayOutputStream()
    val w = StaxSaxWriter.create(out)
    w.startDocument()
    body(w)
    w.endDocument()
    w.flush()
    new String(out.toByteArray, StandardCharsets.UTF_8)

  /** The document after its XML declaration (the two backends spell the declaration apart). */
  private def afterDecl(xml: String): String = xml.dropWhile(_ != '>').drop(1).stripLeading

  // ===== XmlUtil.escapeAttr / compact =====

  test("escapeAttr: named references for &, <, >, \" and character references for TAB, LF, CR") {
    assertEquals(
      XmlUtil.escapeAttr("a&b<c>d\"e\tf\ng\rh'i"),
      "a&amp;b&lt;c&gt;d&quot;e&#9;f&#10;g&#13;h'i"
    )
  }

  test("escapeAttr drops XML-illegal control characters like escapeText does (GH-237)") {
    assertEquals(XmlUtil.escapeAttr("a" + 1.toChar + "b" + 0x0b.toChar), "ab")
  }

  test("compact writes a line break in an attribute value as &#10; (Excel's spelling)") {
    val e = XmlUtil.elem("x", "v" -> "l1\nl2\tt\rr")()
    assertEquals(XmlUtil.compact(e), s"$decl\n<x v=\"l1&#10;l2&#9;t&#13;r\"/>")
  }

  test("compact keeps scala.xml's attribute rendering: chain order, pre:key, &quot;") {
    val e = parse(
      """<a xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" z="1" xr:uid="{u}" b="say &quot;hi&quot;"/>"""
    )
    // the same bytes scala.xml renders when no attribute value carries TAB/LF/CR
    assertEquals(XmlUtil.compact(e), s"$decl\n${e.toString}")
    assert(e.toString.contains("xr:uid=\"{u}\""), e.toString)
    assert(e.toString.contains("b=\"say &quot;hi&quot;\""), e.toString)
  }

  property("compact . parse recovers every attribute value (TAB/LF/CR included)") {
    forAll(genAttrText) { s =>
      val back = parse(XmlUtil.compact(XmlUtil.elem("x", "v" -> s)()))
      Prop(back.attribute("v").map(_.text) == Some(s)) :| s"value: ${s.map(_.toInt).mkString(",")}"
    }
  }

  // ===== StAX backend =====

  test("the StAX backend writes &#10;/&#9;/&#13; in attribute values and agrees with compact") {
    val e = XmlUtil.elem("x", "v" -> "l1\nl2\tt\rr \"q\" <&>")()
    // emptyElement: the minimized `<x .../>` form compact writes for a childless element
    val stax = staxBytes(_.emptyElement("x", Seq("v" -> "l1\nl2\tt\rr \"q\" <&>")))
    assert(stax.contains("v=\"l1&#10;l2&#9;t&#13;r &quot;q&quot; &lt;&amp;&gt;\""), stax)
    assertEquals(afterDecl(stax), afterDecl(XmlUtil.compact(e)))
  }

  test("the StAX backend still escapes text as before: &, <, > only, no &quot;, controls dropped") {
    val stax = staxBytes { w =>
      w.startElement("t")
      w.writeCharacters("a\"b<c&d>e 'f'\n" + 1.toChar)
      w.endElement()
    }
    assertEquals(afterDecl(stax), "<t>a\"b&lt;c&amp;d&gt;e 'f'\n</t>")
  }

  property("StAX . parse recovers every attribute value, and matches compact byte for byte") {
    forAll(genAttrText) { s =>
      val stax = staxBytes(_.emptyElement("x", Seq("v" -> s)))
      val back = parse(stax)
      Prop(back.attribute("v").map(_.text) == Some(s)) :| "round-trip" &&
      Prop(afterDecl(stax) == afterDecl(XmlUtil.compact(XmlUtil.elem("x", "v" -> s)()))) :| "parity"
    }
  }

  // ===== end to end: data-validation prompts and defined-name comments =====

  private val backends = List(
    "ScalaXml" -> WriterConfig(backend = XmlBackend.ScalaXml),
    "SaxStax" -> WriterConfig(backend = XmlBackend.SaxStax)
  )

  private def writeTo(wb: Workbook, label: String, config: WriterConfig): Path =
    val out = Files.createTempFile(s"xl-gh649-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter.writeWith(wb, out, config).fold(e => fail(s"write failed: ${e.message}"), _ => ())
    out

  private def reread(path: Path): Workbook =
    XlsxReader.read(path).fold(e => fail(s"re-read failed: ${e.message}"), identity)

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  private val multilinePrompt = DvMessages(
    showInputMessage = true,
    promptTitle = Some("Two lines"),
    prompt = Some("line 1\nline 2\twith a tab")
  )

  backends.foreach { case (name, config) =>
    test(s"a typed validation's multiline prompt is written as &#10; and reads back ($name)") {
      val sheet = Sheet(SheetName.unsafe("Inputs"))
        .put(ref"A1", CellValue.Number(BigDecimal(1)))
        .withDataValidation(
          ref"B2:B9",
          DataValidation.Rules(Vector.empty, DvKind.List("\"a,b\""), messages = multilinePrompt)
        )
      val wb = Workbook(Vector(sheet))
      val out = writeTo(wb, s"typed-$name", config)
      val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
      assert(sheetXml.contains("prompt=\"line 1&#10;line 2&#9;with a tab\""), sheetXml)
      assert(!sheetXml.contains("_x000A_"), sheetXml)
      assertEquals(reread(out).sheets(0).dataValidations, wb.sheets(0).dataValidations)
    }

    test(s"a defined name's comment keeps its line break ($name)") {
      val sheet = Sheet(SheetName.unsafe("Inputs")).put(ref"A1", CellValue.Number(BigDecimal(1)))
      val base = Workbook(Vector(sheet))
      val named = DefinedName("Rate", "Inputs!$A$1", comment = Some("first\nsecond"))
      val wb = base.copy(metadata = base.metadata.copy(definedNames = Vector(named)))
      val out = writeTo(wb, s"name-$name", config)
      val workbookXml = entryText(out, "xl/workbook.xml")
      assert(workbookXml.contains("comment=\"first&#10;second\""), workbookXml)
      assertEquals(
        reread(out).metadata.definedNames.map(_.comment),
        Vector(Some("first\nsecond"))
      )
    }
  }

  /**
   * The lossy case before GH-649: an Excel-authored entry whose `imeMode` keeps it out of the typed
   * subset rides `DataValidation.Preserved` as canonical XML. Excel wrote `prompt="l1&#10;l2"`; the
   * parse resolved the reference, the Preserved payload re-serialized it raw, and the emission-time
   * re-parse normalized it to `l1 l2`.
   */
  test(
    "an Excel-authored (Preserved) validation keeps its prompt's line break through a dirty write"
  ) {
    val excelDv =
      """<dataValidations count="1">
        |<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Pick" prompt="l1&#10;l2&#9;t" imeMode="hiragana" sqref="D2:D9"><formula1>"a,b"</formula1></dataValidation>
        |</dataValidations>""".stripMargin
    val in = rawDvFixture(excelDv)
    val wb = reread(in)
    wb.sheets(0).dataValidations match
      case Vector(DataValidation.Preserved(xml)) =>
        assert(xml.contains("prompt=\"l1&#10;l2&#9;t\""), xml)
      case other => fail(s"imeMode must force Preserved: $other")
    // an authored entry beside it dirties the slot: the container is regenerated, not passed through
    val updated = wb
      .update(wb.sheets(0).name, _.withDataValidation(ref"B2:B9", DataValidation.listOf("x", "y")))
      .fold(e => fail(s"update failed: $e"), identity)
    backends.foreach { case (name, config) =>
      val out = writeTo(updated, s"preserved-$name", config)
      val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
      assert(sheetXml.contains("prompt=\"l1&#10;l2&#9;t\""), s"$name: $sheetXml")
      assert(!sheetXml.contains("prompt=\"l1 l2 t\""), s"$name: $sheetXml")
      val back = reread(out).sheets(0).dataValidations
      back.headOption match
        case Some(DataValidation.Preserved(xml)) =>
          assert(xml.contains("prompt=\"l1&#10;l2&#9;t\""), s"$name: $xml")
        case other => fail(s"$name: the Excel entry must still ride Preserved: $other")
      // the DOM backend re-emits the payload verbatim; the StAX backend sorts attributes on
      // preserved fragments (the SaxSupport contract), so the payload text is compared only here
      if name == "ScalaXml" then assertEquals(back, updated.sheets(0).dataValidations)
    }
  }

  test(
    "an Excel-authored typed validation's &#10; prompt parses to the newline and is re-spelt &#10;"
  ) {
    val excelDv =
      """<dataValidations count="1">
        |<dataValidation type="list" allowBlank="1" showInputMessage="1" promptTitle="Pick" prompt="l1&#10;l2" sqref="D2:D9"><formula1>"a,b"</formula1></dataValidation>
        |</dataValidations>""".stripMargin
    val wb = reread(rawDvFixture(excelDv))
    wb.sheets(0).dataValidations match
      case Vector(r: DataValidation.Rules) => assertEquals(r.messages.prompt, Some("l1\nl2"))
      case other => fail(s"expected a typed entry: $other")
    val updated = wb
      .update(wb.sheets(0).name, _.withDataValidation(ref"B2:B9", DataValidation.listOf("x")))
      .fold(e => fail(s"update failed: $e"), identity)
    val out = writeTo(updated, "excel-typed", WriterConfig())
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(sheetXml.contains("prompt=\"l1&#10;l2\""), sheetXml)
    assertEquals(reread(out).sheets(0).dataValidations, updated.sheets(0).dataValidations)
  }

  // ===== fixture: a single-sheet book whose worksheet carries the given <dataValidations> =====

  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    out.putNextEntry(new ZipEntry(name))
    out.write(content.getBytes(StandardCharsets.UTF_8))
    out.closeEntry()

  private def rawDvFixture(dataValidationsXml: String): Path =
    val path = Files.createTempFile("gh649-fixture", ".xlsx")
    path.toFile.deleteOnExit()
    val out = new ZipOutputStream(Files.newOutputStream(path))
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
    <sheet name="Inputs" sheetId="1" r:id="rId1"/>
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
        s"""<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
  $dataValidationsXml
</worksheet>"""
      )
    finally out.close()
    path
