package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
import scala.xml.{Elem, Text}

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import com.tjclp.xl.sheets.{DataValidation, DvKind, DvMessages}
import com.tjclp.xl.workbooks.DefinedName

/**
 * GH-649: attribute values escape TAB/LF/CR as the character references `&#9;`, `&#10;`, `&#13;` on
 * BOTH writers. XML 1.0 §3.3.3 attribute-value normalization turns the raw characters into spaces
 * on every re-parse, so a data-validation prompt, a defined-name comment or a table column name
 * with a line break came back as `l1 l2`; Excel and openpyxl write the references. scala.xml's
 * `Utility.escape` and the JDK `XMLStreamWriter` both wrote them raw — the JDK writer even wrote
 * XML-illegal C0 control characters raw, and re-escaped `&` so a reference could not be pushed
 * through it — so the StAX backend now owns its bytes ([[XmlTagWriter]]) and shares
 * [[XmlUtil.escapeAttr]] / [[XmlUtil.escapeText]] with the DOM backend: parity by construction,
 * pinned here against the retired JDK-backed writer ([[JdkStaxReferenceWriter]]) byte for byte. The
 * end-to-end section writes whole workbooks on both backends: a typed validation's multiline
 * prompt, a defined name's comment, and an Excel-authored `Preserved` entry through a dirty write.
 */
class AttributeEscapingSpec extends ScalaCheckSuite:

  private val decl = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"""
  private val staxDecl = """<?xml version="1.0" encoding="UTF-8"?>"""

  private def stax(body: SaxWriter => Unit): String =
    val out = new ByteArrayOutputStream()
    val w = StaxSaxWriter.create(out)
    body(w)
    w.flush()
    new String(out.toByteArray, StandardCharsets.UTF_8)

  private def reference(body: SaxWriter => Unit): String =
    val out = new ByteArrayOutputStream()
    val w = JdkStaxReferenceWriter.create(out)
    body(w)
    w.flush()
    new String(out.toByteArray, StandardCharsets.UTF_8)

  private def parse(xml: String): Elem =
    XmlSecurity.parseSafe(xml, "test").fold(e => fail(s"parse failed: ${e.message}"), identity)

  private def parseAttr(xml: String, name: String): String =
    XmlSecurity.parseSafe(xml, "x.xml").fold(e => fail(s"parse failed: ${e.message}"), _ \@ name)

  // ===== the fix =====

  test("compact writes TAB/LF/CR in an attribute value as character references") {
    val e = XmlUtil.elem("x", "v" -> "l1\nl2\tt\rr")()
    assertEquals(XmlUtil.compact(e), s"$decl\n<x v=\"l1&#10;l2&#9;t&#13;r\"/>")
  }

  test("the StAX backend writes the identical element bytes") {
    val s = stax { w =>
      w.startDocument()
      w.emptyElement("x", Seq("v" -> "l1\nl2\tt\rr"))
      w.endDocument()
    }
    assertEquals(s, s"""$staxDecl<x v="l1&#10;l2&#9;t&#13;r"/>""")
  }

  test("attributes: & < > \" as entities, ' verbatim, C0 controls dropped — both backends") {
    val raw = "a&b<c>d\"e'f" + 1.toChar + 0x1f.toChar + "z"
    val expected = """<x v="a&amp;b&lt;c&gt;d&quot;e'fz"/>"""
    assertEquals(XmlUtil.compact(XmlUtil.elem("x", "v" -> raw)()), s"$decl\n$expected")
    // an empty element is closed by the following call, as with the JDK writer
    assertEquals(stax { w => w.emptyElement("x", Seq("v" -> raw)); w.endDocument() }, expected)
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

  test("escapeAttr: untouched fast path; the escaped form parses back to the original") {
    assertEquals(XmlUtil.escapeAttr("plain r1 s2 'q'"), "plain r1 s2 'q'")
    val original = "l1\nl2\tt\rr \"q\" & <>"
    assertEquals(parseAttr(s"""<x v="${XmlUtil.escapeAttr(original)}"/>""", "v"), original)
  }

  test("text content is unchanged: TAB/LF verbatim, CR verbatim, quotes verbatim (GH-611)") {
    val e = XmlUtil.elem("t")(Text("a\tb\nc\rd \"q\" & <>"))
    assertEquals(XmlUtil.compact(e), s"$decl\n<t>a\tb\nc\rd \"q\" &amp; &lt;&gt;</t>")
    val s = stax { w =>
      w.startElement("t"); w.writeCharacters("a\tb\nc\rd \"q\" & <>"); w.endElement()
    }
    assertEquals(s, "<t>a\tb\nc\rd \"q\" &amp; &lt;&gt;</t>")
  }

  /**
   * Any BMP text: alphanumerics, markup and quote characters, TAB/LF/CR, C0 controls, non-ASCII.
   */
  private val genAttrText: Gen[String] =
    Gen
      .listOf(
        Gen.frequency(
          8 -> Gen.alphaNumChar,
          2 -> Gen.oneOf(' ', '&', '<', '>', '"', '\'', '_', '-', '.', ';', '#'),
          2 -> Gen.oneOf('\t', '\n', '\r'),
          1 -> Gen.oneOf(1.toChar, 0x1f.toChar, 0x0b.toChar),
          1 -> Gen.oneOf('é', 'ß', '日', '€')
        )
      )
      .map(_.mkString)

  property("law: the parsed attribute equals sanitizeXmlText(s) on both backends") {
    forAll(genAttrText) { s =>
      val dom = XmlUtil.compact(XmlUtil.elem("x", "v" -> s)())
      val sax = stax { w =>
        w.startDocument(); w.emptyElement("x", Seq("v" -> s)); w.endDocument()
      }
      parseAttr(dom, "v") == XmlUtil.sanitizeXmlText(s) &&
      parseAttr(sax, "v") == XmlUtil.sanitizeXmlText(s)
    }
  }

  property("law: both backends render an element with attribute and text byte-identically") {
    forAll(genAttrText, genAttrText) { (a, t) =>
      val dom = XmlUtil.compact(XmlUtil.elem("x", "v" -> a)(Text(t))).stripPrefix(decl + "\n")
      val sax = stax { w =>
        w.startElement("x"); w.writeAttribute("v", a); w.writeCharacters(t); w.endElement()
      }
      // compact minimizes only a childless element; a Text("") child keeps <x ...></x> on both
      sax == dom
    }
  }

  // ===== the retired JDK writer as the byte baseline =====

  private sealed trait Ev
  private final case class El(
    name: String,
    ns: Option[String],
    attrs: List[(String, String)],
    kids: List[Ev]
  ) extends Ev
  private final case class Txt(s: String) extends Ev
  private final case class Empty(name: String, attrs: List[(String, String)]) extends Ev

  /** Attribute values the JDK writer rendered correctly: no TAB/LF/CR, no C0 controls. */
  private val genJdkSafeAttr: Gen[String] =
    Gen
      .listOf(
        Gen.frequency(
          8 -> Gen.alphaNumChar,
          2 -> Gen.oneOf(' ', '&', '<', '>', '"', '\'', '_', '-', '.', '=', ';'),
          1 -> Gen.oneOf('é', '日', '€')
        )
      )
      .map(_.mkString)

  /** Text: anything legal, including CR (raw on both), astral characters and quotes. */
  private val genText: Gen[String] =
    Gen
      .listOf(
        Gen.frequency(
          8 -> Gen.alphaNumChar,
          2 -> Gen.oneOf(' ', '&', '<', '>', '"', '\'', '\t', '\n', '\r'),
          1 -> Gen.oneOf("é", "日", "😀", "€").map(_.head)
        )
      )
      .map(_.mkString)
      .map(_.replace("\ud83d", "😀")) // a lone high surrogate becomes a pair

  private val genName: Gen[String] = Gen.oneOf("a", "b", "row", "c", "worksheet", "x14:f")
  private val genAttrName: Gen[String] =
    Gen.frequency(
      6 -> Gen.oneOf("r", "s", "t", "count", "val"),
      1 -> Gen.const("r:id"),
      1 -> Gen.const("x14ac:dyDescent"),
      1 -> Gen.const("xml:space")
    )
  private val genAttrs: Gen[List[(String, String)]] =
    Gen.choose(0, 3).flatMap(n => Gen.listOfN(n, Gen.zip(genAttrName, genJdkSafeAttr)))

  private def genEl(depth: Int): Gen[El] =
    for
      name <- genName
      ns <- Gen.frequency(4 -> Gen.const(None), 1 -> Gen.const(Some("urn:xl:test")))
      attrs <- genAttrs
      kids <-
        if depth == 0 then Gen.const(Nil)
        else
          Gen.choose(0, 3).flatMap { n =>
            Gen.listOfN(
              n,
              Gen.frequency(
                3 -> genText.map(Txt.apply),
                2 -> genEl(depth - 1),
                1 -> Gen.zip(genName, genAttrs).map(Empty.apply.tupled)
              )
            )
          }
    yield
      // the JDK writer mis-nests an empty element immediately followed by its parent's end tag
      // (it pops the empty element and closes the parent's start tag); xl never emits that
      // sequence, so the reference is only asked about sequences it rendered correctly
      val safeKids = kids.lastOption match
        case Some(_: Empty) => kids :+ Txt("")
        case _ => kids
      El(name, ns, attrs, safeKids)

  private def emit(w: SaxWriter, ev: Ev): Unit = ev match
    case El(name, ns, attrs, kids) =>
      ns.fold(w.startElement(name))(w.startElement(name, _))
      attrs.foreach { case (k, v) => w.writeAttribute(k, v) }
      kids.foreach(emit(w, _))
      w.endElement()
    case Txt(s) => w.writeCharacters(s)
    case Empty(name, attrs) => w.emptyElement(name, attrs)

  property("XmlTagWriter renders every event sequence byte-identically to the JDK writer") {
    forAll(genEl(3)) { root =>
      def doc(w: SaxWriter): Unit =
        w.startDocument()
        emit(w, root)
        w.endDocument()
      val ours = stax(doc)
      val theirs = reference(doc)
      if ours != theirs then println(s"ours:   $ours\ntheirs: $theirs")
      ours == theirs
    }
  }

  test("declaration, namespaces, prefixed attributes, empty elements: JDK-identical bytes") {
    def doc(w: SaxWriter): Unit =
      w.startDocument()
      w.startElement("worksheet", XmlUtil.nsSpreadsheetML)
      w.writeAttribute("xmlns:r", XmlUtil.nsRelationships)
      w.writeAttribute("xmlns:x14ac", XmlUtil.nsX14ac)
      w.writeAttribute("xml:space", "preserve")
      w.startElement("sheetData")
      w.startElement("row")
      w.writeAttribute("r", "1")
      w.writeAttribute("x14ac:dyDescent", "0.25")
      w.startElement("c")
      w.writeAttribute("r", "A1")
      w.writeAttribute("t", "s")
      w.emptyElement("f", Seq("t" -> "dataTable", "ref" -> "B2:C3"))
      w.startElement("v")
      w.writeCharacters("0")
      w.endElement()
      w.endElement()
      w.endElement()
      w.endElement()
      w.startElement("drawing")
      w.writeAttribute("r:id", "rId1")
      w.endElement()
      w.startElement("x14:foo", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
      w.endElement()
      w.endElement()
      w.endDocument()
    val ours = stax(doc)
    assertEquals(ours, reference(doc))
    assert(ours.startsWith(staxDecl + "<worksheet xmlns=\"" + XmlUtil.nsSpreadsheetML + "\""), ours)
    assert(ours.contains("""<f t="dataTable" ref="B2:C3"/><v>0</v>"""), ours)
    assert(
      ours.contains(
        """<x14:foo xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"></x14:foo>"""
      ),
      ours
    )
  }

  test(
    "an empty element followed by its parent's end tag nests correctly (the JDK mis-nested it)"
  ) {
    val s = stax { w =>
      w.startElement("c"); w.emptyElement("f", Seq("t" -> "dataTable")); w.endElement()
    }
    assertEquals(s, """<c><f t="dataTable"/></c>""")
  }

  test("endDocument closes every open element; a surplus endElement is ignored (total)") {
    val s = stax { w =>
      w.startDocument(); w.startElement("p"); w.startElement("q"); w.emptyElement("z", Nil)
      w.endDocument()
    }
    assertEquals(s, s"$staxDecl<p><q><z/></q></p>")
    assertEquals(stax(w => w.endElement()), "")
  }

  test("an attribute or xmlns written with no open start tag is dropped, never thrown (total)") {
    // The JDK writer threw IllegalStateException here; xl's emitters never reach it, and the
    // document must stay well-formed either way — pinned so the choice is visible, not accidental.
    val s = stax { w =>
      w.startDocument()
      w.startElement("p")
      w.writeCharacters("x") // closes the start tag: attributes are no longer possible
      w.writeAttribute("late", "v")
      w.writeAttribute("xmlns:z", "urn:z")
      w.endElement()
      w.writeAttribute("after", "v")
      w.endDocument()
    }
    assertEquals(s, s"$staxDecl<p>x</p>")
    assertEquals(stax(w => w.writeAttribute("orphan", "v")), "")
  }

  test("UTF-8: 2-, 3- and 4-byte sequences (emoji) are raw; a lone surrogate becomes '?'") {
    val out = new ByteArrayOutputStream()
    val w = StaxSaxWriter.create(out)
    w.startElement("t"); w.writeCharacters("é日😀\ud83d"); w.endElement()
    val bytes = out.toByteArray.map(_ & 0xff).map(b => f"$b%02x").mkString(" ")
    assertEquals(
      bytes,
      "3c 74 3e c3 a9 e6 97 a5 f0 9f 98 80 3f 3c 2f 74 3e"
    )
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
