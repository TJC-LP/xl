package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
import scala.xml.{Elem, Text}

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import munit.FunSuite

/**
 * GH-611: `"` is emitted verbatim in element text. scala.xml's `Text` serialization escaped it as
 * `&quot;` — needless in element content (only `&`, `<` and `>` must be escaped there), never what
 * Excel writes, and on a model with array constants in defined names, string literals in formulas
 * and quotes in shared strings a measurable size cost and a source of noisy byte diffs. Attribute
 * values keep their `&quot;`. The StAX backend already wrote `"` verbatim; both backends agree now,
 * and the reader accepts either spelling.
 */
class TextQuoteEscapingSpec extends FunSuite:

  private val decl = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"""
  private val fixture = "named-styles-excel.xlsx"
  private val backends = List(
    "ScalaXml" -> WriterConfig(backend = XmlBackend.ScalaXml),
    "SaxStax" -> WriterConfig(backend = XmlBackend.SaxStax)
  )

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  private def parse(xml: String): Elem =
    XmlSecurity.parseSafe(xml, "part.xml").fold(e => fail(s"parse failed: ${e.message}"), identity)

  test("compact keeps '\"' (and ') verbatim in text and escapes only &, < and >") {
    val e = XmlUtil.elem("t")(Text("a\"b<c&d>e 'f'"))
    assertEquals(XmlUtil.compact(e), s"$decl\n<t>a\"b&lt;c&amp;d&gt;e 'f'</t>")
  }

  test("compact still escapes '\"' inside attribute values") {
    val e = XmlUtil.elem("x", "v" -> "a\"b<c&d")()
    assertEquals(XmlUtil.compact(e), s"$decl\n<x v=\"a&quot;b&lt;c&amp;d\"/>")
  }

  test("compact drops XML-illegal control characters from text, as scala.xml did (GH-237)") {
    val e = XmlUtil.elem("t")(Text("a" + 1.toChar + "b\tc\n"))
    assertEquals(XmlUtil.compact(e), s"$decl\n<t>ab\tc\n</t>")
  }

  test("compact equals scala.xml's toString whenever no text node carries a quote") {
    // an Excel styles part: quotes only in attribute values, prefixed attributes (xr:uid),
    // nested namespace declarations (ext/x14), self-closing and mixed elements
    val in = TestFixtures.copyToTemp(fixture)
    val styles = parse(entryText(in, "xl/styles.xml"))
    assertEquals(XmlUtil.compact(styles), s"$decl\n${styles.toString}")
    // a shared string with a quote: the ONLY difference is that &quot; in text
    val sst = parse(entryText(in, "xl/sharedStrings.xml"))
    assert(sst.toString.contains("<t>He said &quot;hi&quot;</t>"), sst.toString)
    assertEquals(XmlUtil.compact(sst), s"$decl\n${sst.toString}".replace("&quot;", "\""))
    assert(XmlUtil.compact(sst).contains("<t>He said \"hi\"</t>"))
  }

  test("the StAX backend writes '\"' verbatim in text and escapes &, <, > like compact") {
    val out = new ByteArrayOutputStream()
    val w = StaxSaxWriter.create(out)
    w.startDocument()
    w.startElement("t")
    w.writeCharacters("a\"b<c&d>e")
    w.endElement()
    w.endDocument()
    w.flush()
    val s = new String(out.toByteArray, StandardCharsets.UTF_8)
    assert(s.contains("<t>a\"b&lt;c&amp;d&gt;e</t>"), s)
    assert(!s.contains("&quot;"), s)
  }

  backends.foreach { case (name, config) =>
    test(s"shared strings, formulas and defined names carry '\"' verbatim on a write ($name)") {
      val in = TestFixtures.copyToTemp(fixture)
      val wb = XlsxReader.read(in).fold(e => fail(s"read failed: ${e.message}"), identity)
      val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
      // a new quoted string forces the SST through the combine path; the put regenerates the sheet
      val edited = wb.put(sheet.put(ref"D1", CellValue.Text("new \"q\"")))
      val out = Files.createTempFile(s"xl-gh611-$name-", ".xlsx")
      out.toFile.deleteOnExit()
      XlsxWriter
        .writeWith(edited, out, config)
        .fold(e => fail(s"write failed: ${e.message}"), identity)

      val sst = entryText(out, "xl/sharedStrings.xml")
      assert(sst.contains("<t>He said \"hi\"</t>"), sst)
      assert(sst.contains("<t>new \"q\"</t>"), sst)
      assert(!sst.contains("&quot;"), sst)

      val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
      assert(sheetXml.contains("<f>\"say \"&amp;\"hi\"</f>"), sheetXml)
      assert(!sheetXml.contains("&quot;"), sheetXml)

      val workbookXml = entryText(out, "xl/workbook.xml")
      assert(
        workbookXml.contains(
          """<definedName name="Tags">{"detail",#N/A,FALSE,"mfg"}</definedName>"""
        ),
        workbookXml
      )

      // and everything reads back
      val back = XlsxReader.read(out).fold(e => fail(s"re-read failed: ${e.message}"), identity)
      val backSheet = back.sheets.headOption.getOrElse(fail("no sheet"))
      assertEquals(backSheet(ref"A1").value, CellValue.Text("He said \"hi\""))
      assertEquals(backSheet(ref"D1").value, CellValue.Text("new \"q\""))
      backSheet(ref"B1").value match
        case f: CellValue.Formula =>
          assertEquals(f.expression, "\"say \"&\"hi\"")
          assertEquals(f.cachedValue, Some(CellValue.Text("say hi")))
        case other => fail(s"B1 lost its formula: $other")
    }
  }

  test("the reader accepts both spellings — &quot; and a literal quote — in <t>, <is> and <f>") {
    val sst =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
        |<si><t>&quot;q&quot;</t></si><si><t>"q"</t></si>
        |</sst>""".stripMargin
    val sheet =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
        |<row r="1">
        |<c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>
        |<c r="C1" t="inlineStr"><is><t>&quot;q&quot;</t></is></c><c r="D1" t="inlineStr"><is><t>"q"</t></is></c>
        |<c r="E1" t="str"><f>&quot;a&quot;&amp;&quot;b&quot;</f><v>ab</v></c><c r="F1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c>
        |</row>
        |</sheetData></worksheet>""".stripMargin
    val wb = XlsxReader
      .readFromBytes(rawXlsx(sst, sheet))
      .fold(e => fail(s"read failed: ${e.message}"), identity)
    val s = wb.sheets.headOption.getOrElse(fail("no sheet"))
    List(ref"A1", ref"B1", ref"C1", ref"D1").foreach { r =>
      assertEquals(s(r).value, CellValue.Text("\"q\""), r.toA1)
    }
    val formula = CellValue.Formula("\"a\"&\"b\"", Some(CellValue.Text("ab")))
    assertEquals(s(ref"E1").value, formula)
    assertEquals(s(ref"F1").value, formula)
  }

  private def rawXlsx(sstXml: String, sheetXml: String): Array[Byte] =
    val contentTypes =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        |<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        |<Default Extension="xml" ContentType="application/xml"/>
        |<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        |<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        |<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
        |</Types>""".stripMargin
    val rootRels =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        |<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        |</Relationships>""".stripMargin
    val workbook =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        |<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
        |</workbook>""".stripMargin
    val workbookRels =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        |<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        |<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
        |</Relationships>""".stripMargin
    val output = new ByteArrayOutputStream()
    val zip = new ZipOutputStream(output)
    try
      List(
        "[Content_Types].xml" -> contentTypes,
        "_rels/.rels" -> rootRels,
        "xl/workbook.xml" -> workbook,
        "xl/_rels/workbook.xml.rels" -> workbookRels,
        "xl/sharedStrings.xml" -> sstXml,
        "xl/worksheets/sheet1.xml" -> sheetXml
      ).foreach { case (name, content) =>
        zip.putNextEntry(new ZipEntry(name))
        zip.write(content.getBytes(StandardCharsets.UTF_8))
        zip.closeEntry()
      }
    finally zip.close()
    output.toByteArray
