package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.styles.font.Font
import munit.FunSuite

/**
 * GH-303: SharedStrings keyed RichText entries by `toPlainText` in the SAME map as plain entries —
 * a RichText whose plain text equals an existing plain string collided with it (one entry dropped,
 * the surviving index shared by both cells: formatting lost or wrongly applied).
 *
 * The fix keys plain entries by exact string (GH-289) and RichText entries by their full run
 * structure (`richIndexMap`), so "Hello" and bold-"Hello" are two SST entries that both round-trip.
 */
class SstRichTextKeyingSpec extends FunSuite:

  private val boldHello = RichText(
    Vector(TextRun("Hello", font = Some(Font.default.withBold(true))))
  )
  private val italicHello = RichText(
    Vector(TextRun("Hello", font = Some(Font.default.withItalic(true))))
  )

  // ========== keying unit tests ==========

  test("GH-303: plain text and RichText with identical plain text are distinct entries") {
    val sst = SharedStrings.fromEntries(Vector(Left("Hello"), Right(boldHello)))
    assertEquals(sst.strings.size, 2, s"collision dropped an entry: ${sst.strings}")
    assertEquals(sst.indexOf("Hello"), Some(0), "plain lookup must hit the plain entry")
    assertEquals(sst.indexOf(boldHello), Some(1), "rich lookup must hit the rich entry")
  }

  test("GH-303: identical RichText dedups; same text with different formatting stays distinct") {
    val sst = SharedStrings.fromEntries(
      Vector(Right(boldHello), Right(boldHello), Right(italicHello))
    )
    assertEquals(sst.strings.size, 2, s"expected bold+italic entries, got: ${sst.strings}")
    assertEquals(sst.indexOf(boldHello), Some(0))
    assertEquals(sst.indexOf(italicHello), Some(1))
    assertEquals(sst.totalCount, 3, "totalCount counts references, not unique entries")
  }

  test("GH-303: fromXml keeps plain and rich same-text entries separately addressable") {
    val xml =
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
        <si><t>Hello</t></si>
        <si><r><rPr><b/></rPr><t>Hello</t></r></si>
      </sst>
    val sst = SharedStrings.fromXml(xml).fold(e => fail(s"parse failed: $e"), identity)
    assertEquals(sst.strings.size, 2)
    // Before the fix the second entry OVERWROTE the "Hello" key, so plain cells silently
    // resolved to the rich entry's index
    assertEquals(sst.indexOf("Hello"), Some(0), "plain lookup poisoned by rich entry")
    val parsedRich = sst.strings.lift(1).collect { case Right(rt) => rt }
    val richIdx = parsedRich.flatMap(sst.indexOf)
    assertEquals(richIdx, Some(1), "rich entry not addressable by its own structure")
  }

  // ========== writer round-trips (both backends) ==========

  private def roundTrip(backend: XmlBackend): Unit =
    val wb = Workbook(
      Sheet("Sheet1")
        .put(ref"A1" -> "Hello")
        .put(ref"B1" -> boldHello)
    )
    val out = Files.createTempFile(s"gh303-$backend", ".xlsx")
    val config = WriterConfig(sstPolicy = SstPolicy.Always, backend = backend)
    XlsxWriter.writeWith(wb, out, config).fold(e => fail(s"write failed: $e"), identity)

    val sstXml = readEntry(out, "xl/sharedStrings.xml")
    assert(
      sstXml.contains("uniqueCount=\"2\""),
      s"[$backend] plain+rich 'Hello' must be TWO SST entries: $sstXml"
    )

    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    sheet(ref"A1").value match
      case CellValue.Text(s) => assertEquals(s, "Hello")
      case other => fail(s"[$backend] plain cell came back as $other (formatting wrongly shared)")
    sheet(ref"B1").value match
      case CellValue.RichText(rt) =>
        assertEquals(rt.toPlainText, "Hello")
        assert(
          rt.runs.exists(_.font.exists(_.bold)),
          s"[$backend] bold formatting lost: ${rt.runs}"
        )
      case other => fail(s"[$backend] rich cell came back as $other (formatting lost)")
    Files.deleteIfExists(out)

  test("GH-303: plain 'Hello' + bold 'Hello' round-trip through the ScalaXml backend") {
    roundTrip(XmlBackend.ScalaXml)
  }

  test("GH-303: plain 'Hello' + bold 'Hello' round-trip through the SaxStax backend") {
    roundTrip(XmlBackend.SaxStax)
  }

  // ========== surgical combine path (GH-277 area) ==========

  test("GH-303: surgical combine adds a rich entry distinct from the preserved equal plain text") {
    val source = createPlainHelloFixture()
    val wb = XlsxReader.read(source).fold(e => fail(s"read failed: $e"), identity)

    val modified = wb("Sheet1")
      .map(sheet => wb.put(sheet.put(ref"B1" -> boldHello)))
      .fold(e => fail(s"modify failed: $e"), identity)

    val output = Files.createTempFile("gh303-combine", ".xlsx")
    XlsxWriter.write(modified, output).fold(e => fail(s"write failed: $e"), identity)

    val sstXml = readEntry(output, "xl/sharedStrings.xml")
    assert(
      sstXml.contains("uniqueCount=\"2\""),
      s"combine must append the rich entry instead of colliding with plain 'Hello': $sstXml"
    )

    val reread = XlsxReader.read(output).fold(e => fail(s"reread failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet(ref"A1").value, CellValue.Text("Hello"), "preserved plain cell drifted")
    sheet(ref"B1").value match
      case CellValue.RichText(rt) =>
        assertEquals(rt.toPlainText, "Hello")
        assert(rt.runs.exists(_.font.exists(_.bold)), s"bold lost via combine: ${rt.runs}")
      case other => fail(s"rich cell resolved to $other through the combine path")
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== helpers ==========

  private def readEntry(path: Path, entryName: String): String =
    val zip = new ZipFile(path.toFile)
    try
      val entry = zip.getEntry(entryName)
      assert(entry != null, s"$entryName missing from $path")
      val is = zip.getInputStream(entry)
      try new String(is.readAllBytes(), "UTF-8")
      finally is.close()
    finally zip.close()

  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    out.putNextEntry(new ZipEntry(name))
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()

  /** Excel-shaped fixture: SST with one plain "Hello" entry referenced from Sheet1!A1. */
  private def createPlainHelloFixture(): Path =
    val path = Files.createTempFile("gh303-fixture", ".xlsx")
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
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
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
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
      )
      writeEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""
      )
      writeEntry(
        out,
        "xl/sharedStrings.xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
          "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">" +
          "<si><t>Hello</t></si>" +
          "</sst>"
      )
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
  </sheetData>
</worksheet>"""
      )
    finally out.close()
    path
