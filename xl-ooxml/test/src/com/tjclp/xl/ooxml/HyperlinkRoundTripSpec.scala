package com.tjclp.xl.ooxml

import java.nio.file.Files
import java.util.zip.ZipFile

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.unsafe.*
import munit.FunSuite

/**
 * GH-235: Cell.hyperlink was modeled with a public API but never serialized (a silent no-op). These
 * tests prove hyperlinks now write (<hyperlinks> + external-URL relationships) and round-trip back
 * into the model.
 */
class HyperlinkRoundTripSpec extends FunSuite:

  private def entryText(zip: ZipFile, name: String): String =
    val is = zip.getInputStream(zip.getEntry(name))
    try new String(is.readAllBytes(), "UTF-8")
    finally is.close()

  test("GH-235: external hyperlink serializes (<hyperlinks> + External rel) and round-trips") {
    val cell = Cell(ref"A1", CellValue.Text("Anthropic")).withHyperlink("https://anthropic.com")
    val wb = Workbook(Sheet("Sheet1").put(cell))
    val out = Files.createTempFile("hl-ext", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    // Round-trip: Cell.hyperlink populated on read
    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet(ref"A1").hyperlink, Some(Hyperlink("https://anthropic.com")))

    // Raw output: <hyperlinks> in the sheet + an External hyperlink relationship in its .rels
    val zip = new ZipFile(out.toFile)
    val sheetXml = entryText(zip, "xl/worksheets/sheet1.xml")
    val relsXml = entryText(zip, "xl/worksheets/_rels/sheet1.xml.rels")
    zip.close()
    assert(sheetXml.contains("<hyperlink "), s"no <hyperlink> in sheet:\n$sheetXml")
    assert(
      relsXml.contains("anthropic.com") && relsXml.contains("External"),
      s"no external hyperlink rel:\n$relsXml"
    )
    Files.deleteIfExists(out)
  }

  test("GH-235: internal location hyperlink uses the location attribute and round-trips") {
    val cell = Cell(ref"A1", CellValue.Text("Go")).withHyperlink("Sheet1!B2")
    val wb = Workbook(Sheet("Sheet1").put(cell))
    val out = Files.createTempFile("hl-int", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet(ref"A1").hyperlink, Some(Hyperlink("Sheet1!B2")))
    Files.deleteIfExists(out)
  }

  test("GH-406: '#'-prefixed internal target normalizes in withHyperlink and round-trips clean") {
    // The '#Sheet1!A1' dialect is common agent/browser spelling; OOXML location wants no '#'.
    val cell = Cell(ref"A1", CellValue.Text("Go")).withHyperlink("#Sheet1!B2")
    assertEquals(cell.hyperlink, Some(Hyperlink("Sheet1!B2")))

    val wb = Workbook(Sheet("Sheet1").put(cell))
    val out = Files.createTempFile("hl-hash", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val zip = new ZipFile(out.toFile)
    val sheetXml = entryText(zip, "xl/worksheets/sheet1.xml")
    zip.close()
    assert(
      sheetXml.contains("""location="Sheet1!B2""""),
      s"location must not carry the leading '#':\n$sheetXml"
    )

    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet(ref"A1").hyperlink, Some(Hyperlink("Sheet1!B2")))
    Files.deleteIfExists(out)
  }

  test("GH-406: writer strips a leading '#' even when the field bypassed withHyperlink") {
    // Direct construction (copy/generators) skips the chokepoint; the location emit defends.
    val cell = Cell(ref"A1", CellValue.Text("Go"), None, Some(Hyperlink("#Sheet1!C3")))
    val wb = Workbook(Sheet("Sheet1").put(cell))
    val out = Files.createTempFile("hl-hash-direct", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val zip = new ZipFile(out.toFile)
    val sheetXml = entryText(zip, "xl/worksheets/sheet1.xml")
    zip.close()
    assert(
      sheetXml.contains("""location="Sheet1!C3""""),
      s"location must not carry the leading '#':\n$sheetXml"
    )
    Files.deleteIfExists(out)
  }

  test("GH-406: external URL fragments are untouched (only a LEADING '#' strips)") {
    val target = "https://example.com/page#section"
    val cell = Cell(ref"A1", CellValue.Text("Docs")).withHyperlink(target)
    assertEquals(cell.hyperlink, Some(Hyperlink(target)))

    val wb = Workbook(Sheet("Sheet1").put(cell))
    val out = Files.createTempFile("hl-frag", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet(ref"A1").hyperlink, Some(Hyperlink(target)))

    val zip = new ZipFile(out.toFile)
    val relsXml = entryText(zip, "xl/worksheets/_rels/sheet1.xml.rels")
    zip.close()
    assert(relsXml.contains("page#section"), s"external rel must keep its fragment:\n$relsXml")
    Files.deleteIfExists(out)
  }

  // ===== GH-895: removing a sheet's last hyperlink sticks through a read-modify-write =====

  /** Write `cells` on sheet "Summary", read it back: a workbook carrying the source's metadata. */
  private def readFromFile(cells: Cell*): (Workbook, java.nio.file.Path) =
    val src = Files.createTempFile("hl-src", ".xlsx")
    val wb = Workbook(cells.foldLeft(Sheet("Summary"))(_.put(_)))
    XlsxWriter.write(wb, src).fold(e => fail(s"write failed: $e"), identity)
    (XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity), src)

  private def clearLink(wb: Workbook, at: com.tjclp.xl.addressing.ARef): Workbook =
    val sheet = wb("Summary").fold(e => fail(s"sheet missing: $e"), identity)
    wb.put(sheet.put(sheet(at).clearHyperlink))

  private def writeAndReopen(wb: Workbook): (Sheet, String, Option[String]) =
    val out = Files.createTempFile("hl-out", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    val sheet = XlsxReader
      .read(out)
      .flatMap(_("Summary"))
      .fold(e => fail(s"reread failed: $e"), identity)
    val zip = new ZipFile(out.toFile)
    val sheetXml = entryText(zip, "xl/worksheets/sheet1.xml")
    val relsName = "xl/worksheets/_rels/sheet1.xml.rels"
    val relsXml = Option(zip.getEntry(relsName)).map(_ => entryText(zip, relsName))
    zip.close()
    Files.deleteIfExists(out)
    (sheet, sheetXml, relsXml)

  test("GH-895: removing a sheet's only external hyperlink drops <hyperlinks> and its rel") {
    val (wb, src) = readFromFile(
      Cell(ref"B7", CellValue.Text("Example link")).withHyperlink("https://example.com/")
    )
    val (sheet, sheetXml, relsXml) = writeAndReopen(clearLink(wb, ref"B7"))
    assertEquals(sheet(ref"B7").hyperlink, None, "B7's hyperlink must stay removed")
    assertEquals(sheet(ref"B7").value, CellValue.Text("Example link"), "B7 keeps its text")
    assert(!sheetXml.contains("<hyperlinks"), s"no <hyperlinks> element expected:\n$sheetXml")
    relsXml.foreach { rels =>
      assert(!rels.contains("example.com"), s"the removed link's rel must not ship:\n$rels")
    }
    Files.deleteIfExists(src)
  }

  test("GH-895: removing a sheet's only internal (location) hyperlink drops <hyperlinks>") {
    val (wb, src) = readFromFile(
      Cell(ref"B7", CellValue.Text("Go")).withHyperlink("Summary!A1")
    )
    val (sheet, sheetXml, _) = writeAndReopen(clearLink(wb, ref"B7"))
    assertEquals(sheet(ref"B7").hyperlink, None, "B7's hyperlink must stay removed")
    assert(!sheetXml.contains("<hyperlinks"), s"no <hyperlinks> element expected:\n$sheetXml")
    Files.deleteIfExists(src)
  }

  test("GH-895: removing one of two hyperlinks keeps the other (and only the other)") {
    val (wb, src) = readFromFile(
      Cell(ref"B7", CellValue.Text("Example link")).withHyperlink("https://example.com/"),
      Cell(ref"B8", CellValue.Text("Other link")).withHyperlink("https://example.org/")
    )
    val (sheet, sheetXml, relsXml) = writeAndReopen(clearLink(wb, ref"B7"))
    assertEquals(sheet(ref"B7").hyperlink, None)
    assertEquals(sheet(ref"B8").hyperlink, Some(Hyperlink("https://example.org/")))
    assert(!sheetXml.contains("""ref="B7""""), s"B7 must not be linked:\n$sheetXml")
    relsXml.foreach(rels => assert(!rels.contains("example.com"), rels))
    Files.deleteIfExists(src)
  }

  test("GH-895: an untouched read-write keeps the sheet's only hyperlink") {
    val (wb, src) = readFromFile(
      Cell(ref"B7", CellValue.Text("Example link")).withHyperlink("https://example.com/")
    )
    val (sheet, sheetXml, relsXml) = writeAndReopen(wb)
    assertEquals(sheet(ref"B7").hyperlink, Some(Hyperlink("https://example.com/")))
    assert(sheetXml.contains("<hyperlinks"), sheetXml)
    assert(relsXml.exists(_.contains("example.com")), s"the link's rel must ship: $relsXml")
    Files.deleteIfExists(src)
  }

  /** A source whose sheet has a cell link on B7 and a range link (unmodeled) on D2:E3. */
  private def sourceWithRangeLink(): java.nio.file.Path =
    val written = Files.createTempFile("hl-range-gen", ".xlsx")
    val wb = Workbook(
      Sheet("Summary")
        .put(Cell(ref"B7", CellValue.Text("Example link")).withHyperlink("https://example.com/"))
        .put(Cell(ref"D2", CellValue.Text("Range link")).withHyperlink("https://example.net/"))
    )
    XlsxWriter.write(wb, written).fold(e => fail(s"write failed: $e"), identity)
    // Excel writes ref="D2:E3" for a link over a multi-cell selection; xl cannot author one, so
    // the source is patched: the reader keeps it as preserved XML, not on a cell.
    val patched = Files.createTempFile("hl-range", ".xlsx")
    val in = new ZipFile(written.toFile)
    val out = new java.util.zip.ZipOutputStream(Files.newOutputStream(patched))
    try
      in.entries().asIterator().forEachRemaining { e =>
        val bytes = in.getInputStream(e).readAllBytes()
        val body =
          if e.getName == "xl/worksheets/sheet1.xml" then
            new String(bytes, "UTF-8").replace("""ref="D2"""", """ref="D2:E3"""").getBytes("UTF-8")
          else bytes
        out.putNextEntry(new java.util.zip.ZipEntry(e.getName))
        out.write(body)
        out.closeEntry()
      }
    finally
      out.close()
      in.close()
      Files.deleteIfExists(written)
    patched

  test("GH-895: removing the last cell link keeps a range link and its rel") {
    val src = sourceWithRangeLink()
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val (sheet, sheetXml, relsXml) = writeAndReopen(clearLink(wb, ref"B7"))
    assertEquals(sheet(ref"B7").hyperlink, None, "B7's hyperlink must stay removed")
    assert(!sheetXml.contains("""ref="B7""""), s"B7 must not be linked:\n$sheetXml")
    assert(sheetXml.contains("""ref="D2:E3""""), s"the range link must survive:\n$sheetXml")
    val rels = relsXml.getOrElse(fail("the range link's rel must ship"))
    assert(rels.contains("example.net"), s"the range link's rel must ship:\n$rels")
    assert(!rels.contains("example.com"), s"the removed link's rel must not ship:\n$rels")
    Files.deleteIfExists(src)
  }
