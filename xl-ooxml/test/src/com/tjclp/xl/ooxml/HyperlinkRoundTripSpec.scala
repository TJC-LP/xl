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
    assertEquals(sheet(ref"A1").hyperlink, Some("https://anthropic.com"))

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
    assertEquals(sheet(ref"A1").hyperlink, Some("Sheet1!B2"))
    Files.deleteIfExists(out)
  }
