package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import scala.xml.XML

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.worksheet.OoxmlCell
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import munit.FunSuite

/**
 * GH-556: Excel stores post-2007 functions as `_xlfn.NAME(...)` inside `<f>` (FILTER/SORT as
 * `_xlfn._xlws.NAME`). A bare `MAXIFS(` in the file is an undefined name to desktop Excel and
 * LibreOffice — `#NAME?` on the first recalculation, invisible to `xl lint` and `--strict` because
 * the cached value is right. Every `<f>` writer must apply the storage prefix, and every reader
 * must hand the model the bare formula-bar spelling, so an Excel-authored book and an xl-authored
 * book hold identical model text and re-serialize byte-identically.
 */
class FutureFunctionPrefixSpec extends FunSuite:

  private val maxifs = CellValue.Formula("=MAXIFS(B1:B3,A1:A3,C1)", Some(CellValue.Number(1)))
  private val filter = CellValue.Formula("FILTER(A1:A3,B1:B3)", None)

  private def tempXlsx(label: String): Path =
    val p = Files.createTempFile(s"xl-xlfn-$label-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found")
    finally zip.close()

  private def workbookWith(values: (ARef, CellValue)*): Workbook =
    val base = Sheet("Data").put(ref"A1" -> "a").put(ref"B1" -> 1).put(ref"C1" -> "a")
    Workbook(Vector(values.foldLeft(base) { case (s, (at, v)) => s.put(at, v) }))

  private def writtenSheetXml(backend: XmlBackend, values: (ARef, CellValue)*): String =
    val path = tempXlsx(backend.toString.toLowerCase)
    XlsxWriter
      .writeWith(workbookWith(values*), path, WriterConfig(backend = backend))
      .fold(err => fail(s"write failed: $err"), identity)
    entryText(path, "xl/worksheets/sheet1.xml")

  test("GH-556: ScalaXml backend writes _xlfn. in front of a post-2007 function") {
    val xml = writtenSheetXml(XmlBackend.ScalaXml, ref"D1" -> maxifs)
    assert(xml.contains("<f>_xlfn.MAXIFS(B1:B3,A1:A3,C1)</f>"), xml)
  }

  test("GH-556: SaxStax backend (DirectSaxEmitter) writes _xlfn. too") {
    val xml = writtenSheetXml(XmlBackend.SaxStax, ref"D1" -> maxifs)
    assert(xml.contains("<f>_xlfn.MAXIFS(B1:B3,A1:A3,C1)</f>"), xml)
  }

  test("GH-556: FILTER/SORT take _xlfn._xlws.; Excel 2007 functions stay bare") {
    val xml = writtenSheetXml(
      XmlBackend.SaxStax,
      ref"D1" -> filter,
      ref"E1" -> CellValue.Formula("SUMIFS(B1:B3,A1:A3,C1)", None)
    )
    assert(xml.contains("<f>_xlfn._xlws.FILTER(A1:A3,B1:B3)</f>"), xml)
    assert(xml.contains("<f>SUMIFS(B1:B3,A1:A3,C1)</f>"), xml)
  }

  test("GH-556: already-prefixed model text (pasted from Excel) is not double-prefixed") {
    val xml = writtenSheetXml(
      XmlBackend.ScalaXml,
      ref"D1" -> CellValue.Formula("_xlfn.XLOOKUP(C1,A1:A3,B1:B3)", None)
    )
    assert(xml.contains("<f>_xlfn.XLOOKUP(C1,A1:A3,B1:B3)</f>"), xml)
    assert(!xml.contains("_xlfn._xlfn."), xml)
  }

  test("GH-556: OoxmlCell.toXml (DOM path) and writeSax apply the prefix") {
    val cell = OoxmlCell(ref"D1", maxifs, None, "")
    assertEquals((cell.toXml \ "f").text, "_xlfn.MAXIFS(B1:B3,A1:A3,C1)")
    val out = new ByteArrayOutputStream()
    val writer = StaxSaxWriter.create(out)
    cell.writeSax(writer)
    writer.flush()
    assertEquals((XML.loadString(out.toString("UTF-8")) \ "f").text, "_xlfn.MAXIFS(B1:B3,A1:A3,C1)")
  }

  test("GH-556: the reader hands the model the bare spelling; re-serialization is byte-identical") {
    val path = tempXlsx("roundtrip")
    XlsxWriter
      .write(workbookWith(ref"D1" -> maxifs, ref"E1" -> filter), path)
      .fold(err => fail(s"write failed: $err"), identity)
    val firstXml = entryText(path, "xl/worksheets/sheet1.xml")
    val wb = XlsxReader.read(path).fold(err => fail(s"read failed: $err"), identity)
    wb.sheets(0)(ref"D1").value match
      case CellValue.Formula(expr, cached, _) =>
        assertEquals(expr, "MAXIFS(B1:B3,A1:A3,C1)")
        assertEquals(cached, Some(CellValue.Number(1)))
      case other => fail(s"expected formula at D1, got $other")
    wb.sheets(0)(ref"E1").value match
      case CellValue.Formula(expr, _, _) => assertEquals(expr, "FILTER(A1:A3,B1:B3)")
      case other => fail(s"expected formula at E1, got $other")
    val again = tempXlsx("roundtrip-2")
    XlsxWriter.write(wb, again).fold(err => fail(s"write failed: $err"), identity)
    assertEquals(entryText(again, "xl/worksheets/sheet1.xml"), firstXml)
  }

  test("GH-556: a prefix on a function outside the known list survives read and write verbatim") {
    val path = tempXlsx("unknown")
    val unknown = CellValue.Formula("_xlfn.NOSUCHFN(A1)", Some(CellValue.Number(7)))
    XlsxWriter.write(workbookWith(ref"D1" -> unknown), path).fold(err => fail(s"$err"), identity)
    assert(entryText(path, "xl/worksheets/sheet1.xml").contains("<f>_xlfn.NOSUCHFN(A1)</f>"))
    val wb = XlsxReader.read(path).fold(err => fail(s"read failed: $err"), identity)
    wb.sheets(0)(ref"D1").value match
      case CellValue.Formula(expr, _, _) => assertEquals(expr, "_xlfn.NOSUCHFN(A1)")
      case other => fail(s"expected formula at D1, got $other")
  }

  test("GH-556: LET is stored with _xlpm. on its parameters and reads back bare") {
    val path = tempXlsx("let")
    val let = CellValue.Formula("=LET(x,B1,x+1)", Some(CellValue.Number(2)))
    XlsxWriter.write(workbookWith(ref"D1" -> let), path).fold(err => fail(s"$err"), identity)
    val firstXml = entryText(path, "xl/worksheets/sheet1.xml")
    assert(firstXml.contains("<f>_xlfn.LET(_xlpm.x,B1,_xlpm.x+1)</f>"), firstXml)
    val wb = XlsxReader.read(path).fold(err => fail(s"read failed: $err"), identity)
    wb.sheets(0)(ref"D1").value match
      case CellValue.Formula(expr, cached, _) =>
        assertEquals(expr, "LET(x,B1,x+1)")
        assertEquals(cached, Some(CellValue.Number(2)))
      case other => fail(s"expected formula at D1, got $other")
    val again = tempXlsx("let-2")
    XlsxWriter.write(wb, again).fold(err => fail(s"write failed: $err"), identity)
    assertEquals(entryText(again, "xl/worksheets/sheet1.xml"), firstXml)
  }
