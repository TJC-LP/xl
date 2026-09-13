package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import scala.xml.XML

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.{fx, ref}
import com.tjclp.xl.ooxml.worksheet.OoxmlCell
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import munit.FunSuite

/**
 * GH-456: `<f>` carries the formula EXPRESSION, never the display form's leading '='. Since GH-479
 * the model itself is canonical — `fx`, `FormulaParser` and `putFormulaInheriting` store the bare
 * text — but a raw `CellValue.Formula("=B4*2")` stays constructible, so EVERY writer still
 * canonicalizes at the `<f>` emission boundary — otherwise such a book ships `<f>=B4*2</f>`
 * (non-spec; openpyxl reads it back as "==B4*2"). Strip ONCE and leading only: interior '='
 * (comparisons) must never be touched.
 */
class FormulaLeadingEqualsSpec extends FunSuite:

  private val leadingEq = CellValue.Formula("=B4*2", None)

  private def tempXlsx(label: String): Path =
    val p = Files.createTempFile(s"xl-leading-eq-$label-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found")
    finally zip.close()

  private def writtenSheetXml(backend: XmlBackend, value: CellValue): String =
    val sheet = Sheet("Data").put(ref"B4" -> 2).put(ref"C4", value)
    val path = tempXlsx(backend.toString.toLowerCase)
    XlsxWriter
      .writeWith(Workbook(Vector(sheet)), path, WriterConfig(backend = backend))
      .fold(err => fail(s"write failed: $err"), identity)
    entryText(path, "xl/worksheets/sheet1.xml")

  private def readBack(value: CellValue, label: String): CellValue =
    val sheet = Sheet("Data").put(ref"B4" -> 2).put(ref"C4", value)
    val path = tempXlsx(label)
    XlsxWriter
      .write(Workbook(Vector(sheet)), path)
      .fold(err => fail(s"write failed: $err"), identity)
    val wb = XlsxReader.read(path).fold(err => fail(s"read failed: $err"), identity)
    wb.sheets(0)(ref"C4").value

  test("GH-456: ScalaXml backend writes fx-style formula without the leading '='") {
    val xml = writtenSheetXml(XmlBackend.ScalaXml, leadingEq)
    assert(xml.contains("<f>B4*2</f>"), xml)
    assert(!xml.contains("<f>="), xml)
  }

  test("GH-456: SaxStax backend (DirectSaxEmitter) writes fx-style formula without the '='") {
    val xml = writtenSheetXml(XmlBackend.SaxStax, leadingEq)
    assert(xml.contains("<f>B4*2</f>"), xml)
    assert(!xml.contains("<f>="), xml)
  }

  test("GH-456: cached value is untouched by the canonicalization") {
    val xml =
      writtenSheetXml(XmlBackend.ScalaXml, CellValue.Formula("=B4*2", Some(CellValue.Number(4))))
    assert(xml.contains("<f>B4*2</f>"), xml)
    assert(xml.contains("<v>4</v>"), xml)
  }

  test("GH-456: interior '=' is never stripped (leading only, once)") {
    val xml = writtenSheetXml(XmlBackend.ScalaXml, CellValue.Formula("=IF(A1=1,2,3)", None))
    assert(xml.contains("<f>IF(A1=1,2,3)</f>"), xml)
  }

  test("GH-456: already-clean expression is written verbatim (CLI putf parity)") {
    val xml = writtenSheetXml(XmlBackend.SaxStax, CellValue.Formula("SUM(A1:A2)", None))
    assert(xml.contains("<f>SUM(A1:A2)</f>"), xml)
  }

  test("GH-456: OoxmlCell.toXml (DOM path) strips the leading '='") {
    val cell = OoxmlCell(ref"C4", leadingEq, None, "")
    assertEquals((cell.toXml \ "f").text, "B4*2")
  }

  test("GH-456: OoxmlCell.writeSax strips the leading '='") {
    val out = new ByteArrayOutputStream()
    val writer = StaxSaxWriter.create(out)
    OoxmlCell(ref"C4", leadingEq, None, "").writeSax(writer)
    writer.flush()
    val xml = XML.loadString(out.toString("UTF-8"))
    assertEquals((xml \ "f").text, "B4*2")
  }

  test("GH-456: a raw '='-prefixed construction reads back canonical (the writer heals it)") {
    readBack(leadingEq, "roundtrip") match
      case CellValue.Formula(expr, _, _) => assertEquals(expr, "B4*2")
      case other => fail(s"expected formula at C4, got $other")
  }

  test("GH-479: an fx-authored cell reads back EQUAL to the authored value (strict, no strip)") {
    val authored = fx"=B4*2"
    assertEquals(authored, CellValue.Formula("B4*2", None))
    assertEquals(readBack(authored, "fx-strict"), authored)
  }

  test(
    "GH-479: a doubled '=' is the documented exception: fx strips once, the writer heals the rest"
  ) {
    // `==B4*2` is not a formula; the canonical entry removes ONE '=' and the `<f>` boundary the
    // other, so this is the one fx-authored shape that does not read back equal — by design.
    val authored = fx"==B4*2"
    assertEquals(authored, CellValue.Formula("=B4*2", None))
    assertEquals(readBack(authored, "fx-doubled"), CellValue.Formula("B4*2", None))
  }

  test("GH-479: a CellValue.formula-authored cell (display form in) reads back EQUAL (strict)") {
    val authored =
      CellValue.formula("=B4*2").fold(err => fail(s"formula refused: ${err.message}"), identity)
    assertEquals(authored, CellValue.Formula("B4*2", None))
    assertEquals(readBack(authored, "formula-strict"), authored)
  }

  test("GH-479: WorkbookEquivalence is strict again — a phantom leading '=' is a diff") {
    val bare = Workbook(Vector(Sheet("Data").put(ref"C4", CellValue.Formula("B4*2", None))))
    val eq = Workbook(Vector(Sheet("Data").put(ref"C4", CellValue.Formula("=B4*2", None))))
    assert(WorkbookEquivalence.diff(bare, bare).isEmpty)
    assert(WorkbookEquivalence.diff(bare, eq).nonEmpty)
    assert(WorkbookEquivalence.diff(eq, bare).nonEmpty)
  }
