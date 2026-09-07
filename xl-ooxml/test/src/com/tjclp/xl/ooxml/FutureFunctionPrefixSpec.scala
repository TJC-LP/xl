package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.jdk.CollectionConverters.*
import scala.xml.XML

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cf.{CfOperator, CfPoint, CfRule, Cfvo, ConditionalFormat}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.metadata.WorkbookMetadataReader
import com.tjclp.xl.ooxml.worksheet.OoxmlCell
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import com.tjclp.xl.sheets.{DataValidation, DvBoundedType, DvKind, DvOperator}
import com.tjclp.xl.styles.Dxf
import com.tjclp.xl.styles.color.Color
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

  // ===== GH-577: the three remaining formula-text boundaries (CF, DV, defined names) =====

  private val red = Color.Rgb(0xffff0000)
  private val green = Color.Rgb(0xff00ff00)

  /**
   * IFS / XLOOKUP / MINIFS / MAXIFS / FILTER spelled bare (the formula-bar form) in every slot
   * outside `<f>`: a CF expression rule, a CF cellIs rule, a colorScale cfvo formula, a custom DV,
   * a list DV, a bounded DV pair, and a workbook-scoped defined name.
   */
  private def gh577Workbook: Workbook =
    val sheet = Sheet("Data")
      .put(ref"A1" -> "a")
      .put(ref"B1" -> 1)
      .put(ref"C1" -> "a")
      .conditionalFormat(
        ref"A1:A3",
        CfRule.Expression("IFS(A1=1,1,TRUE,0)", Some(Dxf.fill(red)), 1),
        CfRule.CellIs(CfOperator.GreaterThan, "XLOOKUP(C1,A1:A3,B1:B3)", None, None, 2)
      )
      .conditionalFormat(
        ref"B1:B3",
        CfRule.ColorScale(
          CfPoint(Cfvo.Formula("MINIFS(B1:B3,A1:A3,C1)"), red),
          None,
          CfPoint(Cfvo.Max, green),
          3
        )
      )
      .withDataValidation(
        ref"D1:D3",
        dv(DvKind.Custom("IFS(D1=1,1,TRUE,0)"))
      )
      .withDataValidation(
        ref"E1:E3",
        dv(DvKind.List("FILTER(A1:A3,B1:B3)"))
      )
      .withDataValidation(
        ref"F1:F3",
        dv(
          DvKind.Bounded(
            DvBoundedType.Whole,
            DvOperator.Between,
            "MINIFS(B1:B3,A1:A3,C1)",
            Some("MAXIFS(B1:B3,A1:A3,C1)")
          )
        )
      )
    Workbook(Vector(sheet))
      .withDefinedName("Best", "MAXIFS(Data!$B$1:$B$3,Data!$A$1:$A$3,Data!$C$1)")

  /** The enum case widens to `DataValidation`; `withDataValidation` wants the `Rules` case. */
  private def dv(kind: DvKind): DataValidation.Rules = DataValidation.Rules(Vector.empty, kind)

  private val gh577SheetFragments = Vector(
    "<formula>_xlfn.IFS(A1=1,1,TRUE,0)</formula>",
    "<formula>_xlfn.XLOOKUP(C1,A1:A3,B1:B3)</formula>",
    "val=\"_xlfn.MINIFS(B1:B3,A1:A3,C1)\"",
    "<formula1>_xlfn.IFS(D1=1,1,TRUE,0)</formula1>",
    "<formula1>_xlfn._xlws.FILTER(A1:A3,B1:B3)</formula1>",
    "<formula1>_xlfn.MINIFS(B1:B3,A1:A3,C1)</formula1>" +
      "<formula2>_xlfn.MAXIFS(B1:B3,A1:A3,C1)</formula2>"
  )

  private val gh577NameFragment =
    "<definedName name=\"Best\">_xlfn.MAXIFS(Data!$B$1:$B$3,Data!$A$1:$A$3,Data!$C$1)</definedName>"

  /** Formula texts of every typed CF rule on the sheet, in document order (cfvo formulas too). */
  private def cfFormulas(sheet: Sheet): Vector[String] =
    sheet.conditionalFormats.flatMap {
      case ConditionalFormat.Rules(_, rules, _) =>
        rules.flatMap {
          case CfRule.CellIs(_, f1, f2, _, _, _) => f1 +: f2.toList.toVector
          case CfRule.Expression(f, _, _, _) => Vector(f)
          case CfRule.ColorScale(min, mid, max, _) =>
            (Vector(min) ++ mid.toList :+ max).map(_.cfvo).collect { case Cfvo.Formula(f) => f }
          case _ => Vector.empty
        }
      case _: ConditionalFormat.Preserved => Vector.empty
    }

  /** Formula texts of every typed data validation on the sheet, in document order. */
  private def dvFormulas(sheet: Sheet): Vector[String] =
    sheet.dataValidations.flatMap {
      case DataValidation.Rules(_, kind, _, _, _) =>
        kind match
          case DvKind.Custom(f) => Vector(f)
          case DvKind.List(f) => Vector(f)
          case DvKind.Bounded(_, _, f1, f2) => f1 +: f2.toList.toVector
          case DvKind.AnyValue => Vector.empty
      case _: DataValidation.Preserved => Vector.empty
    }

  private val bareCfFormulas =
    Vector("IFS(A1=1,1,TRUE,0)", "XLOOKUP(C1,A1:A3,B1:B3)", "MINIFS(B1:B3,A1:A3,C1)")
  private val bareDvFormulas = Vector(
    "IFS(D1=1,1,TRUE,0)",
    "FILTER(A1:A3,B1:B3)",
    "MINIFS(B1:B3,A1:A3,C1)",
    "MAXIFS(B1:B3,A1:A3,C1)"
  )
  private val bareName = "MAXIFS(Data!$B$1:$B$3,Data!$A$1:$A$3,Data!$C$1)"

  private def writeGh577(backend: XmlBackend, label: String): Path =
    val path = tempXlsx(label)
    XlsxWriter
      .writeWith(gh577Workbook, path, WriterConfig(backend = backend))
      .fold(err => fail(s"write failed: $err"), identity)
    path

  private def assertFragments(xml: String, fragments: Vector[String]): Unit =
    fragments.foreach(fragment => assert(xml.contains(fragment), s"missing $fragment in\n$xml"))

  test("GH-577: CF <formula>, cfvo val and DV <formula1>/<formula2> take _xlfn. (ScalaXml)") {
    val xml =
      entryText(writeGh577(XmlBackend.ScalaXml, "cfdv-scalaxml"), "xl/worksheets/sheet1.xml")
    assertFragments(xml, gh577SheetFragments)
    assert(!xml.contains("<formula>IFS("), xml)
    assert(!xml.contains("<formula1>IFS("), xml)
  }

  test("GH-577: CF <formula>, cfvo val and DV <formula1>/<formula2> take _xlfn. (SaxStax)") {
    val xml = entryText(writeGh577(XmlBackend.SaxStax, "cfdv-saxstax"), "xl/worksheets/sheet1.xml")
    assertFragments(xml, gh577SheetFragments)
  }

  test("GH-577: an authored <definedName> takes _xlfn. in workbook.xml") {
    val xml = entryText(writeGh577(XmlBackend.ScalaXml, "name"), "xl/workbook.xml")
    assert(xml.contains(gh577NameFragment), xml)
  }

  test("GH-577: CF/DV/name text reads back bare and re-serializes byte-identically") {
    val path = writeGh577(XmlBackend.ScalaXml, "cfdv-roundtrip")
    val firstSheet = entryText(path, "xl/worksheets/sheet1.xml")
    val firstWorkbook = entryText(path, "xl/workbook.xml")
    val wb = XlsxReader.read(path).fold(err => fail(s"read failed: $err"), identity)
    assertEquals(cfFormulas(wb.sheets(0)), bareCfFormulas)
    assertEquals(dvFormulas(wb.sheets(0)), bareDvFormulas)
    assertEquals(wb.metadata.definedNames.map(_.formula), Vector(bareName))
    // the lightweight metadata reader hands out the same bare spelling
    assertEquals(
      WorkbookMetadataReader.readDefinedNames(path).map(_.map(_.formula)),
      Right(Vector(bareName))
    )
    val again = tempXlsx("cfdv-roundtrip-2")
    XlsxWriter.write(wb, again).fold(err => fail(s"write failed: $err"), identity)
    assertEquals(entryText(again, "xl/worksheets/sheet1.xml"), firstSheet)
    assertEquals(entryText(again, "xl/workbook.xml"), firstWorkbook)
  }

  test("GH-577: regenerating a dirty CF block / DV container / name table keeps the prefixes") {
    val path = writeGh577(XmlBackend.ScalaXml, "cfdv-dirty")
    val wb = XlsxReader.read(path).fold(err => fail(s"read failed: $err"), identity)
    // every slot goes dirty: an extra rule, an extra validation, an extra name, a cell edit
    val sheet = wb
      .sheets(0)
      .put(ref"G1" -> 2)
      .conditionalFormat(ref"C1:C3", CfRule.Expression("LEN(C1)>0", None, 9))
      .withDataValidation(
        ref"H1:H3",
        dv(DvKind.List("\"x,y\""))
      )
    val dirty = wb.put(sheet).withDefinedName("Other", "Data!$A$1")
    val out = tempXlsx("cfdv-dirty-2")
    XlsxWriter.write(dirty, out).fold(err => fail(s"write failed: $err"), identity)
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assertFragments(sheetXml, gh577SheetFragments)
    assert(sheetXml.contains("<formula>LEN(C1)&gt;0</formula>"), sheetXml)
    val wbXml = entryText(out, "xl/workbook.xml")
    assert(wbXml.contains(gh577NameFragment), wbXml)
    assert(wbXml.contains("<definedName name=\"Other\">Data!$A$1</definedName>"), wbXml)
  }

  /** Copy `source` to a new zip, rewriting the text of every `.xml` entry through `patch`. */
  private def patchedZip(source: Path, label: String)(patch: String => String): Path =
    val out = tempXlsx(label)
    val zip = new ZipFile(source.toFile)
    val zos = new ZipOutputStream(Files.newOutputStream(out))
    try
      zip.entries().asScala.foreach { entry =>
        val bytes = zip.getInputStream(entry).readAllBytes()
        val content =
          if entry.getName.endsWith(".xml") then
            patch(new String(bytes, StandardCharsets.UTF_8)).getBytes(StandardCharsets.UTF_8)
          else bytes
        zos.putNextEntry(new ZipEntry(entry.getName))
        zos.write(content)
        zos.closeEntry()
      }
    finally
      zos.close()
      zip.close()
    out

  test("GH-577: third-party bare text reads bare, and a write that regenerates the slot heals it") {
    val written = writeGh577(XmlBackend.ScalaXml, "cfdv-bare-src")
    // strip every storage prefix by hand — the file a bare-writing producer (openpyxl) emits
    val bare = patchedZip(written, "cfdv-bare")(_.replace("_xlfn._xlws.", "").replace("_xlfn.", ""))
    assert(!entryText(bare, "xl/worksheets/sheet1.xml").contains("_xlfn."))
    assert(!entryText(bare, "xl/workbook.xml").contains("_xlfn."))
    val wb = XlsxReader.read(bare).fold(err => fail(s"read failed: $err"), identity)
    assertEquals(cfFormulas(wb.sheets(0)), bareCfFormulas)
    assertEquals(dvFormulas(wb.sheets(0)), bareDvFormulas)
    assertEquals(wb.metadata.definedNames.map(_.formula), Vector(bareName))
    // a write that regenerates the slots (dirty CF, DV and names) emits Excel's storage form
    val sheet = wb
      .sheets(0)
      .conditionalFormat(ref"C1:C3", CfRule.Expression("LEN(C1)>0", None, 9))
      .withDataValidation(
        ref"H1:H3",
        dv(DvKind.List("\"x,y\""))
      )
    val healed = tempXlsx("cfdv-healed")
    XlsxWriter
      .write(wb.put(sheet).withDefinedName("Other", "Data!$A$1"), healed)
      .fold(err => fail(s"write failed: $err"), identity)
    assertFragments(entryText(healed, "xl/worksheets/sheet1.xml"), gh577SheetFragments)
    assert(entryText(healed, "xl/workbook.xml").contains(gh577NameFragment))
  }

  test("GH-577: a prefix on an unknown function in a CF rule or a name survives verbatim") {
    val sheet = Sheet("Data")
      .put(ref"A1" -> 1)
      .conditionalFormat(ref"A1:A3", CfRule.Expression("_xlfn.NOSUCHFN(A1)", None, 1))
    val wb = Workbook(Vector(sheet)).withDefinedName("Odd", "_xlfn.NOSUCHFN(Data!$A$1)")
    val path = tempXlsx("cfdv-unknown")
    XlsxWriter.write(wb, path).fold(err => fail(s"write failed: $err"), identity)
    assert(
      entryText(path, "xl/worksheets/sheet1.xml").contains("<formula>_xlfn.NOSUCHFN(A1)</formula>")
    )
    assert(entryText(path, "xl/workbook.xml").contains(">_xlfn.NOSUCHFN(Data!$A$1)<"))
    val reread = XlsxReader.read(path).fold(err => fail(s"read failed: $err"), identity)
    assertEquals(cfFormulas(reread.sheets(0)), Vector("_xlfn.NOSUCHFN(A1)"))
    assertEquals(reread.metadata.definedNames.map(_.formula), Vector("_xlfn.NOSUCHFN(Data!$A$1)"))
  }
