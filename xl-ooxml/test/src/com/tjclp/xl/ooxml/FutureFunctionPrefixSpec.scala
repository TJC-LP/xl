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
import com.tjclp.xl.ooxml.lint.{Finding, LintCategory, WorkbookLint}
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
    val sheet = gh577Slots(Sheet("Data").put(ref"A1" -> "a").put(ref"B1" -> 1).put(ref"C1" -> "a"))
    Workbook(Vector(sheet))
      .withDefinedName("Best", "MAXIFS(Data!$B$1:$B$3,Data!$A$1:$A$3,Data!$C$1)")

  /** Author the GH-577 CF blocks and DV entries (bare spelling) onto `sheet`, in document order. */
  private def gh577Slots(sheet: Sheet): Sheet =
    sheet
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

  /** The file a bare-writing producer (openpyxl) emits: every storage prefix stripped by hand. */
  private def stripPrefixes(xml: String): String =
    xml.replace("_xlfn._xlws.", "").replace("_xlfn.", "")

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
    val bare = patchedZip(written, "cfdv-bare")(stripPrefixes)
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

  /** The `xlfn-missing` findings of a written file (the lint's rule is the writer's own). */
  private def xlfnFindings(path: Path): Vector[Finding] =
    WorkbookLint
      .lint(path)
      .fold(err => fail(s"lint must not error: ${err.message}"), identity)
      .filter(_.category == LintCategory.XlfnMissing)

  /** The substring from the first `open` to the end of the last `close` (empty when absent). */
  private def slice(xml: String, open: String, close: String): String =
    val start = xml.indexOf(open)
    val end = xml.lastIndexOf(close)
    if start < 0 || end < 0 then "" else xml.substring(start, end + close.length)

  /** Every `dxfId="n"` on the sheet must index into the styles part's `<dxfs count="N">`. */
  private def assertDxfRefsResolve(out: Path): Unit =
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    val dxfCount = """<dxfs count="(\d+)"""".r
      .findFirstMatchIn(entryText(out, "xl/styles.xml"))
      .map(_.group(1).toInt)
      .getOrElse(0)
    val refs = """dxfId="(\d+)"""".r.findAllMatchIn(sheetXml).map(_.group(1).toInt).toVector
    assert(refs.nonEmpty, s"the dxf-bearing rule must keep its dxfId: $sheetXml")
    refs.foreach(id => assert(id < dxfCount, s"dxfId=$id out of range (count $dxfCount)"))

  test("GH-593: an IDENTICAL re-author of bare CF / DV / name text is healed") {
    // The clean-compare gates (planCfWrites / planDvWrites / reconcileDefinedNames) compare models,
    // and bare text parses to the same model as prefixed text — so an identical re-author used to
    // copy the bare source bytes through (the GH-588 pin). The gates are now storage-form aware
    // through FormulaStorage.bareFutureCalls, the lint's own rule: a slot whose source text the
    // writer would still prefix is dirty by construction, and a write heals it.
    val written = writeGh577(XmlBackend.ScalaXml, "cfdv-same-src")
    val bare = patchedZip(written, "cfdv-same")(stripPrefixes)
    assert(xlfnFindings(bare).nonEmpty, "the fixture must lint dirty")
    val wb = XlsxReader.read(bare).fold(err => fail(s"read failed: $err"), identity)
    val source = wb.sheets(0)
    // re-author every CF block and DV entry from empty with the same text; the cell edit forces
    // the worksheet itself to regenerate, so the CF/DV gates decide, not a verbatim worksheet copy
    val reauthored = gh577Slots(
      source.copy(conditionalFormats = Vector.empty, dataValidations = Vector.empty)
    ).put(ref"G1" -> 2)
    assertEquals(reauthored.conditionalFormats, source.conditionalFormats)
    assertEquals(reauthored.dataValidations, source.dataValidations)
    val same = wb.put(reauthored).withDefinedName("Best", bareName)
    assertEquals(same.metadata.definedNames, wb.metadata.definedNames)
    val out = tempXlsx("cfdv-same-out")
    XlsxWriter.write(same, out).fold(err => fail(s"write failed: $err"), identity)
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(sheetXml.contains("r=\"G1\""), sheetXml)
    assertFragments(sheetXml, gh577SheetFragments)
    assert(entryText(out, "xl/workbook.xml").contains(gh577NameFragment))
    assertEquals(xlfnFindings(out), Vector.empty)
  }

  test("GH-593: a bare source with NO re-author but a cell edit heals CF, DV and definedNames") {
    val written = writeGh577(XmlBackend.ScalaXml, "cfdv-edit-src")
    val bare = patchedZip(written, "cfdv-edit")(stripPrefixes)
    val wb = XlsxReader.read(bare).fold(err => fail(s"read failed: $err"), identity)
    // the model is untouched; only a cell changes, so the worksheet regenerates and the gates
    // decide on storage form alone (workbook.xml is regenerated on every non-clean write)
    val out = tempXlsx("cfdv-edit-out")
    XlsxWriter
      .write(wb.put(wb.sheets(0).put(ref"G1" -> 2)), out)
      .fold(err => fail(s"write failed: $err"), identity)
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assertFragments(sheetXml, gh577SheetFragments)
    assert(entryText(out, "xl/workbook.xml").contains(gh577NameFragment))
    assertEquals(xlfnFindings(out), Vector.empty)
    // healing takes the dirty path, which re-plans the dxf table: the refs must still resolve and
    // the model must read back unchanged (rule text, dxf, validations, names)
    assertDxfRefsResolve(out)
    val reread = XlsxReader.read(out).fold(err => fail(s"read failed: $err"), identity)
    assertEquals(reread.sheets(0).conditionalFormats, wb.sheets(0).conditionalFormats)
    assertEquals(reread.sheets(0).dataValidations, wb.sheets(0).dataValidations)
    assertEquals(reread.metadata.definedNames, wb.metadata.definedNames)
    // and the healed file is now CLEAN for the gates: a further edit keeps the slots byte-stable
    val again = tempXlsx("cfdv-edit-again")
    XlsxWriter
      .write(reread.put(reread.sheets(0).put(ref"G2" -> 3)), again)
      .fold(err => fail(s"write failed: $err"), identity)
    val againXml = entryText(again, "xl/worksheets/sheet1.xml")
    assertEquals(
      slice(againXml, "<conditionalFormatting", "</conditionalFormatting>"),
      slice(sheetXml, "<conditionalFormatting", "</conditionalFormatting>")
    )
    assertEquals(
      slice(againXml, "<dataValidations", "</dataValidations>"),
      slice(sheetXml, "<dataValidations", "</dataValidations>")
    )
  }

  test("GH-593: Excel-authored (prefixed) CF / DV / name parts stay byte-identical over an edit") {
    val src = writeGh577(XmlBackend.ScalaXml, "cfdv-prefixed-src")
    val srcSheet = entryText(src, "xl/worksheets/sheet1.xml")
    val srcWorkbook = entryText(src, "xl/workbook.xml")
    assertEquals(xlfnFindings(src), Vector.empty)
    val wb = XlsxReader.read(src).fold(err => fail(s"read failed: $err"), identity)
    val out = tempXlsx("cfdv-prefixed-out")
    XlsxWriter
      .write(wb.put(wb.sheets(0).put(ref"G1" -> 2)), out)
      .fold(err => fail(s"write failed: $err"), identity)
    val outSheet = entryText(out, "xl/worksheets/sheet1.xml")
    assert(outSheet.contains("r=\"G1\""), outSheet) // regenerated ...
    // ... with the gates CLEAN: the source slots ride through byte for byte
    Vector(
      ("<conditionalFormatting", "</conditionalFormatting>"),
      ("<dataValidations", "</dataValidations>")
    ).foreach { case (open, close) =>
      val expected = slice(srcSheet, open, close)
      assert(expected.nonEmpty, s"$open missing from the source: $srcSheet")
      assertEquals(slice(outSheet, open, close), expected)
    }
    val names = slice(srcWorkbook, "<definedNames>", "</definedNames>")
    assert(names.nonEmpty, srcWorkbook)
    assertEquals(
      slice(entryText(out, "xl/workbook.xml"), "<definedNames>", "</definedNames>"),
      names
    )
  }

  test("GH-593: an untouched worksheet and a Preserved (unmodeled) rule keep bare text") {
    // The documented residuals: a worksheet the write does not regenerate rides verbatim, and a
    // CfRule.Preserved payload (here a timePeriod rule, outside the typed subset) is re-emitted
    // as captured — the gate fires on its bare <formula>, the block regenerates, the typed rules
    // heal, the preserved rule's text does not. definedNames heal on every non-clean write.
    val written = tempXlsx("cfdv-residual-src")
    val twoSheets = gh577Workbook.put(Sheet("Other").put(ref"A1" -> 1))
    XlsxWriter
      .writeWith(twoSheets, written, WriterConfig(backend = XmlBackend.ScalaXml))
      .fold(err => fail(s"write failed: $err"), identity)
    val preservedRule =
      "<cfRule type=\"timePeriod\" timePeriod=\"today\" priority=\"9\">" +
        "<formula>IFS(FLOOR(A1,1)=TODAY(),TRUE,FALSE)</formula></cfRule>"
    val blockOpen = "<conditionalFormatting sqref=\"A1:A3\">"
    val bare = patchedZip(written, "cfdv-residual") { xml =>
      stripPrefixes(xml).replace(blockOpen, blockOpen + preservedRule)
    }
    val bareSheet1 = entryText(bare, "xl/worksheets/sheet1.xml")
    assert(bareSheet1.contains(preservedRule), bareSheet1)
    val wb = XlsxReader.read(bare).fold(err => fail(s"read failed: $err"), identity)
    assert(
      wb.sheets(0).conditionalFormats.exists {
        case ConditionalFormat.Rules(_, rules, _) =>
          rules.exists { case _: CfRule.Preserved => true; case _ => false }
        case _ => false
      },
      wb.sheets(0).conditionalFormats.toString
    )
    // 1. edit only the OTHER sheet: sheet1 is copied verbatim (bare), the names still heal
    val untouched = tempXlsx("cfdv-residual-untouched")
    XlsxWriter
      .write(wb.put(wb.sheets(1).put(ref"A2" -> 2)), untouched)
      .fold(err => fail(s"write failed: $err"), identity)
    assertEquals(entryText(untouched, "xl/worksheets/sheet1.xml"), bareSheet1)
    assert(entryText(untouched, "xl/workbook.xml").contains(gh577NameFragment))
    assert(xlfnFindings(untouched).exists(_.part == "xl/worksheets/sheet1.xml"))
    // 2. edit sheet1: the typed rules heal, the Preserved rule's bare text is re-emitted as is
    val edited = tempXlsx("cfdv-residual-edited")
    XlsxWriter
      .write(wb.put(wb.sheets(0).put(ref"G1" -> 2)), edited)
      .fold(err => fail(s"write failed: $err"), identity)
    val editedSheet1 = entryText(edited, "xl/worksheets/sheet1.xml")
    assertFragments(editedSheet1, gh577SheetFragments)
    assert(
      editedSheet1.contains("<formula>IFS(FLOOR(A1,1)=TODAY(),TRUE,FALSE)</formula>"),
      editedSheet1
    )
    val residual = xlfnFindings(edited)
    assertEquals(residual.map(_.part), Vector("xl/worksheets/sheet1.xml"))
    assert(residual.forall(_.message.contains("1 formula(s)")), residual.toString)
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
