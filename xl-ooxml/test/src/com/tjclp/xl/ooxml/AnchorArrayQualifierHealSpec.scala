package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.jdk.CollectionConverters.*

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cf.CfRule
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.lint.{Finding, LintCategory, WorkbookLint}
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import com.tjclp.xl.sheets.{DataValidation, DvKind}
import com.tjclp.xl.workbooks.DefinedName
import munit.FunSuite

/**
 * GH-687, the heal: xl 0.23.0–0.23.1 rewrote Excel's `Sheet!#REF!` into
 * `_xlfn.ANCHORARRAY(Sheet!)REF!` (and `@Sheet!#REF!` into
 * `_xlfn.SINGLE(_xlfn.ANCHORARRAY(Sheet!))REF!`) on every in-memory write — in defined names, cell
 * formulas, CF and DV formulas. A book carrying that corruption must read back as the text Excel
 * wrote, and the next in-memory write must restore it on every part it regenerates: the
 * preserved-bytes gates for defined names, CF and DV compare MODELS, and the corrupt spelling now
 * parses to the healed model, so each gate needs the storage-form check too.
 */
class AnchorArrayQualifierHealSpec extends FunSuite:

  private def tempXlsx(label: String): Path =
    val p = Files.createTempFile(s"xl-gh687-heal-$label-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found")
    finally zip.close()

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

  /**
   * Placeholder → (the text xl 0.23.x wrote, the text Excel wrote). xl writes the placeholders
   * cleanly; the zip is then patched into the 0.23.x spelling, so the source is exactly what a
   * 0.23.x in-memory write left behind.
   */
  private val slots: Vector[(String, String, String)] = Vector(
    ("Support!$Z$99", "_xlfn.ANCHORARRAY(Support!)REF!", "Support!#REF!"),
    ("Sheet1!$Z$98", "_xlfn.ANCHORARRAY(Sheet1!)REF!", "Sheet1!#REF!"),
    ("Sheet1!$Z$97+1", "_xlfn.ANCHORARRAY(Sheet1!)REF!+1", "Sheet1!#REF!+1"),
    ("Sheet1!$Z$96", "_xlfn.SINGLE(_xlfn.ANCHORARRAY(Sheet1!))REF!", "_xlfn.SINGLE(Sheet1!#REF!)"),
    ("Sheet1!$Z$95=1", "_xlfn.ANCHORARRAY(Sheet1!)REF!=1", "Sheet1!#REF!=1"),
    ("'My Sheet'!$Z$94=0", "_xlfn.ANCHORARRAY('My Sheet'!)N/A=0", "'My Sheet'!#N/A=0")
  )

  private def placeholder(i: Int): String = slots(i)._1

  /** The enum case widens to `DataValidation`; `withDataValidation` wants the `Rules` case. */
  private def customDv(formula: String): DataValidation.Rules =
    DataValidation.Rules(Vector.empty, DvKind.Custom(formula))

  /**
   * Three sheets. workbook.xml: `_xlnm._FilterDatabase` (hidden, scoped to Support) and the user
   * name `Old`. Support: B1 and B2 cell formulas, a CF expression rule and a custom DV — every
   * formula-text slot a 0.23.x write corrupted.
   */
  private def corruptedSource(label: String, backend: XmlBackend): Path =
    val support = Sheet("Support")
      .put(ref"A1" -> 1)
      .put(ref"B1", CellValue.Formula(placeholder(2), None))
      .put(ref"B2", CellValue.Formula(placeholder(3), None))
      .conditionalFormat(ref"A1:A3", CfRule.Expression(placeholder(4), None, 1))
      .withDataValidation(ref"D1:D3", customDv(placeholder(5)))
    val base = Workbook(
      Vector(Sheet("Sheet1").put(ref"A1" -> "x"), support, Sheet("My Sheet").put(ref"A1" -> 2))
    )
    val named = base.copy(metadata =
      base.metadata.copy(definedNames =
        Vector(
          DefinedName("_xlnm._FilterDatabase", placeholder(0), Some(1), hidden = true),
          DefinedName("Old", placeholder(1))
        )
      )
    )
    val written = tempXlsx(s"$label-placeholder")
    XlsxWriter
      .writeWith(named, written, WriterConfig(backend = backend))
      .fold(err => fail(s"write failed: $err"), identity)
    patchedZip(written, label)(xml =>
      slots.foldLeft(xml) { case (acc, (from, corrupt, _)) => acc.replace(from, corrupt) }
    )

  private def supportXml(path: Path): String =
    Vector("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml")
      .map(entryText(path, _))
      .find(_.contains("<f>"))
      .getOrElse(fail("no worksheet carries the formulas"))

  private def definedNamesXml(path: Path): String =
    val xml = entryText(path, "xl/workbook.xml")
    val start = xml.indexOf("<definedNames>")
    val end = xml.indexOf("</definedNames>")
    if start < 0 || end < 0 then "" else xml.substring(start, end)

  test("GH-687: the fixture carries the 0.23.x corruption in every slot") {
    val src = corruptedSource("fixture", XmlBackend.ScalaXml)
    val names = definedNamesXml(src)
    assert(names.contains(">_xlfn.ANCHORARRAY(Support!)REF!</definedName>"), names)
    assert(names.contains(">_xlfn.ANCHORARRAY(Sheet1!)REF!</definedName>"), names)
    val sheet = supportXml(src)
    assert(sheet.contains("<f>_xlfn.ANCHORARRAY(Sheet1!)REF!+1</f>"), sheet)
    assert(sheet.contains("<f>_xlfn.SINGLE(_xlfn.ANCHORARRAY(Sheet1!))REF!</f>"), sheet)
    assert(sheet.contains("<formula>_xlfn.ANCHORARRAY(Sheet1!)REF!=1</formula>"), sheet)
    assert(sheet.contains("<formula1>_xlfn.ANCHORARRAY('My Sheet'!)N/A=0</formula1>"), sheet)
  }

  test("GH-687: the reader hands the model Excel's spelling in every slot") {
    val wb = XlsxReader
      .read(corruptedSource("read", XmlBackend.ScalaXml))
      .fold(err => fail(s"read failed: $err"), identity)
    val names = wb.metadata.definedNames.map(dn => dn.name -> dn.formula).toMap
    assertEquals(names.get("Old"), Some("Sheet1!#REF!"))
    assertEquals(names.get("_xlnm._FilterDatabase"), Some("Support!#REF!"))
    val support = wb.sheets.find(_.name.value == "Support").getOrElse(fail("no Support sheet"))
    def formulaAt(r: com.tjclp.xl.addressing.ARef): String =
      support.cells.get(r).map(_.value) match
        case Some(CellValue.Formula(expr, _, _)) => expr.stripPrefix("=")
        case other => fail(s"$r must be a formula, got $other")
    assertEquals(formulaAt(ref"B1"), "Sheet1!#REF!+1")
    assertEquals(formulaAt(ref"B2"), "@Sheet1!#REF!")
    val cf = support.conditionalFormats.collect {
      case com.tjclp.xl.cf.ConditionalFormat.Rules(_, rules, _) =>
        rules.collect { case CfRule.Expression(f, _, _, _) => f }
    }.flatten
    assertEquals(cf, Vector("Sheet1!#REF!=1"))
    val dv = support.dataValidations.collect {
      case DataValidation.Rules(_, DvKind.Custom(f), _, _, _) => f
    }
    assertEquals(dv, Vector("'My Sheet'!#N/A=0"))
  }

  for backend <- Vector(XmlBackend.ScalaXml, XmlBackend.SaxStax) do
    test(s"GH-687: read → put → write restores Sheet!#REF! in every slot ($backend)") {
      val tag = backend.toString.toLowerCase
      val src = corruptedSource(s"rt-$tag", backend)
      val wb = XlsxReader.read(src).fold(err => fail(s"read failed: $err"), identity)
      val support = wb.sheets.find(_.name.value == "Support").getOrElse(fail("no Support sheet"))
      val edited = wb.put(support.put(ref"C5" -> 42))
      val out = tempXlsx(s"rt-out-$tag")
      XlsxWriter
        .writeWith(edited, out, WriterConfig(backend = backend))
        .fold(err => fail(s"write failed: $err"), identity)
      val names = definedNamesXml(out)
      assert(!names.contains("ANCHORARRAY"), names)
      assert(names.contains(">Support!#REF!</definedName>"), names)
      assert(names.contains(">Sheet1!#REF!</definedName>"), names)
      val sheet = supportXml(out)
      assert(!sheet.contains("ANCHORARRAY"), sheet)
      assert(sheet.contains("<f>Sheet1!#REF!+1</f>"), sheet)
      assert(sheet.contains("<f>_xlfn.SINGLE(Sheet1!#REF!)</f>"), sheet)
      assert(sheet.contains("<formula>Sheet1!#REF!=1</formula>"), sheet)
      assert(sheet.contains("<formula1>'My Sheet'!#N/A=0</formula1>"), sheet)
      // the lint that flags the source calls the output clean, in both scanning modes
      assertEquals(corruptFindings(src, streaming = false).map(_.part).sorted, sourceParts(src))
      assertEquals(corruptFindings(out, streaming = false), Vector.empty)
      assertEquals(corruptFindings(out, streaming = true), Vector.empty)
      assertEquals(WorkbookLint.lint(out).map(_.map(_.category)), Right(Vector.empty))
    }

  private def corruptFindings(path: Path, streaming: Boolean): Vector[Finding] =
    val result = if streaming then WorkbookLint.lintStream(path) else WorkbookLint.lint(path)
    result
      .fold(err => fail(s"lint must not error: ${err.message}"), identity)
      .filter(_.category == LintCategory.AnchorArrayQualifierCorrupt)

  /** workbook.xml plus the worksheet part holding Support's formulas, sorted. */
  private def sourceParts(src: Path): Vector[String] =
    val sheetPart =
      Vector("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml")
        .find(p => entryText(src, p).contains("<f>"))
        .getOrElse(fail("no worksheet carries the formulas"))
    Vector("xl/workbook.xml", sheetPart).sorted
