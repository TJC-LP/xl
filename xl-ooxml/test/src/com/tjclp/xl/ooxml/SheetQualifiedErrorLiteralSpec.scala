package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.jdk.CollectionConverters.*

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.lint.{Finding, LintCategory, WorkbookLint}
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import com.tjclp.xl.workbooks.DefinedName
import munit.FunSuite

/**
 * GH-687: Excel spells a defined name (or a formula) whose target was deleted as `Sheet!#REF!` —
 * `_xlnm._FilterDatabase` after the filtered rows go, a user name after its sheet range goes. The
 * `#` after the qualifier's `!` opens an error literal; it is not the spill operator applied to
 * `Sheet!`. The writer used to rewrite such a name into `_xlfn.ANCHORARRAY(Sheet!)REF!` on every
 * in-memory write, and the lint reported it as a bare spill reference.
 */
class SheetQualifiedErrorLiteralSpec extends FunSuite:

  private def tempXlsx(label: String): Path =
    val p = Files.createTempFile(s"xl-gh687-$label-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found")
    finally zip.close()

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

  private def slice(xml: String, open: String, close: String): String =
    val start = xml.indexOf(open)
    val end = xml.lastIndexOf(close)
    if start < 0 || end < 0 then "" else xml.substring(start, end + close.length)

  private def xlfnFindings(path: Path): Vector[Finding] =
    WorkbookLint
      .lint(path)
      .fold(err => fail(s"lint must not error: ${err.message}"), identity)
      .filter(_.category == LintCategory.XlfnMissing)

  // Placeholders xl writes cleanly, patched afterwards into the text Excel itself writes — so the
  // source is an Excel-shaped file, never one xl's own writer produced
  private val filterPlaceholder = "Support!$Z$99"
  private val oldPlaceholder = "Sheet1!$Z$98"
  private val cellPlaceholder = "Sheet1!$Z$97+1"

  /**
   * Two sheets; `_xlnm._FilterDatabase` (hidden, scoped to Support) = `Support!#REF!`, user name
   * `Old` = `Sheet1!#REF!`, and Support!B1 = `Sheet1!#REF!+1`.
   */
  private def excelShapedSource(
    label: String,
    backend: XmlBackend = XmlBackend.ScalaXml
  ): Path =
    val base = Workbook(
      Vector(
        Sheet("Sheet1").put(ref"A1" -> "x"),
        Sheet("Support").put(ref"A1" -> 1).put(ref"B1", CellValue.Formula(cellPlaceholder, None))
      )
    )
    val named = base.copy(metadata =
      base.metadata.copy(definedNames =
        Vector(
          DefinedName("_xlnm._FilterDatabase", filterPlaceholder, Some(1), hidden = true),
          DefinedName("Old", oldPlaceholder)
        )
      )
    )
    val written = tempXlsx(s"$label-placeholder")
    XlsxWriter
      .writeWith(named, written, WriterConfig(backend = backend))
      .fold(err => fail(s"write failed: $err"), identity)
    patchedZip(written, label)(
      _.replace(filterPlaceholder, "Support!#REF!")
        .replace(oldPlaceholder, "Sheet1!#REF!")
        .replace(cellPlaceholder, "Sheet1!#REF!+1")
    )

  private def sheetXmlWithFormula(path: Path): String =
    Vector("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml")
      .map(entryText(path, _))
      .find(_.contains("<f>"))
      .getOrElse(fail("no worksheet carries the formula"))

  test("GH-687: the Excel-shaped source holds Sheet!#REF! and lints clean") {
    val src = excelShapedSource("src")
    val wbXml = entryText(src, "xl/workbook.xml")
    assert(wbXml.contains(">Support!#REF!</definedName>"), wbXml)
    assert(wbXml.contains(">Sheet1!#REF!</definedName>"), wbXml)
    assert(sheetXmlWithFormula(src).contains("<f>Sheet1!#REF!+1</f>"))
    assertEquals(xlfnFindings(src), Vector.empty)
  }

  test("GH-687: the reader keeps Sheet!#REF! names and formulas verbatim") {
    val wb = XlsxReader
      .read(excelShapedSource("read"))
      .fold(err => fail(s"read failed: $err"), identity)
    val formulas = wb.metadata.definedNames.map(dn => dn.name -> dn.formula).toMap
    assertEquals(formulas.get("Old"), Some("Sheet1!#REF!"))
    wb.metadata.definedNames
      .find(_.name == "_xlnm._FilterDatabase")
      .foreach(dn => assertEquals(dn.formula, "Support!#REF!"))
    val support = wb.sheets.find(_.name.value == "Support").getOrElse(fail("no Support sheet"))
    support.cells.get(ref"B1").map(_.value) match
      case Some(CellValue.Formula(expr, _, _)) =>
        assertEquals(expr.stripPrefix("="), "Sheet1!#REF!+1")
      case other => fail(s"B1 must be the formula, got $other")
  }

  // the source is written by the same backend: the two serialize attributes in different orders,
  // which is not what this pins
  for backend <- Vector(XmlBackend.ScalaXml, XmlBackend.SaxStax) do
    test(s"GH-687: read → put → write keeps Sheet!#REF! byte-identical ($backend)") {
      val src = excelShapedSource(s"rt-${backend.toString.toLowerCase}", backend)
      val wb = XlsxReader.read(src).fold(err => fail(s"read failed: $err"), identity)
      val edited = wb.sheets.zipWithIndex.foldLeft(wb) { case (acc, (sheet, _)) =>
        acc.put(sheet.put(ref"C5" -> 42))
      }
      val out = tempXlsx(s"rt-out-${backend.toString.toLowerCase}")
      XlsxWriter
        .writeWith(edited, out, WriterConfig(backend = backend))
        .fold(err => fail(s"write failed: $err"), identity)
      val outWb = entryText(out, "xl/workbook.xml")
      assert(!outWb.contains("ANCHORARRAY"), outWb)
      assert(outWb.contains(">Support!#REF!</definedName>"), outWb)
      assert(outWb.contains(">Sheet1!#REF!</definedName>"), outWb)
      assertEquals(
        slice(outWb, "<definedNames>", "</definedNames>"),
        slice(entryText(src, "xl/workbook.xml"), "<definedNames>", "</definedNames>")
      )
      val outSheet = sheetXmlWithFormula(out)
      assert(outSheet.contains("<f>Sheet1!#REF!+1</f>"), outSheet)
      assert(!outSheet.contains("ANCHORARRAY"), outSheet)
      assertEquals(xlfnFindings(out), Vector.empty)
    }

  test("GH-687: an authored Sheet!#REF! name and formula are written as spelled") {
    val base = Workbook(
      Vector(Sheet("Sheet1").put(ref"A1", CellValue.Formula("=SUM(Sheet1!#REF!,A2)", None)))
    )
    val wb = base.withDefinedName("Old", "Sheet1!#REF!")
    val out = tempXlsx("authored")
    XlsxWriter.write(wb, out).fold(err => fail(s"write failed: $err"), identity)
    assert(entryText(out, "xl/workbook.xml").contains(">Sheet1!#REF!</definedName>"))
    assert(entryText(out, "xl/worksheets/sheet1.xml").contains("<f>SUM(Sheet1!#REF!,A2)</f>"))
    assertEquals(xlfnFindings(out), Vector.empty)
  }
