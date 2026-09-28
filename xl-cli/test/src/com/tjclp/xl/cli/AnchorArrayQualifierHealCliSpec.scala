package com.tjclp.xl.cli

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.jdk.CollectionConverters.*

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cf.CfRule
import com.tjclp.xl.cli.contract.{CliHarness, CliRun}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.XlsxWriter
import com.tjclp.xl.sheets.{DataValidation, DvKind}
import com.tjclp.xl.workbooks.DefinedName

/**
 * GH-687, the heal end to end through the real command tree: a book xl 0.23.0–0.23.1 wrote carries
 * `_xlfn.ANCHORARRAY(Sheet!)REF!` where Excel wrote `Sheet!#REF!` — in defined names, cell
 * formulas, CF and DV formulas. `xl lint` reports it (`anchorarray-qualifier-corrupt`, repair); an
 * in-memory write (`put`, `recalc`) restores Excel's spelling in workbook.xml and in the worksheet
 * it edits, after which the lint is clean. `--stream put` copies untouched parts verbatim and heals
 * nothing — the lint's remedy says so.
 */
class AnchorArrayQualifierHealCliSpec extends CatsEffectSuite:

  private val category = "anchorarray-qualifier-corrupt"

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found")
    finally zip.close()

  private def definedNames(path: Path): String =
    val xml = entryText(path, "xl/workbook.xml")
    val start = xml.indexOf("<definedNames>")
    val end = xml.indexOf("</definedNames>")
    if start < 0 || end < 0 then "" else xml.substring(start, end)

  /** The worksheet part holding Support's formulas. */
  private def supportSheetXml(path: Path): String =
    Vector("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml")
      .map(entryText(path, _))
      .find(_.contains("<f>"))
      .getOrElse(fail("no worksheet carries the formulas"))

  /** Placeholder xl writes cleanly → the text xl 0.23.x left behind. */
  private val corruption: Vector[(String, String)] = Vector(
    "Support!$Z$99" -> "_xlfn.ANCHORARRAY(Support!)REF!",
    "Sheet1!$Z$98" -> "_xlfn.ANCHORARRAY(Sheet1!)REF!",
    "Sheet1!$Z$97+1" -> "_xlfn.ANCHORARRAY(Sheet1!)REF!+1",
    "Sheet1!$Z$95=1" -> "_xlfn.ANCHORARRAY(Sheet1!)REF!=1",
    "Sheet1!$Z$94=0" -> "_xlfn.ANCHORARRAY(Sheet1!)REF!=0"
  )

  private def customDv(formula: String): DataValidation.Rules =
    DataValidation.Rules(Vector.empty, DvKind.Custom(formula))

  /**
   * `_xlnm._FilterDatabase` (hidden, scoped to Support) and `Old` in workbook.xml; Support!B1, a CF
   * expression rule and a custom DV on Support — all in the 0.23.x spelling. `liveFormula` adds an
   * uncached `Support!C1 = A1*2`, so a recalculation has a cache to refresh on the sheet.
   */
  private def corruptedSource(liveFormula: Boolean): IO[Path] = IO.blocking {
    val corrupt = Sheet("Support")
      .put(ref"A1" -> 1)
      .put(ref"B1", CellValue.Formula(corruption(2)._1, None))
      .conditionalFormat(ref"A1:A3", CfRule.Expression(corruption(3)._1, None, 1))
      .withDataValidation(ref"D1:D3", customDv(corruption(4)._1))
    val support =
      if liveFormula then corrupt.put(ref"C1", CellValue.Formula("A1*2", None)) else corrupt
    val base = Workbook(Vector(Sheet("Sheet1").put(ref"A1" -> "x"), support))
    val named = base.copy(metadata =
      base.metadata.copy(definedNames =
        Vector(
          DefinedName("_xlnm._FilterDatabase", corruption(0)._1, Some(1), hidden = true),
          DefinedName("Old", corruption(1)._1)
        )
      )
    )
    val placeholder = Files.createTempFile("xl-gh687-heal-cli-placeholder-", ".xlsx")
    XlsxWriter.write(named, placeholder).fold(err => fail(s"write failed: $err"), identity)
    val out = Files.createTempFile("xl-gh687-heal-cli-src-", ".xlsx")
    val zip = new ZipFile(placeholder.toFile)
    val zos = new ZipOutputStream(Files.newOutputStream(out))
    try
      zip.entries().asScala.foreach { entry =>
        val bytes = zip.getInputStream(entry).readAllBytes()
        val content =
          if entry.getName.endsWith(".xml") then
            corruption
              .foldLeft(new String(bytes, StandardCharsets.UTF_8)) { case (acc, (from, to)) =>
                acc.replace(from, to)
              }
              .getBytes(StandardCharsets.UTF_8)
          else bytes
        zos.putNextEntry(new ZipEntry(entry.getName))
        zos.write(content)
        zos.closeEntry()
      }
    finally
      zos.close()
      zip.close()
    Files.deleteIfExists(placeholder)
    out
  }

  private def findings(run: CliRun): Vector[ujson.Value] =
    ujson.read(run.stdout)("data")("findings").arr.toVector

  private def corruptParts(run: CliRun): Vector[String] =
    findings(run).filter(_("category").str == category).map(_("part").str).sorted

  private def assertHealed(out: Path): Unit =
    val names = definedNames(out)
    assert(!names.contains("ANCHORARRAY"), names)
    assert(names.contains(">Support!#REF!</definedName>"), names)
    assert(names.contains(">Sheet1!#REF!</definedName>"), names)
    val sheet = supportSheetXml(out)
    assert(!sheet.contains("ANCHORARRAY"), sheet)
    assert(sheet.contains("<f>Sheet1!#REF!+1</f>"), sheet)
    assert(sheet.contains("<formula>Sheet1!#REF!=1</formula>"), sheet)
    assert(sheet.contains("<formula1>Sheet1!#REF!=0</formula1>"), sheet)

  private def tempOut(label: String): IO[Path] =
    IO.blocking(Files.createTempFile(s"xl-gh687-heal-cli-$label-", ".xlsx"))

  test("GH-687: xl lint reports the 0.23.x corruption as a repair finding, in both modes") {
    for
      src <- corruptedSource(liveFormula = false)
      dom <- CliHarness.run("--json", "lint", src.toString)
      streamed <- CliHarness.run("--json", "--stream", "lint", src.toString)
    yield
      assertEquals(dom.exit, 1, dom.toString)
      assertEquals(corruptParts(dom).size, 2, dom.stdout)
      assert(corruptParts(dom).contains("xl/workbook.xml"), dom.stdout)
      val all = findings(dom).filter(_("category").str == category)
      assert(all.forall(_("severity").str == "repair"), dom.stdout)
      assert(all.forall(_("message").str.contains("xl 0.23.0–0.23.1")), dom.stdout)
      // the corruption is not a bare post-2007 call
      assert(!findings(dom).exists(_("category").str == "xlfn-missing"), dom.stdout)
      assertEquals(streamed.stdout, dom.stdout)
  }

  test("GH-687: xl put (in memory) restores Sheet!#REF! everywhere and lints clean") {
    for
      src <- corruptedSource(liveFormula = false)
      out <- tempOut("put")
      run <- CliHarness.run(
        "-f",
        src.toString,
        "-s",
        "Support",
        "-o",
        out.toString,
        "put",
        "C5",
        "1"
      )
      lint <- CliHarness.run("--json", "lint", out.toString)
    yield
      assertEquals(run.exit, 0, run.toString)
      assertHealed(out)
      assertEquals(lint.exit, 0, lint.stdout)
      assertEquals(findings(lint), Vector.empty, lint.stdout)
  }

  test("GH-687: xl recalc that refreshes a cache on the sheet restores Sheet!#REF! everywhere") {
    for
      src <- corruptedSource(liveFormula = true)
      out <- tempOut("recalc")
      run <- CliHarness.run("-f", src.toString, "-o", out.toString, "recalc")
      lint <- CliHarness.run("--json", "lint", out.toString)
    yield
      assertEquals(run.exit, 0, run.toString)
      assertHealed(out)
      assertEquals(corruptParts(lint), Vector.empty, lint.stdout)
  }

  test("GH-687: xl recalc that changes nothing copies the book verbatim (the remedy says so)") {
    for
      src <- corruptedSource(liveFormula = false)
      out <- tempOut("recalc-clean")
      run <- CliHarness.run("-f", src.toString, "-o", out.toString, "recalc")
      lint <- CliHarness.run("--json", "lint", out.toString)
    yield
      assertEquals(run.exit, 0, run.toString)
      assert(definedNames(out).contains(">_xlfn.ANCHORARRAY(Support!)REF!</definedName>"))
      assertEquals(corruptParts(lint).size, 2, lint.stdout)
      val message = findings(lint).find(_("category").str == category).map(_("message").str)
      assert(message.exists(_.contains("xl recalc")), lint.stdout)
  }

  test("GH-687: an in-memory put on ANOTHER sheet heals workbook.xml, not the untouched sheet") {
    for
      src <- corruptedSource(liveFormula = false)
      out <- tempOut("put-other")
      run <- CliHarness.run(
        "-f",
        src.toString,
        "-s",
        "Sheet1",
        "-o",
        out.toString,
        "put",
        "B1",
        "1"
      )
      lint <- CliHarness.run("--json", "lint", out.toString)
    yield
      assertEquals(run.exit, 0, run.toString)
      val names = definedNames(out)
      assert(!names.contains("ANCHORARRAY"), names)
      assert(names.contains(">Support!#REF!</definedName>"), names)
      // the untouched worksheet is copied verbatim — the lint still names it, and only it
      assert(supportSheetXml(out).contains("_xlfn.ANCHORARRAY(Sheet1!)REF!+1"))
      assertEquals(corruptParts(lint).size, 1, lint.stdout)
      assert(!corruptParts(lint).contains("xl/workbook.xml"), lint.stdout)
  }

  test("GH-687: xl --stream put copies untouched parts verbatim and heals nothing (documented)") {
    for
      src <- corruptedSource(liveFormula = false)
      out <- tempOut("stream")
      run <- CliHarness.run(
        "-f",
        src.toString,
        "-s",
        "Support",
        "-o",
        out.toString,
        "--stream",
        "put",
        "C5",
        "1"
      )
      lint <- CliHarness.run("--json", "lint", out.toString)
    yield
      assertEquals(run.exit, 0, run.toString)
      assert(definedNames(out).contains(">_xlfn.ANCHORARRAY(Support!)REF!</definedName>"))
      assert(supportSheetXml(out).contains("_xlfn.ANCHORARRAY(Sheet1!)REF!"))
      assertEquals(corruptParts(lint).size, 2, lint.stdout)
      val message = findings(lint).find(_("category").str == category).map(_("message").str)
      assert(message.exists(_.contains("--stream")), lint.stdout)
  }
