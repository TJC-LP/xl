package com.tjclp.xl.cli

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.jdk.CollectionConverters.*

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{CliHarness, CliRun}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.XlsxWriter
import com.tjclp.xl.workbooks.DefinedName

/**
 * GH-687, end to end through the real command tree: Excel spells a name or formula whose target was
 * deleted as `Sheet!#REF!`. `xl -f in.xlsx -o out.xlsx put …` — in memory or `--stream` — must
 * carry it byte for byte (it used to become `_xlfn.ANCHORARRAY(Sheet!)REF!`), and `xl lint` must
 * not report it as a bare spill reference (`xlfn-missing`).
 */
class SheetQualifiedErrorLiteralCliSpec extends CatsEffectSuite:

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found")
    finally zip.close()

  private def slice(xml: String, open: String, close: String): String =
    val start = xml.indexOf(open)
    val end = xml.lastIndexOf(close)
    if start < 0 || end < 0 then "" else xml.substring(start, end + close.length)

  private def definedNames(path: Path): String =
    slice(entryText(path, "xl/workbook.xml"), "<definedNames>", "</definedNames>")

  /** The worksheet part holding Support!B1's formula. */
  private def supportSheetXml(path: Path): String =
    Vector("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml")
      .map(entryText(path, _))
      .find(_.contains("<f>"))
      .getOrElse(fail("no worksheet carries the formula"))

  private val filterPlaceholder = "Support!$Z$99"
  private val oldPlaceholder = "Sheet1!$Z$98"
  private val cellPlaceholder = "Sheet1!$Z$97+1"

  /**
   * `_xlnm._FilterDatabase` (hidden, scoped to Support) = `Support!#REF!`, `Old` = `Sheet1!#REF!`,
   * Support!B1 = `Sheet1!#REF!+1` — xl writes placeholders, then the zip is patched into the text
   * Excel writes, so the source never passed through xl's own formula writer.
   */
  private def excelShapedSource(): IO[Path] = IO.blocking {
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
    val placeholder = Files.createTempFile("xl-gh687-cli-placeholder-", ".xlsx")
    XlsxWriter.write(named, placeholder).fold(err => fail(s"write failed: $err"), identity)
    val out = Files.createTempFile("xl-gh687-cli-src-", ".xlsx")
    val zip = new ZipFile(placeholder.toFile)
    val zos = new ZipOutputStream(Files.newOutputStream(out))
    try
      zip.entries().asScala.foreach { entry =>
        val bytes = zip.getInputStream(entry).readAllBytes()
        val content =
          if entry.getName.endsWith(".xml") then
            new String(bytes, StandardCharsets.UTF_8)
              .replace(filterPlaceholder, "Support!#REF!")
              .replace(oldPlaceholder, "Sheet1!#REF!")
              .replace(cellPlaceholder, "Sheet1!#REF!+1")
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

  private def xlfnCategories(run: CliRun): List[String] =
    ujson
      .read(run.stdout)("data")("findings")
      .arr
      .map(_("category").str)
      .filter(_ == "xlfn-missing")
      .toList

  private def assertCarried(src: Path, out: Path): Unit =
    val names = definedNames(out)
    assert(!names.contains("ANCHORARRAY"), names)
    assertEquals(names, definedNames(src))
    val sheet = supportSheetXml(out)
    assert(sheet.contains("<f>Sheet1!#REF!+1</f>"), sheet)
    assert(!sheet.contains("ANCHORARRAY"), sheet)

  test("GH-687: xl put (in memory) keeps Sheet!#REF! names and formulas byte-identical") {
    for
      src <- excelShapedSource()
      out <- IO.blocking(Files.createTempFile("xl-gh687-cli-out-", ".xlsx"))
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
      other <- IO.blocking(Files.createTempFile("xl-gh687-cli-out1-", ".xlsx"))
      run1 <- CliHarness.run(
        "-f",
        src.toString,
        "-s",
        "Sheet1",
        "-o",
        other.toString,
        "put",
        "A1",
        "1"
      )
    yield
      assertEquals(run.exit, 0, run.toString)
      assertEquals(run1.exit, 0, run1.toString)
      assertCarried(src, out)
      assertCarried(src, other)
  }

  test("GH-687: xl --stream put keeps Sheet!#REF! names and formulas byte-identical") {
    for
      src <- excelShapedSource()
      out <- IO.blocking(Files.createTempFile("xl-gh687-cli-stream-", ".xlsx"))
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
    yield
      assertEquals(run.exit, 0, run.toString)
      assertCarried(src, out)
  }

  test("GH-687: xl lint reports no xlfn-missing on Sheet!#REF!, before and after a write") {
    for
      src <- excelShapedSource()
      out <- IO.blocking(Files.createTempFile("xl-gh687-cli-lint-", ".xlsx"))
      _ <- CliHarness.run("-f", src.toString, "-s", "Support", "-o", out.toString, "put", "C5", "1")
      before <- CliHarness.run("--json", "lint", src.toString)
      after <- CliHarness.run("--json", "lint", out.toString)
      streamed <- CliHarness.run("--json", "--stream", "lint", src.toString)
    yield
      assertEquals(xlfnCategories(before), Nil, before.stdout)
      assertEquals(xlfnCategories(after), Nil, after.stdout)
      assertEquals(xlfnCategories(streamed), Nil, streamed.stdout)
      assert(!before.stdout.contains("ANCHORARRAY"), before.stdout)
  }
