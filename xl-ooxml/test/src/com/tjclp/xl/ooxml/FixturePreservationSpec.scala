package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path}
import java.security.MessageDigest
import java.util.zip.ZipFile

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref

/**
 * Visual-preservation pre-check for GH-240 (ground truth for the GH-221/GH-222 work): what happens
 * to chart (xl/charts/), drawing (xl/drawings/), and media (xl/media/) parts when a real openpyxl
 * file flows through XlsxReader.read -> XlsxWriter.write?
 *
 * These specs pin the CURRENT behavior precisely. If they start failing, the preservation contract
 * changed - update deliberately, not accidentally.
 */
class FixturePreservationSpec extends FunSuite:

  private val visualPrefixes = List("xl/charts/", "xl/drawings/", "xl/media/")

  /** name -> (size, sha256) for every zip entry under the visual prefixes. */
  private def visualParts(path: Path): Map[String, (Long, String)] =
    val zip = new ZipFile(path.toFile)
    try
      val entries = scala.jdk.CollectionConverters.EnumerationHasAsScala(zip.entries()).asScala
      entries
        .filter(e => visualPrefixes.exists(e.getName.startsWith))
        .map { e =>
          val bytes = zip.getInputStream(e).readAllBytes()
          val digest = MessageDigest.getInstance("SHA-256").digest(bytes)
          e.getName -> (bytes.length.toLong, digest.map("%02x".format(_)).mkString)
        }
        .toMap
    finally zip.close()

  private def readFixture(name: String): (Path, Workbook) =
    val path = TestFixtures.copyToTemp(name)
    val wb = XlsxReader
      .read(path)
      .fold(err => fail(s"$name failed to read: ${err.message}"), identity)
    (path, wb)

  private def writeTo(wb: Workbook, label: String): Path =
    val out = Files.createTempFile(s"xl-preserve-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter.write(wb, out).fold(err => fail(s"$label write failed: ${err.message}"), _ => ())
    out

  List("chart-bar.xlsx", "image.xlsx").foreach { fixture =>
    test(s"$fixture: untouched read->write preserves charts/drawings/media byte-identically") {
      val (in, wb) = readFixture(fixture)
      val out = writeTo(wb, "clean")
      val before = visualParts(in)
      val after = visualParts(out)
      assert(before.nonEmpty, s"$fixture has no visual parts - fixture broken")
      assertEquals(after, before, s"$fixture visual parts drifted on untouched write")
    }

    test(s"$fixture: cell-modified write keeps visual part bytes (current behavior)") {
      val (in, wb) = readFixture(fixture)
      val sheetName = wb.sheets(0).name
      val modified = wb
        .update(sheetName, _.put(ref"A20", CellValue.Text("modified after read")))
        .fold(err => fail(s"$fixture update failed: $err"), identity)
      val out = writeTo(modified, "modified")
      val before = visualParts(in)
      val after = visualParts(out)
      assertEquals(after, before, s"$fixture visual parts drifted on modified write")
    }

    test(s"$fixture: KNOWN BUG - cell-modified write emits unbound r: prefix on <drawing>") {
      // openpyxl binds xmlns:r ON the <drawing> element itself:
      //   <drawing xmlns:r="...officeDocument/2006/relationships" r:id="rId1"/>
      // When a cell modification forces worksheet regeneration, the preserved
      // <drawing> is re-emitted as <drawing r:id="rId1"/> WITHOUT the xmlns:r
      // binding, so the output worksheet is not well-formed XML and XlsxReader
      // refuses to re-read the file it just wrote. The chart/image survives in
      // the zip but the workbook is corrupt. This pins the current behavior for
      // GH-221/GH-222; when fixed, flip these assertions to a successful re-read.
      val (_, wb) = readFixture(fixture)
      val sheetName = wb.sheets(0).name
      val modified = wb
        .update(sheetName, _.put(ref"A20", CellValue.Text("modified after read")))
        .fold(err => fail(s"$fixture update failed: $err"), identity)
      val out = writeTo(modified, "nsbug")
      val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
      assert(sheetXml.contains("<drawing"), "drawing element vanished entirely")
      assert(sheetXml.contains("r:id"), "drawing lost its r:id attribute")
      assert(
        !sheetXml.contains("xmlns:r"),
        "regenerated worksheet now binds xmlns:r - the namespace bug may be fixed; " +
          "flip this test to assert a successful re-read instead"
      )
      XlsxReader.read(out) match
        case Left(err) =>
          assert(
            err.message.contains("prefix") || err.message.contains("parse"),
            s"expected namespace parse failure, got: ${err.message}"
          )
        case Right(_) =>
          fail(
            "XlsxReader re-read its own modified-write output - the xmlns:r bug " +
              "may be fixed; flip this test to assert the round-trip instead"
          )
    }
  }

  test("chart-bar.xlsx: workbook relationships to drawings survive a modified write") {
    val (in, wb) = readFixture("chart-bar.xlsx")
    val sheetName = wb.sheets(0).name
    val modified = wb
      .update(sheetName, _.put(ref"A20", CellValue.Number(BigDecimal(99))))
      .fold(err => fail(s"update failed: $err"), identity)
    val out = writeTo(modified, "rels")
    // The worksheet must still reference its drawing, else Excel shows no chart
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(
      sheetXml.contains("<drawing"),
      "worksheet lost its <drawing> element on modified write - chart would vanish in Excel"
    )
    val relsXml = entryText(out, "xl/worksheets/_rels/sheet1.xml.rels")
    assert(
      relsXml.contains("drawing1.xml"),
      "sheet rels lost the drawing relationship on modified write"
    )
  }

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) =>
          new String(zip.getInputStream(e).readAllBytes(), java.nio.charset.StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()
