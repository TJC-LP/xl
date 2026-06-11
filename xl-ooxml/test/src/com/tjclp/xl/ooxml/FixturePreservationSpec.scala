package com.tjclp.xl.ooxml

import java.io.StringReader
import java.nio.file.{Files, Path}
import java.security.MessageDigest
import java.util.zip.ZipFile
import javax.xml.parsers.DocumentBuilderFactory

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig

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
    writeToWith(wb, label, WriterConfig())

  private def writeToWith(wb: Workbook, label: String, config: WriterConfig): Path =
    val out = Files.createTempFile(s"xl-preserve-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter
      .writeWith(wb, out, config)
      .fold(err => fail(s"$label write failed: ${err.message}"), _ => ())
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

    // GH-291: openpyxl binds xmlns:r ON the <drawing> element itself; regenerating a modified
    // worksheet must keep that prefix bound (hoisted to the root) or the workbook is corrupt.
    // Both writer backends regenerate modified sheets, so both must produce valid output.
    List(
      "ScalaXml" -> WriterConfig.scalaXml,
      "SaxStax" -> WriterConfig.saxStax
    ).foreach { case (backend, config) =>
      test(s"$fixture/$backend: cell-modified write stays namespace-valid and re-readable") {
        val (in, wb) = readFixture(fixture)
        val sheetName = wb.sheets(0).name
        val modified = wb
          .update(sheetName, _.put(ref"A20", CellValue.Text("modified after read")))
          .fold(err => fail(s"$fixture update failed: $err"), identity)
        val out = writeToWith(modified, s"valid-$backend", config)

        // Drawing survived regeneration
        val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
        assert(sheetXml.contains("<drawing"), "drawing element vanished entirely")

        // Namespace well-formedness: a namespace-aware parser accepts the worksheet
        // (an unbound r: prefix on <drawing> throws here)
        val doc = parseNamespaceAware(sheetXml, s"$fixture/$backend regenerated worksheet")

        // The drawing's r:id attribute resolves in the relationships namespace
        val drawings = doc.getElementsByTagNameNS("*", "drawing")
        assertEquals(drawings.getLength, 1, "expected exactly one <drawing>")
        val drawingId = Option(drawings.item(0))
          .collect { case e: org.w3c.dom.Element => e.getAttributeNS(nsRel, "id") }
          .filter(_.nonEmpty)
          .getOrElse(fail("drawing r:id missing or not in the relationships namespace"))

        // Zip-level structure: every r:id referenced from the worksheet exists in sheet rels
        val relIds = relationshipIdsOf(out, "xl/worksheets/_rels/sheet1.xml.rels")
        val referenced = relationshipRefs(doc)
        assert(referenced.contains(drawingId), "drawing r:id not among collected references")
        assert(
          referenced.subsetOf(relIds),
          s"worksheet references missing from sheet rels: ${(referenced -- relIds).mkString(", ")}"
        )

        // The output re-reads cleanly (the original GH-291 failure mode)
        val reread = XlsxReader
          .read(out)
          .fold(err => fail(s"re-read of modified write failed: ${err.message}"), identity)
        assertEquals(reread.sheets(0).name, sheetName)

        // Visual parts survive byte-identically
        assertEquals(visualParts(out), visualParts(in), "visual parts drifted")
      }
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

  private val nsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  /** Parse with a namespace-aware parser; an unbound prefix fails the test. */
  private def parseNamespaceAware(xml: String, label: String): org.w3c.dom.Document =
    val factory = DocumentBuilderFactory.newInstance()
    factory.setNamespaceAware(true)
    try factory.newDocumentBuilder().parse(new org.xml.sax.InputSource(new StringReader(xml)))
    catch
      case e: org.xml.sax.SAXException =>
        fail(s"$label is not namespace-well-formed: ${e.getMessage}")

  /** Every r:id value (relationships-namespace `id` attribute) referenced in the document. */
  private def relationshipRefs(doc: org.w3c.dom.Document): Set[String] =
    val all = doc.getElementsByTagNameNS("*", "*")
    (0 until all.getLength).flatMap { i =>
      Option(all.item(i)).collect { case e: org.w3c.dom.Element =>
        Option(e.getAttributeNS(nsRel, "id")).filter(_.nonEmpty)
      }.flatten
    }.toSet

  /** Relationship Ids declared in a rels part of the zip. */
  private def relationshipIdsOf(path: Path, relsEntry: String): Set[String] =
    val relsDoc = parseNamespaceAware(entryText(path, relsEntry), relsEntry)
    val rels = relsDoc.getElementsByTagNameNS("*", "Relationship")
    (0 until rels.getLength).flatMap { i =>
      Option(rels.item(i)).collect { case e: org.w3c.dom.Element =>
        Option(e.getAttribute("Id")).filter(_.nonEmpty)
      }.flatten
    }.toSet
