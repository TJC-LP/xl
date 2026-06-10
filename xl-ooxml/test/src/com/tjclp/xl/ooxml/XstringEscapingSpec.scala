package com.tjclp.xl.ooxml

import java.util.zip.ZipInputStream

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{Cell, CellValue, Comment}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.{SstPolicy, WriterConfig, XmlBackend}
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.styles.font.Font
import munit.FunSuite

/**
 * GH-288: carriage returns in cell text must survive the OOXML round-trip.
 *
 * XML 1.0 line-ending normalization folds raw CR and CRLF in element content to LF on PARSE, so a
 * writer that emits raw `\r` loses it (probe: "line1\rline2\r\nline3" read back as
 * "line1\nline2\nline3"). ECMA-376 (§22.9.2.19 ST_Xstring) requires CR to be stored as the
 * `_x000D_` escape; text that LITERALLY contains an `_xHHHH_` pattern must protect its leading
 * underscore as `_x005F_`.
 *
 * Covers BOTH writer backends (ScalaXml + SaxStax), BOTH string encodings (inline strings + SST),
 * rich text runs, comments, and formula cached string values.
 */
class XstringEscapingSpec extends FunSuite:

  private val crText = "line1\rline2\r\nline3"
  private val literalEscape = "_x000D_ is how Excel stores CR" // must stay LITERAL
  private val configs: List[(String, WriterConfig)] = List(
    "ScalaXml/inline" -> WriterConfig(backend = XmlBackend.ScalaXml, sstPolicy = SstPolicy.Never),
    "ScalaXml/sst" -> WriterConfig(backend = XmlBackend.ScalaXml, sstPolicy = SstPolicy.Always),
    "SaxStax/inline" -> WriterConfig(backend = XmlBackend.SaxStax, sstPolicy = SstPolicy.Never),
    "SaxStax/sst" -> WriterConfig(backend = XmlBackend.SaxStax, sstPolicy = SstPolicy.Always)
  )

  private def writeBytes(wb: Workbook, config: WriterConfig): Array[Byte] =
    val tempPath = java.nio.file.Files.createTempFile("xl-xstring-", ".xlsx")
    try
      XlsxWriter
        .writeWith(wb, tempPath, config)
        .fold(err => fail(s"write failed: ${err.message}"), identity)
      java.nio.file.Files.readAllBytes(tempPath)
    finally java.nio.file.Files.deleteIfExists(tempPath)

  private def roundTrip(wb: Workbook, config: WriterConfig): Workbook =
    XlsxReader
      .readFromBytes(writeBytes(wb, config))
      .fold(err => fail(s"read failed: ${err.message}"), identity)

  private def textAt(wb: Workbook, sheetName: String, r: com.tjclp.xl.addressing.ARef): String =
    wb(sheetName).toOption
      .flatMap(_.cells.get(r))
      .map(_.value)
      .collect { case CellValue.Text(s) => s }
      .getOrElse(fail(s"$sheetName!${r.toA1} is not a Text cell"))

  private def show(s: String): String =
    s.flatMap {
      case '\r' => "\\r"
      case '\n' => "\\n"
      case c => c.toString
    }

  // ========== Plain text: CR survives on every backend × encoding ==========

  configs.foreach { case (label, config) =>
    test(s"GH-288: CR in plain cell text round-trips ($label)") {
      val sheet = Sheet("Data").put(Cell(ref"A1", CellValue.Text(crText)))
      val back = roundTrip(Workbook(Vector(sheet)), config)
      assertEquals(
        show(textAt(back, "Data", ref"A1")),
        show(crText),
        s"$label: CR lost or transformed in round-trip"
      )
    }

    test(s"GH-288: literal _xHHHH_ text stays literal ($label)") {
      val sheet = Sheet("Data").put(Cell(ref"A1", CellValue.Text(literalEscape)))
      val back = roundTrip(Workbook(Vector(sheet)), config)
      assertEquals(textAt(back, "Data", ref"A1"), literalEscape)
    }
  }

  // ========== Rich text runs ==========

  configs.foreach { case (label, config) =>
    test(s"GH-288: CR in rich text run round-trips ($label)") {
      val rich = RichText(
        Vector(
          TextRun("bold\rrun", Some(Font.default.copy(bold = true))),
          TextRun("plain\r\nrun", None)
        )
      )
      val sheet = Sheet("Data").put(Cell(ref"A1", CellValue.RichText(rich)))
      val back = roundTrip(Workbook(Vector(sheet)), config)
      val got = back("Data").toOption
        .flatMap(_.cells.get(ref"A1"))
        .map(_.value)
        .collect { case CellValue.RichText(rt) => rt.toPlainText }
        .getOrElse(fail("A1 is not RichText"))
      assertEquals(show(got), show("bold\rrunplain\r\nrun"), s"$label: CR lost in rich text")
    }
  }

  // ========== Comments ==========

  test("GH-288: CR in comment text round-trips") {
    val sheet = Sheet("Data")
      .put(Cell(ref"A1", CellValue.Text("x")))
      .comment(ref"A1", Comment.plainText("note\rwith CR", Some("Reviewer")))
    val back = roundTrip(Workbook(Vector(sheet)), WriterConfig())
    val got = back("Data").toOption
      .flatMap(_.comments.get(ref"A1"))
      .map(_.text.toPlainText)
      .getOrElse(fail("comment missing"))
    assertEquals(show(got), show("note\rwith CR"), "comment CR lost")
  }

  // ========== Formula cached string values (<v> of t=\"str\") ==========

  test("GH-288: CR in formula cached string value round-trips") {
    val sheet = Sheet("Data").put(
      Cell(ref"A1", CellValue.Formula("CONCATENATE(\"a\",CHAR(13))", Some(CellValue.Text("a\rb"))))
    )
    val back = roundTrip(Workbook(Vector(sheet)), WriterConfig())
    val got = back("Data").toOption
      .flatMap(_.cells.get(ref"A1"))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(CellValue.Text(s))) => s }
      .getOrElse(fail("A1 cached text missing"))
    assertEquals(show(got), show("a\rb"), "cached string CR lost")
  }

  // ========== Raw bytes: the escape is actually on disk, never a raw CR in <t> ==========

  configs.foreach { case (label, config) =>
    test(s"GH-288: written XML stores _x000D_, never a raw CR in text content ($label)") {
      val sheet = Sheet("Data").put(Cell(ref"A1", CellValue.Text(crText)))
      val bytes = writeBytes(Workbook(Vector(sheet)), config)
      val parts = readParts(bytes)
      val textPart =
        if config.sstPolicy == SstPolicy.Always then parts("xl/sharedStrings.xml")
        else parts("xl/worksheets/sheet1.xml")
      assert(textPart.contains("_x000D_"), s"$label: no _x000D_ escape found in:\n$textPart")
      assert(!textPart.contains("\r"), s"$label: raw CR byte still present in:\n$textPart")
    }
  }

  private def readParts(bytes: Array[Byte]): Map[String, String] =
    val zis = new ZipInputStream(new java.io.ByteArrayInputStream(bytes))
    try
      Iterator
        .continually(zis.getNextEntry)
        .takeWhile(_ != null)
        .map(entry => entry.getName -> new String(zis.readAllBytes(), "UTF-8"))
        .toMap
    finally zis.close()
