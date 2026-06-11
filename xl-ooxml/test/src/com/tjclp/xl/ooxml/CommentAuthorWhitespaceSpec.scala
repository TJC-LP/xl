package com.tjclp.xl.ooxml

import java.util.zip.ZipInputStream

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{Cell, CellValue, Comment}
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * GH-290: comment author whitespace must not corrupt the round-trip.
 *
 * Probe: author " Bob " read back as Some("Bob") (the reader trims authors), AND the writer's
 * author-prefix presentation run leaked into the comment text — text "note" returned " Bob :\nnote"
 * because XlsxReader.stripAuthorPrefix compared the TRIMMED author against the UNTRIMMED " Bob :"
 * first run and failed to match.
 *
 * The TRIM IS CANONICAL at write time: XlsxWriter.buildCommentsData normalizes the author once
 * (trim; whitespace-only → unauthored) and builds the <authors> entry AND the bold prefix run from
 * the SAME normalized value, so the reader's strip always matches and text never gains the prefix.
 */
class CommentAuthorWhitespaceSpec extends FunSuite:

  private def roundTrip(wb: Workbook): (Workbook, Map[String, String]) =
    val tempPath = java.nio.file.Files.createTempFile("xl-comment-author-", ".xlsx")
    try
      XlsxWriter.write(wb, tempPath).fold(err => fail(s"write failed: ${err.message}"), identity)
      val bytes = java.nio.file.Files.readAllBytes(tempPath)
      val back = XlsxReader
        .readFromBytes(bytes)
        .fold(err => fail(s"read failed: ${err.message}"), identity)
      (back, readParts(bytes))
    finally java.nio.file.Files.deleteIfExists(tempPath)

  private def readParts(bytes: Array[Byte]): Map[String, String] =
    val zis = new ZipInputStream(new java.io.ByteArrayInputStream(bytes))
    try
      Iterator
        .continually(zis.getNextEntry)
        .takeWhile(_ != null)
        .map(entry => entry.getName -> new String(zis.readAllBytes(), "UTF-8"))
        .toMap
    finally zis.close()

  private def commentAt(wb: Workbook, r: com.tjclp.xl.addressing.ARef): Comment =
    wb("Data").toOption
      .flatMap(_.comments.get(r))
      .getOrElse(fail(s"comment missing at ${r.toA1}"))

  private def wbWithComment(author: Option[String], text: String = "note"): Workbook =
    val sheet = Sheet("Data")
      .put(Cell(ref"A1", CellValue.Text("x")))
      .comment(ref"A1", Comment.plainText(text, author))
    Workbook(Vector(sheet))

  test("GH-290: author ' Bob ' round-trips as 'Bob' and the text gains NO prefix") {
    val (back, parts) = roundTrip(wbWithComment(Some(" Bob ")))
    val comment = commentAt(back, ref"A1")
    assertEquals(comment.author, Some("Bob"), "author should be trimmed (canonical)")
    assertEquals(
      comment.text.toPlainText,
      "note",
      "author-prefix presentation run leaked into the comment text"
    )
    val commentsXml = parts
      .collectFirst {
        case (name, content) if name.startsWith("xl/comments") => content
      }
      .getOrElse(fail("no comments part written"))
    assert(
      commentsXml.contains("<author>Bob</author>"),
      s"<authors> should store the TRIMMED author. XML:\n$commentsXml"
    )
  }

  test("GH-290: already-trimmed author is unchanged and text intact") {
    val (back, _) = roundTrip(wbWithComment(Some("Bob")))
    val comment = commentAt(back, ref"A1")
    assertEquals(comment.author, Some("Bob"))
    assertEquals(comment.text.toPlainText, "note")
  }

  test("GH-290: whitespace-only author is canonicalized to unauthored") {
    val (back, _) = roundTrip(wbWithComment(Some("   ")))
    val comment = commentAt(back, ref"A1")
    assertEquals(comment.author, None, "whitespace-only author should canonicalize to None")
    assertEquals(comment.text.toPlainText, "note")
  }

  test("GH-290: unauthored comment still round-trips") {
    val (back, _) = roundTrip(wbWithComment(None))
    val comment = commentAt(back, ref"A1")
    assertEquals(comment.author, None)
    assertEquals(comment.text.toPlainText, "note")
  }

  test("GH-290: interior whitespace in an author is preserved (only edges trimmed)") {
    val (back, _) = roundTrip(wbWithComment(Some("  Bob  Smith ")))
    val comment = commentAt(back, ref"A1")
    assertEquals(comment.author, Some("Bob  Smith"))
    assertEquals(comment.text.toPlainText, "note")
  }
