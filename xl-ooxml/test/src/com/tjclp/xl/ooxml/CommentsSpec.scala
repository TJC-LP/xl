package com.tjclp.xl.ooxml

import munit.FunSuite
import scala.xml.*
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.api.*
import com.tjclp.xl.macros.ref
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.richtext.RichText.* // Import DSL extensions
import com.tjclp.xl.cells.Comment

/**
 * Tests for OOXML Comments parsing and serialization.
 *
 * Verifies:
 *   - Parse comments XML
 *   - Round-trip (parse → encode → parse)
 *   - Rich text support
 *   - Author attribution
 *   - Forwards compatibility (preserve unknown attributes/children)
 */
class CommentsSpec extends FunSuite:

  test("parse simple comment with single author") {
    val xml =
      <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <authors>
          <author>John Doe</author>
        </authors>
        <commentList>
          <comment ref="A1" authorId="0">
            <text>
              <r><t>This is a comment</t></r>
            </text>
          </comment>
        </commentList>
      </comments>

    val result = OoxmlComments.fromXml(xml)

    assert(result.isRight, s"Expected successful parse, got: $result")
    val comments = result.toOption.getOrElse(fail("Expected Right"))

    assertEquals(comments.authors, Vector("John Doe"))
    assertEquals(comments.comments.size, 1)

    val comment = comments.comments.headOption.getOrElse(fail("Expected comment"))
    assertEquals(comment.ref, ref"A1")
    assertEquals(comment.authorId, 0)
    assertEquals(comment.text.toPlainText, "This is a comment")
  }

  test("parse comment with multiple authors") {
    val xml =
      <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <authors>
          <author>Alice</author>
          <author>Bob</author>
          <author>Charlie</author>
        </authors>
        <commentList>
          <comment ref="B2" authorId="1">
            <text><r><t>Bob's comment</t></r></text>
          </comment>
          <comment ref="C3" authorId="2">
            <text><r><t>Charlie's note</t></r></text>
          </comment>
        </commentList>
      </comments>

    val result = OoxmlComments.fromXml(xml)

    assert(result.isRight)
    val comments = result.toOption.getOrElse(fail("Expected Right"))

    assertEquals(comments.authors.size, 3)
    assertEquals(comments.authors(0), "Alice")
    assertEquals(comments.authors(1), "Bob")
    assertEquals(comments.authors(2), "Charlie")

    assertEquals(comments.comments.size, 2)
    assertEquals(comments.comments(0).authorId, 1)
    assertEquals(comments.comments(1).authorId, 2)
  }

  test("parse comment with rich text formatting") {
    val xml =
      <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <authors>
          <author>Editor</author>
        </authors>
        <commentList>
          <comment ref="D4" authorId="0">
            <text>
              <r>
                <rPr>
                  <b/>
                  <color rgb="FF0000"/>
                </rPr>
                <t>Important: </t>
              </r>
              <r>
                <t>Review this value</t>
              </r>
            </text>
          </comment>
        </commentList>
      </comments>

    val result = OoxmlComments.fromXml(xml)

    assert(result.isRight)
    val comments = result.toOption.getOrElse(fail("Expected Right"))

    val comment = comments.comments.headOption.getOrElse(fail("Expected comment"))
    assertEquals(comment.text.runs.size, 2)

    // First run: bold + red
    val run1 = comment.text.runs(0)
    assertEquals(run1.text, "Important: ")
    assert(run1.font.exists(_.bold))
    assert(run1.font.exists(_.color.isDefined))

    // Second run: plain
    val run2 = comment.text.runs(1)
    assertEquals(run2.text, "Review this value")
  }

  test("round-trip: parse → encode → parse preserves content") {
    val originalXml =
      <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <authors>
          <author>Test Author</author>
        </authors>
        <commentList>
          <comment ref="E5" authorId="0" guid="{12345678-1234-1234-1234-123456789ABC}">
            <text>
              <r><t>Round-trip test</t></r>
            </text>
          </comment>
        </commentList>
      </comments>

    val parsed1 = OoxmlComments.fromXml(originalXml)
    assert(parsed1.isRight)

    val encoded = OoxmlComments.toXml(parsed1.toOption.getOrElse(fail("Expected Right")))
    val parsed2 = OoxmlComments.fromXml(encoded)

    assert(parsed2.isRight)

    val comments1 = parsed1.toOption.getOrElse(fail("Expected Right"))
    val comments2 = parsed2.toOption.getOrElse(fail("Expected Right"))

    assertEquals(comments1.authors, comments2.authors)
    assertEquals(comments1.comments.size, comments2.comments.size)

    val c1 = comments1.comments.headOption.getOrElse(fail("Expected comment"))
    val c2 = comments2.comments.headOption.getOrElse(fail("Expected comment"))

    assertEquals(c1.ref, c2.ref)
    assertEquals(c1.authorId, c2.authorId)
    assertEquals(c1.guid, c2.guid)
    assertEquals(c1.text.toPlainText, c2.text.toPlainText)
  }

  test("preserve unknown attributes for forwards compatibility") {
    val xml =
      <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <authors>
          <author>Author</author>
        </authors>
        <commentList>
          <comment ref="F6" authorId="0" futureAttr="value123">
            <text><r><t>Test</t></r></text>
          </comment>
        </commentList>
      </comments>

    val result = OoxmlComments.fromXml(xml)
    assert(result.isRight)

    val comments = result.toOption.getOrElse(fail("Expected Right"))
    val comment = comments.comments.headOption.getOrElse(fail("Expected comment"))
    assert(comment.otherAttrs.contains("futureAttr"))
    assertEquals(comment.otherAttrs("futureAttr"), "value123")

    // Round-trip should preserve it
    val encoded = OoxmlComments.toXml(comments)
    val reparsed = OoxmlComments.fromXml(encoded)

    assert(reparsed.isRight)
    val comment2 = reparsed.toOption.getOrElse(fail("Expected Right")).comments.headOption.getOrElse(fail("Expected comment"))
    assertEquals(comment2.otherAttrs.get("futureAttr"), Some("value123"))
  }

  test("parse comment with whitespace preservation") {
    val xml =
      <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <authors>
          <author>User</author>
        </authors>
        <commentList>
          <comment ref="G7" authorId="0">
            <text>
              <r><t xml:space="preserve">  Leading and trailing spaces  </t></r>
            </text>
          </comment>
        </commentList>
      </comments>

    val result = OoxmlComments.fromXml(xml)
    assert(result.isRight)

    val comment = result.toOption.getOrElse(fail("Expected Right")).comments.headOption.getOrElse(fail("Expected comment"))
    // Whitespace should be preserved
    assertEquals(comment.text.toPlainText, "  Leading and trailing spaces  ")
  }

  test("error on missing ref attribute") {
    val xml =
      <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <authors><author>A</author></authors>
        <commentList>
          <comment authorId="0">
            <text><r><t>Missing ref</t></r></text>
          </comment>
        </commentList>
      </comments>

    val result = OoxmlComments.fromXml(xml)
    assert(result.isLeft, "Should fail when ref attribute is missing")
  }

  test("error on missing authorId attribute") {
    val xml =
      <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <authors><author>A</author></authors>
        <commentList>
          <comment ref="A1">
            <text><r><t>Missing authorId</t></r></text>
          </comment>
        </commentList>
      </comments>

    val result = OoxmlComments.fromXml(xml)
    assert(result.isLeft, "Should fail when authorId attribute is missing")
  }

  test("error on invalid cell reference") {
    val xml =
      <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <authors><author>A</author></authors>
        <commentList>
          <comment ref="INVALID123!" authorId="0">
            <text><r><t>Bad ref</t></r></text>
          </comment>
        </commentList>
      </comments>

    val result = OoxmlComments.fromXml(xml)
    assert(result.isLeft, "Should fail when cell reference is invalid")
  }

  test("convertToDomainComments fails on out-of-bounds authorId") {
    val comments = OoxmlComments(
      authors = Vector("Author0"),
      comments = Vector(
        OoxmlComment(
          ref = ref"A1",
          authorId = 5,
          text = RichText.plain("Bad authorId"),
          shapeId = 0
        )
      )
    )

    val result = XlsxReader.convertToDomainComments(comments, "xl/comments1.xml")
    assert(result.isLeft, s"Expected failure for invalid authorId, got: $result")
  }

  test("encode simple comment produces valid XML") {
    val comment = OoxmlComment(
      ref = ref"H8",
      authorId = 0,
      text = RichText.plain("Test comment"),
      shapeId = 0,
      guid = None // GUIDs optional per OOXML spec (omitted for deterministic output)
    )

    val comments = OoxmlComments(
      authors = Vector("Test User"),
      comments = Vector(comment)
    )

    val xml = OoxmlComments.toXml(comments)

    // Verify structure
    assertEquals(xml.label, "comments")
    assert((xml \ "authors" \ "author").nonEmpty)
    assert((xml \ "commentList" \ "comment").nonEmpty)

    // Verify required attributes
    val commentElem = (xml \ "commentList" \ "comment").headOption match
      case Some(e: Elem) => e
      case _ => fail("Expected comment element")
    assertEquals((commentElem \ "@ref").text, "H8")
    assertEquals((commentElem \ "@authorId").text, "0")
    assertEquals((commentElem \ "@shapeId").text, "0")

    // Verify GUID omitted (deterministic output)
    val xrNs = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
    val uidAttr = commentElem.attribute(xrNs, "uid").map(_.text)
    assertEquals(uidAttr, None)
  }

  test("encode rich text comment with formatting") {
    val richText = "Bold: ".bold.red + "Normal text"

    val comment = OoxmlComment(
      ref = ref"I9",
      authorId = 1,
      text = richText
    )

    val comments = OoxmlComments(
      authors = Vector("Author1", "Author2"),
      comments = Vector(comment)
    )

    val xml = OoxmlComments.toXml(comments)

    // Verify text structure has multiple runs
    val textElem = (xml \ "commentList" \ "comment" \ "text").headOption match
      case Some(e: Elem) => e
      case _ => fail("Expected text element")
    val runs = textElem \ "r"

    assert(runs.size >= 2, "Should have at least 2 runs (formatted + plain)")

    // First run should have formatting
    val run1 = runs.headOption match
      case Some(e: Elem) => e
      case _ => fail("Expected run element")
    assert((run1 \ "rPr").nonEmpty, "First run should have formatting properties")
    assert((run1 \ "rPr" \ "b").nonEmpty, "First run should be bold")
    assert((run1 \ "rPr" \ "color").nonEmpty, "First run should have color")
  }

  test("GUID preservation when provided (for read round-trips)") {
    val commentWithGuid = OoxmlComment(
      ref = ref"A1",
      authorId = 0,
      text = RichText.plain("Test"),
      shapeId = 0,
      guid = Some("{PRESERVED-GUID}")
    )

    val xml = OoxmlComments.toXml(OoxmlComments(Vector("Author"), Vector(commentWithGuid)))
    val commentElem = (xml \ "commentList" \ "comment").headOption match
      case Some(e: Elem) => e
      case _ => fail("Expected comment element")

    val xrNs = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
    val uidAttr = commentElem.attribute(xrNs, "uid").map(_.text)
    assertEquals(uidAttr, Some("{PRESERVED-GUID}"))
  }

  test("round-trip with authored comments strips author prefix correctly") {
    // Simulate what XlsxWriter produces: author name prepended to text
    val authorName = "John Doe"
    val commentText = "Important note"

    // Create OOXML comment as writer would generate it
    val richTextWithAuthor = com.tjclp.xl.richtext.RichText(Vector(
      com.tjclp.xl.richtext.TextRun(s"$authorName:", Some(com.tjclp.xl.styles.font.Font.default.copy(bold = true))),
      com.tjclp.xl.richtext.TextRun(s"\n$commentText")
    ))

    val ooxmlComment = OoxmlComment(
      ref = ref"A1",
      authorId = 0,
      text = richTextWithAuthor,
      shapeId = 0
    )

    val ooxmlComments = OoxmlComments(
      authors = Vector(authorName),
      comments = Vector(ooxmlComment)
    )

    // Convert to domain (should strip author prefix)
    val domainMap = XlsxReader.convertToDomainComments(ooxmlComments) match
      case Right(m) => m
      case Left(err) => fail(s"Failed to convert: $err")

    val domainComment = domainMap.get(ref"A1").getOrElse(fail("Expected comment at A1"))

    // Verify author extracted correctly
    assertEquals(domainComment.author, Some(authorName))

    // Verify text has author prefix stripped (should just be "Important note")
    assertEquals(domainComment.text.toPlainText, commentText)

    // Verify no double-prefixing after round-trip
    assert(!domainComment.text.toPlainText.contains("John Doe:"),
      "Author prefix should be stripped on read")
  }

  test("stripAuthorPrefix handles Excel-style CRLF author prefix") {
    val authorName = "Excel User"
    val richTextWithAuthor = com.tjclp.xl.richtext.RichText(
      Vector(
        com.tjclp.xl.richtext.TextRun(
          s"$authorName: ",
          Some(com.tjclp.xl.styles.font.Font.default.copy(bold = true))
        ),
        com.tjclp.xl.richtext.TextRun("\r\nActual comment")
      )
    )

    val ooxmlComment = OoxmlComment(
      ref = ref"A1",
      authorId = 0,
      text = richTextWithAuthor
    )

    val ooxmlComments = OoxmlComments(
      authors = Vector(authorName),
      comments = Vector(ooxmlComment)
    )

    val domainMap = XlsxReader.convertToDomainComments(ooxmlComments) match
      case Right(m) => m
      case Left(err) => fail(s"Failed to convert: $err")

    val domainComment = domainMap.get(ref"A1").getOrElse(fail("Comment missing"))
    assertEquals(domainComment.author, Some(authorName))
    assertEquals(domainComment.text.toPlainText, "Actual comment")
  }

  test("empty authors list is preserved") {
    val comments = OoxmlComments(
      authors = Vector.empty,
      comments = Vector.empty
    )

    val xml = OoxmlComments.toXml(comments)
    val reparsed = OoxmlComments.fromXml(xml)

    assert(reparsed.isRight)
    val reparsedComments = reparsed.toOption.getOrElse(fail("Expected Right"))
    assertEquals(reparsedComments.authors, Vector.empty)
    assertEquals(reparsedComments.comments, Vector.empty)
  }

  test("VML shape IDs do not collide across sheets with dense comments") {
    def commentsForSheet(count: Int): OoxmlComments =
      val comments = (0 until count).map { idx =>
        OoxmlComment(
          ref = ARef.from1(1, idx + 1),
          authorId = 0,
          text = RichText.plain(s"Comment $idx")
        )
      }.toVector
      OoxmlComments(
        authors = Vector("Author"),
        comments = comments
      )

    val sheet0 = commentsForSheet(150) // exceeds legacy stride of 100
    val sheet1 = commentsForSheet(5)

    val ids0 = extractShapeIds(VmlDrawing.generateForComments(sheet0, 0))
    val ids1 = extractShapeIds(VmlDrawing.generateForComments(sheet1, 1))

    assertEquals(ids0.size, 150)
    assertEquals(ids1.size, 5)
    assert(ids0.intersect(ids1).isEmpty, s"Shape IDs should not collide: ${ids0.intersect(ids1)}")
  }

  test("round-trip preserves very long author names") {
    val longAuthor = "L" * 300
    val baseSheet = Sheet("Sheet1") match
      case Right(s) => s
      case Left(err) => fail(s"Failed to build sheet: $err")

    val sheet = baseSheet.comment(
      ref"A1",
      Comment.plainText("Long author note", Some(longAuthor))
    )
    val workbook = Workbook(Vector(sheet))

    val bytes = XlsxWriter.writeToBytes(workbook) match
      case Right(b) => b
      case Left(err) => fail(s"Write failed: $err")

    val reread = XlsxReader.readFromBytes(bytes) match
      case Right(wb) => wb
      case Left(err) => fail(s"Read failed: $err")

    val rereadSheet = reread.sheets.headOption.getOrElse(fail("Expected sheet"))
    val comment = rereadSheet.comments.get(ref"A1").getOrElse(
      fail("Missing comment after round-trip")
    )

    assertEquals(comment.author, Some(longAuthor))
  }

  test("comments on merged cells survive round-trip") {
    val baseSheet = Sheet("Sheet1") match
      case Right(s) => s
      case Left(err) => fail(s"Failed to build sheet: $err")

    val mergedRange = CellRange(ref"A1", ref"B1")
    val withMerge = baseSheet.copy(mergedRanges = baseSheet.mergedRanges + mergedRange)
    val sheet = withMerge.comment(ref"A1", Comment.plainText("Merged cell note", Some("Author")))
    val workbook = Workbook(Vector(sheet))

    val bytes = XlsxWriter.writeToBytes(workbook) match
      case Right(b) => b
      case Left(err) => fail(s"Write failed: $err")

    val reread = XlsxReader.readFromBytes(bytes) match
      case Right(wb) => wb
      case Left(err) => fail(s"Read failed: $err")

    val rereadSheet = reread.sheets.headOption.getOrElse(fail("Expected sheet"))
    assert(rereadSheet.mergedRanges.contains(mergedRange), "Merged range should round-trip")
    assertEquals(rereadSheet.comments.get(ref"A1").flatMap(_.author), Some("Author"))
  }

  private def extractShapeIds(vml: String): Set[Int] =
    "_x0000_s(\\d+)".r.findAllMatchIn(vml).map(_.group(1).toInt).toSet
