package com.tjclp.xl.ooxml

import munit.FunSuite
import scala.xml.*
import com.tjclp.xl.api.*
import com.tjclp.xl.macros.ref
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.richtext.RichText.* // Import DSL extensions

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
    val comments = result.toOption.get

    assertEquals(comments.authors, Vector("John Doe"))
    assertEquals(comments.comments.size, 1)

    val comment = comments.comments.head
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
    val comments = result.toOption.get

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
    val comments = result.toOption.get

    val comment = comments.comments.head
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

    val encoded = OoxmlComments.toXml(parsed1.toOption.get)
    val parsed2 = OoxmlComments.fromXml(encoded)

    assert(parsed2.isRight)

    val comments1 = parsed1.toOption.get
    val comments2 = parsed2.toOption.get

    assertEquals(comments1.authors, comments2.authors)
    assertEquals(comments1.comments.size, comments2.comments.size)

    val c1 = comments1.comments.head
    val c2 = comments2.comments.head

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

    val comment = result.toOption.get.comments.head
    assert(comment.otherAttrs.contains("futureAttr"))
    assertEquals(comment.otherAttrs("futureAttr"), "value123")

    // Round-trip should preserve it
    val encoded = OoxmlComments.toXml(result.toOption.get)
    val reparsed = OoxmlComments.fromXml(encoded)

    assert(reparsed.isRight)
    val comment2 = reparsed.toOption.get.comments.head
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

    val comment = result.toOption.get.comments.head
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

  test("encode simple comment produces valid XML") {
    val comment = OoxmlComment(
      ref = ref"H8",
      authorId = 0,
      text = RichText.plain("Test comment"),
      guid = Some("{ABC-123}")
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

    // Verify attributes
    val commentElem = (xml \ "commentList" \ "comment").head.asInstanceOf[Elem]
    assertEquals((commentElem \ "@ref").text, "H8")
    assertEquals((commentElem \ "@authorId").text, "0")
    assertEquals((commentElem \ "@guid").text, "{ABC-123}")
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
    val textElem = (xml \ "commentList" \ "comment" \ "text").head.asInstanceOf[Elem]
    val runs = textElem \ "r"

    assert(runs.size >= 2, "Should have at least 2 runs (formatted + plain)")

    // First run should have formatting
    val run1 = runs(0).asInstanceOf[Elem]
    assert((run1 \ "rPr").nonEmpty, "First run should have formatting properties")
    assert((run1 \ "rPr" \ "b").nonEmpty, "First run should be bold")
    assert((run1 \ "rPr" \ "color").nonEmpty, "First run should have color")
  }

  test("empty authors list is preserved") {
    val comments = OoxmlComments(
      authors = Vector.empty,
      comments = Vector.empty
    )

    val xml = OoxmlComments.toXml(comments)
    val reparsed = OoxmlComments.fromXml(xml)

    assert(reparsed.isRight)
    assertEquals(reparsed.toOption.get.authors, Vector.empty)
    assertEquals(reparsed.toOption.get.comments, Vector.empty)
  }
