package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.error.XLError

/**
 * OOXML comment in xl/commentsN.xml.
 *
 * Maps to Excel's comment model with author attribution and rich text content. Supports forwards
 * compatibility via otherAttrs/otherChildren for unknown properties.
 *
 * OOXML structure:
 * {{{
 * <comment ref="A1" authorId="0" [guid="..."]>
 *   <text>
 *     <r><t>Comment text</t></r>
 *   </text>
 * </comment>
 * }}}
 *
 * @param ref
 *   Cell reference (e.g., A1, B2)
 * @param authorId
 *   Index into authors list (0-based)
 * @param text
 *   Rich text content (may contain multiple formatted runs)
 * @param guid
 *   Optional GUID for comment tracking
 * @param otherAttrs
 *   Unknown attributes for forwards compatibility
 * @param otherChildren
 *   Unknown child elements for forwards compatibility
 */
final case class OoxmlComment(
  ref: ARef,
  authorId: Int,
  text: RichText,
  guid: Option[String] = None,
  otherAttrs: Map[String, String] = Map.empty,
  otherChildren: Seq[Elem] = Seq.empty
)

/**
 * OOXML comments part (xl/commentsN.xml).
 *
 * Contains author list and comment list for a single worksheet.
 *
 * OOXML structure:
 * {{{
 * <comments>
 *   <authors>
 *     <author>John Doe</author>
 *     <author>Jane Smith</author>
 *   </authors>
 *   <commentList>
 *     <comment>...</comment>
 *   </commentList>
 * </comments>
 * }}}
 *
 * @param authors
 *   Author names (indexed by authorId in comments)
 * @param comments
 *   Comments for this worksheet
 * @param otherChildren
 *   Unknown child elements for forwards compatibility
 */
final case class OoxmlComments(
  authors: Vector[String],
  comments: Vector[OoxmlComment],
  otherChildren: Seq[Elem] = Seq.empty
)

object OoxmlComments extends XmlReadable[OoxmlComments]:

  /**
   * Parse comments from XML.
   *
   * REQUIRES: elem is <comments> element from xl/commentsN.xml ENSURES:
   *   - Returns OoxmlComments with authors and comments
   *   - Preserves unknown attributes/children for forwards compatibility
   *   - Returns error if structure is invalid (missing ref, authorId, or text)
   * DETERMINISTIC: Yes (stable iteration order)
   *
   * @param elem
   *   The <comments> root element
   * @return
   *   Either[String, OoxmlComments] with error if parsing fails
   */
  def fromXml(elem: Elem): Either[String, OoxmlComments] =
    if elem.label != "comments" then Left(s"Expected <comments> but found <${elem.label}>")
    else
      // Parse authors
      val authors = (elem \ "authors" \ "author").map(_.text.trim).toVector

      // Parse comments
      val commentsE = (elem \ "commentList" \ "comment").collect { case c: Elem => c }

      val parsedComments = commentsE.toVector.map(decodeComment).partitionMap(identity) match
        case (errors, _) if errors.nonEmpty =>
          Left(s"Comment parse errors: ${errors.mkString("; ")}")
        case (_, comments) => Right(comments)

      parsedComments.map { cs =>
        val knownNodes = (elem \ "authors") ++ (elem \ "commentList")
        val other = elem.child.collect {
          case el: Elem if !knownNodes.contains(el) => el
        }.toVector

        OoxmlComments(
          authors = authors,
          comments = cs,
          otherChildren = other
        )
      }

  /**
   * Parse single comment from XML.
   *
   * REQUIRES: elem is <comment> element ENSURES:
   *   - Returns OoxmlComment with ref, authorId, and text
   *   - Preserves unknown attributes for forwards compatibility
   *   - Returns error if missing required attributes (ref, authorId) or <text> element
   *
   * @param elem
   *   The <comment> element
   * @return
   *   Either[String, OoxmlComment] with error if parsing fails
   */
  private def decodeComment(elem: Elem): Either[String, OoxmlComment] =
    for
      refStr <- getAttr(elem, "ref")
      ref <- ARef.parse(refStr).left.map(err => s"Invalid cell ref '$refStr': $err")
      authorIdStr <- getAttr(elem, "authorId")
      authorId <- authorIdStr.toIntOption.toRight(s"Invalid authorId: $authorIdStr")
      textElem <- getChild(elem, "text")
      text <- parseCommentText(textElem)
    yield
      val known = Set("ref", "authorId", "guid")
      val attrs = elem.attributes.asAttrMap.filterNot { case (k, _) => known.contains(k) }
      val others = elem.child.collect {
        case el: Elem if el.label != "text" => el
      }.toVector

      OoxmlComment(
        ref = ref,
        authorId = authorId,
        text = text,
        guid = getAttrOpt(elem, "guid"),
        otherAttrs = attrs,
        otherChildren = others
      )

  /**
   * Parse comment text from <text> element.
   *
   * Comment text uses same structure as shared strings: <r> elements for rich text runs.
   *
   * REQUIRES: textElem is <text> element from comment ENSURES:
   *   - Returns RichText with one or more TextRun elements
   *   - Handles both plain text (single run) and formatted text (multiple runs)
   *
   * @param textElem
   *   The <text> element
   * @return
   *   Either[String, RichText] with error if parsing fails
   */
  private def parseCommentText(textElem: Elem): Either[String, RichText] =
    val runElems = textElem.child.collect { case e: Elem if e.label == "r" => e }

    if runElems.isEmpty then
      // Plain text comment (no rich formatting) - create single run
      val text = textElem.text.trim
      if text.isEmpty then Left("Comment <text> element is empty")
      else Right(RichText.plain(text))
    else
      // Rich text comment - parse runs
      parseTextRuns(runElems)

  /**
   * Write comments to XML.
   *
   * REQUIRES: comments is valid OoxmlComments ENSURES:
   *   - Returns <comments> element with authors and commentList
   *   - Preserves unknown children for forwards compatibility
   *   - Deterministic attribute ordering for stable diffs
   * DETERMINISTIC: Yes (sorted attributes, stable iteration)
   *
   * @param comments
   *   The comments to serialize
   * @return
   *   XML element <comments>...</comments>
   */
  def toXml(comments: OoxmlComments): Elem =
    <comments xmlns={nsSpreadsheetML}>
      <authors>
        {comments.authors.map(a => <author>{a}</author>)}
      </authors>
      <commentList>
        {comments.comments.map(encodeComment)}
      </commentList>
      {comments.otherChildren}
    </comments>

  /**
   * Write single comment to XML.
   *
   * REQUIRES: c is valid OoxmlComment ENSURES:
   *   - Returns <comment> element with ref, authorId, and text
   *   - Preserves unknown attributes/children for forwards compatibility
   *   - Deterministic attribute ordering for stable diffs
   * DETERMINISTIC: Yes (sorted attributes)
   *
   * @param c
   *   The comment to serialize
   * @return
   *   XML element <comment>...</comment>
   */
  private def encodeComment(c: OoxmlComment): Elem =
    val baseAttrs = Map(
      "ref" -> c.ref.toA1,
      "authorId" -> c.authorId.toString
    ) ++ c.guid.map("guid" -> _).toList.toMap ++ c.otherAttrs

    // Sort attributes for deterministic output
    val sortedAttrs = baseAttrs.toSeq.sortBy(_._1).foldLeft(Null: MetaData) { case (acc, (k, v)) =>
      new UnprefixedAttribute(k, v, acc)
    }

    val textXml = encodeCommentText(c.text)

    Elem(
      null,
      "comment",
      sortedAttrs,
      TopScope,
      minimizeEmpty = true,
      textXml +: c.otherChildren*
    )

  /**
   * Write comment text to XML.
   *
   * REQUIRES: text is valid RichText ENSURES:
   *   - Returns <text> element with <r> run elements
   *   - Each run contains <rPr> (optional) and <t> elements
   *   - Preserves whitespace with xml:space="preserve" when needed
   * DETERMINISTIC: Yes (stable run iteration)
   *
   * @param text
   *   The rich text to serialize
   * @return
   *   XML element <text>...</text>
   */
  private def encodeCommentText(text: RichText): Elem =
    val runElems = text.runs.map { run =>
      // Build <rPr> if font formatting is present
      val rPrElem = run.font.map { f =>
        val props = Seq.newBuilder[Elem]

        if f.bold then props += elem("b")()
        if f.italic then props += elem("i")()
        if f.underline then props += elem("u")()

        f.color.foreach { c =>
          props += elem("color", "rgb" -> c.toHex.drop(1))()
        }

        props += elem("sz", "val" -> f.sizePt.toString)()
        props += elem("name", "val" -> f.name)()

        elem("rPr")(props.result()*)
      }

      // Build <t> with xml:space="preserve" if needed
      val textElem =
        if needsXmlSpacePreserve(run.text) then
          Elem(
            null,
            "t",
            PrefixedAttribute("xml", "space", "preserve", Null),
            TopScope,
            minimizeEmpty = true,
            Text(run.text)
          )
        else elem("t")(Text(run.text))

      elem("r")(rPrElem.toSeq :+ textElem*)
    }

    elem("text")(runElems*)
