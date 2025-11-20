package com.tjclp.xl

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.richtext.RichText

/**
 * Excel comment attached to a cell.
 *
 * Comments are annotations that appear as yellow notes in Excel when a cell is hovered. They
 * support rich text formatting and author attribution.
 *
 * '''Example: Creating comments'''{{{ import com.tjclp.xl.* import com.tjclp.xl.richtext.RichText.*
 *
 * // Plain text comment val simple = Comment( text = RichText.plainText("Review this value"),
 * author = Some("John Doe") )
 *
 * // Rich text comment with formatting val formatted = Comment( text = "Note: ".bold.red +
 * "Critical issue!", author = Some("Jane Smith") ) }}}
 *
 * '''Adding comments to cells'''{{{ val sheet = Sheet("Data") .put(ref"A1", "Revenue")
 * .withComment(ref"A1", Comment.plainText("Q1 2025 data")) }}}
 *
 * @param text
 *   Comment content (supports rich text formatting)
 * @param author
 *   Optional author name (displayed in Excel comment tooltip)
 *
 * @since 0.4.0
 */
final case class Comment(
  text: RichText,
  author: Option[String]
) derives CanEqual

object Comment:
  /**
   * Create plain text comment (no formatting).
   *
   * @param text
   *   Comment text
   * @param author
   *   Optional author name
   * @return
   *   Comment with plain text
   */
  def plainText(text: String, author: Option[String] = None): Comment =
    Comment(RichText.plain(text), author)

  /**
   * Create comment from rich text.
   *
   * @param text
   *   Rich text content
   * @param author
   *   Optional author name
   * @return
   *   Comment with formatted text
   */
  def apply(text: RichText, author: Option[String]): Comment =
    new Comment(text, author)
