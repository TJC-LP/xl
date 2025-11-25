package com.tjclp.xl.html

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.styles.alignment.{HAlign, VAlign}
import com.tjclp.xl.styles.border.{BorderStyle, BorderSide}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

/** Renders Excel sheets to HTML tables with inline CSS styling */
object HtmlRenderer:

  /**
   * Export a sheet range to an HTML table.
   *
   * Generates a `<table>` element with cells converted to `<td>` elements. If includeStyles is
   * true, cell styles are converted to inline CSS. Rich text cells are rendered with HTML
   * formatting tags (<b>, <i>, <u>, <span>). Comments are rendered as HTML title attributes
   * (tooltips on hover).
   *
   * @param sheet
   *   The sheet to export
   * @param range
   *   The cell range to export
   * @param includeStyles
   *   Whether to include inline CSS for cell styles (default: true)
   * @param includeComments
   *   Whether to include comments as HTML title attributes (default: true)
   * @return
   *   HTML table string
   */
  def toHtml(
    sheet: Sheet,
    range: CellRange,
    includeStyles: Boolean = true,
    includeComments: Boolean = true
  ): String =
    // Group cells by row for table structure
    val cellsByRow = range.cells.toSeq
      .map(ref => (ref, sheet.cells.get(ref)))
      .groupBy(_._1.row)
      .toSeq
      .sortBy(_._1.index0)

    val tableRows = cellsByRow
      .map { (rowNum, rowCells) =>
        val cellsHtml = rowCells
          .sortBy(_._1.col.index0)
          .map { (ref, cellOpt) =>
            cellOpt match
              case None =>
                if includeStyles then """<td style="background-color: #FFFFFF"></td>"""
                else "<td></td>"
              case Some(cell) =>
                val style = if includeStyles then cellStyleToInlineCss(cell, sheet) else ""
                // Extract NumFmt from cell's style for proper value formatting
                val numFmt = cell.styleId
                  .flatMap(sheet.styleRegistry.get)
                  .map(_.numFmt)
                  .getOrElse(NumFmt.General)
                val content = cellValueToHtml(cell.value, numFmt)
                // Add default white background if no fill is specified (only when includeStyles)
                val styleAttr =
                  if !includeStyles then ""
                  else
                    val hasBackground = style.contains("background-color")
                    val fullStyle =
                      if hasBackground then style
                      else if style.isEmpty then "background-color: #FFFFFF"
                      else s"background-color: #FFFFFF; $style"
                    s""" style="$fullStyle""""

                // Add comment as title attribute (tooltip) if present
                val commentAttr =
                  if includeComments then
                    sheet
                      .getComment(ref)
                      .map { comment =>
                        val commentText = comment.text.toPlainText
                        val authorPrefix = comment.author.map(a => s"$a: ").getOrElse("")
                        s""" title="${escapeHtml(authorPrefix + commentText)}""""
                      }
                      .getOrElse("")
                  else ""

                s"<td$styleAttr$commentAttr>$content</td>"
          }
          .mkString
        s"  <tr>$cellsHtml</tr>"
      }
      .mkString("\n")

    val tableStyle =
      """style="border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 11pt;""""
    s"""<table $tableStyle>
$tableRows
</table>"""

  /**
   * Convert a CellValue to HTML content with Excel-style number formatting.
   *
   *   - Text: Escaped HTML
   *   - RichText: HTML with <b>, <i>, <u>, <span> tags
   *   - Number/DateTime/Bool: Formatted according to NumFmt, then escaped
   *   - Formula: Shows cached value formatted, or raw formula if no cache
   *   - Error: Excel error code
   */
  private def cellValueToHtml(value: CellValue, numFmt: NumFmt): String = value match
    case CellValue.RichText(richText) =>
      // Rich text has its own formatting, don't apply NumFmt
      richText.runs.map(runToHtml).mkString

    case CellValue.Empty => ""

    case CellValue.Formula(_, Some(cached)) =>
      // Show cached result formatted with NumFmt (matches Excel display)
      escapeHtml(NumFmtFormatter.formatValue(cached, numFmt))

    case CellValue.Formula(expr, None) =>
      // No cached value, show raw formula
      escapeHtml(s"=$expr")

    case other =>
      // Use NumFmtFormatter for Text, Number, Bool, DateTime, Error
      escapeHtml(NumFmtFormatter.formatValue(other, numFmt))

  /**
   * Convert a TextRun to HTML with formatting.
   *
   * Applies <b>, <i>, <u> tags for font styles and <span style="color:"> for colors.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def runToHtml(run: TextRun): String =
    val text = escapeHtml(run.text)
    run.font match
      case None => text
      case Some(f) =>
        var html = text

        // Apply color as innermost wrapper
        f.color.foreach { c =>
          html = s"""<span style="color: ${colorToRgbHex(c)}">$html</span>"""
        }

        // Font size (if different from default)
        if f.sizePt != Font.default.sizePt then
          html = s"""<span style="font-size: ${f.sizePt}pt">$html</span>"""

        // Font family (if different from default)
        if f.name != Font.default.name then
          html = s"""<span style="font-family: '${escapeCss(f.name)}'">$html</span>"""

        // Apply bold/italic/underline as outermost wrappers
        if f.underline then html = s"<u>$html</u>"
        if f.italic then html = s"<i>$html</i>"
        if f.bold then html = s"<b>$html</b>"

        html

  /**
   * Convert cell-level style to inline CSS.
   *
   * Generates CSS properties for font, fill, borders, alignment, etc. Returns empty string if cell
   * has no style.
   */
  private def cellStyleToInlineCss(cell: Cell, sheet: Sheet): String =
    cell.styleId
      .flatMap(sheet.styleRegistry.get)
      .map { style =>
        val css = scala.collection.mutable.ArrayBuffer[String]()

        // Font properties (apply only if not default)
        if style.font.bold then css += "font-weight: bold"
        if style.font.italic then css += "font-style: italic"
        if style.font.underline then css += "text-decoration: underline"
        style.font.color.foreach(c => css += s"color: ${colorToRgbHex(c)}")
        if style.font.sizePt != Font.default.sizePt then css += s"font-size: ${style.font.sizePt}pt"
        if style.font.name != Font.default.name then
          css += s"font-family: '${escapeCss(style.font.name)}'"

        // Fill (background color)
        style.fill match
          case Fill.Solid(color) =>
            css += s"background-color: ${colorToRgbHex(color)}"
          case _ => () // Pattern fill not supported in HTML

        // Borders
        borderSideToCss(style.border.top, "border-top").foreach(css += _)
        borderSideToCss(style.border.right, "border-right").foreach(css += _)
        borderSideToCss(style.border.bottom, "border-bottom").foreach(css += _)
        borderSideToCss(style.border.left, "border-left").foreach(css += _)

        // Alignment
        style.align.horizontal match
          case HAlign.Left => css += "text-align: left"
          case HAlign.Center => css += "text-align: center"
          case HAlign.Right => css += "text-align: right"
          case _ => ()

        style.align.vertical match
          case VAlign.Top => css += "vertical-align: top"
          case VAlign.Middle => css += "vertical-align: middle"
          case VAlign.Bottom => css += "vertical-align: bottom"
          case VAlign.Justify | VAlign.Distributed => () // No direct CSS equivalent

        if style.align.wrapText then css += "white-space: pre-wrap"

        css.mkString("; ")
      }
      .getOrElse("")

  /**
   * Escape HTML special characters.
   *
   * Converts <, >, &, " to HTML entities to prevent XSS and rendering issues.
   */
  private def escapeHtml(s: String): String =
    s.replace("&", "&amp;")
      .replace("<", "&lt;")
      .replace(">", "&gt;")
      .replace("\"", "&quot;")
      .replace("'", "&#39;")

  private def escapeCss(s: String): String =
    s.replace("\\", "\\\\")
      .replace("'", "\\'")
      .replace("\"", "\\\"")

  /**
   * Convert Color to CSS-compatible RGB hex (strips alpha from ARGB).
   *
   * Color.toHex returns #AARRGGBB, CSS needs #RRGGBB.
   */
  private def colorToRgbHex(c: Color): String =
    val hex = c.toHex
    if hex.length == 9 && hex.startsWith("#") then "#" + hex.drop(3) // #AARRGGBB -> #RRGGBB
    else hex

  /**
   * Convert a border side to CSS border property.
   */
  private def borderSideToCss(side: BorderSide, cssProperty: String): Option[String] =
    if side.style == BorderStyle.None then None
    else
      val width = side.style match
        case BorderStyle.Thin => "1px"
        case BorderStyle.Medium => "2px"
        case BorderStyle.Thick => "3px"
        case BorderStyle.Dashed => "1px"
        case BorderStyle.Dotted => "1px"
        case BorderStyle.Double => "3px"
        case BorderStyle.Hair => "1px"
        case BorderStyle.MediumDashed => "2px"
        case BorderStyle.DashDot => "1px"
        case BorderStyle.MediumDashDot => "2px"
        case BorderStyle.DashDotDot => "1px"
        case BorderStyle.SlantDashDot => "2px"
        case _ => "1px"
      val cssStyle = side.style match
        case BorderStyle.Dashed | BorderStyle.MediumDashed => "dashed"
        case BorderStyle.Dotted | BorderStyle.Hair => "dotted"
        case BorderStyle.Double => "double"
        case BorderStyle.DashDot | BorderStyle.MediumDashDot | BorderStyle.DashDotDot |
            BorderStyle.SlantDashDot =>
          "dashed" // CSS doesn't support dash-dot, use dashed as fallback
        case _ => "solid"
      val color = side.color.map(colorToRgbHex).getOrElse("#000000")
      Some(s"$cssProperty: $width $cssStyle $color")
