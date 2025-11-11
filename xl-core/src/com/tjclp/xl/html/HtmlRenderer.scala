package com.tjclp.xl.html

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cell.{Cell, CellValue}
import com.tjclp.xl.style.{Font, Color, BorderStyle, HAlign, VAlign, Fill}

/** Renders Excel sheets to HTML tables with inline CSS styling */
object HtmlRenderer:

  /**
   * Export a sheet range to an HTML table.
   *
   * Generates a `<table>` element with cells converted to `<td>` elements. If includeStyles is
   * true, cell styles are converted to inline CSS. Rich text cells are rendered with HTML
   * formatting tags (<b>, <i>, <u>, <span>).
   *
   * @param sheet
   *   The sheet to export
   * @param range
   *   The cell range to export
   * @param includeStyles
   *   Whether to include inline CSS for cell styles (default: true)
   * @return
   *   HTML table string
   */
  def toHtml(
    sheet: Sheet,
    range: CellRange,
    includeStyles: Boolean = true
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
                "<td></td>" // Empty cell
              case Some(cell) =>
                val style = if includeStyles then cellStyleToInlineCss(cell, sheet) else ""
                val content = cellValueToHtml(cell.value)
                val styleAttr = if style.nonEmpty then s""" style="$style"""" else ""
                s"<td$styleAttr>$content</td>"
          }
          .mkString
        s"  <tr>$cellsHtml</tr>"
      }
      .mkString("\n")

    s"""<table>
$tableRows
</table>"""

  /**
   * Convert a CellValue to HTML content.
   *
   *   - Text: Escaped HTML
   *   - RichText: HTML with <b>, <i>, <u>, <span> tags
   *   - Number/DateTime/Bool: String representation
   *   - Formula/Error: Escaped string
   */
  private def cellValueToHtml(value: CellValue): String = value match
    case CellValue.Text(s) => escapeHtml(s)

    case CellValue.RichText(richText) =>
      richText.runs.map(runToHtml).mkString

    case CellValue.Number(n) => escapeHtml(n.toString)

    case CellValue.Bool(b) => escapeHtml(b.toString)

    case CellValue.DateTime(dt) => escapeHtml(dt.toString)

    case CellValue.Empty => ""

    case CellValue.Formula(expr) => escapeHtml(s"=$expr")

    case CellValue.Error(err) => escapeHtml(err.toExcel)

  /**
   * Convert a TextRun to HTML with formatting.
   *
   * Applies <b>, <i>, <u> tags for font styles and <span style="color:"> for colors.
   */
  private def runToHtml(run: TextRun): String =
    val text = escapeHtml(run.text)
    run.font match
      case None => text
      case Some(f) =>
        var html = text

        // Apply color as innermost wrapper
        f.color.foreach { c =>
          html = s"""<span style="color: ${c.toHex}">$html</span>"""
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
   * Generates CSS properties for font, fill, alignment, etc. Returns empty string if cell has no
   * style.
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
        style.font.color.foreach(c => css += s"color: ${c.toHex}")
        if style.font.sizePt != Font.default.sizePt then css += s"font-size: ${style.font.sizePt}pt"
        if style.font.name != Font.default.name then
          css += s"font-family: '${escapeCss(style.font.name)}'"

        // Fill (background color)
        style.fill match
          case Fill.Solid(color) =>
            css += s"background-color: ${color.toHex}"
          case _ => () // Pattern fill not supported in HTML

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

  private def escapeCss(s: String): String =
    s.replace("\\", "\\\\")
      .replace("'", "\\'")
      .replace("\"", "\\\"")
