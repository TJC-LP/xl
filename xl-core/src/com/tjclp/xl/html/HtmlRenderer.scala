package com.tjclp.xl.html

import java.awt.{Font as AwtFont, Graphics2D}
import java.awt.image.BufferedImage

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.richtext.TextRun
import com.tjclp.xl.styles.alignment.{HAlign, VAlign}
import com.tjclp.xl.styles.border.{BorderStyle, BorderSide}
import com.tjclp.xl.styles.color.{Color, ThemePalette}
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

/** Renders Excel sheets to HTML tables with inline CSS styling */
object HtmlRenderer:
  // Default dimensions
  private val DefaultCellHeightPx = 20 // Excel default ~15pt = 20px
  private val DefaultColumnWidthPx = 64 // Excel default ~8.43 chars = 64px

  // Unit conversion helpers (Excel → pixels)
  /** Convert Excel column width (character units) to pixels. */
  private def excelColWidthToPixels(width: Double): Int =
    (width * 7 + 5).toInt

  /** Convert Excel row height (points) to pixels. 1pt = 4/3 pixels. */
  private def excelRowHeightToPixels(height: Double): Int =
    (height * 4.0 / 3.0).toInt

  // Default font size for measurements
  private val DefaultFontSize = 11

  // Graphics context for text measurement (lazy, with headless fallback)
  private lazy val graphics: Option[Graphics2D] =
    try
      val img = new BufferedImage(1, 1, BufferedImage.TYPE_INT_ARGB)
      Some(img.createGraphics())
    catch case _: Exception => None // Headless environment

  /** Measure text width using AWT FontMetrics, with fallback estimation. */
  private def measureTextWidth(text: String, font: Option[Font]): Int =
    graphics match
      case Some(g) =>
        val awtFont = toAwtFont(font)
        g.setFont(awtFont)
        g.getFontMetrics.stringWidth(text)
      case None =>
        // Fallback: simple estimation for headless environments
        val baseCharWidth = 7
        val sizeFactor = font.map(f => f.sizePt / DefaultFontSize).getOrElse(1.0)
        val boldFactor = if font.exists(_.bold) then 1.1 else 1.0
        (text.length * baseCharWidth * sizeFactor * boldFactor).toInt

  /** Convert our Font to AWT Font. */
  private def toAwtFont(font: Option[Font]): AwtFont =
    font match
      case Some(f) =>
        val style =
          (if f.bold then AwtFont.BOLD else 0) | (if f.italic then AwtFont.ITALIC else 0)
        new AwtFont(f.name, style, f.sizePt.toInt)
      case None =>
        new AwtFont("Calibri", AwtFont.PLAIN, DefaultFontSize)

  /** Measure text width for a CellValue, handling rich text runs. */
  private def measureCellValueWidth(value: CellValue, font: Option[Font]): Int = value match
    case CellValue.RichText(rt) =>
      rt.runs.map(run => measureTextWidth(run.text, run.font.orElse(font))).sum
    case CellValue.Formula(_, Some(cached)) =>
      measureCellValueWidth(cached, font)
    case CellValue.Empty => 0
    case CellValue.Text(s) => measureTextWidth(s, font)
    case CellValue.Number(n) => measureTextWidth(n.toString, font)
    case CellValue.Bool(b) => measureTextWidth(if b then "TRUE" else "FALSE", font)
    case CellValue.DateTime(dt) => measureTextWidth(dt.toString, font)
    case CellValue.Error(e) => measureTextWidth(e.toString, font)
    case _ => 0

  /**
   * Calculate overflow colspan for a cell with text that exceeds its width.
   *
   * For left-aligned/general cells, counts empty cells to the right until:
   *   - A non-empty cell is reached
   *   - The accumulated width covers the text overflow
   *   - The range boundary is reached
   *
   * @param cell
   *   The cell to check for overflow
   * @param cellRef
   *   The cell reference
   * @param cellWidth
   *   The width of the cell in pixels
   * @param colWidths
   *   Vector of column widths (indexed from startCol)
   * @param sheet
   *   The sheet containing the data
   * @param startCol
   *   The starting column index of the range
   * @param endCol
   *   The ending column index of the range
   * @return
   *   The colspan (1 if no overflow, >1 if overflowing into adjacent cells)
   */
  private def calculateOverflowColspan(
    cell: Cell,
    cellRef: ARef,
    cellWidth: Int,
    colWidths: IndexedSeq[Int],
    sheet: Sheet,
    startCol: Int,
    endCol: Int
  ): Int =
    val style = cell.styleId.flatMap(sheet.styleRegistry.get)

    // If wrapText is true, text wraps instead of overflowing
    if style.exists(_.align.wrapText) then return 1

    val font = style.map(_.font)
    val textWidth = measureCellValueWidth(cell.value, font)

    // If text fits within cell, no overflow needed
    if textWidth <= cellWidth then return 1

    // Determine overflow direction based on alignment
    // Default to Left alignment (Excel's "General" behaves like Left for text)
    val align = style.map(_.align.horizontal).getOrElse(HAlign.Left)
    align match
      case HAlign.Left =>
        // Overflow to the right
        countEmptyToRight(cellRef, cellWidth, colWidths, sheet, startCol, endCol, textWidth)
      case HAlign.Right =>
        // For right-aligned, we'd need to overflow left - more complex
        // For now, just use colspan=1 (text clips)
        1
      case HAlign.Center | HAlign.CenterContinuous =>
        // Center alignment would overflow both ways - complex
        // For now, just use colspan=1 (text clips)
        1
      case _ =>
        1

  /**
   * Count how many adjacent empty cells to the right can accommodate text overflow.
   *
   * @return
   *   colspan (1 + number of empty cells needed to fit overflow)
   */
  private def countEmptyToRight(
    cellRef: ARef,
    cellWidth: Int,
    colWidths: IndexedSeq[Int],
    sheet: Sheet,
    startCol: Int,
    endCol: Int,
    textWidth: Int
  ): Int =
    val colIdx = cellRef.col.index0
    var colspan = 1
    var accumulatedWidth = cellWidth

    // Check cells to the right
    var nextCol = colIdx + 1
    while nextCol <= endCol && accumulatedWidth < textWidth do
      val nextRef = ARef.from0(nextCol, cellRef.row.index0)

      // Check if next cell is empty (no content, not part of merge)
      val nextCellOpt = sheet.cells.get(nextRef)
      val nextHasContent = nextCellOpt.exists { c =>
        c.value match
          case CellValue.Empty => false
          case _ => true
      }

      // Check if next cell is part of a merged region
      val nextIsMerged = sheet.getMergedRange(nextRef).isDefined

      if nextHasContent || nextIsMerged then
        // Stop - can't overflow into non-empty or merged cell
        return colspan
      else
        // Include this empty cell in the overflow span
        val widthIdx = nextCol - startCol
        if widthIdx >= 0 && widthIdx < colWidths.length then accumulatedWidth += colWidths(widthIdx)
        colspan += 1
        nextCol += 1

    colspan

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
    includeComments: Boolean = true,
    theme: ThemePalette = ThemePalette.office
  ): String =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0

    // Calculate column widths for <colgroup>
    val colWidths = (startCol to endCol).map { colIdx =>
      val col = Column.from0(colIdx)
      val props = sheet.getColumnProperties(col)
      if props.hidden then 0
      else
        props.width
          .map(excelColWidthToPixels)
          .getOrElse(
            sheet.defaultColumnWidth
              .map(excelColWidthToPixels)
              .getOrElse(DefaultColumnWidthPx)
          )
    }

    // Generate <colgroup> element
    val colgroup = colWidths
      .map(w => s"""  <col style="width: ${w}px">""")
      .mkString("<colgroup>\n", "\n", "\n</colgroup>")

    // Group cells by row for table structure
    val cellsByRow = range.cells.toSeq
      .map(ref => (ref, sheet.cells.get(ref)))
      .groupBy(_._1.row)
      .toSeq
      .sortBy(_._1.index0)

    val tableRows = cellsByRow
      .map { (rowObj, rowCells) =>
        // Calculate row height
        val props = sheet.getRowProperties(rowObj)
        val rowHeight =
          if props.hidden then 0
          else
            props.height
              .map(excelRowHeightToPixels)
              .getOrElse(
                sheet.defaultRowHeight
                  .map(excelRowHeightToPixels)
                  .getOrElse(DefaultCellHeightPx)
              )

        // Track columns covered by text overflow from previous cells in this row
        val overflowSkipCols = scala.collection.mutable.Set[Int]()

        val cellsHtml = rowCells
          .sortBy(_._1.col.index0)
          .flatMap { (ref, cellOpt) =>
            val colIdx = ref.col.index0

            // Skip if this cell is covered by a previous cell's text overflow
            if overflowSkipCols.contains(colIdx) then None
            else
              // Check if this cell is an interior cell of a merged region (skip it)
              val mergeRange = sheet.getMergedRange(ref)
              val isInteriorMergeCell = mergeRange.exists(_.start != ref)
              if isInteriorMergeCell then None
              else
                // Check for colspan/rowspan if this is a merge anchor
                val (mergeColspan, mergeRowspan) = mergeRange match
                  case Some(range) =>
                    // Clamp merge region to visible range
                    val visibleColspan =
                      math.min(range.end.col.index0, endCol) - ref.col.index0 + 1
                    val visibleRowspan =
                      math.min(range.end.row.index0, endRow) - ref.row.index0 + 1
                    (visibleColspan, visibleRowspan)
                  case None => (1, 1)

                cellOpt match
                  case None =>
                    // Empty cell - render normally (no overflow from empty cells)
                    val mergeAttrs =
                      (if mergeColspan > 1 then s""" colspan="$mergeColspan"""" else "") +
                        (if mergeRowspan > 1 then s""" rowspan="$mergeRowspan"""" else "")
                    if includeStyles then
                      Some(
                        s"""<td$mergeAttrs style="background-color: #FFFFFF; white-space: nowrap; overflow: hidden"></td>"""
                      )
                    else Some(s"<td$mergeAttrs></td>")

                  case Some(cell) =>
                    // Calculate overflow colspan (only if no merge colspan)
                    val overflowColspan =
                      if mergeColspan > 1 then 1 // Merged cells take priority
                      else
                        val widthIdx = colIdx - startCol
                        val cellWidth =
                          if widthIdx >= 0 && widthIdx < colWidths.length then colWidths(widthIdx)
                          else DefaultColumnWidthPx
                        calculateOverflowColspan(
                          cell,
                          ref,
                          cellWidth,
                          colWidths,
                          sheet,
                          startCol,
                          endCol
                        )

                    // Mark subsequent columns to skip due to overflow
                    if overflowColspan > 1 then
                      (1 until overflowColspan).foreach(i => overflowSkipCols += (colIdx + i))

                    // Use overflow colspan or merge colspan (merge takes priority)
                    val effectiveColspan =
                      if mergeColspan > 1 then mergeColspan else overflowColspan
                    val mergeAttrs =
                      (if effectiveColspan > 1 then s""" colspan="$effectiveColspan"""" else "") +
                        (if mergeRowspan > 1 then s""" rowspan="$mergeRowspan"""" else "")

                    val style =
                      if includeStyles then cellStyleToInlineCss(cell, sheet, theme) else ""
                    // Extract NumFmt from cell's style for proper value formatting
                    val numFmt = cell.styleId
                      .flatMap(sheet.styleRegistry.get)
                      .map(_.numFmt)
                      .getOrElse(NumFmt.General)
                    val content = cellValueToHtml(cell.value, numFmt, theme)
                    // Add default white background if no fill is specified (only when includeStyles)
                    // Add overflow: hidden only when not spanning (colspan=1)
                    // And white-space: nowrap if not explicitly wrapping (Excel default)
                    val styleAttr =
                      if !includeStyles then ""
                      else
                        val hasBackground = style.contains("background-color")
                        val hasWhitespace = style.contains("white-space")
                        val baseStyle =
                          if hasBackground then style
                          else if style.isEmpty then "background-color: #FFFFFF"
                          else s"background-color: #FFFFFF; $style"
                        // Add nowrap default if cell has no explicit white-space setting
                        val withWhitespace =
                          if hasWhitespace then baseStyle
                          else s"$baseStyle; white-space: nowrap"
                        // Only add overflow: hidden when not spanning (text fits or clips)
                        val finalStyle =
                          if effectiveColspan > 1 then withWhitespace
                          else s"$withWhitespace; overflow: hidden"
                        s""" style="$finalStyle""""

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

                    Some(s"<td$mergeAttrs$styleAttr$commentAttr>$content</td>")
          }
          .mkString

        // Add height style to <tr>
        s"""  <tr style="height: ${rowHeight}px">$cellsHtml</tr>"""
      }
      .mkString("\n")

    // table-layout: fixed ensures column widths from colgroup are respected
    val tableStyle =
      """style="border-collapse: collapse; table-layout: fixed; font-family: Calibri, sans-serif; font-size: 11pt;""""
    s"""<table $tableStyle>
$colgroup
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
  private def cellValueToHtml(value: CellValue, numFmt: NumFmt, theme: ThemePalette): String =
    value match
      case CellValue.RichText(richText) =>
        // Rich text has its own formatting, don't apply NumFmt
        richText.runs.map(run => runToHtml(run, theme)).mkString

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
  private def runToHtml(run: TextRun, theme: ThemePalette): String =
    val text = escapeHtml(run.text)
    run.font match
      case None => text
      case Some(f) =>
        var html = text

        // Apply color as innermost wrapper
        f.color.foreach { c =>
          html = s"""<span style="color: ${colorToRgbHex(c, theme)}">$html</span>"""
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
  private def cellStyleToInlineCss(cell: Cell, sheet: Sheet, theme: ThemePalette): String =
    cell.styleId
      .flatMap(sheet.styleRegistry.get)
      .map { style =>
        val css = scala.collection.mutable.ArrayBuffer[String]()

        // Font properties (apply only if not default)
        if style.font.bold then css += "font-weight: bold"
        if style.font.italic then css += "font-style: italic"
        if style.font.underline then css += "text-decoration: underline"
        style.font.color.foreach(c => css += s"color: ${colorToRgbHex(c, theme)}")
        if style.font.sizePt != Font.default.sizePt then css += s"font-size: ${style.font.sizePt}pt"
        if style.font.name != Font.default.name then
          css += s"font-family: '${escapeCss(style.font.name)}'"

        // Fill (background color)
        style.fill match
          case Fill.Solid(color) =>
            css += s"background-color: ${colorToRgbHex(color, theme)}"
          case _ => () // Pattern fill not supported in HTML

        // Borders
        borderSideToCss(style.border.top, "border-top", theme).foreach(css += _)
        borderSideToCss(style.border.right, "border-right", theme).foreach(css += _)
        borderSideToCss(style.border.bottom, "border-bottom", theme).foreach(css += _)
        borderSideToCss(style.border.left, "border-left", theme).foreach(css += _)

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

        // Excel default is no-wrap; explicit wrapText enables wrapping
        if style.align.wrapText then css += "white-space: pre-wrap"
        else css += "white-space: nowrap"

        // Indentation (Excel uses ~3 characters per indent level, ~21px at 11pt)
        if style.align.indent > 0 then
          val indentPx = style.align.indent * 21
          css += s"padding-left: ${indentPx}px"

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
      .replace("\n", "\\A ")
      .replace("\r", "\\D ")
      .replace("\u0000", "")

  /**
   * Convert Color to CSS-compatible RGB hex using theme for resolution. Returns #RRGGBB format.
   */
  private def colorToRgbHex(c: Color, theme: ThemePalette): String =
    c.toResolvedHex(theme)

  /**
   * Convert a border side to CSS border property.
   */
  private def borderSideToCss(
    side: BorderSide,
    cssProperty: String,
    theme: ThemePalette
  ): Option[String] =
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
      val color = side.color.map(c => colorToRgbHex(c, theme)).getOrElse("#000000")
      Some(s"$cssProperty: $width $cssStyle $color")
