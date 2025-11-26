package com.tjclp.xl.render

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.richtext.TextRun
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.alignment.{HAlign, VAlign}
import com.tjclp.xl.styles.border.{BorderStyle, BorderSide}
import com.tjclp.xl.styles.color.{Color, ThemePalette}
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.styles.CellStyle

/** Renders Excel sheets to HTML tables with inline CSS styling */
object HtmlRenderer:
  import RenderUtils.*

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
   * @param theme
   *   Theme palette for resolving theme colors (default: Office theme)
   * @param applyPrintScale
   *   Whether to apply the sheet's print scale setting (default: false)
   * @param showLabels
   *   Whether to show column letters (A, B, C...) and row numbers (1, 2, 3...) (default: false)
   * @return
   *   HTML table string
   */
  def toHtml(
    sheet: Sheet,
    range: CellRange,
    includeStyles: Boolean = true,
    includeComments: Boolean = true,
    theme: ThemePalette = ThemePalette.office,
    applyPrintScale: Boolean = false,
    showLabels: Boolean = false
  ): String =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0

    // Calculate print scale factor (100 = 100% = 1.0)
    val scaleFactor =
      if applyPrintScale then sheet.pageSetup.map(_.scale / 100.0).getOrElse(1.0)
      else 1.0

    // Calculate column widths for <colgroup>
    val colWidths = (startCol to endCol).map { colIdx =>
      val col = Column.from0(colIdx)
      val props = sheet.getColumnProperties(col)
      if props.hidden then 0
      else
        val baseWidth = props.width
          .map(excelColWidthToPixels)
          .getOrElse(
            sheet.defaultColumnWidth
              .map(excelColWidthToPixels)
              .getOrElse(DefaultColumnWidthPx)
          )
        (baseWidth * scaleFactor).toInt
    }

    // Generate <colgroup> element
    val colgroupCols =
      if showLabels then
        // Include row number column first
        s"""  <col style="width: ${HeaderWidth}px">""" +:
          colWidths.map(w => s"""  <col style="width: ${w}px">""")
      else colWidths.map(w => s"""  <col style="width: ${w}px">""")
    val colgroup = colgroupCols.mkString("<colgroup>\n", "\n", "\n</colgroup>")

    val sb = new StringBuilder

    // Header row with column letters (if showLabels)
    val headerRow =
      if showLabels then
        val headerCells = (startCol to endCol).map { colIdx =>
          val colLetter = Column.from0(colIdx).toLetter
          s"""<td class="xl-header">$colLetter</td>"""
        }
        val cornerCell = """<td class="xl-header"></td>""" // Top-left corner
        val scaledHeaderHeight = (HeaderHeight * scaleFactor).toInt
        s"""  <tr style="height: ${scaledHeaderHeight}px">$cornerCell${headerCells.mkString}</tr>\n"""
      else ""

    // Group cells by row for table structure, filtering out hidden rows
    val cellsByRow = range.cells.toSeq
      .map(ref => (ref, sheet.cells.get(ref)))
      .groupBy(_._1.row)
      .toSeq
      .sortBy(_._1.index0)
      .filterNot { (rowObj, _) => sheet.getRowProperties(rowObj).hidden }

    val tableRows = cellsByRow
      .map { (rowObj, rowCells) =>
        // Calculate row height
        val props = sheet.getRowProperties(rowObj)
        val baseRowHeight =
          props.height
            .map(excelRowHeightToPixels)
            .getOrElse(
              sheet.defaultRowHeight
                .map(excelRowHeightToPixels)
                .getOrElse(DefaultCellHeightPx)
            )
        val rowHeight = (baseRowHeight * scaleFactor).toInt

        // Row number cell (if showLabels)
        val rowNumCell =
          if showLabels then s"""<td class="xl-header">${rowObj.index1}</td>"""
          else ""

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
                    // Calculate cell width (sum of column widths for colspan)
                    val cellWidthPx = (0 until mergeColspan).map { i =>
                      val widthIdx = colIdx - startCol + i
                      if widthIdx >= 0 && widthIdx < colWidths.length then colWidths(widthIdx)
                      else DefaultColumnWidthPx
                    }.sum
                    if includeStyles then
                      Some(
                        s"""<td$mergeAttrs style="width: ${cellWidthPx}px; background-color: #FFFFFF; white-space: nowrap; overflow: hidden"></td>"""
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
                    // Calculate cell width (sum of column widths for colspan)
                    val cellWidthPx = (0 until effectiveColspan).map { i =>
                      val widthIdx = colIdx - startCol + i
                      if widthIdx >= 0 && widthIdx < colWidths.length then colWidths(widthIdx)
                      else DefaultColumnWidthPx
                    }.sum
                    // Add default white background if no fill is specified (only when includeStyles)
                    // Add overflow: hidden only when not spanning (colspan=1)
                    // And white-space: nowrap if not explicitly wrapping (Excel default)
                    val styleAttr =
                      if !includeStyles then ""
                      else
                        val hasBackground = style.contains("background-color")
                        val hasWhitespace = style.contains("white-space")
                        // Start with explicit width to match colgroup
                        val withWidth = s"width: ${cellWidthPx}px"
                        val baseStyle =
                          if hasBackground then s"$withWidth; $style"
                          else if style.isEmpty then s"$withWidth; background-color: #FFFFFF"
                          else s"$withWidth; background-color: #FFFFFF; $style"
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
        s"""  <tr style="height: ${rowHeight}px">$rowNumCell$cellsHtml</tr>"""
      }
      .mkString("\n")

    // table-layout: fixed ensures column widths from colgroup are respected
    val tableStyle =
      """style="border-collapse: collapse; table-layout: fixed; font-family: Calibri, sans-serif; font-size: 11pt;""""

    // Header style (embedded in <style> tag when showLabels)
    val headerStyles =
      if showLabels then """
<style>
  .xl-header {
    background-color: #E0E0E0;
    border: 1px solid #999999;
    text-align: center;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 11px;
    color: #333333;
  }
</style>"""
      else ""

    s"""$headerStyles<table $tableStyle>
$colgroup
$headerRow$tableRows
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
          html = s"""<span style="color: ${colorToHex(c, theme)}">$html</span>"""
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
  /** Determine default horizontal alignment based on cell value type (Excel's General behavior) */
  private def contentBasedAlignment(value: CellValue): HAlign = value match
    case CellValue.Number(_) | CellValue.DateTime(_) => HAlign.Right
    case CellValue.Bool(_) => HAlign.Center
    case CellValue.Formula(_, Some(cached)) => contentBasedAlignment(cached)
    case _ => HAlign.Left

  private def cellStyleToInlineCss(cell: Cell, sheet: Sheet, theme: ThemePalette): String =
    val styleOpt = cell.styleId.flatMap(sheet.styleRegistry.get)
    val css = scala.collection.mutable.ArrayBuffer[String]()

    styleOpt.foreach { style =>
      // Font properties (apply only if not default)
      if style.font.bold then css += "font-weight: bold"
      if style.font.italic then css += "font-style: italic"
      if style.font.underline then css += "text-decoration: underline"
      style.font.color.foreach(c => css += s"color: ${colorToHex(c, theme)}")
      if style.font.sizePt != Font.default.sizePt then css += s"font-size: ${style.font.sizePt}pt"
      if style.font.name != Font.default.name then
        css += s"font-family: '${escapeCss(style.font.name)}'"

      // Fill (background color)
      style.fill match
        case Fill.Solid(color) =>
          css += s"background-color: ${colorToHex(color, theme)}"
        case _ => () // Pattern fill not supported in HTML

      // Borders
      borderSideToCss(style.border.top, "border-top", theme).foreach(css += _)
      borderSideToCss(style.border.right, "border-right", theme).foreach(css += _)
      borderSideToCss(style.border.bottom, "border-bottom", theme).foreach(css += _)
      borderSideToCss(style.border.left, "border-left", theme).foreach(css += _)
    }

    // Alignment - always emit to ensure proper alignment
    // Use explicit alignment from style if set, otherwise use content-based default (General behavior)
    val effectiveHAlign = styleOpt.map(_.align.horizontal).getOrElse(HAlign.General) match
      case HAlign.General => contentBasedAlignment(cell.value)
      case explicit => explicit

    effectiveHAlign match
      case HAlign.Left => css += "text-align: left"
      case HAlign.Center => css += "text-align: center"
      case HAlign.Right => css += "text-align: right"
      case _ => css += "text-align: left" // Fallback for Justify, Fill, etc.

    val effectiveVAlign = styleOpt.map(_.align.vertical).getOrElse(VAlign.Bottom)
    effectiveVAlign match
      case VAlign.Top => css += "vertical-align: top"
      case VAlign.Middle => css += "vertical-align: middle"
      case VAlign.Bottom => css += "vertical-align: bottom"
      case VAlign.Justify | VAlign.Distributed => () // No direct CSS equivalent

    // Excel default is no-wrap; explicit wrapText enables wrapping
    val wrapText = styleOpt.map(_.align.wrapText).getOrElse(false)
    if wrapText then css += "white-space: pre-wrap"
    else css += "white-space: nowrap"

    // Indentation (Excel uses ~3 characters per indent level, ~21px at 11pt)
    styleOpt.foreach { style =>
      if style.align.indent > 0 then
        val indentPx = style.align.indent * 21
        css += s"padding-left: ${indentPx}px"
    }

    css.mkString("; ")

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
      val color = side.color.map(c => colorToHex(c, theme)).getOrElse("#000000")
      Some(s"$cssProperty: $width $cssStyle $color")
