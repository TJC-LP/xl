package com.tjclp.xl.render

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.cf.{CfBar, CfOverlay}
import com.tjclp.xl.richtext.TextRun
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.alignment.{HAlign, VAlign}
import com.tjclp.xl.styles.border.{BorderStyle, BorderSide}
import com.tjclp.xl.styles.color.{Color, ThemePalette}
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.{Font, Underline}
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
   *   HTML table string, without conditional formatting (see the overload taking a [[CfOverlay]])
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
    toHtml(
      sheet,
      range,
      includeStyles,
      includeComments,
      theme,
      applyPrintScale,
      showLabels,
      CfOverlay.empty
    )

  /**
   * [[toHtml]] with conditional formatting painted from a precomputed `overlay` (GH-497): each
   * painted cell's dxf is laid over its own style (fill, font, strike, borders, number format) and
   * its data bar drawn as a gradient background as wide as SvgRenderer's; a painted cell the sheet
   * does not hold is drawn as an empty one. The overlay comes from xl-evaluator
   * (`sheet.conditionalFormatOverlay(range)`); an empty overlay renders byte-identically to the
   * CF-blind [[toHtml]], and `includeStyles = false` ignores it.
   */
  def toHtml(
    sheet: Sheet,
    range: CellRange,
    includeStyles: Boolean,
    includeComments: Boolean,
    theme: ThemePalette,
    applyPrintScale: Boolean,
    showLabels: Boolean,
    overlay: CfOverlay
  ): String =
    toHtmlResolving(ResolvedCell(_, _))(
      sheet,
      range,
      includeStyles,
      includeComments,
      theme,
      applyPrintScale,
      showLabels,
      overlay
    )

  /**
   * [[toHtml]] with the per-cell resolver injected. Each rendered cell is resolved exactly once
   * (`RenderUtilsSpec` counts through this seam): every overflow, alignment, hash and text decision
   * reads the one [[ResolvedCell]], so a Custom format code is parsed once per cell.
   */
  private[render] def toHtmlResolving(resolve: (Cell, Sheet) => ResolvedCell)(
    sheet: Sheet,
    range: CellRange,
    includeStyles: Boolean,
    includeComments: Boolean,
    theme: ThemePalette,
    applyPrintScale: Boolean,
    showLabels: Boolean,
    overlay: CfOverlay
  ): String =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0
    // Conditional formatting is styling: an unstyled render paints none of it
    val paints = if includeStyles then overlay else CfOverlay.empty

    // Calculate print scale factor (100 = 100% = 1.0)
    val scaleFactor =
      if applyPrintScale then sheet.pageSetup.map(_.scale / 100.0).getOrElse(1.0)
      else 1.0

    // Calculate column widths for <colgroup>
    val colWidths = calculateColumnWidths(sheet, range, scaleFactor)
    val widthAt: Int => Int = col => colWidths.lift(col - startCol).getOrElse(DefaultColumnWidthPx)

    // Row heights, autofitted exactly as SvgRenderer's (row labels share each <tr>)
    val rowHeights = calculateRowHeights(sheet, range, scaleFactor)

    // Generate <colgroup> element, the row number column first (if showLabels)
    val colgroupWidths = if showLabels then HeaderWidth +: colWidths else colWidths
    val colgroup = colgroupWidths
      .map(w => s"""  <col style="width: ${w}px">""")
      .mkString("<colgroup>\n", "\n", "\n</colgroup>")

    val sb = new StringBuilder

    // Header row with column letters (if showLabels)
    val headerRow =
      if showLabels then
        val headerCells = (startCol to endCol).map { colIdx =>
          // A hidden column keeps its zero-width cell but not its letter, which the fixed layout
          // would draw over the next column's header
          val colLetter = if widthAt(colIdx) > 0 then Column.from0(colIdx).toLetter else ""
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
        val rowHeight = rowHeights(rowObj.index0 - startRow)

        // Row number cell (if showLabels)
        val rowNumCell =
          if showLabels then s"""<td class="xl-header">${rowObj.index1}</td>"""
          else ""

        val cellsHtml = rowCells
          .sortBy(_._1.col.index0)
          .flatMap { (ref, cellOpt) =>
            val colIdx = ref.col.index0

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
              val mergeAttrs =
                (if mergeColspan > 1 then s""" colspan="$mergeColspan"""" else "") +
                  (if mergeRowspan > 1 then s""" rowspan="$mergeRowspan"""" else "")
              // Cell width: its column, or the sum of its merge's columns
              val cellWidthPx = (0 until mergeColspan).map(i => widthAt(colIdx + i)).sum
              // A CF-painted cell the sheet does not hold is drawn as an empty one
              val paint = paints.at(ref)

              cellOpt.orElse(paint.map(_ => Cell.empty(ref))) match
                case None =>
                  if includeStyles then
                    Some(
                      s"""<td$mergeAttrs style="width: ${cellWidthPx}px; background-color: #FFFFFF; white-space: nowrap; overflow: hidden"></td>"""
                    )
                  else Some(s"<td$mergeAttrs></td>")

                case Some(cell) =>
                  // Style and rendered content, resolved ONCE: the span, alignment, hash and
                  // body below all read them, and resolving re-formats the value and
                  // re-parses a Custom code. Then laid under the cell's CF paint.
                  val resolved = painted(resolve(cell, sheet), paint)

                  // Text spilling into empty neighbours keeps ONE <td> per cell (a colspan
                  // would repaint its fill over theirs and drop their borders): it is drawn in
                  // a box laid over them. Merged cells never spill.
                  val span = overflowSpan(
                    resolved,
                    ref,
                    widthAt(colIdx),
                    colWidths,
                    sheet,
                    startCol,
                    endCol
                  )
                  val spills = includeStyles && span.spills

                  val style =
                    if includeStyles then cellStyleToInlineCss(resolved, strikes(paint), theme)
                    else ""
                  val cellStyle = resolved.style
                  // Excel's #### marker: a clipped numeral reads as a different, plausible
                  // number (GH-459). Must agree with SvgRenderer.
                  val text = hashOverflowText(resolved.content, cellStyle, cellWidthPx)
                    .getOrElse(cellValueToHtml(resolved, theme))
                  // A merge's single line is anchored in a box as wide as the merge, as
                  // SvgRenderer anchors it: a td's own line box overflows only at its end
                  val boxedInMerge = includeStyles && mergeRange.isDefined &&
                    !resolved.style.exists(_.align.wrapText)
                  // A hidden column's content is never shown: a zero-width cell only clips it,
                  // and a table without inline CSS would not even do that
                  val content =
                    if cellWidthPx <= 0 || hidesValue(paint) then ""
                    else if spills then
                      textBox("xl-overflow", text, resolved, span, colIdx, cellWidthPx, widthAt)
                    else if boxedInMerge then
                      val noSpan = OverflowSpan.none
                      textBox("xl-merge", text, resolved, noSpan, colIdx, cellWidthPx, widthAt)
                    else text
                  // Add default white background if no fill is specified (only when includeStyles)
                  // Add overflow: hidden unless spilling: a merge clips to itself, as in Excel
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
                      val finalStyle =
                        if spills then withWhitespace
                        else s"$withWhitespace; overflow: hidden"
                      val barCss = paint
                        .flatMap(_.bar)
                        .flatMap(dataBarCss(_, cellWidthPx, theme))
                        .fold("")(css => s"; $css")
                      s""" style="$finalStyle$barCss""""

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

    // table-layout: fixed keeps the colgroup's widths, but only on a table with a width (CSS 2.1
    // §17.5.2.1): an auto-width table is laid out automatically, and a spilled text box would widen
    // its own column instead of lying over the neighbours
    val tableStyle =
      s"""style="border-collapse: collapse; table-layout: fixed; font-family: Calibri, sans-serif; font-size: 11pt; width: ${colgroupWidths.sum}px""""

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
   * The box a single line of text is drawn in: SvgRenderer's clip box. For a spill (`xl-overflow`)
   * that is exactly the columns [[RenderUtils.overflowSpan]] grants it, laid over its neighbours
   * from inside the source `<td>`, which is left without `overflow: hidden`. CSS paints every table
   * cell's background and border before any cell's text, so the text lies above the neighbours'
   * fills, each still painted by its own `<td>`. For a merge (`xl-merge`, with no span) it is the
   * merge itself.
   *
   * The box is pulled left, out of the source cell's content box (its indent padding included),
   * over the leftward span and clips to the whole span. A flex row anchors the text: unlike a line
   * box, which always overflows towards its end, a `flex-end` or `center` item overflows towards
   * the start too — right-aligned text spills (or is clipped) left, centred text both ways. Padding
   * on the longer side keeps centred text centred on its own cell when one side is blocked.
   */
  private def textBox(
    cssClass: String,
    content: String,
    cell: ResolvedCell,
    span: OverflowSpan,
    col: Int,
    cellWidth: Int,
    widthAt: Int => Int
  ): String =
    val leftPx = (col - span.left until col).map(widthAt).sum
    val rightPx = (col + 1 to col + span.right).map(widthAt).sum
    val indentPx = cell.style.map(_.align.indent).getOrElse(0) * IndentPxPerLevel
    // Left of the box to the text's centre, minus the centre to the box's right: the cell's
    // centre is half an indent right of its middle, as in SvgRenderer
    val lean = leftPx - rightPx + indentPx
    val (justify, padding) = cell.align match
      case HAlign.Right => ("flex-end", "")
      case HAlign.Center | HAlign.CenterContinuous =>
        val pad =
          if lean > 0 then s" padding-left: ${lean}px;"
          else if lean < 0 then s" padding-right: ${-lean}px;"
          else ""
        ("center", pad)
      case _ => ("flex-start", if indentPx > 0 then s" padding-left: ${indentPx}px;" else "")
    val boxStyle =
      s"margin-left: ${-(leftPx + indentPx)}px; width: ${leftPx + cellWidth + rightPx}px;" +
        s"$padding box-sizing: border-box; overflow: hidden; display: flex; justify-content: $justify"
    s"""<div class="$cssClass" style="$boxStyle"><span>$content</span></div>"""

  /**
   * Convert a CellValue to HTML content with Excel-style number formatting.
   *
   *   - Text: Escaped HTML
   *   - RichText: HTML with <b>, <i>, <u>, <span> tags
   *   - Number/DateTime/Bool: Formatted according to NumFmt, then escaped
   *   - Formula: Shows cached value formatted, or raw formula if no cache
   *   - Error: Excel error code
   */
  private def cellValueToHtml(cell: ResolvedCell, theme: ThemePalette): String =
    cell.cell.value match
      case CellValue.RichText(richText) =>
        // Rich text has its own formatting, don't apply NumFmt
        richText.runs.map(run => runToHtml(run, theme)).mkString

      case _ =>
        // The resolved content: a cached formula's value formatted with the NumFmt (matches
        // Excel display), an uncached formula's source, the formatted text of everything else
        escapeHtml(cell.content.text)

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
        if f.underline != Underline.None then html = s"<u>$html</u>"
        if f.italic then html = s"<i>$html</i>"
        if f.bold then html = s"<b>$html</b>"

        html

  /**
   * A CF data bar as a td background (GH-497): a no-repeat gradient from the bar colour to white,
   * inset 2px and as wide as SvgRenderer's bar ([[RenderUtils.dataBarWidth]]). None when the cell
   * leaves it no width.
   */
  private def dataBarCss(bar: CfBar, cellWidth: Int, theme: ThemePalette): Option[String] =
    val width = dataBarWidth(bar.fraction, cellWidth)
    Option.when(width > 0) {
      val hex = colorToHex(bar.color, theme)
      val inset = DataBarInsetPx
      s"background-image: linear-gradient(to right, $hex, #FFFFFF); background-size: ${width}px calc(100% - ${2 * inset}px); background-position: ${inset}px center; background-repeat: no-repeat"
    }

  /**
   * Convert cell-level style to inline CSS.
   *
   * Generates CSS properties for font, fill, borders, alignment, etc. Returns empty string if cell
   * has no style. `strike` (a CF dxf attribute `Font` cannot carry) adds line-through.
   */
  private def cellStyleToInlineCss(
    cell: ResolvedCell,
    strike: Boolean,
    theme: ThemePalette
  ): String =
    val styleOpt = cell.style
    val css = scala.collection.mutable.ArrayBuffer[String]()

    styleOpt.foreach { style =>
      // Font properties (apply only if not default)
      if style.font.bold then css += "font-weight: bold"
      if style.font.italic then css += "font-style: italic"
      textDecoration(style.font.underline != Underline.None, strike).foreach { d =>
        css += s"text-decoration: $d"
      }
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

    // Alignment - always emit to ensure proper alignment. The style's explicit alignment wins;
    // General resolves from the rendered kind exactly as SvgRenderer anchors it (GH-500, GH-501).
    cell.align match
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

    // Indentation (Excel uses ~3 characters per indent level)
    styleOpt.foreach { style =>
      if style.align.indent > 0 then
        val indentPx = style.align.indent * IndentPxPerLevel
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
