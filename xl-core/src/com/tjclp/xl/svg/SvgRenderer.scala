package com.tjclp.xl.svg

import java.awt.{Font as AwtFont, Graphics2D}
import java.awt.image.BufferedImage

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.richtext.TextRun
import com.tjclp.xl.styles.alignment.{HAlign, VAlign}
import com.tjclp.xl.styles.border.{BorderStyle, BorderSide}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font

/**
 * Renders Excel sheets to SVG images with styled cells.
 *
 * SVG output can be viewed directly by Claude as images, making it useful for visualizing formatted
 * Excel data including colors, fonts, and borders.
 */
object SvgRenderer:

  // Default dimensions
  private val DefaultCellHeight = 24
  private val HeaderWidth = 40 // Row number column width
  private val HeaderHeight = 24 // Column letter row height
  private val MinCellWidth = 60
  private val CellPaddingX = 6
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

  /**
   * Export a sheet range to an SVG image.
   *
   * @param sheet
   *   The sheet to export
   * @param range
   *   The cell range to export
   * @param includeStyles
   *   Whether to include cell styling (colors, fonts, borders)
   * @return
   *   SVG string
   */
  def toSvg(
    sheet: Sheet,
    range: CellRange,
    includeStyles: Boolean = true
  ): String =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0

    // Calculate column widths based on content
    val colWidths = calculateColumnWidths(sheet, range)

    // Calculate total dimensions
    val totalWidth = HeaderWidth + colWidths.sum
    val totalHeight = HeaderHeight + (endRow - startRow + 1) * DefaultCellHeight

    val sb = new StringBuilder

    // SVG header
    sb.append(
      s"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 $totalWidth $totalHeight" """
    )
    sb.append(s"""width="$totalWidth" height="$totalHeight">\n""")

    // Embedded styles
    sb.append("""  <style>
    .header { fill: #E0E0E0; stroke: #999999; stroke-width: 1; }
    .header-text { font-family: 'Segoe UI', Arial, sans-serif; font-size: 11px; fill: #333333; }
    .cell { stroke: #D0D0D0; stroke-width: 0.5; }
    .cell-text { font-family: 'Calibri', 'Segoe UI', Arial, sans-serif; font-size: 11px; }
  </style>
""")

    // Pre-calculate column x positions (cumulative widths)
    val colXPositions = colWidths.scanLeft(HeaderWidth)(_ + _).dropRight(1)

    // Column headers (A, B, C...)
    sb.append("  <g class=\"col-headers\">\n")
    (startCol to endCol).foreach { col =>
      val colIdx = col - startCol
      val colLetter = Column.from0(col).toLetter
      val width = colWidths(colIdx)
      val xOffset = colXPositions(colIdx)
      sb.append(
        s"""    <rect x="$xOffset" y="0" width="$width" height="$HeaderHeight" class="header"/>\n"""
      )
      val textX = xOffset + width / 2
      val textY = HeaderHeight / 2 + 4
      sb.append(
        s"""    <text x="$textX" y="$textY" text-anchor="middle" class="header-text">$colLetter</text>\n"""
      )
    }
    sb.append("  </g>\n")

    // Row headers (1, 2, 3...)
    sb.append("  <g class=\"row-headers\">\n")
    (startRow to endRow).foreach { row =>
      val rowNum = row + 1
      val y = HeaderHeight + (row - startRow) * DefaultCellHeight
      sb.append(
        s"""    <rect x="0" y="$y" width="$HeaderWidth" height="$DefaultCellHeight" class="header"/>\n"""
      )
      val textX = HeaderWidth / 2
      val textY = y + DefaultCellHeight / 2 + 4
      sb.append(
        s"""    <text x="$textX" y="$textY" text-anchor="middle" class="header-text">$rowNum</text>\n"""
      )
    }
    sb.append("  </g>\n")

    // Cells
    sb.append("  <g class=\"cells\">\n")
    (startRow to endRow).foreach { row =>
      val y = HeaderHeight + (row - startRow) * DefaultCellHeight
      (startCol to endCol).foreach { col =>
        val colIdx = col - startCol
        val xOffset = colXPositions(colIdx)
        val ref = ARef.from0(col, row)
        val width = colWidths(colIdx)
        val cellOpt = sheet.cells.get(ref)

        // Cell background and border
        val (fillAttr, strokeAttr) = cellOpt
          .flatMap(c => if includeStyles then cellStyleToSvg(c, sheet) else None)
          .getOrElse(("fill=\"#FFFFFF\"", ""))

        sb.append(s"""    <rect x="$xOffset" y="$y" width="$width" height="$DefaultCellHeight" """)
        sb.append(s"""$fillAttr class="cell" $strokeAttr/>\n""")

        // Cell text
        cellOpt.foreach { cell =>
          val (textX, anchor) = textAlignment(cell, sheet, xOffset, width)
          val textY = y + DefaultCellHeight / 2 + 4

          cell.value match
            case CellValue.RichText(rt) if includeStyles && rt.runs.nonEmpty =>
              // Rich text: render each run as a tspan with explicit x positioning
              sb.append(s"""    <text y="$textY" class="cell-text">""")
              rt.runs.foldLeft(textX) { (currentX, run) =>
                val runStyle = runToSvgStyle(run)
                val escapedText = escapeXml(run.text)
                sb.append(s"""<tspan x="$currentX"$runStyle>$escapedText</tspan>""")
                currentX + measureTextWidth(run.text, run.font)
              }
              sb.append("</text>\n")

            case other =>
              val text = cellValueToText(other)
              if text.nonEmpty then
                val textStyle = if includeStyles then cellTextStyle(cell, sheet) else ""
                val escapedText = escapeXml(text)
                sb.append(
                  s"""    <text x="$textX" y="$textY" text-anchor="$anchor" class="cell-text"$textStyle>"""
                )
                sb.append(s"""$escapedText</text>\n""")
        }
      }
    }
    sb.append("  </g>\n")

    sb.append("</svg>")
    sb.toString

  /**
   * Calculate column widths based on cell content.
   */
  private def calculateColumnWidths(sheet: Sheet, range: CellRange): Vector[Int] =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0

    (startCol to endCol).map { col =>
      val headerWidth = measureTextWidth(Column.from0(col).toLetter, None) + CellPaddingX * 2
      val maxContentWidth = (startRow to endRow)
        .map { row =>
          val ref = ARef.from0(col, row)
          sheet.cells.get(ref) match
            case Some(cell) => measureCellValueWidth(cell.value) + CellPaddingX * 2
            case None => 0
        }
        .maxOption
        .getOrElse(0)
      math.max(MinCellWidth, math.max(headerWidth, maxContentWidth))
    }.toVector

  /** Measure the width of a cell value, handling rich text runs. */
  private def measureCellValueWidth(value: CellValue): Int = value match
    case CellValue.RichText(rt) =>
      rt.runs.map(run => measureTextWidth(run.text, run.font)).sum
    case CellValue.Formula(_, Some(cached)) =>
      measureCellValueWidth(cached)
    case other =>
      measureTextWidth(cellValueToText(other), None)

  /**
   * Convert cell value to plain text for display.
   */
  private def cellValueToText(value: CellValue): String = value match
    case CellValue.Text(s) => s
    case CellValue.RichText(rt) => rt.toPlainText
    case CellValue.Number(n) => formatNumber(n)
    case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
    case CellValue.DateTime(dt) => dt.toLocalDate.toString
    case CellValue.Empty => ""
    case CellValue.Formula(_, Some(cv)) => cellValueToText(cv)
    case CellValue.Formula(expr, None) => s"=$expr"
    case CellValue.Error(err) => err.toExcel

  private def formatNumber(n: BigDecimal): String =
    if n.isWhole then n.toBigInt.toString
    else n.underlying.stripTrailingZeros.toPlainString

  /**
   * Get SVG fill and stroke attributes for a cell's background.
   */
  private def cellStyleToSvg(cell: Cell, sheet: Sheet): Option[(String, String)] =
    cell.styleId.flatMap(sheet.styleRegistry.get).map { style =>
      val fill = style.fill match
        case Fill.Solid(color) => s"""fill="${colorToSvgHex(color)}""""
        case _ => """fill="#FFFFFF""""

      val stroke = borderToStroke(style.border)
      (fill, stroke)
    }

  /**
   * Convert border to SVG stroke attributes.
   */
  private def borderToStroke(border: com.tjclp.xl.styles.border.Border): String =
    // For simplicity, use the bottom border style for all borders
    // A more complete implementation would draw each side separately
    val side = border.bottom
    if side.style == BorderStyle.None then ""
    else
      val width = side.style match
        case BorderStyle.Thin => 1
        case BorderStyle.Medium => 2
        case BorderStyle.Thick => 3
        case _ => 1
      val color = side.color.map(colorToSvgHex).getOrElse("#000000")
      s"""stroke="$color" stroke-width="$width""""

  /**
   * Get SVG text style attributes for a cell.
   */
  private def cellTextStyle(cell: Cell, sheet: Sheet): String =
    cell.styleId
      .flatMap(sheet.styleRegistry.get)
      .map { style =>
        val attrs = scala.collection.mutable.ArrayBuffer[String]()

        // Font color
        style.font.color.foreach(c => attrs += s"""fill="${colorToSvgHex(c)}"""")

        // Font weight
        if style.font.bold then attrs += """font-weight="bold""""

        // Font style
        if style.font.italic then attrs += """font-style="italic""""

        // Font size (if different from default)
        if style.font.sizePt != Font.default.sizePt then
          attrs += s"""font-size="${style.font.sizePt}px""""

        // Font family (if different from default)
        if style.font.name != Font.default.name then
          attrs += s"""font-family="${escapeCss(style.font.name)}""""

        if attrs.nonEmpty then " " + attrs.mkString(" ") else ""
      }
      .getOrElse("")

  /**
   * Convert a TextRun to SVG tspan style attributes.
   */
  private def runToSvgStyle(run: TextRun): String =
    run.font match
      case None => ""
      case Some(f) =>
        val attrs = scala.collection.mutable.ArrayBuffer[String]()

        // Font color
        f.color.foreach(c => attrs += s"""fill="${colorToSvgHex(c)}"""")

        // Font weight
        if f.bold then attrs += """font-weight="bold""""

        // Font style
        if f.italic then attrs += """font-style="italic""""

        // Underline (SVG uses text-decoration)
        if f.underline then attrs += """text-decoration="underline""""

        // Font size (if different from default)
        if f.sizePt != Font.default.sizePt then attrs += s"""font-size="${f.sizePt}px""""

        // Font family (if different from default)
        if f.name != Font.default.name then attrs += s"""font-family="${escapeCss(f.name)}""""

        if attrs.nonEmpty then " " + attrs.mkString(" ") else ""

  /**
   * Calculate text x position and anchor based on alignment.
   */
  private def textAlignment(cell: Cell, sheet: Sheet, cellX: Int, cellWidth: Int): (Int, String) =
    val align = cell.styleId
      .flatMap(sheet.styleRegistry.get)
      .map(_.align.horizontal)
      .getOrElse(HAlign.Left)

    align match
      case HAlign.Center | HAlign.CenterContinuous =>
        (cellX + cellWidth / 2, "middle")
      case HAlign.Right =>
        (cellX + cellWidth - CellPaddingX, "end")
      case _ =>
        (cellX + CellPaddingX, "start")

  /**
   * Escape XML special characters.
   */
  private def escapeXml(s: String): String =
    s.replace("&", "&amp;")
      .replace("<", "&lt;")
      .replace(">", "&gt;")
      .replace("\"", "&quot;")
      .replace("'", "&apos;")

  private def escapeCss(s: String): String =
    s.replace("\\", "\\\\")
      .replace("'", "\\'")
      .replace("\"", "\\\"")

  /**
   * Convert Color to SVG-compatible RGB hex (strips alpha from ARGB). Color.toHex returns
   * #AARRGGBB, SVG needs #RRGGBB.
   */
  private def colorToSvgHex(c: Color): String =
    val hex = c.toHex
    if hex.length == 9 && hex.startsWith("#") then "#" + hex.drop(3) // #AARRGGBB -> #RRGGBB
    else hex
