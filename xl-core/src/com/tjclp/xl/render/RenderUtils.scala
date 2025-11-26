package com.tjclp.xl.render

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.alignment.{HAlign, VAlign}
import com.tjclp.xl.styles.color.{Color, ThemePalette}
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

import java.awt.{Font as AwtFont, Graphics2D}
import java.awt.image.BufferedImage

/**
 * Shared utilities for rendering.
 *
 * Contains common constants, unit conversions, font measurement, escaping, and color resolution
 * functions used by all renderers (HTML, SVG, etc.).
 */
object RenderUtils:

  // ========== Constants (Harmonized) ==========

  /** Default cell height in pixels (Excel default ~15pt). */
  val DefaultCellHeightPx: Int = 20

  /** Default column width in pixels (Excel default ~8.43 chars). */
  val DefaultColumnWidthPx: Int = 64

  /** Default font size in points (Calibri 11pt). */
  val DefaultFontSize: Int = 11

  /** Horizontal text padding in pixels. */
  val CellPaddingX: Int = 6

  /** Row label column width in pixels. */
  val HeaderWidth: Int = 40

  /** Column label row height in pixels. */
  val HeaderHeight: Int = 24

  // ========== Unit Conversion ==========

  /** Convert Excel column width (character units) to pixels. */
  def excelColWidthToPixels(width: Double): Int =
    (width * 7 + 5).toInt

  /** Convert Excel row height (points) to pixels. 1pt = 4/3 pixels. */
  def excelRowHeightToPixels(height: Double): Int =
    (height * 4.0 / 3.0).toInt

  // ========== AWT Font Measurement ==========

  /** Graphics context for text measurement (lazy, with headless fallback). */
  private lazy val graphics: Option[Graphics2D] =
    try
      val img = new BufferedImage(1, 1, BufferedImage.TYPE_INT_ARGB)
      Some(img.createGraphics())
    catch case _: Exception => None // Headless environment

  /** Measure text width using AWT FontMetrics, with fallback estimation. */
  def measureTextWidth(text: String, font: Option[Font]): Int =
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
  def toAwtFont(font: Option[Font]): AwtFont =
    font match
      case Some(f) =>
        val style =
          (if f.bold then AwtFont.BOLD else 0) | (if f.italic then AwtFont.ITALIC else 0)
        new AwtFont(f.name, style, f.sizePt.toInt)
      case None =>
        new AwtFont("Calibri", AwtFont.PLAIN, DefaultFontSize)

  /** Measure text width for a CellValue, handling rich text runs. */
  def measureCellValueWidth(value: CellValue, font: Option[Font]): Int = value match
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

  // ========== Text Overflow Calculation ==========

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
  def calculateOverflowColspan(
    cell: com.tjclp.xl.cells.Cell,
    cellRef: ARef,
    cellWidth: Int,
    colWidths: IndexedSeq[Int],
    sheet: Sheet,
    startCol: Int,
    endCol: Int
  ): Int =
    import scala.util.boundary, boundary.break

    boundary:
      val style = cell.styleId.flatMap(sheet.styleRegistry.get)

      // If wrapText is true, text wraps instead of overflowing
      if style.exists(_.align.wrapText) then break(1)

      val font = style.map(_.font)
      val textWidth = measureCellValueWidth(cell.value, font)

      // If text fits within cell, no overflow needed
      if textWidth <= cellWidth then break(1)

      // Determine overflow direction based on alignment
      // General alignment behaves like Left for text overflow in Excel
      val align = style.map(_.align.horizontal).getOrElse(HAlign.General)
      align match
        case HAlign.Left | HAlign.General =>
          // Overflow to the right (General alignment for text behaves like Left)
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

    // Use tail recursion to iterate through cells
    @scala.annotation.tailrec
    def loop(nextCol: Int, colspan: Int, accumulatedWidth: Int): Int =
      if nextCol > endCol || accumulatedWidth >= textWidth then colspan
      else
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
          colspan
        else
          // Include this empty cell in the overflow span
          val widthIdx = nextCol - startCol
          val newWidth =
            if widthIdx >= 0 && widthIdx < colWidths.length then
              accumulatedWidth + colWidths(widthIdx)
            else accumulatedWidth
          loop(nextCol + 1, colspan + 1, newWidth)

    loop(colIdx + 1, 1, cellWidth)

  // ========== Escaping ==========

  /** Escape HTML special characters. */
  def escapeHtml(s: String): String =
    s.replace("&", "&amp;")
      .replace("<", "&lt;")
      .replace(">", "&gt;")
      .replace("\"", "&quot;")
      .replace("'", "&#39;")

  /** Escape XML special characters. */
  def escapeXml(s: String): String =
    s.replace("&", "&amp;")
      .replace("<", "&lt;")
      .replace(">", "&gt;")
      .replace("\"", "&quot;")
      .replace("'", "&apos;")

  /** Escape CSS special characters. */
  def escapeCss(s: String): String =
    s.replace("\\", "\\\\")
      .replace("'", "\\'")
      .replace("\"", "\\\"")
      .replace("\n", "\\A ")
      .replace("\r", "\\D ")
      .replace("\u0000", "")

  // ========== Color Resolution ==========

  /** Convert Color to CSS/SVG-compatible RGB hex using theme for resolution. Returns #RRGGBB. */
  def colorToHex(c: Color, theme: ThemePalette): String =
    c.toResolvedHex(theme)

  /**
   * Convert Color to SVG fill attributes with opacity support. Returns both fill and fill-opacity
   * for translucent ARGB colors, just fill for fully opaque colors.
   */
  def colorToFillAttrsWithOpacity(c: Color, theme: ThemePalette): String =
    val argb = c.toResolvedArgb(theme)
    val alpha = ((argb >> 24) & 0xff) / 255.0
    val hex = f"#${argb & 0xffffff}%06X"
    if alpha >= 1.0 then s"""fill="$hex""""
    else s"""fill="$hex" fill-opacity="${alpha}""""

  // ========== Common Rendering Logic ==========

  /** Determine default horizontal alignment based on cell value type (Excel's General behavior). */
  def contentBasedAlignment(value: CellValue): HAlign = value match
    case CellValue.Number(_) | CellValue.DateTime(_) => HAlign.Right
    case CellValue.Bool(_) => HAlign.Center
    case CellValue.Formula(_, Some(cached)) => contentBasedAlignment(cached)
    case _ => HAlign.Left

  /** Calculate column widths for a range, respecting sheet properties. */
  def calculateColumnWidths(
    sheet: Sheet,
    range: CellRange,
    scaleFactor: Double = 1.0
  ): IndexedSeq[Int] =
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0

    (startCol to endCol).map { colIdx =>
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

  /** Calculate row heights for a range, respecting sheet properties. */
  def calculateRowHeights(
    sheet: Sheet,
    range: CellRange,
    scaleFactor: Double = 1.0
  ): IndexedSeq[Int] =
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0

    (startRow to endRow).map { rowIdx =>
      val row = Row.from0(rowIdx)
      getRowHeight(sheet, row, scaleFactor)
    }

  /** Get the height of a single row in pixels. */
  def getRowHeight(sheet: Sheet, row: Row, scaleFactor: Double = 1.0): Int =
    val props = sheet.getRowProperties(row)
    if props.hidden then 0
    else
      val baseHeight = props.height
        .map(excelRowHeightToPixels)
        .getOrElse(
          sheet.defaultRowHeight
            .map(excelRowHeightToPixels)
            .getOrElse(DefaultCellHeightPx)
        )
      (baseHeight * scaleFactor).toInt

  /** Get cell value as plain text with formatting. */
  def cellValueToText(value: CellValue, numFmt: NumFmt): String = value match
    case CellValue.RichText(rt) => rt.toPlainText
    case CellValue.Empty => ""
    case CellValue.Formula(_, Some(cached)) => NumFmtFormatter.formatValue(cached, numFmt)
    case CellValue.Formula(expr, None) => s"=$expr"
    case other => NumFmtFormatter.formatValue(other, numFmt)

/**
 * Base trait for cell renderers.
 *
 * Provides a common interface and default implementations for rendering Excel sheets to various
 * output formats. Extend this trait to create new renderers (HTML, SVG, PDF, Markdown, etc.).
 *
 * The trait defines:
 *   - Common configuration via RenderConfig
 *   - Shared rendering logic for column/row sizing, overflow calculation
 *   - Abstract methods for format-specific output generation
 */
trait CellRenderer:
  import RenderUtils.*

  /** Configuration for rendering. */
  case class RenderConfig(
    includeStyles: Boolean = true,
    theme: ThemePalette = ThemePalette.office,
    showLabels: Boolean = false,
    scaleFactor: Double = 1.0
  )

  /** Output type for the renderer (e.g., String for text formats, Array[Byte] for binary). */
  type Output

  /** Render a sheet range to the output format. */
  def render(sheet: Sheet, range: CellRange, config: RenderConfig): Output

  // ========== Helper Methods Available to Subclasses ==========

  /** Calculate effective column widths for rendering. */
  protected def getColumnWidths(
    sheet: Sheet,
    range: CellRange,
    config: RenderConfig
  ): IndexedSeq[Int] =
    calculateColumnWidths(sheet, range, config.scaleFactor)

  /** Calculate effective row heights for rendering. */
  protected def getRowHeights(
    sheet: Sheet,
    range: CellRange,
    config: RenderConfig
  ): IndexedSeq[Int] =
    calculateRowHeights(sheet, range, config.scaleFactor)

  /** Get total width for rendering (with optional label column). */
  protected def getTotalWidth(colWidths: IndexedSeq[Int], showLabels: Boolean): Int =
    val dataWidth = colWidths.sum
    if showLabels then HeaderWidth + dataWidth else dataWidth

  /** Get total height for rendering (with optional label row). */
  protected def getTotalHeight(
    rowHeights: IndexedSeq[Int],
    showLabels: Boolean,
    scaleFactor: Double
  ): Int =
    val dataHeight = rowHeights.sum
    val headerHeight = if showLabels then (HeaderHeight * scaleFactor).toInt else 0
    headerHeight + dataHeight

  /** Calculate column x positions (cumulative widths). */
  protected def getColumnXPositions(
    colWidths: IndexedSeq[Int],
    showLabels: Boolean
  ): IndexedSeq[Int] =
    val xOffset = if showLabels then HeaderWidth else 0
    colWidths.scanLeft(xOffset)(_ + _).dropRight(1)

  /** Calculate row y positions (cumulative heights). */
  protected def getRowYPositions(
    rowHeights: IndexedSeq[Int],
    showLabels: Boolean,
    scaleFactor: Double
  ): IndexedSeq[Int] =
    val yOffset = if showLabels then (HeaderHeight * scaleFactor).toInt else 0
    rowHeights.scanLeft(yOffset)(_ + _).dropRight(1)
