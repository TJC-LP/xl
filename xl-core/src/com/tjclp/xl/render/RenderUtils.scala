package com.tjclp.xl.render

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.display.{FormatCodeParser, NumFmtFormatter}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.{HAlign, VAlign}
import com.tjclp.xl.styles.color.{Color, ThemePalette}
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

import scala.util.boundary, boundary.break

/**
 * What a cell lays out as once its number format is applied.
 *
 * The three overflow decisions — whether a too-wide value hashes, how General alignment anchors it,
 * and how wide it is when sizing an overflow span — are taken from this kind and its text, never
 * from the raw `CellValue`: the renderers draw the formatted text, so deciding from the raw value
 * measures one string and rules on another (GH-500, GH-501, GH-502).
 */
enum RenderedKind derives CanEqual:
  /** A number or date under a numeric format: right-aligned, `####` when too wide. */
  case Numeric

  /** Text, rich text, an uncached formula's source, or a number under a text-only format. */
  case Text

  /** `TRUE` / `FALSE`: centred, `####` when too wide. */
  case Bool

  /** An Excel error code such as `#DIV/0!`: centred like a logical, `####` when too wide. */
  case Error

  /** Nothing to draw. */
  case Empty

/** The effective rendered content of a cell: its [[RenderedKind]] and the text it draws. */
final case class RenderedContent(kind: RenderedKind, text: String)

/**
 * A cell with its style and [[RenderedContent]] resolved — once, at the top of a renderer's
 * per-cell path. Every decision that follows (overflow span, anchor, `####`, the text itself) reads
 * these, and resolving re-formats the value and re-parses a Custom format code, so the renderers
 * thread this through instead of resolving at each decision.
 */
final case class ResolvedCell(cell: Cell, style: Option[CellStyle], content: RenderedContent):
  /** The cell's number format: the style's, or General. */
  def numFmt: NumFmt = style.map(_.numFmt).getOrElse(NumFmt.General)

  /** The alignment the renderers anchor by (see [[RenderUtils.resolveHAlign]]). */
  def align: HAlign = RenderUtils.resolveHAlign(style, content)

object ResolvedCell:
  /** Resolve `cell` against `sheet`'s style registry. */
  def apply(cell: Cell, sheet: Sheet): ResolvedCell =
    val style = cell.styleId.flatMap(sheet.styleRegistry.get)
    val numFmt = style.map(_.numFmt).getOrElse(NumFmt.General)
    ResolvedCell(cell, style, RenderUtils.renderedContent(cell.value, numFmt))

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

  /** Default column width in pixels (Excel default ~8.43 chars × 8 + 5 ≈ 72). */
  val DefaultColumnWidthPx: Int = 72

  /** Default font size in points (Calibri 11pt). */
  val DefaultFontSize: Int = 11

  /** Horizontal text padding in pixels. */
  val CellPaddingX: Int = 6

  /** Row label column width in pixels. */
  val HeaderWidth: Int = 40

  /** Column label row height in pixels. */
  val HeaderHeight: Int = 24

  /** Pixels per indent level (~3 characters at ~7px each, matches Excel behavior). */
  val IndentPxPerLevel: Int = 21

  /** Inter-run gap for rich text (compensates for AWT vs SVG font metric differences). */
  val InterRunGapPx: Int = 4

  // ========== Unit Conversion ==========

  /**
   * Convert Excel column width (character units) to pixels.
   *
   * Excel's width unit is based on the "0" character in the default font. For SVG rendering with
   * 15px Calibri (11pt × 4/3), characters average ~8px wide. Using 8× multiplier + 5px padding
   * provides better fidelity than the previous 7× factor.
   */
  def excelColWidthToPixels(width: Double): Int =
    (width * 8 + 5).toInt

  /** Convert Excel row height (points) to pixels. 1pt = 4/3 pixels. */
  def excelRowHeightToPixels(height: Double): Int =
    (height * 4.0 / 3.0).toInt

  // ========== Font Measurement (platform-specific backend) ==========

  /** Measure text width via the platform text measurer (AWT on JVM, heuristic elsewhere). */
  def measureTextWidth(text: String, font: Option[Font]): Int =
    TextMeasure.measureTextWidth(text, font)

  /** Pure width estimation used when no platform font metrics are available. */
  private[render] def estimateTextWidth(text: String, font: Option[Font]): Int =
    val baseCharWidth = 7
    val sizeFactor = font.map(f => f.sizePt / DefaultFontSize).getOrElse(1.0)
    val boldFactor = if font.exists(_.bold) then 1.1 else 1.0
    (text.length * baseCharWidth * sizeFactor * boldFactor).toInt

  /**
   * Measure the width of what a value draws under a number format. Rich text sums its runs (each
   * carries its own font); everything else measures its formatted text — `$1,234.50` rather than
   * `1234.5`, `0.12` rather than `0.123456789`, `#DIV/0!` rather than the enum case name — because
   * that is the string the renderers emit (GH-502).
   */
  def measureCellValueWidth(value: CellValue, numFmt: NumFmt, font: Option[Font]): Int =
    measureContentWidth(value, renderedContent(value, numFmt), font)

  /** [[measureCellValueWidth]] with the content already resolved (`content` is `value`'s). */
  private def measureContentWidth(
    value: CellValue,
    content: RenderedContent,
    font: Option[Font]
  ): Int =
    value match
      case CellValue.RichText(rt) =>
        rt.runs.map(run => measureTextWidth(run.text, run.font.orElse(font))).sum
      case CellValue.Formula(_, Some(cached), _) => measureContentWidth(cached, content, font)
      case _ => measureTextWidth(content.text, font)

  /** [[measureCellValueWidth]] under the General number format. */
  def measureCellValueWidth(value: CellValue, font: Option[Font]): Int =
    measureCellValueWidth(value, NumFmt.General, font)

  // ========== Text Overflow Calculation ==========

  /**
   * Check if a cell is both empty (no content) and unstyled (no borders/fills).
   *
   * Excel prevents text overflow into cells with styling even if they're empty. This helper checks
   * both content and styling to match Excel's overflow behavior.
   *
   * @param ref
   *   Cell reference to check
   * @param sheet
   *   Sheet containing the cell
   * @return
   *   true if cell is empty and has no styling (allows overflow)
   */
  private def isCellEmptyAndUnstyled(ref: ARef, sheet: Sheet): Boolean =
    boundary:
      import com.tjclp.xl.styles.border.Border
      import com.tjclp.xl.styles.fill.Fill

      val cellOpt = sheet.cells.get(ref)

      // Check if cell has content
      val isEmpty = cellOpt.forall { c =>
        c.value match
          case CellValue.Empty => true
          case _ => false
      }
      if !isEmpty then break(false)

      // Check if cell has styling (borders or fills) - Excel blocks overflow into styled cells
      val hasNoStyle = cellOpt.flatMap(_.styleId).flatMap(sheet.styleRegistry.get).forall { style =>
        style.border == Border.none && style.fill == Fill.None
      }
      if !hasNoStyle then break(false)

      // Check if cell is part of a merged region
      val notMerged = sheet.getMergedRange(ref).isEmpty

      isEmpty && hasNoStyle && notMerged

  /**
   * Calculate overflow colspan for a cell with text that exceeds its width.
   *
   * Only TEXT overflows. A number, date, logical or error never borrows a neighbour, however empty
   * and whatever the cell's alignment: Excel confines it to its own column and shows `####` when it
   * does not fit (GH-459, GH-500). For left/centre-aligned text, counts empty cells to the right
   * until:
   *   - A non-empty cell is reached
   *   - The accumulated width covers the text overflow
   *   - The range boundary is reached
   *
   * @param cell
   *   The cell to check for overflow, with its style and rendered content resolved
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
    cell: ResolvedCell,
    cellRef: ARef,
    cellWidth: Int,
    colWidths: IndexedSeq[Int],
    sheet: Sheet,
    startCol: Int,
    endCol: Int
  ): Int =
    import scala.util.boundary, boundary.break

    boundary:
      val style = cell.style

      // If wrapText is true, text wraps instead of overflowing
      if style.exists(_.align.wrapText) then break(1)

      // A kind that hashes never spans: Excel draws a too-wide number, date, logical or error
      // as #### inside its own column, empty neighbour or not, Left/Center alignment or not.
      // Gating here — before the alignment match — is what lets hashOverflowText see the
      // cell's OWN width instead of an already widened span (GH-500).
      if hashesOnOverflow(cell.content.kind) then break(1)

      val font = style.map(_.font)
      // Size the span from what the renderers DRAW — the formatted text — never the raw value:
      // a rounding format must not claim neighbours for digits it never shows, and a widening
      // one must get the room its text needs (GH-502).
      val textWidth = measureContentWidth(cell.cell.value, cell.content, font)

      // If text fits within cell, no overflow needed
      if textWidth <= cellWidth then break(1)

      // Determine overflow direction based on alignment: the same resolution the renderers use
      // when they anchor the text.
      cell.align match
        case HAlign.Left | HAlign.General =>
          // Overflow to the right (General alignment for text behaves like Left)
          countEmptyToRight(cellRef, cellWidth, colWidths, sheet, startCol, endCol, textWidth)
        case HAlign.Center | HAlign.CenterContinuous =>
          // Center-aligned text can overflow right (like Excel)
          // Excel allows center text to bleed into empty cells on the right
          countEmptyToRight(cellRef, cellWidth, colWidths, sheet, startCol, endCol, textWidth)
        case HAlign.Right =>
          // Right-aligned text clips in Excel (doesn't overflow left)
          // This matches Excel's actual behavior
          1
        case _ =>
          1

  /** [[calculateOverflowColspan]] resolving `cell`'s style and content against `sheet` first. */
  def calculateOverflowColspan(
    cell: Cell,
    cellRef: ARef,
    cellWidth: Int,
    colWidths: IndexedSeq[Int],
    sheet: Sheet,
    startCol: Int,
    endCol: Int
  ): Int =
    calculateOverflowColspan(
      ResolvedCell(cell, sheet),
      cellRef,
      cellWidth,
      colWidths,
      sheet,
      startCol,
      endCol
    )

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

        // Use helper to check if cell allows overflow (empty AND unstyled)
        if !isCellEmptyAndUnstyled(nextRef, sheet) then
          // Stop - can't overflow into non-empty, styled, or merged cell
          colspan
        else
          // Include this empty unstyled cell in the overflow span
          val widthIdx = nextCol - startCol
          val newWidth =
            if widthIdx >= 0 && widthIdx < colWidths.length then
              accumulatedWidth + colWidths(widthIdx)
            else accumulatedWidth
          loop(nextCol + 1, colspan + 1, newWidth)

    loop(colIdx + 1, 1, cellWidth)

  // ========== Numeric Overflow Marker (####) ==========

  /**
   * Kinds Excel replaces with `#` when they do not fit: numbers and dates (GH-459), logicals and
   * errors (GH-500). Text is never hashed — it bleeds into empty neighbours or clips, which loses
   * no information the reader can misread — and a number under a text-only format IS text (GH-501).
   */
  private def hashesOnOverflow(kind: RenderedKind): Boolean = kind match
    case RenderedKind.Numeric | RenderedKind.Bool | RenderedKind.Error => true
    case RenderedKind.Text | RenderedKind.Empty => false

  /**
   * The pixel box a cell's text actually occupies inside its clip, given the alignment the
   * renderers resolve for it and its indent.
   *
   *   - right-anchored text ends `CellPaddingX` inside the box, but the anchor is clamped back to
   *     the box's left edge when the text is wider than that (SvgRenderer), so the whole width is
   *     usable;
   *   - left-anchored text starts at `CellPaddingX + indent` and runs to the clip's right edge, so
   *     the indent and the leading pad are gone from its box;
   *   - centred text is pushed right by half the indent, so it loses the indent overall.
   *
   * Deriving the box here rather than in each renderer is what keeps HTML and SVG hashing the same
   * cells: HTML's geometry differs (no horizontal padding on data cells, `padding-left` for indent)
   * but the decision must not.
   */
  private def textBoxWidth(align: HAlign, style: Option[CellStyle], cellWidth: Int): Int =
    val indentPx = style.map(_.align.indent).getOrElse(0) * IndentPxPerLevel
    val box = align match
      case HAlign.Right => cellWidth
      case HAlign.Center | HAlign.CenterContinuous => cellWidth - indentPx
      case _ => cellWidth - CellPaddingX - indentPx
    math.max(0, box)

  /**
   * Excel's `####` overflow marker for a numeric, date, logical or error cell whose formatted text
   * is wider than the space it renders in.
   *
   * Clipping a numeral is not a cosmetic defect: shearing the leading digits off `1,234,567.9`
   * leaves `4,567.9`, a different number that still looks like a real one (GH-459). Excel refuses
   * to show a partial numeral and fills the column with `#` instead. The cut is just as misleading
   * from the other end: a left-aligned or indented number whose tail runs past the clip renders
   * `1234567.9` as `1234`, so the fit is tested against the cell's real text box (`textBoxWidth`),
   * not against the whole column. `TRUE` and `#DIV/0!` get the same treatment (GH-500): Excel
   * hashes them too, and a clipped `#DIV/0!` is a fragment no reader can place.
   *
   * The decision is taken from the value's rendered content and its style rather than from
   * renderer-local text, so SVG and HTML hash the same cells with the same marker.
   *
   * @param content
   *   the cell's rendered content — hashed when its kind is Numeric, Bool or Error
   * @param style
   *   the cell's resolved style: number format, font, alignment, indent and wrapText
   * @param availableWidth
   *   the pixel width of the cell, after merge/overflow expansion
   * @return
   *   the `#` run to render in place of the text, or None to render the text unchanged
   */
  def hashOverflowText(
    content: RenderedContent,
    style: Option[CellStyle],
    availableWidth: Int
  ): Option[String] =
    val wrapText = style.exists(_.align.wrapText)
    if wrapText || availableWidth <= 0 || !hashesOnOverflow(content.kind) then None
    else
      val font = style.map(_.font)
      val text = content.text
      val box = textBoxWidth(resolveHAlign(style, content), style, availableWidth)
      if text.isEmpty || measureTextWidth(text, font) <= box then None
      else
        // The marker must itself fit where the text would have gone: bounded by the cell's
        // inner width (both pads) and by the text box (indent, single pad).
        val innerWidth = math.max(0, math.min(box, availableWidth - CellPaddingX * 2))
        val hashWidth = math.max(1, measureTextWidth("#", font))
        // At least one '#': a column too narrow for even one marker must still refuse to
        // show digits.
        Some("#" * math.max(1, innerWidth / hashWidth))

  /** [[hashOverflowText]] resolving `value`'s content under `style`'s number format first. */
  def hashOverflowText(
    value: CellValue,
    style: Option[CellStyle],
    availableWidth: Int
  ): Option[String] =
    val numFmt = style.map(_.numFmt).getOrElse(NumFmt.General)
    hashOverflowText(renderedContent(value, numFmt), style, availableWidth)

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
      // explicit filter, not String.replace: Scala Native's javalib String.replace
      // mishandles the NUL pattern (ADR-016 spike); the JVM result is identical either way
      .filterNot(_ == '\u0000')

  /**
   * Format a font family name for an SVG presentation attribute (`font-family="..."`).
   *
   * SVG presentation attributes accept multi-word family names unquoted; CSS-style single quotes
   * would become part of the family name, so fontconfig-based rasterizers (rsvg-convert, resvg)
   * fail to match the family and silently substitute sans-serif (GH-255). Only XML escaping is
   * applied. Quotes are still required in CSS contexts (HTML `style=""` attributes and `<style>`
   * blocks) — use `escapeCss` with explicit quotes there.
   */
  def svgFontFamily(name: String): String = escapeXml(name)

  // ========== Color Resolution ==========

  /** Convert Color to CSS/SVG-compatible RGB hex using theme for resolution. Returns #RRGGBB. */
  def colorToHex(c: Color, theme: ThemePalette): String =
    try c.toResolvedHex(theme)
    catch
      case e: Exception =>
        // Fallback to black if theme color resolution fails
        System.err.println(
          s"Warning: Color resolution failed for $c, using black. Error: ${e.getMessage}"
        )
        "#000000"

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

  /**
   * Resolve the effective rendered content of a value under a number format — the one input to
   * every overflow decision and the text both renderers draw (`cellValueToText` is this `.text`).
   *
   * A formula resolves through its cached value, or to its source when uncached. A number or date
   * under a text-only format (`@`, or a Custom code with no numeric section) is Text: Excel shows
   * its General digits but lays the cell out as text — left-aligned, overflowing into empty
   * neighbours, never `####` (GH-501). A 4-section code whose last arm is `@` is still Numeric.
   */
  def renderedContent(value: CellValue, numFmt: NumFmt): RenderedContent = value match
    case CellValue.RichText(rt) => RenderedContent(RenderedKind.Text, rt.toPlainText)
    case CellValue.Empty => RenderedContent(RenderedKind.Empty, "")
    case CellValue.Formula(_, Some(cached), _) => renderedContent(cached, numFmt)
    case CellValue.Formula(expr, None, _) => RenderedContent(RenderedKind.Text, s"=$expr")
    case CellValue.Text(_) =>
      RenderedContent(RenderedKind.Text, NumFmtFormatter.formatValue(value, numFmt))
    case CellValue.Bool(_) =>
      RenderedContent(RenderedKind.Bool, NumFmtFormatter.formatValue(value, numFmt))
    case CellValue.Error(_) =>
      RenderedContent(RenderedKind.Error, NumFmtFormatter.formatValue(value, numFmt))
    case CellValue.Number(_) | CellValue.DateTime(_) =>
      val kind = if isTextOnlyFormat(numFmt) then RenderedKind.Text else RenderedKind.Numeric
      RenderedContent(kind, NumFmtFormatter.formatValue(value, numFmt))

  /** `@`, or a Custom code with no numeric section: Excel lays a number out as text under it. */
  private def isTextOnlyFormat(numFmt: NumFmt): Boolean = numFmt match
    case NumFmt.Text => true
    case NumFmt.Custom(code) => FormatCodeParser.parse(code).exists(FormatCodeParser.isTextOnly)
    case _ => false

  /**
   * Excel's General alignment by rendered kind: numbers and dates right, logicals AND errors
   * centred, text left. Errors used to fall to Left, which is where `#DIV/0!` was anchored and
   * clipped (GH-500).
   */
  def alignmentFor(kind: RenderedKind): HAlign = kind match
    case RenderedKind.Numeric => HAlign.Right
    case RenderedKind.Bool | RenderedKind.Error => HAlign.Center
    case RenderedKind.Text | RenderedKind.Empty => HAlign.Left

  /**
   * The horizontal alignment the renderers anchor by: the style's explicit alignment, or General
   * resolved from the rendered content. Both renderers and every overflow decision go through here,
   * so a cell cannot be anchored one way and hashed or spanned another.
   */
  def resolveHAlign(style: Option[CellStyle], content: RenderedContent): HAlign =
    style.map(_.align.horizontal).getOrElse(HAlign.General) match
      case HAlign.General => alignmentFor(content.kind)
      case explicit => explicit

  /** General alignment of a value under the General number format (see [[alignmentFor]]). */
  def contentBasedAlignment(value: CellValue): HAlign =
    alignmentFor(renderedContent(value, NumFmt.General).kind)

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

  /** Get cell value as plain text with formatting: the text of [[renderedContent]]. */
  def cellValueToText(value: CellValue, numFmt: NumFmt): String =
    renderedContent(value, numFmt).text

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
