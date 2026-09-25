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
import com.tjclp.xl.styles.units.StyleId

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
 * How many neighbour columns a cell's text spills into on each side (see
 * [[RenderUtils.overflowSpan]]): its clip box runs from `left` columns before the cell to `right`
 * columns after it.
 */
final case class OverflowSpan(left: Int, right: Int):
  /** Whether the text reaches past its own cell at all. */
  def spills: Boolean = left > 0 || right > 0

object OverflowSpan:
  /** Text that stays inside its own cell. */
  val none: OverflowSpan = OverflowSpan(0, 0)

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
   * Whether text may spill into the cell at `ref`: Excel lets it into a cell that holds no value,
   * whatever that cell's fill, borders, number format or font. Any value blocks — an empty string,
   * a formula whose result is `""` — and so does every cell of a merged range.
   */
  private def allowsOverflowInto(ref: ARef, sheet: Sheet): Boolean =
    val holdsNoValue = sheet.cells.get(ref).forall { c =>
      c.value match
        case CellValue.Empty => true
        case _ => false
    }
    holdsNoValue && sheet.getMergedRange(ref).isEmpty

  /**
   * The empty neighbour columns a cell's text spills into (Excel's text overflow): how far past
   * each edge of its own cell the text's clip box reaches. The renderers clip the text to that box
   * and nothing else — every cell under it keeps its own fill and borders, and the text is drawn
   * above them.
   *
   * Only TEXT spills. A number, date, logical or error never borrows a neighbour, however empty and
   * whatever the cell's alignment: Excel confines it to its own column and shows `####` when it
   * does not fit (GH-459, GH-500). Wrapped text wraps instead, a merged cell clips to its merge,
   * and a cell in a hidden column (`cellWidth` 0) spills nowhere: Excel never shows it. The
   * direction follows the alignment the renderers anchor by: left and General text spills right,
   * right-aligned text spills LEFT, centred text spills both ways while staying centred on its own
   * cell (clipped on a side whose neighbour blocks). A side stops at the first neighbour that holds
   * a value or is merged, at the render window's edge, or once it covers the text.
   *
   * @param cell
   *   the cell, with its style and rendered content resolved
   * @param cellRef
   *   the cell's reference
   * @param cellWidth
   *   the cell's width in pixels
   * @param colWidths
   *   the window's column widths, indexed from `startCol`
   * @param sheet
   *   the sheet holding the neighbours
   * @param startCol
   *   the window's first column (0-based)
   * @param endCol
   *   the window's last column (0-based)
   */
  def overflowSpan(
    cell: ResolvedCell,
    cellRef: ARef,
    cellWidth: Int,
    colWidths: IndexedSeq[Int],
    sheet: Sheet,
    startCol: Int,
    endCol: Int
  ): OverflowSpan =
    val style = cell.style
    val spills = cell.content.kind match
      case RenderedKind.Text => !style.exists(_.align.wrapText)
      case _ => false
    if !spills || cellWidth <= 0 || sheet.getMergedRange(cellRef).isDefined then OverflowSpan.none
    else
      // Size the span from what the renderers DRAW — the formatted text — never the raw value:
      // a rounding format must not claim neighbours for digits it never shows, and a widening
      // one must get the room its text needs (GH-502).
      val textWidth = measureContentWidth(cell.cell.value, cell.content, style.map(_.font))
      val (pastLeft, pastRight) = textOverhang(cell.align, style, cellWidth, textWidth)
      def side(step: Int, need: Double): Int =
        emptyNeighbours(cellRef, step, need, colWidths, sheet, startCol, endCol)
      OverflowSpan(side(-1, pastLeft), side(1, pastRight))

  /**
   * How far (px) text of `textWidth` reaches past the left and right edges of its cell, placed the
   * way SvgRenderer places it: left text starts `CellPaddingX` plus the indent in; right text ends
   * `CellPaddingX` in, but is clamped to start at the cell's left edge while it fits the cell;
   * centred text is centred half an indent right of the middle.
   */
  private def textOverhang(
    align: HAlign,
    style: Option[CellStyle],
    cellWidth: Int,
    textWidth: Int
  ): (Double, Double) =
    val indentPx = style.map(_.align.indent).getOrElse(0) * IndentPxPerLevel
    align match
      case HAlign.Left | HAlign.General =>
        (0.0, (CellPaddingX + indentPx + textWidth - cellWidth).toDouble)
      case HAlign.Right =>
        if textWidth <= cellWidth then (0.0, 0.0)
        else ((textWidth + CellPaddingX - cellWidth).toDouble, 0.0)
      case HAlign.Center | HAlign.CenterContinuous =>
        val middle = cellWidth / 2 + indentPx / 2
        (textWidth / 2.0 - middle, middle + textWidth / 2.0 - cellWidth)
      case _ => (0.0, 0.0)

  /**
   * How many consecutive neighbours from `cellRef` in direction `step` (-1 left, +1 right) the text
   * takes to cover `need` px: it stops at a neighbour that blocks, at the window's edge, or once
   * the columns taken cover the need.
   */
  private def emptyNeighbours(
    cellRef: ARef,
    step: Int,
    need: Double,
    colWidths: IndexedSeq[Int],
    sheet: Sheet,
    startCol: Int,
    endCol: Int
  ): Int =
    val row = cellRef.row.index0

    @scala.annotation.tailrec
    def loop(col: Int, taken: Int, covered: Int): Int =
      if covered >= need || col < startCol || col > endCol then taken
      else if !allowsOverflowInto(ARef.from0(col, row), sheet) then taken
      else loop(col + step, taken + 1, covered + colWidths.lift(col - startCol).getOrElse(0))

    if need <= 0 then 0 else loop(cellRef.col.index0 + step, 0, 0)

  /**
   * The rightward extent of a cell's text overflow as a colspan: 1 plus the empty neighbours its
   * text spills into on the RIGHT (see [[overflowSpan]], which also reports a leftward spill).
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
    1 + overflowSpan(cell, cellRef, cellWidth, colWidths, sheet, startCol, endCol).right

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
    (range.start.col.index0 to range.end.col.index0).map(columnWidthPx(sheet, _, scaleFactor))

  /** One column's width in pixels: 0 when hidden, else its own width or the sheet's default. */
  private def columnWidthPx(sheet: Sheet, colIdx: Int, scaleFactor: Double): Int =
    val props = sheet.getColumnProperties(Column.from0(colIdx))
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

  // ========== Row Heights (Excel autofit) ==========

  /**
   * Calculate row heights for a range. A row with an explicit height keeps it exactly and a hidden
   * row is 0; every other row takes the height its content needs, as Excel autofits it, and never
   * less than the sheet's default row height (see [[autofitRowHeights]]).
   */
  def calculateRowHeights(
    sheet: Sheet,
    range: CellRange,
    scaleFactor: Double = 1.0
  ): IndexedSeq[Int] =
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0
    val content = autofitRowHeights(sheet, startRow, endRow)
    (startRow to endRow).map(r => rowHeightPx(sheet, Row.from0(r), content.get(r), scaleFactor))

  /**
   * Get the height of a single row in pixels (see [[calculateRowHeights]]). Autofit reads the row's
   * cells, so a renderer sizes its window with one [[calculateRowHeights]] call instead.
   */
  def getRowHeight(sheet: Sheet, row: Row, scaleFactor: Double = 1.0): Int =
    val r = row.index0
    rowHeightPx(sheet, row, autofitRowHeights(sheet, r, r).get(r), scaleFactor)

  private def rowHeightPx(
    sheet: Sheet,
    row: Row,
    contentPx: Option[Int],
    scaleFactor: Double
  ): Int =
    val props = sheet.getRowProperties(row)
    if props.hidden then 0
    else
      val baseHeight = props.height
        .map(excelRowHeightToPixels)
        .getOrElse(math.max(defaultRowHeightPx(sheet), contentPx.getOrElse(0)))
      (baseHeight * scaleFactor).toInt

  private def defaultRowHeightPx(sheet: Sheet): Int =
    sheet.defaultRowHeight.map(excelRowHeightToPixels).getOrElse(DefaultCellHeightPx)

  /** The size of the book's base font: the style in registry slot 0 (Excel's Normal style). */
  private def baseFontPt(sheet: Sheet): Double =
    sheet.styleRegistry.get(StyleId(0)).fold(DefaultFontSize.toDouble)(_.font.sizePt)

  /**
   * Excel's autofit height, in whole pixels at 96 DPI, of one line of text at `sizePt`: the rounded
   * Calibri win ascent and descent (1950 and 550 of 2048 font units) plus a pixel above and below.
   * It reproduces Excel's Calibri autofit rows — 11pt 15pt (20px), 14pt 18.75pt (25px), 18pt
   * 23.25pt (31px), 24pt 31.5pt (42px) — and depends on no installed font, so a headless sandbox
   * renders the heights a desktop does. Other families are sized with Calibri's metrics.
   */
  private[render] def autofitLineHeightPx(sizePt: Double): Int =
    val px = sizePt * 4.0 / 3.0
    val ascent = math.ceil(px * 1950.0 / 2048.0 - 1e-9).toInt
    val descent = math.round(px * 550.0 / 2048.0).toInt
    ascent + descent + 2

  /**
   * The height one line of text at `sizePt` takes in `sheet`: [[autofitLineHeightPx]], except that
   * the book's base font never needs more than the default row the sheet stores — Excel sized that
   * row for it, and Calibri's metrics would grow every row of, say, an Arial 10 book by two pixels.
   * With no stored default the fallback row is Calibri 11's, which says nothing of the base font.
   */
  private[render] def lineHeightPx(sheet: Sheet, sizePt: Double): Int =
    val autofit = autofitLineHeightPx(sizePt)
    sheet.defaultRowHeight match
      case Some(stored) if sizePt == baseFontPt(sheet) =>
        math.min(autofit, excelRowHeightToPixels(stored))
      case _ => autofit

  /** The largest font a value draws in: its largest rich-text run, else the cell's own font. */
  private[render] def largestFontPt(value: CellValue, cellFontPt: Double): Double =
    value match
      case CellValue.RichText(rt) =>
        rt.runs.map(_.font.fold(cellFontPt)(_.sizePt)).maxOption.getOrElse(cellFontPt)
      case _ => cellFontPt

  /**
   * How far above a cell's bottom edge a bottom-aligned baseline sits: the line's descent (a
   * quarter of its em), and never less than 4px, so the descenders of a large font stay inside an
   * autofitted row.
   */
  private[render] def baselineInsetPx(sizePt: Double): Int =
    math.max(4, math.ceil(sizePt * 4.0 / 3.0 / 4.0 - 1e-9).toInt)

  /**
   * The height each row in `startRow..endRow` needs for its content, for the rows Excel autofits
   * (no explicit height, not hidden): the tallest cell's need, where a cell needs the line height
   * of its largest font (rich-text runs included) times its line count — one, or for wrapped text
   * the lines [[wrapText]] breaks it into at its column's width, exactly as SvgRenderer draws it.
   *
   * Every cell of a row counts, not only the rendered window's columns, in ONE pass over the
   * sheet's cells per call. Cells without a value, cells in a merged range (Excel's autofit ignores
   * merges) and cells in hidden columns do not count. A row with no counting cell is absent.
   */
  private[render] def autofitRowHeights(sheet: Sheet, startRow: Int, endRow: Int): Map[Int, Int] =
    val merges = sheet.mergedRanges.filter { m =>
      m.start.row.index0 <= endRow && m.end.row.index0 >= startRow
    }
    sheet.cells.valuesIterator.foldLeft(Map.empty[Int, Int]) { (acc, cell) =>
      val r = cell.ref.row.index0
      if r < startRow || r > endRow then acc
      else
        val props = sheet.getRowProperties(cell.ref.row)
        if props.height.isDefined || props.hidden || merges.exists(_.contains(cell.ref)) then acc
        else
          contentHeightPx(cell, sheet) match
            case Some(h) if h > acc.getOrElse(r, 0) => acc.updated(r, h)
            case _ => acc
    }

  /** The height a cell's content needs (see [[autofitRowHeights]]); None when it needs none. */
  private def contentHeightPx(cell: Cell, sheet: Sheet): Option[Int] =
    val width = columnWidthPx(sheet, cell.ref.col.index0, 1.0)
    cell.value match
      case CellValue.Empty => None
      case _ if width <= 0 => None
      case value =>
        val style = cell.styleId.flatMap(sheet.styleRegistry.get)
        val fontPt = largestFontPt(value, style.fold(baseFontPt(sheet))(_.font.sizePt))
        val lines = value match
          // SvgRenderer draws rich text on one line, wrapText or not
          case CellValue.RichText(_) => 1
          case _ if style.exists(_.align.wrapText) =>
            val text = renderedContent(value, style.fold(NumFmt.General)(_.numFmt)).text
            wrapText(text, width - CellPaddingX * 2, style.map(_.font)).size
          case _ => 1
        Some(lines * lineHeightPx(sheet, fontPt))

  // ========== Text Wrapping ==========

  /**
   * Break `text` into lines no wider than `maxWidth` px, at whitespace; a too-long word stands
   * alone.
   */
  private[render] def wrapText(text: String, maxWidth: Int, font: Option[Font]): List[String] =
    if maxWidth <= 0 || text.isEmpty then List(text)
    else
      val words = text.split("\\s+").toList
      if words.isEmpty then List("")
      else wrapWords(words, maxWidth, font)

  @scala.annotation.tailrec
  private def wrapWords(
    words: List[String],
    maxWidth: Int,
    font: Option[Font],
    lines: List[String] = Nil,
    currentLine: String = ""
  ): List[String] =
    words match
      case Nil =>
        if currentLine.isEmpty then lines.reverse
        else (currentLine :: lines).reverse
      case word :: rest =>
        val testLine = if currentLine.isEmpty then word else s"$currentLine $word"
        if measureTextWidth(testLine, font) <= maxWidth then
          wrapWords(rest, maxWidth, font, lines, testLine)
        else if currentLine.isEmpty then
          // Word is too long to fit on a line, force it
          wrapWords(rest, maxWidth, font, word :: lines, "")
        else
          // Start a new line with this word
          wrapWords(rest, maxWidth, font, currentLine :: lines, word)

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
