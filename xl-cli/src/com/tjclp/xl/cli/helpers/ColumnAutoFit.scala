package com.tjclp.xl.cli.helpers

import scala.util.Try

import com.tjclp.xl.addressing.Column
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.render.RenderUtils
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Column auto-fit width calculation shared by `col --auto-fit`, `autofit`, and batch `autofit`
 * (GH-156).
 *
 * Each cell's formatted display text is measured with its resolved font via AWT font metrics
 * (`RenderUtils.measureTextWidth`), then converted from pixels to Excel column-width units using
 * the OOXML/POI convention for the Calibri-11 default: max digit width (MDW) = 7px at 96 DPI plus
 * 5px of cell padding (2px left margin + 2px right margin + 1px gridline), i.e.
 * `width = (textPx + 5) / 7`. `RenderUtils.toAwtFont` sizes fonts at pt × 4/3 px — the
 * 96-DPI-equivalent space in which MDW = 7 holds.
 *
 * Note: `RenderUtils.excelColWidthToPixels` deliberately uses a wider 8px/unit factor for SVG
 * display fidelity (so rendered text never clips); auto-fit needs the Excel-accurate MDW = 7 so
 * produced widths match Excel's own autofit.
 *
 * Fallbacks: `RenderUtils.measureTextWidth` already estimates from character counts (scaled by font
 * size and bold) when no graphics context can be created, so headless CI still measures. If
 * measurement throws anyway (exotic environments without fontconfig), the pre-0.12 char-count
 * heuristic (`chars × 0.90 + 1.5`) is used per cell.
 */
object ColumnAutoFit:

  /** Excel default column width in character units, used for empty columns. */
  val DefaultColumnWidth: Double = 8.43

  /** Lower bound on auto-fit width (pre-existing CLI floor; Excel allows narrow columns). */
  private val MinWidth: Double = 5.0

  /** Max digit width of Calibri 11 at 96 DPI (OOXML/POI column-width convention). */
  private val MaxDigitWidthPx: Double = 7.0

  /** Cell padding in px: 2 left margin + 2 right margin + 1 gridline (OOXML/POI convention). */
  private val CellPaddingPx: Double = 5.0

  /** Calculate the optimal width for a column in Excel character units. */
  def calculateWidth(sheet: Sheet, col: Column): Double =
    val cellsInColumn = sheet.cells.filter { case (ref, _) => ref.col == col }
    if cellsInColumn.isEmpty then DefaultColumnWidth
    else
      val maxWidth = cellsInColumn.values.map(cellWidth(_, sheet)).maxOption.getOrElse(0.0)
      val rounded = BigDecimal(maxWidth).setScale(2, BigDecimal.RoundingMode.HALF_UP).toDouble
      math.max(rounded, MinWidth)

  /** Width one cell's formatted content needs, in Excel character units. */
  private def cellWidth(cell: Cell, sheet: Sheet): Double =
    val styleOpt = cell.styleId.flatMap(sheet.styleRegistry.get)
    val numFmt = styleOpt.map(_.numFmt).getOrElse(NumFmt.General)
    val text = RenderUtils.cellValueToText(cell.value, numFmt)
    if text.isEmpty then 0.0
    else
      Try(RenderUtils.measureTextWidth(text, styleOpt.map(_.font))).fold(
        _ => heuristicWidth(text, styleOpt),
        px => (px + CellPaddingPx) / MaxDigitWidthPx
      )

  /** Pre-GH-156 char-count heuristic, kept as the missing-font fallback. */
  private def heuristicWidth(text: String, styleOpt: Option[CellStyle]): Double =
    val boldFactor = if styleOpt.exists(_.font.bold) then 1.1 else 1.0
    text.length * boldFactor * 0.90 + 1.5
