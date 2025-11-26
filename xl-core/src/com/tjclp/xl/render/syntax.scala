package com.tjclp.xl.render

import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.styles.color.ThemePalette

/**
 * Extension methods for exporting sheets to HTML and SVG.
 *
 * These extensions add theme-aware rendering capabilities to Sheet.
 */
object syntax:

  extension (sheet: Sheet)

    /**
     * Export sheet range to HTML table.
     *
     * @param range
     *   Cell range to export
     * @param includeStyles
     *   Include inline CSS for cell formatting (default: true)
     * @param includeComments
     *   Include comments as HTML tooltips (default: true)
     * @param theme
     *   Theme palette for resolving theme colors (default: Office theme)
     * @param applyPrintScale
     *   Apply print scaling from pageSetup (default: false)
     * @param showLabels
     *   Show column letters (A, B, C...) and row numbers (1, 2, 3...) (default: false)
     * @return
     *   HTML string
     */
    @annotation.targetName("toHtmlWithTheme")
    def toHtml(
      range: CellRange,
      includeStyles: Boolean = true,
      includeComments: Boolean = true,
      theme: ThemePalette = ThemePalette.office,
      applyPrintScale: Boolean = false,
      showLabels: Boolean = false
    ): String =
      HtmlRenderer.toHtml(
        sheet,
        range,
        includeStyles,
        includeComments,
        theme,
        applyPrintScale,
        showLabels
      )

    /**
     * Export sheet range to SVG.
     *
     * @param range
     *   Cell range to export
     * @param includeStyles
     *   Include cell styling (colors, fonts, borders) (default: true)
     * @param theme
     *   Theme palette for resolving theme colors (default: Office theme)
     * @param showLabels
     *   Show column letters (A, B, C...) and row numbers (1, 2, 3...) (default: false)
     * @param showGridlines
     *   Show cell gridlines (default: false, matches HTML behavior)
     * @return
     *   SVG string
     */
    @annotation.targetName("toSvgWithTheme")
    def toSvg(
      range: CellRange,
      includeStyles: Boolean = true,
      theme: ThemePalette = ThemePalette.office,
      showLabels: Boolean = false,
      showGridlines: Boolean = false
    ): String =
      SvgRenderer.toSvg(sheet, range, includeStyles, theme, showLabels, showGridlines)
