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
     * @return
     *   HTML string
     */
    @annotation.targetName("toHtmlWithTheme")
    def toHtml(
      range: CellRange,
      includeStyles: Boolean = true,
      includeComments: Boolean = true,
      theme: ThemePalette = ThemePalette.office
    ): String =
      HtmlRenderer.toHtml(sheet, range, includeStyles, includeComments, theme)

    /**
     * Export sheet range to SVG.
     *
     * @param range
     *   Cell range to export
     * @param includeStyles
     *   Include cell styling (colors, fonts, borders) (default: true)
     * @param theme
     *   Theme palette for resolving theme colors (default: Office theme)
     * @return
     *   SVG string
     */
    @annotation.targetName("toSvgWithTheme")
    def toSvg(
      range: CellRange,
      includeStyles: Boolean = true,
      theme: ThemePalette = ThemePalette.office
    ): String =
      SvgRenderer.toSvg(sheet, range, includeStyles, theme)
