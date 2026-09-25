package com.tjclp.xl.render

import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cf.CfOverlay
import com.tjclp.xl.sheets.Sheet
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

    // GH-497: the conditional-formatting overloads take the overlay explicitly and carry NO
    // default arguments — this block is wildcard-exported (api.*, the scripting prelude), where a
    // second defaulted alternative would crash the compiler. The overlay comes from xl-evaluator:
    // `sheet.toSvg(range, sheet.conditionalFormatOverlay(range))`.

    /** [[toHtml]] with conditional formatting painted from `overlay` (styles, comments on). */
    @annotation.targetName("toHtmlWithOverlay")
    def toHtml(range: CellRange, overlay: CfOverlay): String =
      HtmlRenderer.toHtml(sheet, range, true, true, ThemePalette.office, false, false, overlay)

    /** [[toHtml]] with every option explicit and conditional formatting painted from `overlay`. */
    @annotation.targetName("toHtmlWithThemeAndOverlay")
    def toHtml(
      range: CellRange,
      includeStyles: Boolean,
      includeComments: Boolean,
      theme: ThemePalette,
      applyPrintScale: Boolean,
      showLabels: Boolean,
      overlay: CfOverlay
    ): String =
      HtmlRenderer.toHtml(
        sheet,
        range,
        includeStyles,
        includeComments,
        theme,
        applyPrintScale,
        showLabels,
        overlay
      )

    /** [[toSvg]] with conditional formatting painted from `overlay` (styles on). */
    @annotation.targetName("toSvgWithOverlay")
    def toSvg(range: CellRange, overlay: CfOverlay): String =
      SvgRenderer.toSvg(sheet, range, true, ThemePalette.office, false, false, overlay)

    /** [[toSvg]] with every option explicit and conditional formatting painted from `overlay`. */
    @annotation.targetName("toSvgWithThemeAndOverlay")
    def toSvg(
      range: CellRange,
      includeStyles: Boolean,
      theme: ThemePalette,
      showLabels: Boolean,
      showGridlines: Boolean,
      overlay: CfOverlay
    ): String =
      SvgRenderer.toSvg(sheet, range, includeStyles, theme, showLabels, showGridlines, overlay)
