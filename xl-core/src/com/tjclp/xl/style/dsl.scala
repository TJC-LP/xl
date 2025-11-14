package com.tjclp.xl.style

import com.tjclp.xl.style.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.style.border.{Border, BorderStyle}
import com.tjclp.xl.style.color.Color
import com.tjclp.xl.style.fill.Fill
import com.tjclp.xl.style.font.Font
import com.tjclp.xl.style.numfmt.NumFmt

/**
 * Fluent DSL for building CellStyle instances with ergonomic shortcuts.
 *
 * Provides extension methods on CellStyle for common styling operations, enabling method chaining
 * without verbose constructors. All methods are inline for zero runtime overhead.
 *
 * Usage:
 * {{{
 *   import com.tjclp.xl.*
 *
 *   // Fluent builder with presets
 *   val style = CellStyle.default.bold.size(14.0).red.bgBlue.center
 *
 *   // Custom colors with RGB
 *   val brandStyle = CellStyle.default.rgb(68, 114, 196).bgRgb(240, 240, 240)
 *
 *   // Custom colors with hex
 *   val themeStyle = CellStyle.default.hex("#003366").bgHex("#F5F5F5")
 *
 *   // Prebuilt constants
 *   val header = Style.header  // bold, 14pt, centered, blue bg, white text
 * }}}
 */
object dsl:

  extension (style: CellStyle)
    // ========== Font Styling ==========

    /** Set font to bold */
    inline def bold: CellStyle =
      style.withFont(style.font.withBold(true))

    /** Set font to italic */
    inline def italic: CellStyle =
      style.withFont(style.font.withItalic(true))

    /** Set font to underline */
    inline def underline: CellStyle =
      style.withFont(style.font.withUnderline(true))

    /** Set font size in points */
    inline def size(pt: Double): CellStyle =
      style.withFont(style.font.withSize(pt))

    /** Set font family name */
    inline def fontFamily(name: String): CellStyle =
      style.withFont(style.font.withName(name))

    // ========== Preset Font Colors ==========

    /** Set font color to red (255, 0, 0) */
    inline def red: CellStyle =
      style.withFont(style.font.withColor(Color.fromRgb(255, 0, 0)))

    /** Set font color to green (0, 255, 0) */
    inline def green: CellStyle =
      style.withFont(style.font.withColor(Color.fromRgb(0, 255, 0)))

    /** Set font color to blue (0, 0, 255) */
    inline def blue: CellStyle =
      style.withFont(style.font.withColor(Color.fromRgb(0, 0, 255)))

    /** Set font color to white (255, 255, 255) */
    inline def white: CellStyle =
      style.withFont(style.font.withColor(Color.fromRgb(255, 255, 255)))

    /** Set font color to black (0, 0, 0) */
    inline def black: CellStyle =
      style.withFont(style.font.withColor(Color.fromRgb(0, 0, 0)))

    /** Set font color to yellow (255, 255, 0) */
    inline def yellow: CellStyle =
      style.withFont(style.font.withColor(Color.fromRgb(255, 255, 0)))

    // ========== Custom Font Colors ==========

    /** Set font color from RGB components (0-255) */
    inline def rgb(r: Int, g: Int, b: Int): CellStyle =
      style.withFont(style.font.withColor(Color.fromRgb(r, g, b)))

    /** Set font color from hex code (#RRGGBB or #AARRGGBB). Invalid codes are silently ignored. */
    inline def hex(code: String): CellStyle =
      Color.fromHex(code) match
        case Right(c) => style.withFont(style.font.withColor(c))
        case Left(_) => style // Silently ignore invalid hex

    // ========== Preset Background Colors ==========

    /** Set background to red (255, 0, 0) */
    inline def bgRed: CellStyle =
      style.withFill(Fill.Solid(Color.fromRgb(255, 0, 0)))

    /** Set background to green (0, 255, 0) */
    inline def bgGreen: CellStyle =
      style.withFill(Fill.Solid(Color.fromRgb(0, 255, 0)))

    /** Set background to blue (68, 114, 196) - Excel default blue */
    inline def bgBlue: CellStyle =
      style.withFill(Fill.Solid(Color.fromRgb(68, 114, 196)))

    /** Set background to yellow (255, 255, 0) */
    inline def bgYellow: CellStyle =
      style.withFill(Fill.Solid(Color.fromRgb(255, 255, 0)))

    /** Set background to white (255, 255, 255) */
    inline def bgWhite: CellStyle =
      style.withFill(Fill.Solid(Color.fromRgb(255, 255, 255)))

    /** Set background to light gray (200, 200, 200) */
    inline def bgGray: CellStyle =
      style.withFill(Fill.Solid(Color.fromRgb(200, 200, 200)))

    /** Set background to no fill */
    inline def bgNone: CellStyle =
      style.withFill(Fill.None)

    // ========== Custom Background Colors ==========

    /** Set background from RGB components (0-255) */
    inline def bgRgb(r: Int, g: Int, b: Int): CellStyle =
      style.withFill(Fill.Solid(Color.fromRgb(r, g, b)))

    /** Set background from hex code (#RRGGBB or #AARRGGBB). Invalid codes are silently ignored. */
    inline def bgHex(code: String): CellStyle =
      Color.fromHex(code) match
        case Right(c) => style.withFill(Fill.Solid(c))
        case Left(_) => style

    // ========== Alignment ==========

    /** Set horizontal alignment to center */
    inline def center: CellStyle =
      style.withAlign(style.align.withHAlign(HAlign.Center))

    /** Set horizontal alignment to left */
    inline def left: CellStyle =
      style.withAlign(style.align.withHAlign(HAlign.Left))

    /** Set horizontal alignment to right */
    inline def right: CellStyle =
      style.withAlign(style.align.withHAlign(HAlign.Right))

    /** Set vertical alignment to top */
    inline def top: CellStyle =
      style.withAlign(style.align.withVAlign(VAlign.Top))

    /** Set vertical alignment to middle */
    inline def middle: CellStyle =
      style.withAlign(style.align.withVAlign(VAlign.Middle))

    /** Set vertical alignment to bottom */
    inline def bottom: CellStyle =
      style.withAlign(style.align.withVAlign(VAlign.Bottom))

    /** Enable text wrapping */
    inline def wrap: CellStyle =
      style.withAlign(style.align.withWrap(true))

    // ========== Borders ==========

    /** Set all borders to thin */
    inline def bordered: CellStyle =
      style.withBorder(Border.all(BorderStyle.Thin))

    /** Set all borders to medium */
    inline def borderedMedium: CellStyle =
      style.withBorder(Border.all(BorderStyle.Medium))

    /** Set all borders to thick */
    inline def borderedThick: CellStyle =
      style.withBorder(Border.all(BorderStyle.Thick))

    /** Remove all borders */
    inline def borderNone: CellStyle =
      style.withBorder(Border.none)

    // ========== Number Formats ==========

    /** Set number format to currency ($#,##0.00) */
    inline def currency: CellStyle =
      style.withNumFmt(NumFmt.Currency)

    /** Set number format to percent (0.00%) */
    inline def percent: CellStyle =
      style.withNumFmt(NumFmt.Percent)

    /** Set number format to decimal (#,##0.00) */
    inline def decimal: CellStyle =
      style.withNumFmt(NumFmt.Decimal)

    /** Set number format to date (m/d/yyyy) */
    inline def dateFormat: CellStyle =
      style.withNumFmt(NumFmt.Date)

    /** Set number format to datetime (m/d/yyyy h:mm) */
    inline def dateTime: CellStyle =
      style.withNumFmt(NumFmt.DateTime)

  // ========== Prebuilt Style Constants ==========

  /**
   * Common style presets for quick access.
   *
   * These constants provide starting points for frequently-used styles that can be further
   * customized via chaining.
   */
  object Style:
    /** Bold, 14pt, centered, blue background, white text - typical table header */
    val header: CellStyle =
      CellStyle.default.bold.size(14.0).center.middle.bgBlue.white

    /** Currency format, right-aligned - typical for financial data */
    val currencyCell: CellStyle =
      CellStyle.default.currency.right

    /** Percent format, right-aligned - typical for ratios and rates */
    val percentCell: CellStyle =
      CellStyle.default.percent.right

    /** Date format, centered - typical for date columns */
    val dateCell: CellStyle =
      CellStyle.default.dateFormat.center

    /** Red, bold, italic - typical for error messages */
    val errorText: CellStyle =
      CellStyle.default.red.bold.italic

    /** Green, bold - typical for success messages */
    val successText: CellStyle =
      CellStyle.default.green.bold

    /** Bold, centered - simple column header */
    val columnHeader: CellStyle =
      CellStyle.default.bold.center

export dsl.*
