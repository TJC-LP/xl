package com.tjclp.xl.cli

import com.tjclp.xl.styles.color.{Color, ThemeSlot}

/**
 * Color parser for CLI input.
 *
 * Supports:
 *   - Named colors: red, blue, yellow, green, white, black, etc.
 *   - Hex: #RGB, #RRGGBB, #AARRGGBB
 *   - RGB: rgb(r,g,b)
 *   - Theme: theme:accent1 or theme:accent1:0.25 (slot + optional tint in [-1.0, 1.0])
 */
object ColorParser:

  /** Named color mappings (CSS-like) */
  private val namedColors: Map[String, Color] = Map(
    "black" -> Color.fromRgb(0, 0, 0),
    "white" -> Color.fromRgb(255, 255, 255),
    "red" -> Color.fromRgb(255, 0, 0),
    "green" -> Color.fromRgb(0, 128, 0),
    "blue" -> Color.fromRgb(0, 0, 255),
    "yellow" -> Color.fromRgb(255, 255, 0),
    "orange" -> Color.fromRgb(255, 165, 0),
    "purple" -> Color.fromRgb(128, 0, 128),
    "pink" -> Color.fromRgb(255, 192, 203),
    "cyan" -> Color.fromRgb(0, 255, 255),
    "magenta" -> Color.fromRgb(255, 0, 255),
    "gray" -> Color.fromRgb(128, 128, 128),
    "grey" -> Color.fromRgb(128, 128, 128),
    "lightgray" -> Color.fromRgb(211, 211, 211),
    "lightgrey" -> Color.fromRgb(211, 211, 211),
    "darkgray" -> Color.fromRgb(169, 169, 169),
    "darkgrey" -> Color.fromRgb(169, 169, 169),
    "brown" -> Color.fromRgb(139, 69, 19),
    "navy" -> Color.fromRgb(0, 0, 128),
    "teal" -> Color.fromRgb(0, 128, 128),
    "olive" -> Color.fromRgb(128, 128, 0),
    "maroon" -> Color.fromRgb(128, 0, 0),
    "silver" -> Color.fromRgb(192, 192, 192),
    "gold" -> Color.fromRgb(255, 215, 0),
    "lime" -> Color.fromRgb(0, 255, 0)
  )

  /** Theme slot names for the `theme:<slot>[:<tint>]` syntax (GH-358). */
  private val themeSlots: Map[String, ThemeSlot] = Map(
    "dark1" -> ThemeSlot.Dark1,
    "light1" -> ThemeSlot.Light1,
    "dark2" -> ThemeSlot.Dark2,
    "light2" -> ThemeSlot.Light2,
    "accent1" -> ThemeSlot.Accent1,
    "accent2" -> ThemeSlot.Accent2,
    "accent3" -> ThemeSlot.Accent3,
    "accent4" -> ThemeSlot.Accent4,
    "accent5" -> ThemeSlot.Accent5,
    "accent6" -> ThemeSlot.Accent6
  )

  private val rgbPattern = """rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)""".r
  private val shortHexPattern = """#([0-9A-Fa-f]{3})""".r

  /** Parse `theme:<slot>[:<tint>]`, e.g. `theme:accent1` or `theme:accent2:0.25`. */
  private def parseTheme(input: String): Either[String, Color] =
    input.split(":", -1).toList match
      case "theme" :: slotName :: rest if rest.length <= 1 =>
        for
          slot <- themeSlots
            .get(slotName)
            .toRight(
              s"Unknown theme slot: $slotName. Use ${themeSlots.keys.toList.sorted.mkString(", ")}"
            )
          tint <- rest.headOption match
            case None => Right(0.0)
            case Some(t) =>
              t.toDoubleOption
                .toRight(s"Invalid theme tint: $t (expected a number in [-1.0, 1.0])")
                .flatMap(Color.validTint)
        yield Color.Theme(slot, tint)
      case _ =>
        Left(s"Invalid theme color: $input. Use theme:<slot>[:<tint>], e.g. theme:accent1:0.25")

  /**
   * Parse a color string.
   *
   * @param s
   *   Color string (name, #hex, or rgb(r,g,b))
   * @return
   *   Either error message or parsed Color
   */
  def parse(s: String): Either[String, Color] =
    val input = s.trim.toLowerCase
    namedColors.get(input) match
      case Some(color) => Right(color)
      case None =>
        input match
          // theme:<slot>[:<tint>] format (GH-358)
          case theme if theme.startsWith("theme:") =>
            parseTheme(theme)

          // rgb(r,g,b) format
          case rgbPattern(r, g, b) =>
            try
              val ri = r.toInt
              val gi = g.toInt
              val bi = b.toInt
              if ri < 0 || ri > 255 || gi < 0 || gi > 255 || bi < 0 || bi > 255 then
                Left(s"RGB values must be 0-255: $s")
              else Right(Color.fromRgb(ri, gi, bi))
            catch case _: NumberFormatException => Left(s"Invalid RGB values: $s")

          // Short hex #RGB -> #RRGGBB
          case shortHexPattern(hex) =>
            val expanded = hex.flatMap(c => s"$c$c")
            Color.fromHex(s"#$expanded")

          // Standard hex #RRGGBB or #AARRGGBB
          case hex if hex.startsWith("#") =>
            Color.fromHex(hex)

          case _ =>
            Left(
              s"Unknown color: $s. Use named (red, blue, ...), hex (#RRGGBB), rgb(r,g,b), " +
                "or theme:<slot>[:<tint>] (e.g. theme:accent1:0.25)"
            )

  /** List available named colors */
  def availableNames: List[String] = namedColors.keys.toList.sorted
