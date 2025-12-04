package com.tjclp.xl.cli

import com.tjclp.xl.styles.color.Color

/**
 * Color parser for CLI input.
 *
 * Supports:
 *   - Named colors: red, blue, yellow, green, white, black, etc.
 *   - Hex: #RGB, #RRGGBB, #AARRGGBB
 *   - RGB: rgb(r,g,b)
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

  private val rgbPattern = """rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)""".r
  private val shortHexPattern = """#([0-9A-Fa-f]{3})""".r

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
              s"Unknown color: $s. Use named (red, blue, ...), hex (#RRGGBB), or rgb(r,g,b)"
            )

  /** List available named colors */
  def availableNames: List[String] = namedColors.keys.toList.sorted
