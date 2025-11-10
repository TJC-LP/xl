package com.tjclp.xl.style

/**
 * Color representation for Excel cell styling.
 *
 * Excel supports both RGB colors and theme-based colors with tints.
 */

// ========== Colors ==========

/** Theme color slots from Office theme */
enum ThemeSlot:
  case Dark1, Light1, Dark2, Light2
  case Accent1, Accent2, Accent3, Accent4, Accent5, Accent6

/** Color representation: either RGB or theme-based with tint */
enum Color:
  /** RGB color with alpha channel (ARGB format: 0xAARRGGBB) */
  case Rgb(argb: Int)

  /** Theme color with optional tint (-1.0 to 1.0, where 0 is theme default) */
  case Theme(slot: ThemeSlot, tint: Double)

object Color:
  /** Create RGB color from components (0-255) */
  def fromRgb(r: Int, g: Int, b: Int, a: Int = 255): Color =
    Rgb((a << 24) | (r << 16) | (g << 8) | b)

  /** Create color from hex string (#RRGGBB or #AARRGGBB) */
  def fromHex(hex: String): Either[String, Color] =
    val cleaned = if hex.startsWith("#") then hex.substring(1) else hex
    cleaned.length match
      case 6 => // #RRGGBB
        try Right(Rgb(0xff000000 | Integer.parseInt(cleaned, 16)))
        catch case _: NumberFormatException => Left(s"Invalid hex color: $hex")
      case 8 => // #AARRGGBB
        try Right(Rgb(Integer.parseUnsignedInt(cleaned, 16)))
        catch case _: NumberFormatException => Left(s"Invalid hex color: $hex")
      case _ => Left(s"Invalid hex color length: $hex")

  /** Validate tint is in valid range */
  def validTint(tint: Double): Either[String, Double] =
    if tint >= -1.0 && tint <= 1.0 then Right(tint)
    else Left(s"Tint must be in [-1.0, 1.0], got: $tint")

  extension (color: Color)
    /** Convert to ARGB integer */
    def toArgb: Int = color match
      case Rgb(argb) => argb
      case Theme(_, _) => 0 // Needs theme resolution at IO boundary

    /** Get hex string representation (#AARRGGBB) */
    def toHex: String = color match
      case Rgb(argb) => f"#$argb%08X"
      case Theme(slot, tint) => s"Theme($slot, $tint)"
