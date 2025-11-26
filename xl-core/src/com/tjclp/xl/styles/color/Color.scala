package com.tjclp.xl.styles.color

/**
 * Color representation for Excel cell styling.
 *
 * Excel supports both RGB colors and theme-based colors with tints.
 */

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
    /** Convert to ARGB integer. For theme colors, use toResolvedArgb. */
    def toArgb: Int = color match
      case Rgb(argb) => argb
      case Theme(slot, tint) =>
        throw IllegalStateException(
          s"Theme color $slot (tint=$tint) requires resolution via toResolvedArgb(theme)"
        )

    /** Resolve to ARGB integer using the provided theme palette */
    def toResolvedArgb(theme: ThemePalette): Int = color match
      case Rgb(argb) => argb
      case Theme(slot, tint) => ThemePalette.resolve(theme, slot, tint)

    /** Get hex string (#AARRGGBB). For theme colors, use toResolvedHex. */
    def toHex: String = color match
      case Rgb(argb) => f"#$argb%08X"
      case Theme(slot, tint) =>
        throw IllegalStateException(
          s"Theme color $slot (tint=$tint) requires resolution via toResolvedHex(theme)"
        )

    /** Resolve to RGB hex string (#RRGGBB) using the provided theme palette */
    def toResolvedHex(theme: ThemePalette): String = color match
      case Rgb(argb) => f"#${argb & 0xffffff}%06X"
      case Theme(slot, tint) =>
        val resolved = ThemePalette.resolve(theme, slot, tint)
        f"#${resolved & 0xffffff}%06X"
