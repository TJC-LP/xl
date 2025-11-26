package com.tjclp.xl.styles.color

/**
 * Workbook theme color palette containing the 12 standard Office theme colors.
 *
 * Each color is stored as ARGB integer (0xAARRGGBB format). Theme colors can be modified by a tint
 * value (-1.0 to 1.0) when applied to cells.
 *
 * @see
 *   ECMA-376 Part 1, §20.1.6.2 (clrScheme)
 */
final case class ThemePalette(
  dark1: Int,
  light1: Int,
  dark2: Int,
  light2: Int,
  accent1: Int,
  accent2: Int,
  accent3: Int,
  accent4: Int,
  accent5: Int,
  accent6: Int,
  hyperlink: Int = 0xff0563c1,
  followedHyperlink: Int = 0xff954f72
)

object ThemePalette:

  /** Default Office theme (Office 2007-2019 standard colors) */
  val office: ThemePalette = ThemePalette(
    dark1 = 0xff000000, // Black
    light1 = 0xffffffff, // White
    dark2 = 0xff1f497d, // Dark blue
    light2 = 0xffeeece1, // Tan/cream
    accent1 = 0xff4f81bd, // Blue
    accent2 = 0xffc0504d, // Red
    accent3 = 0xff9bbb59, // Green
    accent4 = 0xff8064a2, // Purple
    accent5 = 0xff4bacc6, // Teal
    accent6 = 0xfff79646 // Orange
  )

  /**
   * Resolve a theme color slot to ARGB with tint applied.
   *
   * @param palette
   *   The theme palette containing base colors
   * @param slot
   *   The theme slot to resolve
   * @param tint
   *   Tint modifier (-1.0 to 1.0): positive lightens, negative darkens
   * @return
   *   ARGB integer with tint applied
   */
  def resolve(palette: ThemePalette, slot: ThemeSlot, tint: Double): Int =
    val baseColor = slot match
      case ThemeSlot.Dark1 => palette.dark1
      case ThemeSlot.Light1 => palette.light1
      case ThemeSlot.Dark2 => palette.dark2
      case ThemeSlot.Light2 => palette.light2
      case ThemeSlot.Accent1 => palette.accent1
      case ThemeSlot.Accent2 => palette.accent2
      case ThemeSlot.Accent3 => palette.accent3
      case ThemeSlot.Accent4 => palette.accent4
      case ThemeSlot.Accent5 => palette.accent5
      case ThemeSlot.Accent6 => palette.accent6
    applyTint(baseColor, tint)

  /**
   * Apply tint to an ARGB color per ECMA-376 §20.1.2.3.40.
   *
   *   - tint > 0: Blend toward white (lighten)
   *   - tint < 0: Blend toward black (darken)
   *   - tint = 0: No change
   */
  def applyTint(argb: Int, tint: Double): Int =
    if tint == 0.0 then argb
    else
      val a = (argb >> 24) & 0xff
      val r = applyTintComponent((argb >> 16) & 0xff, tint)
      val g = applyTintComponent((argb >> 8) & 0xff, tint)
      val b = applyTintComponent(argb & 0xff, tint)
      (a << 24) | (r << 16) | (g << 8) | b

  private def applyTintComponent(value: Int, tint: Double): Int =
    val result =
      if tint > 0 then value + (255 - value) * tint
      else value * (1 + tint)
    math.max(0, math.min(255, result.toInt))
