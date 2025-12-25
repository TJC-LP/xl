package com.tjclp.xl.ooxml.style

import com.tjclp.xl.styles.color.{Color, ThemeSlot}

/** Color resolution helpers for OOXML style parsing */
private[ooxml] object ColorHelpers:

  // OOXML theme color indices (ECMA-376 Part 1, 18.8.3):
  // 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4-9=accent1-6, 10=hlink, 11=folHlink
  def themeSlotFromIndex(idx: Int): Option[ThemeSlot] = idx match
    case 0 => Some(ThemeSlot.Light1)
    case 1 => Some(ThemeSlot.Dark1)
    case 2 => Some(ThemeSlot.Light2)
    case 3 => Some(ThemeSlot.Dark2)
    case 4 => Some(ThemeSlot.Accent1)
    case 5 => Some(ThemeSlot.Accent2)
    case 6 => Some(ThemeSlot.Accent3)
    case 7 => Some(ThemeSlot.Accent4)
    case 8 => Some(ThemeSlot.Accent5)
    case 9 => Some(ThemeSlot.Accent6)
    case _ => None

  // Standard Excel indexed color palette (0-63)
  // Based on ECMA-376 Part 1, 18.8.27 and legacy BIFF8 color table
  private val indexedColors: Vector[Int] = Vector(
    0x000000, // 0 - Black
    0xffffff, // 1 - White
    0xff0000, // 2 - Red
    0x00ff00, // 3 - Bright Green
    0x0000ff, // 4 - Blue
    0xffff00, // 5 - Yellow
    0xff00ff, // 6 - Pink
    0x00ffff, // 7 - Turquoise
    0x000000, // 8 - Black
    0xffffff, // 9 - White
    0xff0000, // 10 - Red
    0x00ff00, // 11 - Bright Green
    0x0000ff, // 12 - Blue
    0xffff00, // 13 - Yellow
    0xff00ff, // 14 - Pink
    0x00ffff, // 15 - Turquoise
    0x800000, // 16 - Dark Red
    0x008000, // 17 - Green
    0x000080, // 18 - Dark Blue (Navy)
    0x808000, // 19 - Dark Yellow (Olive)
    0x800080, // 20 - Purple
    0x008080, // 21 - Teal
    0xc0c0c0, // 22 - Silver
    0x808080, // 23 - Gray
    0x9999ff, // 24 - Periwinkle
    0x993366, // 25 - Plum
    0xffffcc, // 26 - Ivory
    0xccffff, // 27 - Light Turquoise
    0x660066, // 28 - Dark Purple
    0xff8080, // 29 - Coral
    0x0066cc, // 30 - Ocean Blue
    0xccccff, // 31 - Ice Blue
    0x000080, // 32 - Dark Blue
    0xff00ff, // 33 - Pink
    0xffff00, // 34 - Yellow
    0x00ffff, // 35 - Turquoise
    0x800080, // 36 - Violet
    0x800000, // 37 - Dark Red
    0x008080, // 38 - Teal
    0x0000ff, // 39 - Blue
    0x00ccff, // 40 - Sky Blue
    0xccffff, // 41 - Light Turquoise
    0xccffcc, // 42 - Light Green
    0xffff99, // 43 - Light Yellow
    0x99ccff, // 44 - Pale Blue
    0xff99cc, // 45 - Rose
    0xcc99ff, // 46 - Lavender
    0xffcc99, // 47 - Tan
    0x3366ff, // 48 - Light Blue
    0x33cccc, // 49 - Aqua
    0x99cc00, // 50 - Lime
    0xffcc00, // 51 - Gold
    0xff9900, // 52 - Light Orange
    0xff6600, // 53 - Orange
    0x666699, // 54 - Blue-Gray
    0x969696, // 55 - Gray 40%
    0x003366, // 56 - Dark Teal
    0x339966, // 57 - Sea Green
    0x003300, // 58 - Dark Green
    0x333300, // 59 - Olive Green
    0x993300, // 60 - Brown
    0x993366, // 61 - Plum
    0x333399, // 62 - Indigo
    0x333333 // 63 - Gray 80%
  )

  /** Convert indexed color (0-63) to RGB Color */
  def indexedColorToRgb(idx: Int): Option[Color] =
    if idx >= 0 && idx < indexedColors.length then
      // ARGB format with full opacity
      Some(Color.Rgb(0xff000000 | indexedColors(idx)))
    else None // Indices 64+ are custom colors defined in workbook, not supported yet
