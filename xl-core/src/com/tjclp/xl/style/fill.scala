package com.tjclp.xl.style

/**
 * Cell background fill patterns and colors.
 */

// ========== Fill ==========

/** Fill pattern types */
enum PatternType:
  case None, Solid, Gray125, Gray0625
  case DarkGray, MediumGray, LightGray
  case DarkHorizontal, DarkVertical, DarkDown, DarkUp
  case DarkGrid, DarkTrellis
  case LightHorizontal, LightVertical, LightDown, LightUp
  case LightGrid, LightTrellis

/** Cell background fill */
enum Fill:
  case None
  case Solid(color: Color)
  case Pattern(foreground: Color, background: Color, pattern: PatternType)

object Fill:
  val default: Fill = None
