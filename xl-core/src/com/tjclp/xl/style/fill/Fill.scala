package com.tjclp.xl.style.fill

import com.tjclp.xl.style.color.Color

/** Cell background fill */
enum Fill:
  case None
  case Solid(color: Color)
  case Pattern(foreground: Color, background: Color, pattern: PatternType)

object Fill:
  val default: Fill = None
