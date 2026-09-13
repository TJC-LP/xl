package com.tjclp.xl.styles.fill

import com.tjclp.xl.styles.color.Color

/**
 * Cell fill (ECMA-376 Part 1, §18.8.32 `CT_PatternFill`).
 *
 * `Pattern` mirrors the schema: `fgColor` and `bgColor` are both optional children, so a texture
 * may carry either colour, both or neither. An absent colour is Excel's "automatic" colour —
 * written as no child at all, as `auto="1"`, or as the system palette indices 64/65 — which Excel
 * renders as the default foreground (black hatch) over the window background. Requiring both
 * colours made every openpyxl- or Excel-authored hatch unrepresentable and dropped it on write
 * (GH-566). `Solid` keeps its mandatory colour: a solid fill without one is no fill.
 */
enum Fill derives CanEqual:
  case None
  case Solid(color: Color)
  case Pattern(foreground: Option[Color], background: Option[Color], pattern: PatternType)

object Fill:
  val default: Fill = None

  /** A texture with both colours given — the common authoring case. */
  def pattern(foreground: Color, background: Color, pattern: PatternType): Fill =
    Pattern(Some(foreground), Some(background), pattern)
