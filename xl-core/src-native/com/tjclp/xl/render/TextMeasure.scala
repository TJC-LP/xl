package com.tjclp.xl.render

import com.tjclp.xl.styles.font.Font

/** Scala Native text measurement: no AWT — pure character-count heuristic. */
private[xl] object TextMeasure:

  def measureTextWidth(text: String, font: Option[Font]): Int =
    RenderUtils.estimateTextWidth(text, font)
