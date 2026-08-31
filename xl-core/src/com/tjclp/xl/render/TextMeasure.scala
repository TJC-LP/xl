package com.tjclp.xl.render

import com.tjclp.xl.styles.font.Font

import java.awt.{Font as AwtFont, Graphics2D}
import java.awt.image.BufferedImage

/**
 * JVM text measurement: AWT FontMetrics with a pure heuristic fallback.
 *
 * The ONLY place `java.awt` may be imported in xl-core (ADR-016): renderers reach font metrics
 * through [[RenderUtils.measureTextWidth]], so this backend can be swapped for the pure
 * [[RenderUtils.estimateTextWidth]] heuristic on platforms without AWT (Scala Native / Scala.js,
 * wave B1).
 */
private[xl] object TextMeasure:

  /** Graphics context for text measurement (lazy, with headless fallback). */
  private lazy val graphics: Option[Graphics2D] =
    try
      // Default AWT to headless unless the embedder explicitly chose otherwise. Font metrics
      // don't need a display, and non-headless AWT (notably on macOS) starts a non-daemon
      // AWT-Shutdown thread that keeps script JVMs alive after main completes — any script
      // calling toHtml/toSvg would hang at exit. GUI embedders that need a display can set
      // -Djava.awt.headless=false (or initialize their toolkit before calling xl rendering).
      if System.getProperty("java.awt.headless") == null then
        System.setProperty("java.awt.headless", "true")
      val img = new BufferedImage(1, 1, BufferedImage.TYPE_INT_ARGB)
      Some(img.createGraphics())
    catch case _: Throwable => None // Headless environment (catches UnsatisfiedLinkError too)

  /** Measure text width using AWT FontMetrics, with fallback estimation. */
  def measureTextWidth(text: String, font: Option[Font]): Int =
    graphics match
      case Some(g) =>
        val awtFont = toAwtFont(font)
        g.setFont(awtFont)
        g.getFontMetrics.stringWidth(text)
      case None =>
        RenderUtils.estimateTextWidth(text, font)

  /**
   * Convert our Font to AWT Font for text measurement.
   *
   * Uses pixel-equivalent sizes (pt * 4/3) to match SVG rendering, which uses pixel sizes. This
   * ensures text measurements match actual SVG output widths for accurate wrapping.
   */
  def toAwtFont(font: Option[Font]): AwtFont =
    font match
      case Some(f) =>
        val style =
          (if f.bold then AwtFont.BOLD else 0) | (if f.italic then AwtFont.ITALIC else 0)
        // Convert pt to px (pt * 4/3) to match SVG rendering
        val fontSizePx = (f.sizePt * 4.0 / 3.0).toInt
        new AwtFont(f.name, style, fontSizePx)
      case None =>
        // DefaultFontSize is 11pt = ~15px
        val defaultSizePx = (RenderUtils.DefaultFontSize * 4.0 / 3.0).toInt
        new AwtFont("Calibri", AwtFont.PLAIN, defaultSizePx)
