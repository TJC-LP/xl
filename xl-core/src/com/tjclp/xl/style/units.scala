package com.tjclp.xl.style

/**
 * Units for Excel dimensions and style identifiers.
 *
 * Excel uses multiple unit systems for different purposes:
 *   - Points (Pt): Primary unit for fonts and dimensions (1/72 inch)
 *   - Pixels (Px): Screen unit (assumes 96 DPI)
 *   - EMUs (Emu): OOXML precise unit (1/914400 inch)
 */

// ========== Units ==========

/** Point (1/72 inch) - Excel's primary unit for fonts and dimensions */
opaque type Pt = Double

object Pt:
  def apply(value: Double): Pt = value

  extension (pt: Pt)
    def value: Double = pt

    /** Convert to pixels (assumes 96 DPI) */
    def toPx: Px = Px(pt * 96.0 / 72.0)

    /** Convert to EMUs (English Metric Units) */
    def toEmu: Emu = Emu((pt * 914400.0 / 72.0).toLong)

/** Pixel - screen unit (assumes 96 DPI for conversions) */
opaque type Px = Double

object Px:
  def apply(value: Double): Px = value

  extension (px: Px)
    def value: Double = px

    /** Convert to points */
    def toPt: Pt = Pt(px * 72.0 / 96.0)

    /** Convert to EMUs */
    def toEmu: Emu = Emu((px * 914400.0 / 96.0).toLong)

/** English Metric Unit - OOXML's precise unit (1/914400 inch) */
opaque type Emu = Long

object Emu:
  def apply(value: Long): Emu = value

  extension (emu: Emu)
    def value: Long = emu

    /** Convert to points */
    def toPt: Pt = Pt((emu * 72.0) / 914400.0)

    /** Convert to pixels */
    def toPx: Px = Px((emu * 96.0) / 914400.0)

// ========== StyleId ==========

/** Style identifier for type-safe style indices */
opaque type StyleId = Int

object StyleId:
  def apply(i: Int): StyleId = i

  extension (s: StyleId) inline def value: Int = s
