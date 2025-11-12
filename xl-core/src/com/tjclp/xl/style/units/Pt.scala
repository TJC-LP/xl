package com.tjclp.xl.style.units

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
