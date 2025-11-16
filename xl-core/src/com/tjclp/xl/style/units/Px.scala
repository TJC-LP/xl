package com.tjclp.xl.style.units

/** Pixel - screen unit (assumes 96 DPI for conversions) */
opaque type Px = Double

object Px:
  inline def apply(value: Double): Px = value

  extension (px: Px)
    inline def value: Double = px

    /** Convert to points */
    inline def toPt: Pt = Pt(px * 72.0 / 96.0)

    /** Convert to EMUs */
    inline def toEmu: Emu = Emu((px * 914400.0 / 96.0).toLong)
