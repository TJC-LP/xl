package com.tjclp.xl.styles.units

/** English Metric Unit - OOXML's precise unit (1/914400 inch) */
opaque type Emu = Long

object Emu:
  inline def apply(value: Long): Emu = value

  extension (emu: Emu)
    inline def value: Long = emu

    /** Convert to points */
    inline def toPt: Pt = Pt((emu * 72.0) / 914400.0)

    /** Convert to pixels */
    inline def toPx: Px = Px((emu * 96.0) / 914400.0)
