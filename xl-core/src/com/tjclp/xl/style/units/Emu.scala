package com.tjclp.xl.style.units

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
