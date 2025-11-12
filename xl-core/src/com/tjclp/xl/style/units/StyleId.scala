package com.tjclp.xl.style.units

/** Style identifier for type-safe style indices */
opaque type StyleId = Int

object StyleId:
  def apply(i: Int): StyleId = i

  extension (s: StyleId) inline def value: Int = s
