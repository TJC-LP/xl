package com.tjclp.xl.styles.units

/** Style identifier for type-safe style indices */
opaque type StyleId = Int

object StyleId:
  inline def apply(i: Int): StyleId = i

  extension (s: StyleId) inline def value: Int = s
