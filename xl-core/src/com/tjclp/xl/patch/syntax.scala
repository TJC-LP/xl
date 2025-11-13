package com.tjclp.xl.patch

/** Syntax for patch composition using cats Monoid */
object syntax:
  export cats.syntax.monoid.given
  export cats.syntax.semigroup.given

export syntax.*
