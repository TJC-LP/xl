package com.tjclp.xl

/**
 * Unified syntax facade (core helpers + macros).
 *
 * This lives in `xl-macros` (which depends on `xl-core`) so that importing
 * `com.tjclp.xl.syntax.*` or even `com.tjclp.xl.*` brings everything in
 * with one line, similar to Cats' `cats.syntax.all.*`.
 */
object syntax:
  // Core helpers (col/row/ref constructors, String parsing, addressing ops)
  export com.tjclp.xl.coreSyntax.*

  // Unified reference literal
  export com.tjclp.xl.macros.RefLiteral.*

  // Formula literal
  export com.tjclp.xl.macros.CellRangeLiterals.fx

  // Formatted literals
  export com.tjclp.xl.macros.FormattedLiterals.{money, percent, date, accounting}

  // Batch put macro
  export com.tjclp.xl.macros.BatchPutMacro.put

export syntax.*
