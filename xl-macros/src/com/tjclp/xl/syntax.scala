package com.tjclp.xl

/**
 * Macro facade for convenient top-level imports.
 *
 * Import all macros via wildcard:
 * {{{
 * import com.tjclp.xl.syntax.*  // Gets ref, fx, money, percent, put, etc.
 * }}}
 *
 * Or selectively:
 * {{{
 * import com.tjclp.xl.syntax.{ref, money, percent}
 * }}}
 *
 * Or import directly from package level (via export syntax.*):
 * {{{
 * import com.tjclp.xl.{ref, money, percent, put}
 * }}}
 *
 * Or import directly from source:
 * {{{
 * import com.tjclp.xl.macros.RefLiteral.ref
 * import com.tjclp.xl.macros.FormattedLiterals.{money, percent}
 * }}}
 */
object syntax:
  // Unified reference literal
  export com.tjclp.xl.macros.RefLiteral.ref
  export com.tjclp.xl.addressing.RefType

  // Formula literal
  export com.tjclp.xl.macros.CellRangeLiterals.fx

  // Formatted literals
  export com.tjclp.xl.macros.FormattedLiterals.{money, percent, date, accounting}

  // Batch put macro
  export com.tjclp.xl.macros.BatchPutMacro.put

export syntax.*
