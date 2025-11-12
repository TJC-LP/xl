package com.tjclp.xl

/**
 * Macro facade for convenient top-level imports.
 *
 * Re-exports all compile-time validated literals and macros so they can be imported directly:
 * {{{
 * import com.tjclp.xl.macros.*  // Gets ref, money, percent, put, etc.
 * }}}
 *
 * Or selectively:
 * {{{
 * import com.tjclp.xl.macros.{ref, money, percent}
 * }}}
 *
 * Note: Deprecated `cell` and `range` literals are NOT exported at top-level due to naming
 * conflicts with the `cell` package. Use `ref` instead, or import explicitly:
 * {{{
 * import com.tjclp.xl.macros.CellRangeLiterals.{cell, range}  // Deprecated
 * }}}
 */
object syntax:
  // Unified reference literal (recommended)
  export com.tjclp.xl.macros.RefLiteral.ref
  export com.tjclp.xl.addressing.RefType

  // Formula literal
  export com.tjclp.xl.macros.CellRangeLiterals.fx

  // Formatted literals
  export com.tjclp.xl.macros.FormattedLiterals.{money, percent, date, accounting}

  // Batch put macro
  export com.tjclp.xl.macros.BatchPutMacro.put

  // NOTE: cell and range are NOT exported here due to naming conflict with cell package
  // Users must use ref"A1" instead, or import deprecated macros explicitly

export syntax.*
