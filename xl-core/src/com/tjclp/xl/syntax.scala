package com.tjclp.xl

/**
 * Unified syntax facade bringing together all XL extensions and macros.
 *
 * Import `com.tjclp.xl.*` or `com.tjclp.xl.syntax.*` for access to:
 * - Domain model (Sheet, Cell, Workbook, CellStyle, etc.)
 * - Extension methods (DSL operators, codec operations, optics)
 * - Compile-time validated literals (ref, money, date, fx)
 *
 * This provides ZIO-style single-import ergonomics.
 *
 * Example:
 * {{{
 * import com.tjclp.xl.*
 *
 * val sheet = Sheet("Demo")
 *   .put(ref"A1", money"$1,234.56")
 *   .put(ref"B1", date"2025-11-13")
 *   .applyPatch(ref"A1:B1".merge)
 * }}}
 */
object syntax:
  // Core syntax extensions from various subpackages
  export dsl.syntax.*
  export codec.syntax.*
  export optics.syntax.*
  export patch.syntax.*
  export sheet.syntax.*

  // RichText string extensions (.red, .bold, etc.)
  export richtext.RichText.{given, *}

  // Compile-time validated literals (macros)
  export macros.RefLiteral.*
  export macros.CellRangeLiterals.fx
  export macros.FormattedLiterals.{money, percent, date, accounting}
  export macros.BatchPutMacro.put

export syntax.*
