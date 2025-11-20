package com.tjclp.xl

/**
 * Unified syntax facade bringing together all XL extensions and macros.
 *
 * Import `com.tjclp.xl.*` or `com.tjclp.xl.syntax.*` for access to:
 *   - Domain model (Sheet, Cell, Workbook, CellStyle, etc.)
 *   - Extension methods (DSL operators, codec operations, optics)
 *   - Compile-time validated literals (ref, money, date, fx)
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
  // Easy Mode utilities
  export error.XLException // Exception wrapper for .unsafe boundary

  // Core syntax extensions with selective aliasing to avoid export conflicts
  export dsl.syntax as dslSyntax
  export dslSyntax.*

  export codec.syntax as codecSyntax
  export codecSyntax.*

  export optics.syntax as opticsSyntax
  export opticsSyntax.*

  export patch.syntax as patchSyntax
  export patchSyntax.*

  export sheets.syntax as sheetSyntax
  export sheetSyntax.*

  export styles.dsl as styleDsl
  export styleDsl.*

  // Easy Mode: String-based extensions (Sheet.put("A1", value), Sheet.style("A1:B1", style))
  export extensions.*

  // RichText string extensions (.red, .bold, etc.)
  export richtext.RichText.{given, *}

  // Compile-time validated literals (macros)
  export macros.RefLiteral.*
  export macros.CellRangeLiterals.fx
  export macros.FormattedLiterals.{money, percent, date, accounting}
  // BatchPutMacro.put removed - dead code (shadowed by Sheet.put member)

export syntax.*
