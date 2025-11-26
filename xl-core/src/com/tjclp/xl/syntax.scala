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

  // CellCodec given instances for type class-based put() methods
  // Note: Scala 3's export doesn't transitively propagate givens from re-exports,
  // so we need a direct export here even though extensions.scala also exports them.
  export codec.CellCodec.given

  // Auto-conversions (String→Text, Int→Number, Double→Number, Boolean→Bool, etc.)
  // Enables: sheet.put("A1", "Hello") without explicit CellValue.Text wrapper
  export conversions.given

  // Display module: formatted display with string interpolation
  // Enables: given Sheet = mySheet; println(excel"Revenue: ${ref"A1"}")
  // Note: Uses type aliases and re-exports to avoid package name conflict
  type DisplayWrapper = com.tjclp.xl.display.DisplayWrapper
  val DisplayWrapper = com.tjclp.xl.display.DisplayWrapper
  type FormulaDisplayStrategy = com.tjclp.xl.display.FormulaDisplayStrategy
  val NumFmtFormatter = com.tjclp.xl.display.NumFmtFormatter

  // Re-export givens and extension methods from display module
  export com.tjclp.xl.display.DisplayConversions.given
  export com.tjclp.xl.display.ExcelInterpolator.excel
  export com.tjclp.xl.display.FormulaDisplayStrategy.given
  export com.tjclp.xl.display.syntax.{displayCell, displayFormula}

  // Compile-time validated literals (macros)
  export macros.RefLiteral.*
  export macros.ColumnLiteral.col
  export macros.CellRangeLiterals.fx
  export macros.FormattedLiterals.{money, percent, date, accounting}
  // BatchPutMacro.put removed - dead code (shadowed by Sheet.put member)

export syntax.*
