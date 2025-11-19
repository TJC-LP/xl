package com.tjclp.xl

/**
 * Easy Mode: Single import for everything including IO.
 *
 * Provides the complete Easy Mode API with LLM-friendly ergonomics:
 *   - String-based cell references (no macros required)
 *   - Simplified IO operations
 *   - Compile-time validated literals (ref, money, date, fx)
 *   - Style DSL and RichText formatting
 *
 * '''Example: Complete workflow'''
 * {{{
 * import com.tjclp.xl.easy.*
 *
 * // Read existing workbooks
 * val wb = Excel.read("data.xlsx")
 *
 * // Create/modify sheets with string refs
 * val sheets = Sheet("Sales")
 *   .put("A1", "Product")                    // String ref
 *   .put("B1", "Revenue", CellStyle.header)  // Inline styling
 *   .applyStyle("A1:B1", CellStyle.bold)     // Template styling
 *
 * // Save
 * Excel.write(wb.put(sheets), "output.xlsx")
 * }}}
 *
 * '''Example: Modify in-place'''
 * {{{
 * import com.tjclp.xl.easy.*
 *
 * Excel.modify("report.xlsx") { wb =>
 *   val sales = wb.sheets("Sales")
 *     .put("A1", "Updated: " + LocalDate.now)
 *   wb.put(sales)
 * }
 * }}}
 *
 * '''Two import options:'''
 *   - `import com.tjclp.xl.*` - Core only (no IO)
 *   - `import com.tjclp.xl.easy.*` - Core + IO (this module)
 *
 * @note
 *   For pure functional code with Cats Effect, use [[com.tjclp.xl.io.ExcelIO]] directly instead of
 *   the `Excel` object.
 * @since 0.3.0
 */
object easy:
  // Re-export core syntax (domain, macros, extensions, styles DSL)
  export com.tjclp.xl.syntax.*

  // Export simplified Excel IO (aliased from EasyExcel)
  export com.tjclp.xl.io.EasyExcel as Excel

export easy.*
