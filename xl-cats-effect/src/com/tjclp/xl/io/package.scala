package com.tjclp.xl

/**
 * IO package providing Excel file operations.
 *
 * '''Import pattern:'''
 * {{{
 * import com.tjclp.xl.io.Excel  // Synchronous IO (scripts, REPL)
 * // or
 * import com.tjclp.xl.io.ExcelIO  // Cats Effect (production)
 * }}}
 */
package object io:
  /** Excel IO alias for convenient synchronous operations. */
  val Excel: EasyExcel.type = EasyExcel
