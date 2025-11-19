package com.tjclp.xl.cells

import com.tjclp.xl.errors.XLError
import scala.util.boundary, boundary.break

/**
 * Pure runtime parser for formula strings.
 *
 * Currently performs minimal validation (parentheses balance only). Full formula parsing is planned
 * for Phase 7 (Formula Evaluator).
 *
 * Limitations:
 *   - Does not validate function names
 *   - Does not validate formula syntax beyond parentheses
 */
object FormulaParser:

  /** Excel's maximum cell content length per ECMA-376 specification */
  private val ExcelCellLimit = 32767

  /**
   * Parse formula string with minimal validation.
   *
   * Accepts:
   *   - Any string with balanced parentheses
   *   - With or without leading '=' (both are valid)
   *
   * Rejects:
   *   - Empty string
   *   - Unbalanced parentheses
   */
  def parse(s: String): Either[XLError, CellValue] =
    if s.isEmpty then Left(XLError.FormulaError(s, "Formula cannot be empty"))
    else if s.length > ExcelCellLimit then
      Left(
        XLError.FormulaError(
          s.take(50) + "...",
          s"Formula exceeds Excel cell limit ($ExcelCellLimit chars)"
        )
      )
    else if !validateParentheses(s) then Left(XLError.FormulaError(s, "Unbalanced parentheses"))
    else Right(CellValue.Formula(s))

  /**
   * Check parentheses are balanced, respecting string literals.
   *
   * String literals in Excel formulas use double quotes ("text"). Characters inside quotes are
   * ignored when balancing parentheses. Escaped quotes ("") within strings are handled correctly.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def validateParentheses(s: String): Boolean =
    boundary:
      var depth = 0
      var i = 0
      val n = s.length
      var inString = false

      while i < n do
        val c = s.charAt(i)

        if c == '"' then
          // Check for escaped quote ("")
          if inString && i + 1 < n && s.charAt(i + 1) == '"' then
            i += 2 // Skip both quotes, stay in string
          else
            inString = !inString
            i += 1
        else if !inString then
          // Only count parens outside of strings
          if c == '(' then depth += 1
          else if c == ')' then
            depth -= 1
            if depth < 0 then break(false)
          i += 1
        else
          // Inside string, just skip character
          i += 1

      // Must have balanced parens AND not be inside an unclosed string
      depth == 0 && !inString
