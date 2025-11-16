package com.tjclp.xl.cell

import com.tjclp.xl.error.XLError
import scala.util.boundary, boundary.break

/**
 * Pure runtime parser for formula strings.
 *
 * Currently performs minimal validation (parentheses balance only). Full formula parsing is planned
 * for Phase 7 (Formula Evaluator).
 *
 * Limitations:
 *   - Does not validate function names
 *   - Does not detect string literals ("text") which may contain parens
 *   - Does not validate formula syntax beyond parentheses
 */
object FormulaParser:

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
          // Toggle string state, but handle escaped quotes ("")
          if inString && i + 1 < n && s.charAt(i + 1) == '"' then i += 1 // Skip the escaped quote
          else inString = !inString
        else if !inString then
          // Only count parens outside of strings
          if c == '(' then depth += 1
          else if c == ')' then
            depth -= 1
            if depth < 0 then break(false)

        i += 1

      // Must have balanced parens AND not be inside an unclosed string
      depth == 0 && !inString
