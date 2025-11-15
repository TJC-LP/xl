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
   * Check parentheses are balanced.
   *
   * Limitation: Does not handle string literals ("text") which may contain parens. This is
   * acceptable for Phase 1; full parsing is Phase 7.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def validateParentheses(s: String): Boolean =
    boundary:
      var depth = 0
      var i = 0
      val n = s.length

      while i < n do
        val c = s.charAt(i)
        if c == '(' then depth += 1
        else if c == ')' then
          depth -= 1
          if depth < 0 then break(false)
        i += 1

      depth == 0
