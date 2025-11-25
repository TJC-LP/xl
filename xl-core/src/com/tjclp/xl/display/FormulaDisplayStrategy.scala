package com.tjclp.xl.display

import com.tjclp.xl.sheets.Sheet

/**
 * Type class for displaying formula cells.
 *
 * This provides capability-based formula display:
 *   - '''xl-core only''': Shows raw formula text `"=SUM(A1:A10)"`
 *   - '''xl-core + xl-evaluator''': Evaluates and formats `"$1,000,000"`
 *
 * The active strategy is determined by which given instances are in scope.
 *
 * @since 0.2.0
 */
trait FormulaDisplayStrategy:
  /**
   * Format a formula cell for display.
   *
   * @param formula
   *   The formula expression (without leading =)
   * @param sheet
   *   The sheet context for evaluation (if evaluating)
   * @return
   *   Formatted display string
   */
  def format(formula: String, sheet: Sheet): String

/**
 * Low-priority given instances for FormulaDisplayStrategy.
 *
 * This trait provides the default (non-evaluating) strategy at low priority. Objects that extend
 * this trait can provide higher-priority givens that override the default.
 *
 * This pattern (used by Cats, Scalaz) resolves ambiguity via inheritance depth: givens defined in
 * more-specific objects win over givens in parent traits.
 */
trait LowPriorityFormulaDisplay:
  /**
   * Default strategy: Show raw formula text (no evaluation).
   *
   * This is the fallback when xl-evaluator is not imported. Formulas display as `"=SUM(A1:A10)"`.
   */
  given default: FormulaDisplayStrategy with
    def format(formula: String, sheet: Sheet): String =
      // Formula expression already includes "=" prefix
      if formula.startsWith("=") then formula else s"=$formula"

object FormulaDisplayStrategy extends LowPriorityFormulaDisplay
