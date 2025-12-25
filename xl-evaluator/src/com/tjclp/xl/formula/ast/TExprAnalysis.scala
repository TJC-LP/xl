package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.{FunctionSpecs, ArgValue}
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import TExpr.*

trait TExprAnalysis:
  // ===== Date Function Detection =====
  // Used by CLI to auto-apply date formatting when writing formulas

  /**
   * Check if expression contains any date-returning functions.
   *
   * Used to determine if a formula result should be formatted as a date in Excel. Recursively
   * checks compound expressions (arithmetic, conditionals, etc.).
   */
  def containsDateFunction(expr: TExpr[?]): Boolean = expr match
    case call: Call[?] =>
      call.spec.flags.returnsDate ||
      call.spec.argSpec
        .toValues(call.args)
        .collect { case ArgValue.Expr(e) => containsDateFunction(e) }
        .exists(identity)
    // Date-to-serial wrappers (for arithmetic)
    case DateToSerial(_) | DateTimeToSerial(_) => true
    // Arithmetic - recursively check operands
    case Add(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Sub(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Mul(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Div(l, r) => containsDateFunction(l) || containsDateFunction(r)
    // Conditionals and logical functions handled via Call args
    // Comparisons
    case Eq(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Neq(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Lt(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Lte(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Gt(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Gte(l, r) => containsDateFunction(l) || containsDateFunction(r)
    // Error handling
    // Type conversion
    case ToInt(e) => containsDateFunction(e)
    // Default: no date function
    case _ => false

  /**
   * Check if expression contains time-returning functions (NOW).
   *
   * Used to distinguish between Date format (m/d/yy) and DateTime format (m/d/yy h:mm). If true,
   * use DateTime format; otherwise use Date format.
   */
  def containsTimeFunction(expr: TExpr[?]): Boolean = expr match
    case call: Call[?] =>
      call.spec.flags.returnsTime ||
      call.spec.argSpec
        .toValues(call.args)
        .collect { case ArgValue.Expr(e) => containsTimeFunction(e) }
        .exists(identity)
    case DateTimeToSerial(_) => true
    // Arithmetic - recursively check operands
    case Add(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Sub(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Mul(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Div(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    // Conditionals and logical functions handled via Call args
    // Comparisons
    case Eq(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Neq(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Lt(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Lte(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Gt(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Gte(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    // Error handling
    // Type conversion
    case ToInt(e) => containsTimeFunction(e)
    // Default: no time function
    case _ => false
