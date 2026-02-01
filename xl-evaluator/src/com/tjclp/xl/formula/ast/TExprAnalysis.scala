package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.{FunctionSpecs, ArgValue}
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.{CellRange, SheetName}

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

  // ===== Range Collection and Transformation =====
  // GH-197: Used by SUMPRODUCT to bound full-column ranges in array expressions

  /**
   * Collect all CellRange references from an expression.
   *
   * Returns local ranges as (None, range) and cross-sheet ranges as (Some(sheetName), range). Used
   * by SUMPRODUCT to compute shared bounds across all ranges in array expressions.
   */
  def collectRanges(expr: TExpr[?]): List[(Option[SheetName], CellRange)] = expr match
    case RangeRef(range) => List((None, range))
    case SheetRange(sheet, range) => List((Some(sheet), range))
    case call: Call[?] =>
      call.spec.argSpec
        .toValues(call.args)
        .flatMap {
          case ArgValue.Expr(e) => collectRanges(e)
          case ArgValue.Range(l) => List((l.sheetName, l.range))
          case _ => Nil
        }
    // Arithmetic - recursively collect from operands
    case Add(l, r) => collectRanges(l) ++ collectRanges(r)
    case Sub(l, r) => collectRanges(l) ++ collectRanges(r)
    case Mul(l, r) => collectRanges(l) ++ collectRanges(r)
    case Div(l, r) => collectRanges(l) ++ collectRanges(r)
    // Comparisons
    case Eq(l, r) => collectRanges(l) ++ collectRanges(r)
    case Neq(l, r) => collectRanges(l) ++ collectRanges(r)
    case Lt(l, r) => collectRanges(l) ++ collectRanges(r)
    case Lte(l, r) => collectRanges(l) ++ collectRanges(r)
    case Gt(l, r) => collectRanges(l) ++ collectRanges(r)
    case Gte(l, r) => collectRanges(l) ++ collectRanges(r)
    // Type conversion
    case ToInt(e) => collectRanges(e)
    // Default: no ranges (Lit, Ref, PolyRef, SheetRef, SheetPolyRef, etc.)
    case _ => Nil

  /**
   * Transform all range references in an expression using a mapping function.
   *
   * Used to constrain full-column/row ranges to shared bounds before evaluation.
   *
   * @param expr
   *   The expression to transform
   * @param f
   *   Function that takes (optional sheet name, original range) and returns constrained range
   * @return
   *   New expression with transformed ranges
   */
  // GADT type erasure requires asInstanceOf; @nowarn for exhaustive pattern match false positives
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  @annotation.nowarn("msg=Unreachable case")
  def transformRanges[A](expr: TExpr[A], f: (Option[SheetName], CellRange) => CellRange): TExpr[A] =
    (expr match
      case RangeRef(range) => RangeRef(f(None, range))
      case SheetRange(sheet, range) => SheetRange(sheet, f(Some(sheet), range))
      // Arithmetic - recursively transform operands
      case Add(l, r) =>
        Add(transformRanges(l, f), transformRanges(r, f))
      case Sub(l, r) =>
        Sub(transformRanges(l, f), transformRanges(r, f))
      case Mul(l, r) =>
        Mul(transformRanges(l, f), transformRanges(r, f))
      case Div(l, r) =>
        Div(transformRanges(l, f), transformRanges(r, f))
      // Comparisons
      case Eq(l, r) =>
        Eq(transformRanges(l, f), transformRanges(r, f))
      case Neq(l, r) =>
        Neq(transformRanges(l, f), transformRanges(r, f))
      case Lt(l, r) =>
        Lt(transformRanges(l, f), transformRanges(r, f))
      case Lte(l, r) =>
        Lte(transformRanges(l, f), transformRanges(r, f))
      case Gt(l, r) =>
        Gt(transformRanges(l, f), transformRanges(r, f))
      case Gte(l, r) =>
        Gte(transformRanges(l, f), transformRanges(r, f))
      // Type conversion
      case ToInt(e) =>
        ToInt(transformRanges(e, f))
      // Default: return unchanged (Lit, Ref, PolyRef, SheetRef, SheetPolyRef, Call, etc.)
      case other => other
    ).asInstanceOf[TExpr[A]]
