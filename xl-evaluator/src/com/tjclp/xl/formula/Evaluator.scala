package com.tjclp.xl.formula

import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.syntax.* // Extension methods for Sheet.get, CellRange.cells, ARef.toA1
import scala.math.BigDecimal

/**
 * Pure functional formula evaluator.
 *
 * Evaluates TExpr AST against a Sheet, returning either an error or the computed value. All
 * evaluation is total - no exceptions thrown, no side effects.
 *
 * Laws satisfied:
 *   1. Literal identity: eval(Lit(x)) == Right(x)
 *   2. Arithmetic laws: eval(Add(Lit(a), Lit(b))) == Right(a + b)
 *   3. Short-circuit: And(Lit(false), error) == Right(false) (no error raised)
 *   4. Totality: eval always returns Either[EvalError, A] (never throws)
 *
 * Example:
 * {{{
 * val expr = TExpr.Add(TExpr.Lit(BigDecimal(10)), TExpr.Ref(ref"A1", TExpr.decodeNumeric))
 * val evaluator = Evaluator.instance
 * evaluator.eval(expr, sheet) match
 *   case Right(result) => println(s"Result: $$result")
 *   case Left(error) => println(s"Error: $$error")
 * }}}
 */
trait Evaluator:
  /**
   * Evaluate expression against sheet.
   *
   * @param expr
   *   The expression to evaluate
   * @param sheet
   *   The sheet providing cell values
   * @return
   *   Either evaluation error or computed value
   */
  def eval[A](expr: TExpr[A], sheet: Sheet): Either[EvalError, A]

object Evaluator:
  /**
   * Default evaluator instance.
   *
   * Pure functional implementation with short-circuit evaluation for And/Or.
   */
  def instance: Evaluator = new EvaluatorImpl

  /**
   * Convenience method for direct evaluation (forwards to instance.eval).
   */
  def eval[A](expr: TExpr[A], sheet: Sheet): Either[EvalError, A] =
    instance.eval(expr, sheet)

/**
 * Private implementation of Evaluator.
 *
 * Implements all 17 TExpr cases with proper error handling and short-circuit semantics.
 */
private class EvaluatorImpl extends Evaluator:
  def eval[A](expr: TExpr[A], sheet: Sheet): Either[EvalError, A] =
    expr match
      // ===== Literals =====
      case TExpr.Lit(value) =>
        // Literal: return value directly (identity law)
        Right(value)

      // ===== Cell References =====
      case TExpr.Ref(at, decode) =>
        // Ref: resolve cell, decode value with codec
        // Note: sheet(at) returns empty cell if not present, decode handles empty cells
        val cell = sheet(at)
        decode(cell).left.map(codecErr => EvalError.CodecFailed(at, codecErr))

      // ===== Conditional =====
      case TExpr.If(cond, ifTrue, ifFalse) =>
        // If: evaluate condition, then branch based on result
        eval(cond, sheet).flatMap { condValue =>
          if condValue then eval(ifTrue, sheet)
          else eval(ifFalse, sheet)
        }

      // ===== Arithmetic Operators =====
      case TExpr.Add(x, y) =>
        // Add: evaluate both operands, sum results
        for
          xv <- eval(x, sheet)
          yv <- eval(y, sheet)
        yield xv + yv

      case TExpr.Sub(x, y) =>
        // Subtract: evaluate both operands, subtract second from first
        for
          xv <- eval(x, sheet)
          yv <- eval(y, sheet)
        yield xv - yv

      case TExpr.Mul(x, y) =>
        // Multiply: evaluate both operands, multiply results
        for
          xv <- eval(x, sheet)
          yv <- eval(y, sheet)
        yield xv * yv

      case TExpr.Div(x, y) =>
        // Divide: evaluate both operands, check for division by zero
        for
          xv <- eval(x, sheet)
          yv <- eval(y, sheet)
          result <-
            if yv == BigDecimal(0) then
              // Division by zero: provide helpful error message with expressions
              Left(
                EvalError.DivByZero(
                  FormulaPrinter.print(x, includeEquals = false),
                  FormulaPrinter.print(y, includeEquals = false)
                )
              )
            else Right(xv / yv)
        yield result

      // ===== Logical Operators =====
      case TExpr.And(x, y) =>
        // And: short-circuit evaluation (if x is false, don't evaluate y)
        eval(x, sheet).flatMap {
          case false =>
            // Short-circuit: x is false, result is false (don't evaluate y)
            Right(false)
          case true =>
            // x is true, evaluate y to determine final result
            eval(y, sheet)
        }

      case TExpr.Or(x, y) =>
        // Or: short-circuit evaluation (if x is true, don't evaluate y)
        eval(x, sheet).flatMap {
          case true =>
            // Short-circuit: x is true, result is true (don't evaluate y)
            Right(true)
          case false =>
            // x is false, evaluate y to determine final result
            eval(y, sheet)
        }

      case TExpr.Not(x) =>
        // Not: logical negation
        eval(x, sheet).map(xv => !xv)

      // ===== Comparison Operators =====
      case TExpr.Lt(x, y) =>
        // Less than: numeric comparison
        for
          xv <- eval(x, sheet)
          yv <- eval(y, sheet)
        yield xv < yv

      case TExpr.Lte(x, y) =>
        // Less than or equal: numeric comparison
        for
          xv <- eval(x, sheet)
          yv <- eval(y, sheet)
        yield xv <= yv

      case TExpr.Gt(x, y) =>
        // Greater than: numeric comparison
        for
          xv <- eval(x, sheet)
          yv <- eval(y, sheet)
        yield xv > yv

      case TExpr.Gte(x, y) =>
        // Greater than or equal: numeric comparison
        for
          xv <- eval(x, sheet)
          yv <- eval(y, sheet)
        yield xv >= yv

      case TExpr.Eq(x, y) =>
        // Equality: polymorphic comparison
        for
          xv <- eval(x, sheet)
          yv <- eval(y, sheet)
        yield xv == yv

      case TExpr.Neq(x, y) =>
        // Inequality: polymorphic comparison
        for
          xv <- eval(x, sheet)
          yv <- eval(y, sheet)
        yield xv != yv

      // ===== Range Aggregation =====
      case foldExpr: TExpr.FoldRange[a, b] =>
        // FoldRange: iterate cells in range, apply step function with accumulator
        // Note: Use pattern match with type parameters to preserve types
        // Excel behavior: SUM/COUNT/AVERAGE skip cells that can't be decoded (empty cells, text in numeric context)
        val cells = foldExpr.range.cells.map(cellRef => sheet(cellRef))
        val result: Either[EvalError, b] = cells.foldLeft[Either[EvalError, b]](
          Right(foldExpr.z)
        ) { (accEither, cellInstance) =>
          accEither.flatMap { acc =>
            foldExpr.decode(cellInstance) match
              case Right(value) =>
                // Successfully decoded, apply step function
                Right(foldExpr.step(acc, value))
              case Left(_codecErr) =>
                // Failed to decode: skip this cell (Excel behavior for SUM/COUNT/AVERAGE)
                // Empty cells and type mismatches are silently ignored
                Right(acc)
          }
        }
        result.asInstanceOf[Either[EvalError, A]]
