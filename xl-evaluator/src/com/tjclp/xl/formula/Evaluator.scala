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
   * @param clock
   *   Clock for date/time functions (defaults to system clock)
   * @return
   *   Either evaluation error or computed value
   */
  def eval[A](expr: TExpr[A], sheet: Sheet, clock: Clock = Clock.system): Either[EvalError, A]

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
  def eval[A](expr: TExpr[A], sheet: Sheet, clock: Clock = Clock.system): Either[EvalError, A] =
    instance.eval(expr, sheet, clock)

/**
 * Private implementation of Evaluator.
 *
 * Implements all 17 TExpr cases with proper error handling and short-circuit semantics.
 */
private class EvaluatorImpl extends Evaluator:
  // Suppress asInstanceOf warning for FoldRange GADT type handling (required for type parameter erasure)
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def eval[A](expr: TExpr[A], sheet: Sheet, clock: Clock = Clock.system): Either[EvalError, A] =
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

      case TExpr.PolyRef(at) =>
        // PolyRef should never reach evaluator - function parsers must convert to typed Ref
        // If this is hit, it's a parser bug (forgot to convert PolyRef in function parser)
        scala.util.Left(
          EvalError.EvalFailed(
            "PolyRef reached evaluator - parser bug (function did not convert PolyRef to typed Ref)",
            Some("This is an internal error - please report this issue")
          )
        )

      // ===== Conditional =====
      case TExpr.If(cond, ifTrue, ifFalse) =>
        // If: evaluate condition, then branch based on result
        eval(cond, sheet, clock).flatMap { condValue =>
          if condValue then eval(ifTrue, sheet, clock)
          else eval(ifFalse, sheet, clock)
        }

      // ===== Arithmetic Operators =====
      case TExpr.Add(x, y) =>
        // Add: evaluate both operands, sum results
        for
          xv <- eval(x, sheet, clock)
          yv <- eval(y, sheet, clock)
        yield xv + yv

      case TExpr.Sub(x, y) =>
        // Subtract: evaluate both operands, subtract second from first
        for
          xv <- eval(x, sheet, clock)
          yv <- eval(y, sheet, clock)
        yield xv - yv

      case TExpr.Mul(x, y) =>
        // Multiply: evaluate both operands, multiply results
        for
          xv <- eval(x, sheet, clock)
          yv <- eval(y, sheet, clock)
        yield xv * yv

      case TExpr.Div(x, y) =>
        // Divide: evaluate both operands, check for division by zero
        for
          xv <- eval(x, sheet, clock)
          yv <- eval(y, sheet, clock)
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
        eval(x, sheet, clock).flatMap {
          case false =>
            // Short-circuit: x is false, result is false (don't evaluate y)
            Right(false)
          case true =>
            // x is true, evaluate y to determine final result
            eval(y, sheet, clock)
        }

      case TExpr.Or(x, y) =>
        // Or: short-circuit evaluation (if x is true, don't evaluate y)
        eval(x, sheet, clock).flatMap {
          case true =>
            // Short-circuit: x is true, result is true (don't evaluate y)
            Right(true)
          case false =>
            // x is false, evaluate y to determine final result
            eval(y, sheet, clock)
        }

      case TExpr.Not(x) =>
        // Not: logical negation
        eval(x, sheet, clock).map(xv => !xv)

      // ===== Comparison Operators =====
      case TExpr.Lt(x, y) =>
        // Less than: numeric comparison
        for
          xv <- eval(x, sheet, clock)
          yv <- eval(y, sheet, clock)
        yield xv < yv

      case TExpr.Lte(x, y) =>
        // Less than or equal: numeric comparison
        for
          xv <- eval(x, sheet, clock)
          yv <- eval(y, sheet, clock)
        yield xv <= yv

      case TExpr.Gt(x, y) =>
        // Greater than: numeric comparison
        for
          xv <- eval(x, sheet, clock)
          yv <- eval(y, sheet, clock)
        yield xv > yv

      case TExpr.Gte(x, y) =>
        // Greater than or equal: numeric comparison
        for
          xv <- eval(x, sheet, clock)
          yv <- eval(y, sheet, clock)
        yield xv >= yv

      case TExpr.Eq(x, y) =>
        // Equality: polymorphic comparison
        for
          xv <- eval(x, sheet, clock)
          yv <- eval(y, sheet, clock)
        yield xv == yv

      case TExpr.Neq(x, y) =>
        // Inequality: polymorphic comparison
        for
          xv <- eval(x, sheet, clock)
          yv <- eval(y, sheet, clock)
        yield xv != yv

      // ===== Type Conversions =====
      case TExpr.ToInt(expr) =>
        // ToInt: Convert BigDecimal to Int (validates integer range)
        eval(expr, sheet, clock).flatMap { bd =>
          if bd.isValidInt then Right(bd.toInt)
          else
            Left(
              EvalError.TypeMismatch(
                "ToInt",
                "valid integer",
                s"$bd (out of Int range)"
              )
            )
        }

      // ===== Text Functions =====
      case TExpr.Concatenate(xs) =>
        // Concatenate: evaluate all expressions, concat results
        xs.foldLeft[Either[EvalError, String]](Right("")) { (accEither, expr) =>
          for
            acc <- accEither
            value <- eval(expr, sheet, clock)
          yield acc + value
        }

      case TExpr.Left(text, n) =>
        // Left: extract left n characters
        for
          textValue <- eval(text, sheet, clock)
          nValue <- eval(n, sheet, clock)
          result <-
            if nValue < 0 then
              Left(EvalError.EvalFailed(s"LEFT: n must be non-negative, got $nValue"))
            else if nValue >= textValue.length then Right(textValue)
            else Right(textValue.take(nValue))
        yield result

      case TExpr.Right(text, n) =>
        // Right: extract right n characters
        for
          textValue <- eval(text, sheet, clock)
          nValue <- eval(n, sheet, clock)
          result <-
            if nValue < 0 then
              Left(EvalError.EvalFailed(s"RIGHT: n must be non-negative, got $nValue"))
            else if nValue >= textValue.length then Right(textValue)
            else Right(textValue.takeRight(nValue))
        yield result

      case TExpr.Len(text) =>
        // Len: text length (returns BigDecimal to match Excel and enable arithmetic)
        eval(text, sheet, clock).map(s => BigDecimal(s.length))

      case TExpr.Upper(text) =>
        // Upper: convert to uppercase
        eval(text, sheet, clock).map(_.toUpperCase)

      case TExpr.Lower(text) =>
        // Lower: convert to lowercase
        eval(text, sheet, clock).map(_.toLowerCase)

      // ===== Date/Time Functions =====
      case TExpr.Today() =>
        // Today: get current date from clock
        Right(clock.today())

      case TExpr.Now() =>
        // Now: get current date and time from clock
        Right(clock.now())

      case TExpr.Date(year, month, day) =>
        // Date: construct date from components
        for
          y <- eval(year, sheet, clock)
          m <- eval(month, sheet, clock)
          d <- eval(day, sheet, clock)
          result <-
            scala.util.Try(java.time.LocalDate.of(y, m, d)).toEither.left.map { ex =>
              EvalError.EvalFailed(
                s"DATE: invalid date components (year=$y, month=$m, day=$d): ${ex.getMessage}"
              )
            }
        yield result

      case TExpr.Year(date) =>
        // Year: extract year from date (returns BigDecimal to match Excel)
        eval(date, sheet, clock).map(d => BigDecimal(d.getYear))

      case TExpr.Month(date) =>
        // Month: extract month from date (returns BigDecimal to match Excel)
        eval(date, sheet, clock).map(d => BigDecimal(d.getMonthValue))

      case TExpr.Day(date) =>
        // Day: extract day from date (returns BigDecimal to match Excel)
        eval(date, sheet, clock).map(d => BigDecimal(d.getDayOfMonth))

      // ===== Arithmetic Range Functions =====
      case TExpr.Min(range) =>
        // Min: find minimum value in range
        val cells = range.cells.map(cellRef => sheet(cellRef))
        val values = cells.flatMap { cell =>
          TExpr.decodeNumeric(cell).toOption
        }
        if values.isEmpty then
          // Excel behavior: MIN of empty range returns 0
          Right(BigDecimal(0))
        else Right(values.min)

      case TExpr.Max(range) =>
        // Max: find maximum value in range
        val cells = range.cells.map(cellRef => sheet(cellRef))
        val values = cells.flatMap { cell =>
          TExpr.decodeNumeric(cell).toOption
        }
        if values.isEmpty then
          // Excel behavior: MAX of empty range returns 0
          Right(BigDecimal(0))
        else Right(values.max)

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
