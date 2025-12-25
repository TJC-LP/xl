package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs, EvalContext}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.Clock

import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.SheetName
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
   * @param workbook
   *   Optional workbook for cross-sheet references (defaults to None)
   * @return
   *   Either evaluation error or computed value
   */
  def eval[A](
    expr: TExpr[A],
    sheet: Sheet,
    clock: Clock = Clock.system,
    workbook: Option[Workbook] = None
  ): Either[EvalError, A]

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
  def eval[A](
    expr: TExpr[A],
    sheet: Sheet,
    clock: Clock = Clock.system,
    workbook: Option[Workbook] = None
  ): Either[EvalError, A] =
    instance.eval(expr, sheet, clock, workbook)

  // Helper methods for consistent cross-sheet error messages
  private[formula] def missingWorkbookError(refStr: String, isRange: Boolean = false): EvalError =
    val refType = if isRange then "range" else "reference"
    EvalError.EvalFailed(
      s"Cross-sheet $refType $refStr requires workbook context, but none was provided.",
      None
    )

  private[formula] def sheetNotFoundError(
    sheetName: SheetName,
    err: com.tjclp.xl.error.XLError
  ): EvalError =
    EvalError.EvalFailed(
      s"Sheet '${sheetName.value}' not found in workbook: ${err.message}",
      None
    )

  /**
   * Resolve a RangeLocation to the target sheet.
   *
   * For Local ranges, returns the current sheet. For CrossSheet ranges, looks up the target sheet
   * in the workbook context.
   */
  private[formula] def resolveRangeLocation(
    location: TExpr.RangeLocation,
    currentSheet: Sheet,
    workbook: Option[Workbook]
  ): Either[EvalError, Sheet] =
    location match
      case TExpr.RangeLocation.Local(_) =>
        Right(currentSheet)
      case TExpr.RangeLocation.CrossSheet(sheetName, range) =>
        workbook match
          case None =>
            val refStr = s"${sheetName.value}!${range.toA1}"
            Left(missingWorkbookError(refStr, isRange = true))
          case Some(wb) =>
            wb(sheetName) match
              case Left(err) => Left(sheetNotFoundError(sheetName, err))
              case Right(targetSheet) => Right(targetSheet)

/**
 * Private implementation of Evaluator.
 *
 * Implements all TExpr cases with proper error handling and short-circuit semantics.
 */
private class EvaluatorImpl extends Evaluator:
  // Suppress asInstanceOf warning for GADT type handling (required for type parameter erasure)
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def eval[A](
    expr: TExpr[A],
    sheet: Sheet,
    clock: Clock = Clock.system,
    workbook: Option[Workbook] = None
  ): Either[EvalError, A] =
    // @unchecked: GADT exhaustivity - PolyRef should be resolved before evaluation
    (expr: @unchecked) match
      // ===== PolyRef Handling (Same-Sheet Reference) =====
      //
      // PolyRef should be resolved to typed Ref during parsing (see resolveTopLevelPolyRef
      // in FormulaParser). If we reach this case, it means a PolyRef escaped resolution,
      // which is a programming error. Return an error instead of using unsafe asInstanceOf.
      //
      case TExpr.PolyRef(at, _) =>
        Left(
          EvalError.EvalFailed(
            s"Unresolved PolyRef at ${(at: ARef).toA1} - should have been resolved during parsing",
            None
          )
        )

      // ===== Sheet-Qualified References (Cross-Sheet) =====
      //
      // SheetPolyRef should be resolved to typed SheetRef during parsing (see
      // resolveTopLevelPolyRef in FormulaParser). If we reach this case, it means
      // a SheetPolyRef escaped resolution, which is a programming error.
      // Return an error instead of using unsafe asInstanceOf.
      //
      case TExpr.SheetPolyRef(sheetName, at, _) =>
        Left(
          EvalError.EvalFailed(
            s"Unresolved SheetPolyRef at ${sheetName.value}!${(at: ARef).toA1} - should have been resolved during parsing",
            None
          )
        )

      case TExpr.SheetRef(sheetName, at, _, decode) =>
        // SheetRef: resolve cell from target sheet in workbook
        workbook match
          case None =>
            val refStr = s"${sheetName.value}!${(at: ARef).toA1}"
            Left(Evaluator.missingWorkbookError(refStr))
          case Some(wb) =>
            wb(sheetName) match
              case Left(err) =>
                Left(Evaluator.sheetNotFoundError(sheetName, err))
              case Right(targetSheet) =>
                val cell = targetSheet(at)
                decode(cell).left.map(codecErr => EvalError.CodecFailed(at, codecErr))

      case TExpr.SheetRange(sheetName, range) =>
        // SheetRange should be wrapped in a function (SUM, COUNT, etc.) before evaluation
        val refStr = s"${sheetName.value}!${range.toA1}"
        Left(
          EvalError.EvalFailed(
            s"Cross-sheet range $refStr must be used within a function like SUM or COUNT.",
            None
          )
        )

      case TExpr.RangeRef(range) =>
        Left(
          EvalError.EvalFailed(
            s"Range ${range.toA1} must be used within a function like SUM or COUNT.",
            None
          )
        )

      // ===== Literals =====
      case TExpr.Lit(value) =>
        // Literal: return value directly (identity law)
        Right(value)

      // ===== Cell References =====
      case TExpr.Ref(at, _, decode) =>
        // Ref: resolve cell, decode value with codec
        // Note: sheet(at) returns empty cell if not present, decode handles empty cells
        val cell = sheet(at)
        decode(cell).left.map(codecErr => EvalError.CodecFailed(at, codecErr))

      // ===== Arithmetic Operators =====
      case TExpr.Add(x, y) =>
        // Add: evaluate both operands, sum results
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv + yv

      case TExpr.Sub(x, y) =>
        // Subtract: evaluate both operands, subtract second from first
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv - yv

      case TExpr.Mul(x, y) =>
        // Multiply: evaluate both operands, multiply results
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv * yv

      case TExpr.Div(x, y) =>
        // Divide: evaluate both operands, check for division by zero
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
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

      // ===== Comparison Operators =====
      case TExpr.Lt(x, y) =>
        // Less than: numeric comparison
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv < yv

      case TExpr.Lte(x, y) =>
        // Less than or equal: numeric comparison
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv <= yv

      case TExpr.Gt(x, y) =>
        // Greater than: numeric comparison
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv > yv

      case TExpr.Gte(x, y) =>
        // Greater than or equal: numeric comparison
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv >= yv

      case TExpr.Eq(x, y) =>
        // Equality: polymorphic comparison
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv == yv

      case TExpr.Neq(x, y) =>
        // Inequality: polymorphic comparison
        for
          xv <- eval(x, sheet, clock, workbook)
          yv <- eval(y, sheet, clock, workbook)
        yield xv != yv

      // ===== Type Conversions =====
      case TExpr.ToInt(expr) =>
        // ToInt: Convert BigDecimal to Int (validates integer range)
        eval(expr, sheet, clock, workbook).flatMap { bd =>
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

      // ===== Date/Time Conversions =====
      case TExpr.DateToSerial(dateExpr) =>
        eval(dateExpr, sheet, clock, workbook).map { date =>
          BigDecimal(CellValue.dateTimeToExcelSerial(date.atStartOfDay()))
        }

      case TExpr.DateTimeToSerial(dtExpr) =>
        eval(dtExpr, sheet, clock, workbook).map { dt =>
          BigDecimal(CellValue.dateTimeToExcelSerial(dt))
        }

      case TExpr.Aggregate(aggregatorId, location) =>
        // Use Aggregator typeclass to evaluate any registered aggregate function
        Aggregator.lookup(aggregatorId) match
          case None =>
            Left(EvalError.EvalFailed(s"Unknown aggregator: $aggregatorId", None))
          case Some(agg) =>
            Evaluator.resolveRangeLocation(location, sheet, workbook).flatMap { targetSheet =>
              val cells = location.range.cells.map(cellRef => targetSheet(cellRef))
              // Fold over cells using the aggregator's combine function
              val result = cells.foldLeft(agg.empty) { (acc, cell) =>
                if agg.countsNonEmpty then
                  // COUNTA mode: count any non-empty cell
                  cell.value match
                    case CellValue.Empty => acc
                    case _ => agg.combine(acc, BigDecimal(1))
                else if agg.countsEmpty then
                  // COUNTBLANK mode: count only empty cells
                  cell.value match
                    case CellValue.Empty => agg.combine(acc, BigDecimal(1))
                    case _ => acc
                else
                  // Standard mode: only process numeric values
                  TExpr.decodeNumeric(cell) match
                    case Right(value) => agg.combine(acc, value)
                    case Left(_) if agg.skipNonNumeric => acc
                    case Left(_) => acc // Skip non-numeric cells
              }
              // Finalize and return the result (may return error for AVERAGE on empty range)
              agg.finalizeWithError(result)
            }

      case call: TExpr.Call[?] =>
        val ctx = EvalContext(
          sheet,
          clock,
          workbook,
          [A] => (expr: TExpr[A]) => eval(expr, sheet, clock, workbook)
        )
        call.spec.eval(call.args, ctx)
