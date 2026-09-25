package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.{BindingCoercion, TExpr}
import com.tjclp.xl.formula.functions.LiftSlot

import com.tjclp.xl.cells.{CellError, CellValue}

/**
 * The pure half of Excel array lifting ([[com.tjclp.xl.formula.functions.FunctionFlags.lift]]); the
 * evaluator-dependent half (evaluating and classifying slots) is `EvaluatorImpl.evalLiftedCall`.
 *
 * A lifted call evaluates each lifted scalar slot at most once, then calls the function once per
 * element of the broadcast shape with every array slot replaced by that element, typed for the slot
 * ([[substitute]]). Lifting law: `f(range)[i] == f(cell_i)` under the element table
 * ([[elementResult]]), except where the single-reference decoders predate lifting — numeric text
 * and cached formula cells — which ArrayLiftingLawsSpec's carve-out test pins.
 */
private[formula] object ArrayLifting:

  /** What one lifted slot evaluated to before the element loop. */
  enum SlotValue:
    /**
     * Not pre-evaluated: the function evaluates the slot's own expression, lazily, exactly as an
     * unlifted call would — a scalar-certain slot, a computed slot of a plain (scalar-mode) cell,
     * or one that failed with a host failure (IFERROR's fallback must stay lazy and catching).
     */
    case Keep

    /** One value for every element. `fromCell`: read from a cell, so a value slot resolves it. */
    case Scalar(value: Any, fromCell: Boolean)

    /** An array the call lifts over (at least two elements). */
    case Elements(array: ArrayResult, fromCells: Boolean)

    /** The slot evaluated to an Excel error value, substituted as that error literal. */
    case ErrorArg(code: CellError)

  /**
   * Strip the conversion wrapper the parser put around a slot; the wrapper's target becomes the
   * slot kind, so an element substituted for the peeled expression still converts as the slot did
   * (`ToInt(B2:B4+5)` substitutes `Coerced(Lit(element), Integer)`). UnaryPlus is not peeled:
   * `+A2:A10` is a computed array, not a reference (the Analysis ToolPak rule depends on it).
   */
  def peel(expr: TExpr[?], slot: LiftSlot): (TExpr[?], LiftSlot) = expr match
    case TExpr.Coerced(inner, target) => (inner, LiftSlot.Typed(target))
    case TExpr.CoercedBindingRef(name, target) => (TExpr.BindingRef(name), LiftSlot.Typed(target))
    case TExpr.ToInt(inner) => (inner, LiftSlot.Typed(BindingCoercion.Integer))
    case TExpr.DateToSerial(inner) => (inner, LiftSlot.Typed(BindingCoercion.Numeric))
    case TExpr.DateTimeToSerial(inner) => (inner, LiftSlot.Typed(BindingCoercion.Numeric))
    case other => (other, slot)

  /** The range a reference-shaped slot names — a range, or a defined name that may denote one. */
  def referenceLocation(expr: TExpr[?]): Option[TExpr.RangeLocation] = expr match
    case TExpr.RangeRef(range, form) => Some(TExpr.RangeLocation.Local(range, form))
    case TExpr.SheetRange(sheet, range, form) =>
      Some(TExpr.RangeLocation.CrossSheet(sheet, range, form))
    case TExpr.NameRef(name) => Some(TExpr.RangeLocation.Name(name, None))
    case TExpr.SheetNameRef(sheet, name) => Some(TExpr.RangeLocation.Name(name, Some(sheet)))
    case _ => None

  /**
   * Conservative and sound: true only for slots that can never evaluate to an array, so the call
   * takes its unlifted path untouched (ROUND(A1,2) allocates nothing, and Missing/Ref shapes reach
   * the function as parsed). Calls, names and LET bodies may yield arrays and answer false.
   */
  def isScalarCertain(expr: TExpr[?], bindings: Map[String, Any]): Boolean = expr match
    case TExpr.Lit(_: ArrayResult) => false
    case TExpr.Lit(_) => true
    case TExpr.Ref(_, _, _) | TExpr.SheetRef(_, _, _, _) | TExpr.PolyRef(_, _) |
        TExpr.SheetPolyRef(_, _, _) | TExpr.ExternalRef(_, _, _, _) | TExpr.ErrorLit(_) |
        TExpr.Missing | TExpr.Aggregate(_, _) =>
      true
    case TExpr.BindingRef(name) => boundScalar(bindings, name)
    case TExpr.CoercedBindingRef(name, _) => boundScalar(bindings, name)
    case TExpr.Coerced(inner, _) => isScalarCertain(inner, bindings)
    case TExpr.ToInt(inner) => isScalarCertain(inner, bindings)
    case TExpr.DateToSerial(inner) => isScalarCertain(inner, bindings)
    case TExpr.DateTimeToSerial(inner) => isScalarCertain(inner, bindings)
    case TExpr.UnaryPlus(inner) => isScalarCertain(inner, bindings)
    case TExpr.Percent(inner) => isScalarCertain(inner, bindings)
    case TExpr.Add(x, y) => isScalarCertain(x, bindings) && isScalarCertain(y, bindings)
    case TExpr.Sub(x, y) => isScalarCertain(x, bindings) && isScalarCertain(y, bindings)
    case TExpr.Mul(x, y) => isScalarCertain(x, bindings) && isScalarCertain(y, bindings)
    case TExpr.Div(x, y) => isScalarCertain(x, bindings) && isScalarCertain(y, bindings)
    case TExpr.Pow(x, y) => isScalarCertain(x, bindings) && isScalarCertain(y, bindings)
    case TExpr.Concat(x, y) => isScalarCertain(x, bindings) && isScalarCertain(y, bindings)
    case TExpr.Eq(x, y) => isScalarCertain(x, bindings) && isScalarCertain(y, bindings)
    case TExpr.Neq(x, y) => isScalarCertain(x, bindings) && isScalarCertain(y, bindings)
    case TExpr.Lt(x, y) => isScalarCertain(x, bindings) && isScalarCertain(y, bindings)
    case TExpr.Lte(x, y) => isScalarCertain(x, bindings) && isScalarCertain(y, bindings)
    case TExpr.Gt(x, y) => isScalarCertain(x, bindings) && isScalarCertain(y, bindings)
    case TExpr.Gte(x, y) => isScalarCertain(x, bindings) && isScalarCertain(y, bindings)
    case _ => false

  private def boundScalar(bindings: Map[String, Any], name: String): Boolean =
    bindings.get(name) match
      case Some(_: ArrayResult) => false
      case _ => true

  /**
   * An evaluated array as a slot value: a 1×1 array is its element and an empty one is Empty (the
   * GH-302 collapse), so only a genuine array broadcasts and a call over scalars keeps its scalar
   * result (IF over it keeps its lazy branch selection).
   */
  def fromArray(array: ArrayResult, fromCells: Boolean): SlotValue =
    if array.isEmpty then SlotValue.Scalar(CellValue.Empty, fromCells)
    else if array.rows == 1 && array.cols == 1 then SlotValue.Scalar(array(0, 0), fromCells)
    else SlotValue.Elements(array, fromCells)

  /** One value, typed for its slot: what the function reads in place of the slot's expression. */
  def substitute(value: Any, slot: LiftSlot, fromCell: Boolean): TExpr[?] = slot match
    case LiftSlot.Typed(target) => TExpr.Coerced[Any](TExpr.Lit(value), target)
    case LiftSlot.Cell => TExpr.Lit(ArrayArithmetic.anyToCellValue(value))
    case LiftSlot.Value =>
      value match
        case cv: CellValue if fromCell => TExpr.Lit(TExpr.resolvedValue(cv))
        case other => TExpr.Lit(other)

  /**
   * One output element from one call: a value (a nested array keeps its top-left element); an Excel
   * error value keeps its code; an element that failed to coerce (TypeMismatch/CodecFailed) is
   * `#VALUE!`, Excel's answer for it. Host failures (EvalFailed, RefError, CircularRef) fail the
   * whole lifted call — a missing sheet is never laundered into `#VALUE!` elements.
   */
  def elementResult(result: Either[EvalError, Any]): Either[EvalError, CellValue] =
    result match
      case Right(value) => Right(EvalResult.toCellValue(value))
      case Left(error) =>
        EvalError.toErrorValue(error) match
          case Some(code) => Right(CellValue.Error(code))
          case None =>
            error match
              case EvalError.TypeMismatch(_, _, _) | EvalError.CodecFailed(_, _) =>
                Right(CellValue.Error(CellError.Value))
              case other => Left(other)

  /**
   * The Analysis ToolPak refusal: a multi-cell range REFERENCE in a scalar slot of EDATE, EOMONTH,
   * WORKDAY, NETWORKDAYS, YEARFRAC or MROUND is `#VALUE!` in Excel, while an array lifts.
   */
  def rangeReferenceRefused(fnName: String, location: TExpr.RangeLocation): EvalError =
    EvalError.ErrorValue(
      CellError.Value,
      Some(
        s"$fnName: a multi-cell range reference (${location.toA1}) in a scalar argument is " +
          s"#VALUE! in Excel; pass an array instead, e.g. +${location.toA1}"
      )
    )
