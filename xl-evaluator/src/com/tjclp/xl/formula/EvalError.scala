package com.tjclp.xl.formula

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.error.XLError
import com.tjclp.xl.codec.CodecError

/**
 * Formula evaluation errors with detailed diagnostics.
 *
 * All error types provide enough context to understand what went wrong during formula evaluation
 * and where. These errors are pure data - no exceptions thrown.
 *
 * Design follows ParseError pattern: enum with derives CanEqual, companion object with conversion
 * utilities.
 */
enum EvalError derives CanEqual:
  /**
   * Cell reference could not be resolved.
   *
   * @param ref
   *   The cell reference that failed
   * @param reason
   *   Why the reference failed (e.g., "cell not found", "cell is empty")
   *
   * Example: Formula references A1, but A1 doesn't exist → RefError(A1, "cell not found")
   */
  case RefError(ref: ARef, reason: String)

  /**
   * Cell value could not be decoded to expected type.
   *
   * @param ref
   *   The cell reference with the type mismatch
   * @param codecError
   *   The underlying codec error with details
   *
   * Example: Formula expects number in A1, but A1 contains text → CodecFailed(A1,
   * TypeMismatch(...))
   */
  case CodecFailed(ref: ARef, codecError: CodecError)

  /**
   * Division by zero attempted.
   *
   * @param numerator
   *   String representation of numerator expression
   * @param denominator
   *   String representation of denominator expression
   *
   * Example: Formula "=10/0" → DivByZero("10", "0") Example: Formula "=A1/(B1-B1)" →
   * DivByZero("Ref(A1)", "Sub(Ref(B1), Ref(B1))")
   */
  case DivByZero(numerator: String, denominator: String)

  /**
   * Circular reference detected in formula dependencies.
   *
   * @param cycle
   *   The cycle of cell references (A → B → C → A)
   *
   * Example: A1="=B1", B1="=C1", C1="=A1" → CircularRef(List(A1, B1, C1, A1))
   *
   * Note: Detection requires dependency graph analysis (WI-09b). For now, this error type exists
   * but is not yet raised during evaluation.
   */
  case CircularRef(cycle: List[ARef])

  /**
   * Type mismatch in operation.
   *
   * @param operation
   *   The operation that failed (e.g., "Add", "Lt", "And")
   * @param expected
   *   Description of expected type
   * @param actual
   *   Description of actual type received
   *
   * Example: Trying to add boolean to number → TypeMismatch("Add", "BigDecimal", "Boolean")
   */
  case TypeMismatch(operation: String, expected: String, actual: String)

  /**
   * Generic evaluation failure.
   *
   * @param reason
   *   Human-readable error description
   * @param context
   *   Optional context (expression that failed, cell reference, etc.)
   *
   * Used for errors that don't fit other categories.
   *
   * Example: Unsupported operation, internal error, etc.
   */
  case EvalFailed(reason: String, context: Option[String] = None)

object EvalError:
  /**
   * Convert EvalError to XLError for integration with existing error handling.
   *
   * @param error
   *   The evaluation error to convert
   * @param formula
   *   Optional formula string for context (used in error message)
   * @return
   *   XLError with descriptive message
   */
  def toXLError(error: EvalError, formula: Option[String] = None): XLError =
    val message = error match
      case RefError(ref, reason) =>
        s"Cell reference error at ${ref.toA1}: $reason"

      case CodecFailed(ref, codecErr) =>
        val codecMsg = codecErr match
          case CodecError.TypeMismatch(expected, actual) =>
            s"Type mismatch: expected $expected, got $actual"
          case CodecError.ParseError(value, targetType, detail) =>
            s"Parse failed for '$value' as $targetType: $detail"
        s"Failed to decode cell ${ref.toA1}: $codecMsg"

      case DivByZero(num, denom) =>
        s"Division by zero: $num / $denom"

      case CircularRef(cycle) =>
        val refs = cycle.map(_.toA1).mkString(" → ")
        s"Circular reference detected: $refs"

      case TypeMismatch(op, expected, actual) =>
        s"Type mismatch in $op: expected $expected, got $actual"

      case EvalFailed(reason, contextOpt) =>
        contextOpt.fold(s"Evaluation failed: $reason")(ctx =>
          s"Evaluation failed: $reason (context: $ctx)"
        )

    formula match
      case Some(f) => XLError.FormulaError(f, message)
      case None => XLError.Other(s"Formula evaluation error: $message")

  /**
   * Create a RefError with standard "cell not found" message.
   */
  def cellNotFound(ref: ARef): EvalError =
    RefError(ref, "cell not found or is empty")

  /**
   * Create a RefError with standard "cell is empty" message.
   */
  def cellEmpty(ref: ARef): EvalError =
    RefError(ref, "cell is empty")

  /**
   * Create a DivByZero error with pretty-printed expressions.
   */
  def divisionByZero(numExpr: String, denomExpr: String): EvalError =
    DivByZero(numExpr, denomExpr)

  /**
   * Create a CodecFailed error from codec error.
   */
  def codecFailed(ref: ARef, err: CodecError): EvalError =
    CodecFailed(ref, err)

  /**
   * Create a generic evaluation failure.
   */
  def failed(reason: String): EvalError =
    EvalFailed(reason, None)

  /**
   * Create a generic evaluation failure with context.
   */
  def failedWithContext(reason: String, context: String): EvalError =
    EvalFailed(reason, Some(context))
