package com.tjclp.xl.formula.eval

import scala.util.boundary
import scala.util.control.NonFatal

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.error.{XLError, XLResult}

/**
 * Totality at the evaluation boundary (#681 item 5). The evaluator is total by contract, so a
 * throwable escaping it — a ClassCastException, a failing Clock or Rng capability, a stack
 * exhausted by a pathological nesting — is an internal defect, not an Excel error value. It becomes
 * [[EvalError.EvalFailed]] naming the throwable and the cell, which [[EvalError.toErrorValue]]
 * never promotes: the cell stays a loud per-cell failure, never a `#VALUE!`.
 *
 * Only NonFatal throwables and StackOverflowError are contained. OutOfMemoryError (MemoryGuard's
 * RESOURCE_LIMIT contract), the other VirtualMachineErrors, InterruptedException (parallel
 * recalculation's cancellation) and ControlThrowable propagate, and so does `boundary.Break`: it is
 * a RuntimeException in Scala 3, so NonFatal alone would turn a caller's Clock or Rng breaking to
 * its own boundary into a defect.
 *
 * Placement is load-bearing: [[guard]] runs only in the outermost evaluator frame (the one every
 * `Evaluator` factory returns), never inside the recursive evaluation, so IFERROR/ISERROR and the
 * array broadcasts never see a throwable they could demote. The stack has unwound by the time a
 * StackOverflowError reaches that frame, so recovering there is safe.
 */
private[formula] object EvalDefect:

  /** Longest slice of a throwable's own message a defect message quotes. */
  private val MaxDetail = 200

  def guard[A](cell: Option[ARef])(body: => Either[EvalError, A]): Either[EvalError, A] =
    try body
    catch
      case escape: boundary.Break[?] => throw escape
      case overflow: StackOverflowError => Left(EvalError.EvalFailed(message(overflow, cell)))
      case NonFatal(t) => Left(EvalError.EvalFailed(message(t, cell)))

  /**
   * [[guard]] for a per-cell recalculation site holding an XLResult: defence in depth around work
   * outside the evaluator proper (pinned-cache checks, result conversion), with the same wording
   * the guarded evaluator's failure renders with.
   */
  def xlGuard[A](formulaText: String, cell: Option[ARef])(body: => XLResult[A]): XLResult[A] =
    try body
    catch
      case escape: boundary.Break[?] => throw escape
      case overflow: StackOverflowError => Left(failure(formulaText, overflow, cell))
      case NonFatal(t) => Left(failure(formulaText, t, cell))

  def message(t: Throwable, cell: Option[ARef]): String =
    val at = cell.fold("")(ref => s" at ${ref.toA1}")
    t match
      case _: StackOverflowError =>
        s"Evaluation exhausted the stack$at: the formula's nesting or reference chain is too deep"
      case _ =>
        val detail = Option(t.getMessage).filter(_.nonEmpty).fold("")(m => s": ${capped(m)}")
        s"Evaluation threw ${t.getClass.getName}$detail$at — an internal evaluator defect, not " +
          "an Excel error value; please report it with the formula"

  private def failure(formulaText: String, t: Throwable, cell: Option[ARef]): XLError =
    XLError.FormulaError(formulaText, s"Evaluation failed: ${message(t, cell)}")

  private def capped(detail: String): String =
    if detail.length <= MaxDetail then detail else detail.take(MaxDetail) + "…"
