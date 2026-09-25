package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.{ExprCoercer, FunctionSpecs}
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.{ARef, Anchor}
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.codec.CodecError
import TExpr.*

trait TExprConstructors:
  /**
   * Smart constructor for literals.
   *
   * Example: TExpr.lit(42)
   */
  def lit[A](value: A): TExpr[A] = Lit(value)

  /**
   * Smart constructor for cell references.
   *
   * Example: TExpr.ref(ARef("A1"), Anchor.Relative, codec)
   */
  def ref[A](at: ARef, anchor: Anchor, decode: Cell => Either[CodecError, A]): TExpr[A] =
    Ref(at, anchor, decode)

  /**
   * Smart constructor for cell references with default Relative anchor.
   *
   * Example: TExpr.ref(ARef("A1"), codec)
   */
  def ref[A](at: ARef, decode: Cell => Either[CodecError, A]): TExpr[A] =
    Ref(at, Anchor.Relative, decode)

  /**
   * Smart constructor for conditionals.
   *
   * IF's branches are Any-typed, so the call is brought to `A` the way a typed argument slot is:
   * through the slot's coercer (a Coerced wrapper for numeric, text, boolean and date results), so
   * a branch of another runtime type is a clean Left at evaluation, never a mistyped value.
   *
   * Example: TExpr.cond(test, ifTrue, ifFalse)
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def cond[A](test: TExpr[Boolean], ifTrue: TExpr[A], ifFalse: TExpr[A])(using
    coercer: ExprCoercer[A]
  ): TExpr[A] =
    coercer.coerce(
      Call(
        FunctionSpecs.ifFn,
        (test, ifTrue.asInstanceOf[TExpr[Any]], ifFalse.asInstanceOf[TExpr[Any]])
      )
    )
