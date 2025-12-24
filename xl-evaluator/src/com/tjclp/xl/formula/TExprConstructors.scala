package com.tjclp.xl.formula

import com.tjclp.xl.{ARef, Anchor}
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.codec.CodecError
import com.tjclp.xl.formula.TExpr.*

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
   * Example: TExpr.cond(test, ifTrue, ifFalse)
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def cond[A](test: TExpr[Boolean], ifTrue: TExpr[A], ifFalse: TExpr[A]): TExpr[A] =
    Call(
      FunctionSpecs.ifFn,
      (test, ifTrue.asInstanceOf[TExpr[Any]], ifFalse.asInstanceOf[TExpr[Any]])
    ).asInstanceOf[TExpr[A]]
