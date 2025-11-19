package com.tjclp.xl.macros

import scala.quoted.*

/**
 * Shared utilities for macro compile-time optimization (Phase 2).
 *
 * Provides detection and reconstruction helpers used by all macros to enable zero-overhead
 * compile-time optimization for constant interpolations.
 */
object MacroUtil:

  /**
   * Detect if all interpolation arguments are compile-time constants.
   *
   * Returns None if any argument is a runtime expression, or Some(literals) if all are constants.
   *
   * Uses Scala 3's FromExpr mechanism to properly detect constants (including inline vals).
   *
   * Example:
   * {{{
   * inline val x = "A1"       // Compile-time constant (inline)
   * ref"$x"                   // allLiterals → Some(Seq("A1"))
   *
   * val y = "A1"              // Runtime variable
   * ref"$y"                   // allLiterals → None
   *
   * def getUserInput() = "A1" // Runtime expression
   * ref"${getUserInput()}"    // allLiterals → None
   * }}}
   */
  def allLiterals(args: Expr[Seq[Any]])(using Quotes): Option[Seq[Any]] =
    args match
      case Varargs(exprs) =>
        val values: Seq[Option[Any]] = exprs.map(extractConst)

        if values.forall(_.isDefined) then Some(values.flatten.toSeq)
        else None

      case _ => None

  /**
   * Extract compile-time constant value from an Expr[Any].
   *
   * Tries common types in order using Scala 3's FromExpr mechanism. Returns None if the expression
   * is not a compile-time constant.
   */
  private def extractConst(expr: Expr[Any])(using Quotes): Option[Any] =
    // Try string-like first (most common in refs)
    constOf[String](expr)
      // Numeric types used in refs
      .orElse(constOf[Int](expr))
      .orElse(constOf[Long](expr))
      .orElse(constOf[Double](expr))
      // Boolean for completeness
      .orElse(constOf[Boolean](expr))

  /**
   * Try to extract a compile-time constant of specific type T.
   *
   * Uses Scala 3's Expr.unapply which leverages FromExpr[T] instances.
   */
  private def constOf[T: Type: FromExpr](expr: Expr[Any])(using Quotes): Option[T] =
    expr match
      case '{ $x: T } => Expr.unapply[T](x)
      case _ => None

  /**
   * Reconstruct interpolated string from parts and literal values.
   *
   * Interleaves StringContext.parts with argument values to produce the full string.
   *
   * Algorithm: parts(0) + literals(0) + parts(1) + literals(1) + ... + parts(n)
   *
   * Invariant: parts.length == literals.length + 1 (enforced by Scala's StringContext)
   *
   * Example:
   * {{{
   * parts = Seq("", "!A1")
   * literals = Seq("Sales")
   * reconstructString(parts, literals) → "Sales!A1"
   *
   * parts = Seq("", "", "")
   * literals = Seq("B", 42)
   * reconstructString(parts, literals) → "B42"
   * }}}
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def reconstructString(parts: Seq[String], literals: Seq[Any]): String =
    require(
      parts.length == literals.length + 1,
      s"parts.length (${parts.length}) must equal literals.length + 1 (${literals.length + 1})"
    )

    val sb = new StringBuilder
    var i = 0
    val n = literals.length

    while i < n do
      sb.append(parts(i))
      sb.append(literals(i).toString)
      i += 1

    sb.append(parts(i)) // Final part (always exists)
    sb.toString

  /**
   * Format compile errors for invalid compile-time interpolation.
   *
   * Provides helpful errors message with the reconstructed string and parse errors.
   *
   * Example output:
   * {{{
   * Invalid ref literal in interpolation: 'INVALID!@#$'
   * Error: Invalid characters in sheets name
   * Hint: Check that all interpolated parts form a valid Excel ref
   * }}}
   */
  def formatCompileError(macroName: String, fullString: String, parseError: String): String =
    s"Invalid $macroName literal in interpolation: '$fullString'\n" +
      s"Error: $parseError\n" +
      s"Hint: Check that all interpolated parts form a valid Excel $macroName"

end MacroUtil
