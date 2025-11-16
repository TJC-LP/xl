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
   * Uses reflection to check if each argument is a literal constant.
   *
   * Example:
   * {{{
   * val x = "A1"              // Compile-time constant
   * ref"$x"                   // allLiterals → Some(Seq("A1"))
   *
   * def getUserInput() = "A1" // Runtime expression
   * ref"${getUserInput()}"    // allLiterals → None
   * }}}
   */
  def allLiterals(args: Expr[Seq[Any]])(using Quotes): Option[Seq[Any]] =
    import quotes.reflect.*

    args match
      case Varargs(exprs) =>
        val literalValues = exprs.map { expr =>
          // Check if expression is a literal constant using reflection
          expr.asTerm match
            case Inlined(_, _, Literal(constant)) => Some(constant.value)
            case Literal(constant) => Some(constant.value)
            case _ => None
        }
        // Check if ALL are defined (all are compile-time constants)
        if literalValues.forall(_.isDefined) then Some(literalValues.flatten.toSeq)
        else None
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
   * Format compile error for invalid compile-time interpolation.
   *
   * Provides helpful error message with the reconstructed string and parse error.
   *
   * Example output:
   * {{{
   * Invalid ref literal in interpolation: 'INVALID!@#$'
   * Error: Invalid characters in sheet name
   * Hint: Check that all interpolated parts form a valid Excel ref
   * }}}
   */
  def formatCompileError(macroName: String, fullString: String, parseError: String): String =
    s"Invalid $macroName literal in interpolation: '$fullString'\n" +
      s"Error: $parseError\n" +
      s"Hint: Check that all interpolated parts form a valid Excel $macroName"

end MacroUtil
