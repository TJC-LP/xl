package com.tjclp.xl.formula

import com.tjclp.xl.formula.parser.ParseError

/**
 * Arity specification for function arguments.
 *
 * Captures argument count constraints in a type-safe way.
 */
enum Arity derives CanEqual:
  /** Exactly N arguments required (e.g., NOT takes exactly 1) */
  case Exact(n: Int)

  /** Between min and max arguments inclusive (e.g., IF with optional 4th arg for error value) */
  case Range(min: Int, max: Int)

  /** At least N arguments, no upper bound (e.g., AND, OR, CONCATENATE) */
  case AtLeast(n: Int)

object Arity:
  /** Zero arguments (TODAY, NOW) */
  def none: Arity = Exact(0)

  /** Exactly one argument (NOT, LEN, UPPER, etc.) */
  def one: Arity = Exact(1)

  /** Exactly two arguments (LEFT, RIGHT) */
  def two: Arity = Exact(2)

  /** Exactly three arguments (IF, DATE) */
  def three: Arity = Exact(3)

  /** At least one argument (AND, OR, CONCATENATE) */
  def atLeastOne: Arity = AtLeast(1)

  extension (arity: Arity)
    /**
     * Validate argument count matches arity.
     *
     * @return
     *   Either error message or unit
     */
    def validate(argCount: Int, functionName: String, pos: Int): Either[ParseError, Unit] =
      arity match
        case Exact(n) if argCount != n =>
          val expected =
            if n == 0 then "0 arguments"
            else if n == 1 then "1 argument"
            else s"$n arguments"
          scala.util.Left(
            ParseError.InvalidArguments(functionName, pos, expected, s"$argCount arguments")
          )
        case Range(min, max) if argCount < min || argCount > max =>
          scala.util.Left(
            ParseError.InvalidArguments(
              functionName,
              pos,
              s"between $min and $max arguments",
              s"$argCount arguments"
            )
          )
        case AtLeast(min) if argCount < min =>
          val expected =
            if min == 1 then "at least 1 argument" else s"at least $min arguments"
          scala.util.Left(
            ParseError.InvalidArguments(functionName, pos, expected, s"$argCount arguments")
          )
        case _ => scala.util.Right(())
