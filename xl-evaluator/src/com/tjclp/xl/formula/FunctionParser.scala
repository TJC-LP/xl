package com.tjclp.xl.formula

import com.tjclp.xl.addressing.CellRange

/**
 * Type class for parsing Excel formula functions.
 *
 * Each function (SUM, IF, MIN, etc.) has its own FunctionParser instance that encapsulates:
 *   - Function metadata (name, arity)
 *   - Argument validation logic
 *   - TExpr construction
 *
 * This enables:
 *   - Extensible function library (users can add custom functions via given instances)
 *   - Centralized function registry for introspection
 *   - Consistent error reporting
 *   - Type-safe construction of TExpr nodes
 *
 * Design follows CellWriter pattern: trait + given instances + companion object registry.
 */
trait FunctionParser[F]:
  /** Function name (uppercase, e.g., "SUM") */
  def name: String

  /** Arity specification for this function */
  def arity: Arity

  /**
   * Parse function arguments into TExpr.
   *
   * @param args
   *   Parsed argument expressions
   * @param pos
   *   Position in input (for error reporting)
   * @return
   *   Either ParseError or constructed TExpr
   */
  def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]]

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

/**
 * FunctionParser companion object and given instances.
 *
 * @note
 *   Suppression rationale:
 *   - AsInstanceOf: Type casts restore GADT type parameters lost during runtime parsing. Each cast
 *     is safe based on parser validation context.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
object FunctionParser:
  /**
   * Summon a parser for function type F.
   *
   * Example: FunctionParser[TExpr.Min]
   */
  def apply[F](using p: FunctionParser[F]): FunctionParser[F] = p

  /**
   * Registry of all available functions (populated by given instances).
   *
   * Key: Function name (uppercase) Value: FunctionParser instance (type-erased)
   */
  private lazy val registry: Map[String, FunctionParser[?]] = buildRegistry()

  /**
   * Lookup function parser by name (case-insensitive).
   *
   * @param name
   *   Function name (e.g., "sum", "SUM", "Sum")
   * @return
   *   Option[FunctionParser] if found
   */
  def lookup(name: String): Option[FunctionParser[?]] =
    registry.get(name.toUpperCase)

  /**
   * All registered function names (sorted alphabetically).
   *
   * Useful for introspection, documentation generation, and autocomplete.
   */
  def allFunctions: List[String] =
    registry.keys.toList.sorted

  /**
   * Check if function name is registered.
   */
  def isKnown(name: String): Boolean =
    registry.contains(name.toUpperCase)

  /**
   * Build registry from all given instances.
   *
   * Collects all FunctionParser instances in scope at initialization time.
   */
  private def buildRegistry(): Map[String, FunctionParser[?]] =
    // Directly collect instances by name (avoids summoning ambiguity)
    List(
      sumFunctionParser,
      countFunctionParser,
      averageFunctionParser,
      minFunctionParser,
      maxFunctionParser,
      ifFunctionParser,
      andFunctionParser,
      orFunctionParser,
      notFunctionParser,
      concatenateFunctionParser,
      leftFunctionParser,
      rightFunctionParser,
      lenFunctionParser,
      upperFunctionParser,
      lowerFunctionParser,
      todayFunctionParser,
      nowFunctionParser,
      dateFunctionParser,
      yearFunctionParser,
      monthFunctionParser,
      dayFunctionParser
    ).map(p => p.name -> p).toMap

  // ========== Given Instances for All Functions ==========

  // === Aggregate Functions ===

  /** SUM function: SUM(range) */
  given sumFunctionParser: FunctionParser[Unit] with
    def name: String = "SUM"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?]) =>
          scala.util.Right(fold) // Already created by parseRange
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("SUM", pos, "1 range argument", s"${args.length} arguments")
          )

  /** COUNT function: COUNT(range) */
  given countFunctionParser: FunctionParser[Unit] with
    def name: String = "COUNT"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?]) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) => scala.util.Right(TExpr.count(range))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "COUNT",
              pos,
              "1 range argument",
              s"${args.length} arguments"
            )
          )

  /** AVERAGE function: AVERAGE(range) */
  given averageFunctionParser: FunctionParser[Unit] with
    def name: String = "AVERAGE"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?]) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) => scala.util.Right(TExpr.average(range))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "AVERAGE",
              pos,
              "1 range argument",
              s"${args.length} arguments"
            )
          )

  /** MIN function: MIN(range) */
  given minFunctionParser: FunctionParser[Unit] with
    def name: String = "MIN"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?]) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) => scala.util.Right(TExpr.min(range))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("MIN", pos, "1 range argument", s"${args.length} arguments")
          )

  /** MAX function: MAX(range) */
  given maxFunctionParser: FunctionParser[Unit] with
    def name: String = "MAX"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?]) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) => scala.util.Right(TExpr.max(range))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("MAX", pos, "1 range argument", s"${args.length} arguments")
          )

  // === Logical Functions ===

  /** IF function: IF(condition, ifTrue, ifFalse) */
  given ifFunctionParser: FunctionParser[Unit] with
    def name: String = "IF"
    def arity: Arity = Arity.three

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(cond, ifTrue, ifFalse) =>
          scala.util.Right(
            TExpr.If(
              TExpr.asBooleanExpr(cond), // Convert PolyRef to Boolean
              ifTrue.asInstanceOf[TExpr[Any]], // ifTrue/ifFalse can be any type
              ifFalse.asInstanceOf[TExpr[Any]]
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("IF", pos, "3 arguments", s"${args.length} arguments")
          )

  /** AND function: AND(expr1, expr2, ...) - variadic */
  given andFunctionParser: FunctionParser[Unit] with
    def name: String = "AND"
    def arity: Arity = Arity.atLeastOne

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case Nil =>
          scala.util.Left(
            ParseError.InvalidArguments("AND", pos, "at least 1 argument", "0 arguments")
          )
        case head :: tail =>
          // Convert all arguments to Boolean (handles PolyRef)
          val result = tail.foldLeft(TExpr.asBooleanExpr(head)) { (acc, expr) =>
            TExpr.And(acc, TExpr.asBooleanExpr(expr))
          }
          scala.util.Right(result)

  /** OR function: OR(expr1, expr2, ...) - variadic */
  given orFunctionParser: FunctionParser[Unit] with
    def name: String = "OR"
    def arity: Arity = Arity.atLeastOne

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case Nil =>
          scala.util.Left(
            ParseError.InvalidArguments("OR", pos, "at least 1 argument", "0 arguments")
          )
        case head :: tail =>
          // Convert all arguments to Boolean (handles PolyRef)
          val result = tail.foldLeft(TExpr.asBooleanExpr(head)) { (acc, expr) =>
            TExpr.Or(acc, TExpr.asBooleanExpr(expr))
          }
          scala.util.Right(result)

  /** NOT function: NOT(expr) */
  given notFunctionParser: FunctionParser[Unit] with
    def name: String = "NOT"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(expr) => scala.util.Right(TExpr.Not(TExpr.asBooleanExpr(expr)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("NOT", pos, "1 argument", s"${args.length} arguments")
          )

  // === Text Functions ===

  /** CONCATENATE function: CONCATENATE(text1, text2, ...) - variadic */
  given concatenateFunctionParser: FunctionParser[Unit] with
    def name: String = "CONCATENATE"
    def arity: Arity = Arity.atLeastOne

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case Nil =>
          scala.util.Left(
            ParseError.InvalidArguments("CONCATENATE", pos, "at least 1 argument", "0 arguments")
          )
        case _ =>
          // Convert all arguments to String with coercion (handles PolyRef, numbers, etc.)
          scala.util.Right(TExpr.concatenate(args.map(TExpr.asStringExpr)))

  /** LEFT function: LEFT(text, n) */
  given leftFunctionParser: FunctionParser[Unit] with
    def name: String = "LEFT"
    def arity: Arity = Arity.two

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(text, n) =>
          // Convert PolyRef to String with coercion (handles numeric cells, etc.)
          val textStr = TExpr.asStringExpr(text)
          // Convert n to Int (handles PolyRef and BigDecimal literals)
          val nInt = TExpr.asIntExpr(n)
          scala.util.Right(TExpr.left(textStr, nInt))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("LEFT", pos, "2 arguments", s"${args.length} arguments")
          )

  /** RIGHT function: RIGHT(text, n) */
  given rightFunctionParser: FunctionParser[Unit] with
    def name: String = "RIGHT"
    def arity: Arity = Arity.two

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(text, n) =>
          // Convert PolyRef to String with coercion
          val textStr = TExpr.asStringExpr(text)
          // Convert n to Int (handles PolyRef and BigDecimal literals)
          val nInt = TExpr.asIntExpr(n)
          scala.util.Right(TExpr.right(textStr, nInt))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("RIGHT", pos, "2 arguments", s"${args.length} arguments")
          )

  /** LEN function: LEN(text) */
  given lenFunctionParser: FunctionParser[Unit] with
    def name: String = "LEN"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(text) => scala.util.Right(TExpr.len(TExpr.asStringExpr(text)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("LEN", pos, "1 argument", s"${args.length} arguments")
          )

  /** UPPER function: UPPER(text) */
  given upperFunctionParser: FunctionParser[Unit] with
    def name: String = "UPPER"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(text) => scala.util.Right(TExpr.upper(TExpr.asStringExpr(text)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("UPPER", pos, "1 argument", s"${args.length} arguments")
          )

  /** LOWER function: LOWER(text) */
  given lowerFunctionParser: FunctionParser[Unit] with
    def name: String = "LOWER"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(text) => scala.util.Right(TExpr.lower(TExpr.asStringExpr(text)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("LOWER", pos, "1 argument", s"${args.length} arguments")
          )

  // === Date/Time Functions ===

  /** TODAY function: TODAY() - no arguments */
  given todayFunctionParser: FunctionParser[Unit] with
    def name: String = "TODAY"
    def arity: Arity = Arity.none

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case Nil => scala.util.Right(TExpr.today())
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("TODAY", pos, "0 arguments", s"${args.length} arguments")
          )

  /** NOW function: NOW() - no arguments */
  given nowFunctionParser: FunctionParser[Unit] with
    def name: String = "NOW"
    def arity: Arity = Arity.none

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case Nil => scala.util.Right(TExpr.now())
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("NOW", pos, "0 arguments", s"${args.length} arguments")
          )

  /** DATE function: DATE(year, month, day) */
  given dateFunctionParser: FunctionParser[Unit] with
    def name: String = "DATE"
    def arity: Arity = Arity.three

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(year, month, day) =>
          // Convert arguments to Int (handles PolyRef, BigDecimal literals, etc.)
          scala.util.Right(
            TExpr.date(
              TExpr.asIntExpr(year),
              TExpr.asIntExpr(month),
              TExpr.asIntExpr(day)
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("DATE", pos, "3 arguments", s"${args.length} arguments")
          )

  /** YEAR function: YEAR(date) */
  given yearFunctionParser: FunctionParser[Unit] with
    def name: String = "YEAR"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(date) =>
          // Convert PolyRef to LocalDate with coercion
          scala.util.Right(TExpr.year(TExpr.asDateExpr(date)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("YEAR", pos, "1 argument", s"${args.length} arguments")
          )

  /** MONTH function: MONTH(date) */
  given monthFunctionParser: FunctionParser[Unit] with
    def name: String = "MONTH"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(date) =>
          // Convert PolyRef to LocalDate with coercion
          scala.util.Right(TExpr.month(TExpr.asDateExpr(date)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("MONTH", pos, "1 argument", s"${args.length} arguments")
          )

  /** DAY function: DAY(date) */
  given dayFunctionParser: FunctionParser[Unit] with
    def name: String = "DAY"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(date) =>
          // Convert PolyRef to LocalDate with coercion
          scala.util.Right(TExpr.day(TExpr.asDateExpr(date)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("DAY", pos, "1 argument", s"${args.length} arguments")
          )
