package com.tjclp.xl.formula

import com.tjclp.xl.addressing.CellRange
import scala.util.boundary
import boundary.break

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
      countaFunctionParser,
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
      piFunctionParser,
      dateFunctionParser,
      yearFunctionParser,
      monthFunctionParser,
      dayFunctionParser,
      npvFunctionParser,
      irrFunctionParser,
      vlookupFunctionParser,
      sumIfFunctionParser,
      countIfFunctionParser,
      sumIfsFunctionParser,
      countIfsFunctionParser,
      sumProductFunctionParser,
      xlookupFunctionParser,
      // Error handling functions
      iferrorFunctionParser,
      iserrorFunctionParser,
      // Rounding and math functions
      roundFunctionParser,
      roundUpFunctionParser,
      roundDownFunctionParser,
      absFunctionParser,
      // Lookup functions
      indexFunctionParser,
      matchFunctionParser,
      // Date-based financial functions
      xnpvFunctionParser,
      xirrFunctionParser,
      // Date calculation functions
      eomonthFunctionParser,
      edateFunctionParser,
      datedifFunctionParser,
      networkdaysFunctionParser,
      workdayFunctionParser,
      yearfracFunctionParser
    ).map(p => p.name -> p).toMap

  // ========== Given Instances for All Functions ==========

  // === Aggregate Functions ===

  /** SUM function: SUM(range) - uses unified Aggregate case with Aggregator typeclass */
  given sumFunctionParser: FunctionParser[Unit] with
    def name: String = "SUM"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?]) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              scala.util.Right(TExpr.Aggregate("SUM", TExpr.RangeLocation.Local(range)))
        case List(TExpr.SheetFoldRange(sheet, range, _, _, _)) =>
          scala.util.Right(TExpr.Aggregate("SUM", TExpr.RangeLocation.CrossSheet(sheet, range)))
        case List(TExpr.SheetRange(sheet, range)) =>
          scala.util.Right(TExpr.Aggregate("SUM", TExpr.RangeLocation.CrossSheet(sheet, range)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("SUM", pos, "1 range argument", s"${args.length} arguments")
          )

  /** COUNT function: COUNT(range) - uses unified Aggregate case with Aggregator typeclass */
  given countFunctionParser: FunctionParser[Unit] with
    def name: String = "COUNT"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?]) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              scala.util.Right(TExpr.Aggregate("COUNT", TExpr.RangeLocation.Local(range)))
        case List(TExpr.SheetFoldRange(sheet, range, _, _, _)) =>
          scala.util.Right(TExpr.Aggregate("COUNT", TExpr.RangeLocation.CrossSheet(sheet, range)))
        case List(TExpr.SheetRange(sheet, range)) =>
          scala.util.Right(TExpr.Aggregate("COUNT", TExpr.RangeLocation.CrossSheet(sheet, range)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "COUNT",
              pos,
              "1 range argument",
              s"${args.length} arguments"
            )
          )

  /** COUNTA function: COUNTA(range) - counts non-empty cells using Aggregator typeclass */
  given countaFunctionParser: FunctionParser[Unit] with
    def name: String = "COUNTA"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?]) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              scala.util.Right(TExpr.Aggregate("COUNTA", TExpr.RangeLocation.Local(range)))
        case List(TExpr.SheetFoldRange(sheet, range, _, _, _)) =>
          scala.util.Right(TExpr.Aggregate("COUNTA", TExpr.RangeLocation.CrossSheet(sheet, range)))
        case List(TExpr.SheetRange(sheet, range)) =>
          scala.util.Right(TExpr.Aggregate("COUNTA", TExpr.RangeLocation.CrossSheet(sheet, range)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "COUNTA",
              pos,
              "1 range argument",
              s"${args.length} arguments"
            )
          )

  /** AVERAGE function: AVERAGE(range) - uses unified Aggregate case with Aggregator typeclass */
  given averageFunctionParser: FunctionParser[Unit] with
    def name: String = "AVERAGE"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?]) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              scala.util.Right(TExpr.Aggregate("AVERAGE", TExpr.RangeLocation.Local(range)))
        case List(TExpr.SheetFoldRange(sheet, range, _, _, _)) =>
          scala.util.Right(TExpr.Aggregate("AVERAGE", TExpr.RangeLocation.CrossSheet(sheet, range)))
        case List(TExpr.SheetRange(sheet, range)) =>
          scala.util.Right(TExpr.Aggregate("AVERAGE", TExpr.RangeLocation.CrossSheet(sheet, range)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "AVERAGE",
              pos,
              "1 range argument",
              s"${args.length} arguments"
            )
          )

  /** MIN function: MIN(range) - uses unified Aggregate case with Aggregator typeclass */
  given minFunctionParser: FunctionParser[Unit] with
    def name: String = "MIN"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?]) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              scala.util.Right(TExpr.Aggregate("MIN", TExpr.RangeLocation.Local(range)))
        case List(TExpr.SheetFoldRange(sheet, range, _, _, _)) =>
          scala.util.Right(TExpr.Aggregate("MIN", TExpr.RangeLocation.CrossSheet(sheet, range)))
        case List(TExpr.SheetRange(sheet, range)) =>
          scala.util.Right(TExpr.Aggregate("MIN", TExpr.RangeLocation.CrossSheet(sheet, range)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("MIN", pos, "1 range argument", s"${args.length} arguments")
          )

  /** MAX function: MAX(range) - uses unified Aggregate case with Aggregator typeclass */
  given maxFunctionParser: FunctionParser[Unit] with
    def name: String = "MAX"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?]) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              scala.util.Right(TExpr.Aggregate("MAX", TExpr.RangeLocation.Local(range)))
        case List(TExpr.SheetFoldRange(sheet, range, _, _, _)) =>
          scala.util.Right(TExpr.Aggregate("MAX", TExpr.RangeLocation.CrossSheet(sheet, range)))
        case List(TExpr.SheetRange(sheet, range)) =>
          scala.util.Right(TExpr.Aggregate("MAX", TExpr.RangeLocation.CrossSheet(sheet, range)))
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

  // === Math Constants ===

  /** PI function: PI() - returns mathematical constant pi */
  given piFunctionParser: FunctionParser[Unit] with
    def name: String = "PI"
    def arity: Arity = Arity.none

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case Nil => scala.util.Right(TExpr.pi())
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("PI", pos, "0 arguments", s"${args.length} arguments")
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

  // === Financial Functions ===

  /** NPV function: NPV(rate, range) */
  given npvFunctionParser: FunctionParser[Unit] with
    def name: String = "NPV"
    def arity: Arity = Arity.two // rate + range

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case rateExpr :: (fold: TExpr.FoldRange[?, ?]) :: Nil =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              // Coerce rate to numeric; range becomes CellRange
              scala.util.Right(
                TExpr.npv(
                  TExpr.asNumericExpr(rateExpr),
                  range
                )
              )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "NPV",
              pos,
              "2 arguments (rate, range)",
              s"${args.length} arguments"
            )
          )

  /** IRR function: IRR(range, [guess]) */
  given irrFunctionParser: FunctionParser[Unit] with
    def name: String = "IRR"
    def arity: Arity = Arity.Range(1, 2)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?]) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              scala.util.Right(TExpr.irr(range, None))

        case List(fold: TExpr.FoldRange[?, ?], guessExpr) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              scala.util.Right(
                TExpr.irr(
                  range,
                  Some(TExpr.asNumericExpr(guessExpr))
                )
              )

        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "IRR",
              pos,
              "1 or 2 arguments (range, [guess])",
              s"${args.length} arguments"
            )
          )

  /** VLOOKUP function: VLOOKUP(lookup, table, colIndex, [rangeLookup]) */
  given vlookupFunctionParser: FunctionParser[Unit] with
    def name: String = "VLOOKUP"
    def arity: Arity = Arity.Range(3, 4)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        // Local range, no explicit range_lookup → default TRUE
        case List(lookupExpr, fold: TExpr.FoldRange[?, ?], colIndexExpr) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              scala.util.Right(
                TExpr.vlookupWithLocation(
                  TExpr.asCellValueExpr(lookupExpr), // Resolve PolyRef, preserve value type
                  TExpr.RangeLocation.Local(range),
                  TExpr.asIntExpr(colIndexExpr),
                  TExpr.Lit(true)
                )
              )

        // Local range, all four arguments provided
        case List(lookupExpr, fold: TExpr.FoldRange[?, ?], colIndexExpr, rangeLookupExpr) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              scala.util.Right(
                TExpr.vlookupWithLocation(
                  TExpr.asCellValueExpr(lookupExpr), // Resolve PolyRef, preserve value type
                  TExpr.RangeLocation.Local(range),
                  TExpr.asIntExpr(colIndexExpr),
                  TExpr.asBooleanExpr(rangeLookupExpr)
                )
              )

        // Cross-sheet range (SheetRange), no explicit range_lookup → default TRUE
        case List(lookupExpr, TExpr.SheetRange(sheet, range), colIndexExpr) =>
          scala.util.Right(
            TExpr.vlookupWithLocation(
              TExpr.asCellValueExpr(lookupExpr), // Resolve PolyRef, preserve value type
              TExpr.RangeLocation.CrossSheet(sheet, range),
              TExpr.asIntExpr(colIndexExpr),
              TExpr.Lit(true)
            )
          )

        // Cross-sheet range (SheetRange), all four arguments provided
        case List(lookupExpr, TExpr.SheetRange(sheet, range), colIndexExpr, rangeLookupExpr) =>
          scala.util.Right(
            TExpr.vlookupWithLocation(
              TExpr.asCellValueExpr(lookupExpr), // Resolve PolyRef, preserve value type
              TExpr.RangeLocation.CrossSheet(sheet, range),
              TExpr.asIntExpr(colIndexExpr),
              TExpr.asBooleanExpr(rangeLookupExpr)
            )
          )

        // Cross-sheet range (SheetFoldRange from function context), no explicit range_lookup
        case List(lookupExpr, sheetFold: TExpr.SheetFoldRange[?, ?], colIndexExpr) =>
          sheetFold match
            case TExpr.SheetFoldRange(sheet, range, _, _, _) =>
              scala.util.Right(
                TExpr.vlookupWithLocation(
                  TExpr.asCellValueExpr(lookupExpr), // Resolve PolyRef, preserve value type
                  TExpr.RangeLocation.CrossSheet(sheet, range),
                  TExpr.asIntExpr(colIndexExpr),
                  TExpr.Lit(true)
                )
              )

        // Cross-sheet range (SheetFoldRange from function context), all four arguments
        case List(
              lookupExpr,
              sheetFold: TExpr.SheetFoldRange[?, ?],
              colIndexExpr,
              rangeLookupExpr
            ) =>
          sheetFold match
            case TExpr.SheetFoldRange(sheet, range, _, _, _) =>
              scala.util.Right(
                TExpr.vlookupWithLocation(
                  TExpr.asCellValueExpr(lookupExpr), // Resolve PolyRef, preserve value type
                  TExpr.RangeLocation.CrossSheet(sheet, range),
                  TExpr.asIntExpr(colIndexExpr),
                  TExpr.asBooleanExpr(rangeLookupExpr)
                )
              )

        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "VLOOKUP",
              pos,
              "3 or 4 arguments (lookup_value, table_array, col_index_num, [range_lookup])",
              s"${args.length} arguments"
            )
          )

  // === Conditional Aggregation Functions ===

  /** SUMIF function: SUMIF(range, criteria, [sum_range]) */
  given sumIfFunctionParser: FunctionParser[Unit] with
    def name: String = "SUMIF"
    def arity: Arity = Arity.Range(2, 3)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        // 2 arguments: SUMIF(range, criteria) - sum range is same as criteria range
        case List(fold: TExpr.FoldRange[?, ?], criteriaExpr) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              scala.util.Right(TExpr.sumIf(range, criteriaExpr, None))

        // 3 arguments: SUMIF(range, criteria, sum_range)
        case List(fold: TExpr.FoldRange[?, ?], criteriaExpr, sumFold: TExpr.FoldRange[?, ?]) =>
          (fold, sumFold) match
            case (TExpr.FoldRange(range, _, _, _), TExpr.FoldRange(sumRange, _, _, _)) =>
              scala.util.Right(TExpr.sumIf(range, criteriaExpr, Some(sumRange)))

        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "SUMIF",
              pos,
              "2 or 3 arguments (range, criteria, [sum_range])",
              s"${args.length} arguments"
            )
          )

  /** COUNTIF function: COUNTIF(range, criteria) */
  given countIfFunctionParser: FunctionParser[Unit] with
    def name: String = "COUNTIF"
    def arity: Arity = Arity.two

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(fold: TExpr.FoldRange[?, ?], criteriaExpr) =>
          fold match
            case TExpr.FoldRange(range, _, _, _) =>
              scala.util.Right(TExpr.countIf(range, criteriaExpr))

        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "COUNTIF",
              pos,
              "2 arguments (range, criteria)",
              s"${args.length} arguments"
            )
          )

  /** SUMIFS function: SUMIFS(sum_range, criteria_range1, criteria1, ...) */
  given sumIfsFunctionParser: FunctionParser[Unit] with
    def name: String = "SUMIFS"
    def arity: Arity = Arity.AtLeast(3) // At least sum_range + one (range, criteria) pair

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      boundary {
        // Validate: need odd number of args >= 3 (sum_range + pairs of range,criteria)
        if args.length < 3 || args.length % 2 == 0 then
          break(
            scala.util.Left(
              ParseError.InvalidArguments(
                "SUMIFS",
                pos,
                "odd number of arguments >= 3 (sum_range, range1, criteria1, ...)",
                s"${args.length} arguments"
              )
            )
          )

        args.headOption match
          case Some(sumFold: TExpr.FoldRange[?, ?]) =>
            sumFold match
              case TExpr.FoldRange(sumRange, _, _, _) =>
                // Parse pairs of (range, criteria) from remaining args
                val pairs = args.drop(1).grouped(2).toList
                val conditions: Either[ParseError, List[(CellRange, TExpr[?])]] =
                  pairs.zipWithIndex.foldLeft[Either[ParseError, List[(CellRange, TExpr[?])]]](
                    scala.util.Right(List.empty)
                  ) { case (acc, (pair, idx)) =>
                    acc.flatMap { list =>
                      pair match
                        case List(fold: TExpr.FoldRange[?, ?], criteria) =>
                          fold match
                            case TExpr.FoldRange(range, _, _, _) =>
                              scala.util.Right(list :+ (range, criteria))
                        case _ =>
                          scala.util.Left(
                            ParseError.InvalidArguments(
                              "SUMIFS",
                              pos,
                              s"range at position ${2 + idx * 2}",
                              "non-range expression"
                            )
                          )
                    }
                  }
                conditions.map(conds => TExpr.sumIfs(sumRange, conds))

          case _ =>
            scala.util.Left(
              ParseError.InvalidArguments(
                "SUMIFS",
                pos,
                "first argument must be a range",
                "non-range expression"
              )
            )
      }

  /** COUNTIFS function: COUNTIFS(criteria_range1, criteria1, ...) */
  given countIfsFunctionParser: FunctionParser[Unit] with
    def name: String = "COUNTIFS"
    def arity: Arity = Arity.AtLeast(2) // At least one (range, criteria) pair

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      boundary {
        // Validate: need even number of args >= 2 (pairs of range,criteria)
        if args.length < 2 || args.length % 2 != 0 then
          break(
            scala.util.Left(
              ParseError.InvalidArguments(
                "COUNTIFS",
                pos,
                "even number of arguments >= 2 (range1, criteria1, ...)",
                s"${args.length} arguments"
              )
            )
          )

        // Parse pairs of (range, criteria)
        val pairs = args.grouped(2).toList
        val conditions: Either[ParseError, List[(CellRange, TExpr[?])]] =
          pairs.zipWithIndex.foldLeft[Either[ParseError, List[(CellRange, TExpr[?])]]](
            scala.util.Right(List.empty)
          ) { case (acc, (pair, idx)) =>
            acc.flatMap { list =>
              pair match
                case List(fold: TExpr.FoldRange[?, ?], criteria) =>
                  fold match
                    case TExpr.FoldRange(range, _, _, _) =>
                      scala.util.Right(list :+ (range, criteria))
                case _ =>
                  scala.util.Left(
                    ParseError.InvalidArguments(
                      "COUNTIFS",
                      pos,
                      s"range at position ${1 + idx * 2}",
                      "non-range expression"
                    )
                  )
            }
          }
        conditions.map(conds => TExpr.countIfs(conds))
      }

  // === Array and Advanced Lookup Functions ===

  /** SUMPRODUCT function: SUMPRODUCT(array1, [array2], ...) */
  given sumProductFunctionParser: FunctionParser[Unit] with
    def name: String = "SUMPRODUCT"
    def arity: Arity = Arity.AtLeast(1) // At least one array

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      boundary {
        if args.isEmpty then
          break(
            scala.util.Left(
              ParseError.InvalidArguments(
                "SUMPRODUCT",
                pos,
                "at least 1 array argument",
                "0 arguments"
              )
            )
          )

        // All arguments must be ranges
        val ranges: Either[ParseError, List[CellRange]] =
          args.zipWithIndex.foldLeft[Either[ParseError, List[CellRange]]](
            scala.util.Right(List.empty)
          ) { case (acc, (arg, idx)) =>
            acc.flatMap { list =>
              arg match
                case fold: TExpr.FoldRange[?, ?] =>
                  fold match
                    case TExpr.FoldRange(range, _, _, _) =>
                      scala.util.Right(list :+ range)
                case _ =>
                  scala.util.Left(
                    ParseError.InvalidArguments(
                      "SUMPRODUCT",
                      pos,
                      s"range at position ${idx + 1}",
                      "non-range expression"
                    )
                  )
            }
          }

        ranges.map(TExpr.sumProduct)
      }

  /**
   * XLOOKUP function: XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found],
   * [match_mode], [search_mode])
   */
  given xlookupFunctionParser: FunctionParser[Unit] with
    def name: String = "XLOOKUP"
    def arity: Arity = Arity.Range(3, 6)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      // Extract CellRange from FoldRange
      def extractRange(expr: TExpr[?], argName: String): Either[ParseError, CellRange] =
        expr match
          case fold: TExpr.FoldRange[?, ?] =>
            fold match
              case TExpr.FoldRange(range, _, _, _) => scala.util.Right(range)
          case _ =>
            scala.util.Left(
              ParseError.InvalidArguments(
                "XLOOKUP",
                pos,
                s"$argName must be a range",
                "non-range expression"
              )
            )

      args match
        // 3 args: XLOOKUP(lookup_value, lookup_array, return_array)
        case List(lookupValue, lookupArrayExpr, returnArrayExpr) =>
          for
            lookupArray <- extractRange(lookupArrayExpr, "lookup_array")
            returnArray <- extractRange(returnArrayExpr, "return_array")
          yield TExpr.xlookup(lookupValue, lookupArray, returnArray)

        // 4 args: with if_not_found
        case List(lookupValue, lookupArrayExpr, returnArrayExpr, ifNotFound) =>
          for
            lookupArray <- extractRange(lookupArrayExpr, "lookup_array")
            returnArray <- extractRange(returnArrayExpr, "return_array")
          yield TExpr.xlookup(lookupValue, lookupArray, returnArray, Some(ifNotFound))

        // 5 args: with if_not_found and match_mode
        case List(lookupValue, lookupArrayExpr, returnArrayExpr, ifNotFound, matchMode) =>
          for
            lookupArray <- extractRange(lookupArrayExpr, "lookup_array")
            returnArray <- extractRange(returnArrayExpr, "return_array")
          yield TExpr.xlookup(
            lookupValue,
            lookupArray,
            returnArray,
            Some(ifNotFound),
            TExpr.asIntExpr(matchMode)
          )

        // 6 args: all parameters
        case List(
              lookupValue,
              lookupArrayExpr,
              returnArrayExpr,
              ifNotFound,
              matchMode,
              searchMode
            ) =>
          for
            lookupArray <- extractRange(lookupArrayExpr, "lookup_array")
            returnArray <- extractRange(returnArrayExpr, "return_array")
          yield TExpr.xlookup(
            lookupValue,
            lookupArray,
            returnArray,
            Some(ifNotFound),
            TExpr.asIntExpr(matchMode),
            TExpr.asIntExpr(searchMode)
          )

        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "XLOOKUP",
              pos,
              "3 to 6 arguments (lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])",
              s"${args.length} arguments"
            )
          )

  // === Error Handling Functions ===

  /** IFERROR function: IFERROR(value, value_if_error) */
  given iferrorFunctionParser: FunctionParser[Unit] with
    def name: String = "IFERROR"
    def arity: Arity = Arity.two

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(valueExpr, errorExpr) =>
          scala.util.Right(
            TExpr.iferror(
              TExpr.asCellValueExpr(valueExpr),
              TExpr.asCellValueExpr(errorExpr)
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "IFERROR",
              pos,
              "2 arguments (value, value_if_error)",
              s"${args.length} arguments"
            )
          )

  /** ISERROR function: ISERROR(value) */
  given iserrorFunctionParser: FunctionParser[Unit] with
    def name: String = "ISERROR"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(valueExpr) =>
          scala.util.Right(TExpr.iserror(TExpr.asCellValueExpr(valueExpr)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "ISERROR",
              pos,
              "1 argument",
              s"${args.length} arguments"
            )
          )

  // === Rounding and Math Functions ===

  /** ROUND function: ROUND(number, num_digits) */
  given roundFunctionParser: FunctionParser[Unit] with
    def name: String = "ROUND"
    def arity: Arity = Arity.two

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(valueExpr, digitsExpr) =>
          scala.util.Right(
            TExpr.round(
              TExpr.asNumericExpr(valueExpr),
              TExpr.asNumericExpr(digitsExpr)
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "ROUND",
              pos,
              "2 arguments (number, num_digits)",
              s"${args.length} arguments"
            )
          )

  /** ROUNDUP function: ROUNDUP(number, num_digits) */
  given roundUpFunctionParser: FunctionParser[Unit] with
    def name: String = "ROUNDUP"
    def arity: Arity = Arity.two

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(valueExpr, digitsExpr) =>
          scala.util.Right(
            TExpr.roundUp(
              TExpr.asNumericExpr(valueExpr),
              TExpr.asNumericExpr(digitsExpr)
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "ROUNDUP",
              pos,
              "2 arguments (number, num_digits)",
              s"${args.length} arguments"
            )
          )

  /** ROUNDDOWN function: ROUNDDOWN(number, num_digits) */
  given roundDownFunctionParser: FunctionParser[Unit] with
    def name: String = "ROUNDDOWN"
    def arity: Arity = Arity.two

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(valueExpr, digitsExpr) =>
          scala.util.Right(
            TExpr.roundDown(
              TExpr.asNumericExpr(valueExpr),
              TExpr.asNumericExpr(digitsExpr)
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "ROUNDDOWN",
              pos,
              "2 arguments (number, num_digits)",
              s"${args.length} arguments"
            )
          )

  /** ABS function: ABS(number) */
  given absFunctionParser: FunctionParser[Unit] with
    def name: String = "ABS"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(valueExpr) =>
          scala.util.Right(TExpr.abs(TExpr.asNumericExpr(valueExpr)))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "ABS",
              pos,
              "1 argument",
              s"${args.length} arguments"
            )
          )

  // === Lookup Functions ===

  /** INDEX function: INDEX(array, row_num, [column_num]) */
  given indexFunctionParser: FunctionParser[Unit] with
    def name: String = "INDEX"
    def arity: Arity = Arity.Range(2, 3)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(arrayExpr, rowNumExpr) =>
          arrayExpr match
            case fold: TExpr.FoldRange[?, ?] =>
              fold match
                case TExpr.FoldRange(range, _, _, _) =>
                  scala.util.Right(
                    TExpr.index(range, TExpr.asNumericExpr(rowNumExpr), None)
                  )
            case _ =>
              scala.util.Left(
                ParseError.InvalidArguments(
                  "INDEX",
                  pos,
                  "first argument must be a cell range",
                  "non-range expression"
                )
              )
        case List(arrayExpr, rowNumExpr, colNumExpr) =>
          arrayExpr match
            case fold: TExpr.FoldRange[?, ?] =>
              fold match
                case TExpr.FoldRange(range, _, _, _) =>
                  scala.util.Right(
                    TExpr.index(
                      range,
                      TExpr.asNumericExpr(rowNumExpr),
                      Some(TExpr.asNumericExpr(colNumExpr))
                    )
                  )
            case _ =>
              scala.util.Left(
                ParseError.InvalidArguments(
                  "INDEX",
                  pos,
                  "first argument must be a cell range",
                  "non-range expression"
                )
              )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "INDEX",
              pos,
              "2-3 arguments (array, row_num, [column_num])",
              s"${args.length} arguments"
            )
          )

  /** MATCH function: MATCH(lookup_value, lookup_array, [match_type]) */
  given matchFunctionParser: FunctionParser[Unit] with
    def name: String = "MATCH"
    def arity: Arity = Arity.Range(2, 3)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(lookupValueExpr, lookupArrayExpr) =>
          lookupArrayExpr match
            case fold: TExpr.FoldRange[?, ?] =>
              fold match
                case TExpr.FoldRange(range, _, _, _) =>
                  scala.util.Right(
                    TExpr.matchExpr(lookupValueExpr, range, TExpr.Lit(BigDecimal(1)))
                  )
            case _ =>
              scala.util.Left(
                ParseError.InvalidArguments(
                  "MATCH",
                  pos,
                  "second argument must be a cell range",
                  "non-range expression"
                )
              )
        case List(lookupValueExpr, lookupArrayExpr, matchTypeExpr) =>
          lookupArrayExpr match
            case fold: TExpr.FoldRange[?, ?] =>
              fold match
                case TExpr.FoldRange(range, _, _, _) =>
                  scala.util.Right(
                    TExpr.matchExpr(lookupValueExpr, range, TExpr.asNumericExpr(matchTypeExpr))
                  )
            case _ =>
              scala.util.Left(
                ParseError.InvalidArguments(
                  "MATCH",
                  pos,
                  "second argument must be a cell range",
                  "non-range expression"
                )
              )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "MATCH",
              pos,
              "2-3 arguments (lookup_value, lookup_array, [match_type])",
              s"${args.length} arguments"
            )
          )

  // === Date-based Financial Functions ===

  /** XNPV function: XNPV(rate, values, dates) */
  given xnpvFunctionParser: FunctionParser[Unit] with
    def name: String = "XNPV"
    def arity: Arity = Arity.Exact(3)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(rateExpr, valuesFold: TExpr.FoldRange[?, ?], datesFold: TExpr.FoldRange[?, ?]) =>
          (valuesFold, datesFold) match
            case (TExpr.FoldRange(valuesRange, _, _, _), TExpr.FoldRange(datesRange, _, _, _)) =>
              scala.util.Right(
                TExpr.xnpv(TExpr.asNumericExpr(rateExpr), valuesRange, datesRange)
              )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "XNPV",
              pos,
              "3 arguments (rate, values, dates)",
              s"${args.length} arguments"
            )
          )

  /** XIRR function: XIRR(values, dates, [guess]) */
  given xirrFunctionParser: FunctionParser[Unit] with
    def name: String = "XIRR"
    def arity: Arity = Arity.Range(2, 3)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(valuesFold: TExpr.FoldRange[?, ?], datesFold: TExpr.FoldRange[?, ?]) =>
          (valuesFold, datesFold) match
            case (TExpr.FoldRange(valuesRange, _, _, _), TExpr.FoldRange(datesRange, _, _, _)) =>
              scala.util.Right(TExpr.xirr(valuesRange, datesRange, None))
        case List(
              valuesFold: TExpr.FoldRange[?, ?],
              datesFold: TExpr.FoldRange[?, ?],
              guessExpr
            ) =>
          (valuesFold, datesFold) match
            case (TExpr.FoldRange(valuesRange, _, _, _), TExpr.FoldRange(datesRange, _, _, _)) =>
              scala.util.Right(
                TExpr.xirr(valuesRange, datesRange, Some(TExpr.asNumericExpr(guessExpr)))
              )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "XIRR",
              pos,
              "2-3 arguments (values, dates, [guess])",
              s"${args.length} arguments"
            )
          )

  // === Date Calculation Functions ===

  /** EOMONTH function: EOMONTH(start_date, months) - end of month N months from start */
  given eomonthFunctionParser: FunctionParser[Unit] with
    def name: String = "EOMONTH"
    def arity: Arity = Arity.two

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(startDateExpr, monthsExpr) =>
          scala.util.Right(
            TExpr.eomonth(
              TExpr.asDateExpr(startDateExpr),
              TExpr.asIntExpr(monthsExpr)
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "EOMONTH",
              pos,
              "2 arguments (start_date, months)",
              s"${args.length} arguments"
            )
          )

  /** EDATE function: EDATE(start_date, months) - same day N months later */
  given edateFunctionParser: FunctionParser[Unit] with
    def name: String = "EDATE"
    def arity: Arity = Arity.two

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(startDateExpr, monthsExpr) =>
          scala.util.Right(
            TExpr.edate(
              TExpr.asDateExpr(startDateExpr),
              TExpr.asIntExpr(monthsExpr)
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "EDATE",
              pos,
              "2 arguments (start_date, months)",
              s"${args.length} arguments"
            )
          )

  /** DATEDIF function: DATEDIF(start_date, end_date, unit) - difference between dates */
  given datedifFunctionParser: FunctionParser[Unit] with
    def name: String = "DATEDIF"
    def arity: Arity = Arity.three

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(startDateExpr, endDateExpr, unitExpr) =>
          scala.util.Right(
            TExpr.datedif(
              TExpr.asDateExpr(startDateExpr),
              TExpr.asDateExpr(endDateExpr),
              TExpr.asStringExpr(unitExpr)
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "DATEDIF",
              pos,
              "3 arguments (start_date, end_date, unit)",
              s"${args.length} arguments"
            )
          )

  /** NETWORKDAYS function: NETWORKDAYS(start_date, end_date, [holidays]) - count working days */
  given networkdaysFunctionParser: FunctionParser[Unit] with
    def name: String = "NETWORKDAYS"
    def arity: Arity = Arity.Range(2, 3)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(startDateExpr, endDateExpr) =>
          scala.util.Right(
            TExpr.networkdays(
              TExpr.asDateExpr(startDateExpr),
              TExpr.asDateExpr(endDateExpr),
              None
            )
          )
        case List(startDateExpr, endDateExpr, holidaysFold: TExpr.FoldRange[?, ?]) =>
          holidaysFold match
            case TExpr.FoldRange(holidaysRange, _, _, _) =>
              scala.util.Right(
                TExpr.networkdays(
                  TExpr.asDateExpr(startDateExpr),
                  TExpr.asDateExpr(endDateExpr),
                  Some(holidaysRange)
                )
              )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "NETWORKDAYS",
              pos,
              "2-3 arguments (start_date, end_date, [holidays])",
              s"${args.length} arguments"
            )
          )

  /** WORKDAY function: WORKDAY(start_date, days, [holidays]) - add working days */
  given workdayFunctionParser: FunctionParser[Unit] with
    def name: String = "WORKDAY"
    def arity: Arity = Arity.Range(2, 3)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(startDateExpr, daysExpr) =>
          scala.util.Right(
            TExpr.workday(
              TExpr.asDateExpr(startDateExpr),
              TExpr.asIntExpr(daysExpr),
              None
            )
          )
        case List(startDateExpr, daysExpr, holidaysFold: TExpr.FoldRange[?, ?]) =>
          holidaysFold match
            case TExpr.FoldRange(holidaysRange, _, _, _) =>
              scala.util.Right(
                TExpr.workday(
                  TExpr.asDateExpr(startDateExpr),
                  TExpr.asIntExpr(daysExpr),
                  Some(holidaysRange)
                )
              )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "WORKDAY",
              pos,
              "2-3 arguments (start_date, days, [holidays])",
              s"${args.length} arguments"
            )
          )

  /** YEARFRAC function: YEARFRAC(start_date, end_date, [basis]) - year fraction between dates */
  given yearfracFunctionParser: FunctionParser[Unit] with
    def name: String = "YEARFRAC"
    def arity: Arity = Arity.Range(2, 3)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(startDateExpr, endDateExpr) =>
          scala.util.Right(
            TExpr.yearfrac(
              TExpr.asDateExpr(startDateExpr),
              TExpr.asDateExpr(endDateExpr),
              TExpr.Lit(0) // Default basis is 0 (US 30/360)
            )
          )
        case List(startDateExpr, endDateExpr, basisExpr) =>
          scala.util.Right(
            TExpr.yearfrac(
              TExpr.asDateExpr(startDateExpr),
              TExpr.asDateExpr(endDateExpr),
              TExpr.asIntExpr(basisExpr)
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "YEARFRAC",
              pos,
              "2-3 arguments (start_date, end_date, [basis])",
              s"${args.length} arguments"
            )
          )
