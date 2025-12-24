package com.tjclp.xl.formula

import com.tjclp.xl.addressing.CellRange
import scala.util.boundary
import boundary.break

/**
 * Type class for parsing Excel formula functions.
 *
 * Each parser-based function (TODAY, VLOOKUP, etc.) has its own FunctionParser instance that
 * encapsulates:
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
      sumProductFunctionParser,
      xlookupFunctionParser,
      // Reference functions
      rowFunctionParser,
      columnFunctionParser,
      rowsFunctionParser,
      columnsFunctionParser,
      addressFunctionParser,
      // Lookup functions
      indexFunctionParser,
      matchFunctionParser,
      // Date-based financial functions
      xnpvFunctionParser,
      xirrFunctionParser,
      // TVM functions
      pmtFunctionParser,
      fvFunctionParser,
      pvFunctionParser,
      nperFunctionParser,
      rateFunctionParser
    ).map(p => p.name -> p).toMap

  // ========== Given Instances for All Functions ==========

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

  // === Reference Functions ===

  /** ROW function: ROW(reference) - returns 1-based row number */
  given rowFunctionParser: FunctionParser[Unit] with
    def name: String = "ROW"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(refExpr) =>
          scala.util.Right(TExpr.Row_(refExpr))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("ROW", pos, "1 argument", s"${args.length} arguments")
          )

  /** COLUMN function: COLUMN(reference) - returns 1-based column number */
  given columnFunctionParser: FunctionParser[Unit] with
    def name: String = "COLUMN"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(refExpr) =>
          scala.util.Right(TExpr.Column_(refExpr))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("COLUMN", pos, "1 argument", s"${args.length} arguments")
          )

  /** ROWS function: ROWS(range) - returns number of rows in range */
  given rowsFunctionParser: FunctionParser[Unit] with
    def name: String = "ROWS"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(rangeExpr) =>
          scala.util.Right(TExpr.Rows(rangeExpr))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("ROWS", pos, "1 argument", s"${args.length} arguments")
          )

  /** COLUMNS function: COLUMNS(range) - returns number of columns in range */
  given columnsFunctionParser: FunctionParser[Unit] with
    def name: String = "COLUMNS"
    def arity: Arity = Arity.one

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(rangeExpr) =>
          scala.util.Right(TExpr.Columns(rangeExpr))
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments("COLUMNS", pos, "1 argument", s"${args.length} arguments")
          )

  /** ADDRESS function: ADDRESS(row, column, [abs_num], [a1], [sheet_text]) */
  given addressFunctionParser: FunctionParser[Unit] with
    def name: String = "ADDRESS"
    def arity: Arity = Arity.Range(2, 5)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(rowExpr, colExpr) =>
          scala.util.Right(
            TExpr.Address(
              TExpr.asNumericExpr(rowExpr),
              TExpr.asNumericExpr(colExpr),
              TExpr.Lit(BigDecimal(1)),
              TExpr.Lit(true),
              None
            )
          )
        case List(rowExpr, colExpr, absNumExpr) =>
          scala.util.Right(
            TExpr.Address(
              TExpr.asNumericExpr(rowExpr),
              TExpr.asNumericExpr(colExpr),
              TExpr.asNumericExpr(absNumExpr),
              TExpr.Lit(true),
              None
            )
          )
        case List(rowExpr, colExpr, absNumExpr, a1Expr) =>
          scala.util.Right(
            TExpr.Address(
              TExpr.asNumericExpr(rowExpr),
              TExpr.asNumericExpr(colExpr),
              TExpr.asNumericExpr(absNumExpr),
              TExpr.asBooleanExpr(a1Expr),
              None
            )
          )
        case List(rowExpr, colExpr, absNumExpr, a1Expr, sheetExpr) =>
          scala.util.Right(
            TExpr.Address(
              TExpr.asNumericExpr(rowExpr),
              TExpr.asNumericExpr(colExpr),
              TExpr.asNumericExpr(absNumExpr),
              TExpr.asBooleanExpr(a1Expr),
              Some(TExpr.asStringExpr(sheetExpr))
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "ADDRESS",
              pos,
              "2 to 5 arguments",
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

  // === TVM (Time Value of Money) Functions ===

  /** PMT function: PMT(rate, nper, pv, [fv], [type]) - payment per period */
  given pmtFunctionParser: FunctionParser[Unit] with
    def name: String = "PMT"
    def arity: Arity = Arity.Range(3, 5)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(rateExpr, nperExpr, pvExpr) =>
          scala.util.Right(
            TExpr.pmt(
              TExpr.asNumericExpr(rateExpr),
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pvExpr),
              None,
              None
            )
          )
        case List(rateExpr, nperExpr, pvExpr, fvExpr) =>
          scala.util.Right(
            TExpr.pmt(
              TExpr.asNumericExpr(rateExpr),
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pvExpr),
              Some(TExpr.asNumericExpr(fvExpr)),
              None
            )
          )
        case List(rateExpr, nperExpr, pvExpr, fvExpr, pmtTypeExpr) =>
          scala.util.Right(
            TExpr.pmt(
              TExpr.asNumericExpr(rateExpr),
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pvExpr),
              Some(TExpr.asNumericExpr(fvExpr)),
              Some(TExpr.asNumericExpr(pmtTypeExpr))
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "PMT",
              pos,
              "3-5 arguments (rate, nper, pv, [fv], [type])",
              s"${args.length} arguments"
            )
          )

  /** FV function: FV(rate, nper, pmt, [pv], [type]) - future value */
  given fvFunctionParser: FunctionParser[Unit] with
    def name: String = "FV"
    def arity: Arity = Arity.Range(3, 5)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(rateExpr, nperExpr, pmtExpr) =>
          scala.util.Right(
            TExpr.fv(
              TExpr.asNumericExpr(rateExpr),
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pmtExpr),
              None,
              None
            )
          )
        case List(rateExpr, nperExpr, pmtExpr, pvExpr) =>
          scala.util.Right(
            TExpr.fv(
              TExpr.asNumericExpr(rateExpr),
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pmtExpr),
              Some(TExpr.asNumericExpr(pvExpr)),
              None
            )
          )
        case List(rateExpr, nperExpr, pmtExpr, pvExpr, pmtTypeExpr) =>
          scala.util.Right(
            TExpr.fv(
              TExpr.asNumericExpr(rateExpr),
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pmtExpr),
              Some(TExpr.asNumericExpr(pvExpr)),
              Some(TExpr.asNumericExpr(pmtTypeExpr))
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "FV",
              pos,
              "3-5 arguments (rate, nper, pmt, [pv], [type])",
              s"${args.length} arguments"
            )
          )

  /** PV function: PV(rate, nper, pmt, [fv], [type]) - present value */
  given pvFunctionParser: FunctionParser[Unit] with
    def name: String = "PV"
    def arity: Arity = Arity.Range(3, 5)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(rateExpr, nperExpr, pmtExpr) =>
          scala.util.Right(
            TExpr.pv(
              TExpr.asNumericExpr(rateExpr),
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pmtExpr),
              None,
              None
            )
          )
        case List(rateExpr, nperExpr, pmtExpr, fvExpr) =>
          scala.util.Right(
            TExpr.pv(
              TExpr.asNumericExpr(rateExpr),
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pmtExpr),
              Some(TExpr.asNumericExpr(fvExpr)),
              None
            )
          )
        case List(rateExpr, nperExpr, pmtExpr, fvExpr, pmtTypeExpr) =>
          scala.util.Right(
            TExpr.pv(
              TExpr.asNumericExpr(rateExpr),
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pmtExpr),
              Some(TExpr.asNumericExpr(fvExpr)),
              Some(TExpr.asNumericExpr(pmtTypeExpr))
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "PV",
              pos,
              "3-5 arguments (rate, nper, pmt, [fv], [type])",
              s"${args.length} arguments"
            )
          )

  /** NPER function: NPER(rate, pmt, pv, [fv], [type]) - number of periods */
  given nperFunctionParser: FunctionParser[Unit] with
    def name: String = "NPER"
    def arity: Arity = Arity.Range(3, 5)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(rateExpr, pmtExpr, pvExpr) =>
          scala.util.Right(
            TExpr.nper(
              TExpr.asNumericExpr(rateExpr),
              TExpr.asNumericExpr(pmtExpr),
              TExpr.asNumericExpr(pvExpr),
              None,
              None
            )
          )
        case List(rateExpr, pmtExpr, pvExpr, fvExpr) =>
          scala.util.Right(
            TExpr.nper(
              TExpr.asNumericExpr(rateExpr),
              TExpr.asNumericExpr(pmtExpr),
              TExpr.asNumericExpr(pvExpr),
              Some(TExpr.asNumericExpr(fvExpr)),
              None
            )
          )
        case List(rateExpr, pmtExpr, pvExpr, fvExpr, pmtTypeExpr) =>
          scala.util.Right(
            TExpr.nper(
              TExpr.asNumericExpr(rateExpr),
              TExpr.asNumericExpr(pmtExpr),
              TExpr.asNumericExpr(pvExpr),
              Some(TExpr.asNumericExpr(fvExpr)),
              Some(TExpr.asNumericExpr(pmtTypeExpr))
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "NPER",
              pos,
              "3-5 arguments (rate, pmt, pv, [fv], [type])",
              s"${args.length} arguments"
            )
          )

  /** RATE function: RATE(nper, pmt, pv, [fv], [type], [guess]) - interest rate per period */
  given rateFunctionParser: FunctionParser[Unit] with
    def name: String = "RATE"
    def arity: Arity = Arity.Range(3, 6)

    def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
      args match
        case List(nperExpr, pmtExpr, pvExpr) =>
          scala.util.Right(
            TExpr.rate(
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pmtExpr),
              TExpr.asNumericExpr(pvExpr),
              None,
              None,
              None
            )
          )
        case List(nperExpr, pmtExpr, pvExpr, fvExpr) =>
          scala.util.Right(
            TExpr.rate(
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pmtExpr),
              TExpr.asNumericExpr(pvExpr),
              Some(TExpr.asNumericExpr(fvExpr)),
              None,
              None
            )
          )
        case List(nperExpr, pmtExpr, pvExpr, fvExpr, pmtTypeExpr) =>
          scala.util.Right(
            TExpr.rate(
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pmtExpr),
              TExpr.asNumericExpr(pvExpr),
              Some(TExpr.asNumericExpr(fvExpr)),
              Some(TExpr.asNumericExpr(pmtTypeExpr)),
              None
            )
          )
        case List(nperExpr, pmtExpr, pvExpr, fvExpr, pmtTypeExpr, guessExpr) =>
          scala.util.Right(
            TExpr.rate(
              TExpr.asNumericExpr(nperExpr),
              TExpr.asNumericExpr(pmtExpr),
              TExpr.asNumericExpr(pvExpr),
              Some(TExpr.asNumericExpr(fvExpr)),
              Some(TExpr.asNumericExpr(pmtTypeExpr)),
              Some(TExpr.asNumericExpr(guessExpr))
            )
          )
        case _ =>
          scala.util.Left(
            ParseError.InvalidArguments(
              "RATE",
              pos,
              "3-6 arguments (nper, pmt, pv, [fv], [type], [guess])",
              s"${args.length} arguments"
            )
          )
