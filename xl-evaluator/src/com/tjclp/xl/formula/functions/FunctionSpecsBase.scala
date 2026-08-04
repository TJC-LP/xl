package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue, BindingCoercion}
import com.tjclp.xl.formula.eval.{
  EvalError,
  Evaluator,
  ArrayArithmetic,
  ArrayResult,
  ScalarCoercion
}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import java.time.LocalDate

trait FunctionSpecsBase:
  protected given numericExpr: ArgSpec[TExpr[BigDecimal]] = ArgSpec.expr[BigDecimal]
  protected given stringExpr: ArgSpec[TExpr[String]] = ArgSpec.expr[String]
  protected given intExpr: ArgSpec[TExpr[Int]] = ArgSpec.expr[Int]
  protected given booleanExpr: ArgSpec[TExpr[Boolean]] = ArgSpec.expr[Boolean]
  protected given cellValueExpr: ArgSpec[TExpr[CellValue]] = ArgSpec.expr[CellValue]
  protected given dateExpr: ArgSpec[TExpr[LocalDate]] = ArgSpec.expr[LocalDate]
  protected given rangeLocation: ArgSpec[TExpr.RangeLocation] = ArgSpec.rangeLocation
  @deprecated(
    "Use rangeLocation (ArgSpec[TExpr.RangeLocation]) instead — see ArgSpec.cellRange",
    "0.18.0"
  )
  @annotation.nowarn("cat=deprecation") // forwarder must keep referencing the deprecated given
  protected given cellRange: ArgSpec[CellRange] = ArgSpec.cellRange

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  protected given anyExpr: ArgSpec[TExpr[Any]] with
    def describeParts: List[String] = List("value")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (TExpr[Any], List[TExpr[?]])] =
      args match
        case head :: tail => Right((head.asInstanceOf[TExpr[Any]], tail))
        case Nil =>
          Left(ParseError.InvalidArguments(fnName, pos, describe, "0 arguments"))

    def toValues(args: TExpr[Any]): List[ArgValue] =
      List(ArgValue.Expr(args))

    def map(
      args: TExpr[Any]
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): TExpr[Any] =
      mapExpr(args).asInstanceOf[TExpr[Any]]

  type UnaryNumeric = TExpr[BigDecimal]
  type BinaryNumeric = (TExpr[BigDecimal], TExpr[BigDecimal])
  type BinaryNumericOpt = (TExpr[BigDecimal], Option[TExpr[BigDecimal]])
  type UnaryText = TExpr[String]
  type BinaryTextInt = (TExpr[String], TExpr[Int])
  type TextIntInt = (TExpr[String], TExpr[Int], TExpr[Int])
  type FindArgs = (TExpr[String], TExpr[String], Option[TExpr[Int]])
  type SubstituteArgs = (TExpr[String], TExpr[String], TExpr[String], Option[TExpr[Int]])
  type TextArgs = (TExpr[Any], TExpr[String])
  type TextList = List[TExpr[String]]
  type UnaryBoolean = TExpr[Boolean]
  type BooleanList = List[TExpr[Boolean]]
  type UnaryCellValue = TExpr[CellValue]
  type UnaryRange = TExpr.RangeLocation

  // Variadic numeric: Either a range (aggregated) or a single numeric expression
  // Used for Excel-compatible SUM(1,2,3) / SUM(A1:A5,B1:B5) / SUM(A1,5,B1:B3)
  type NumericArg = Either[TExpr.RangeLocation, TExpr[BigDecimal]]
  type VariadicNumeric = List[NumericArg]
  type RangeCriteriaList = List[(TExpr.RangeLocation, TExpr[Any])]
  type SumIfArgs = (TExpr.RangeLocation, TExpr[Any], Option[TExpr.RangeLocation])
  type CountIfArgs = (TExpr.RangeLocation, TExpr[Any])
  type SumIfsArgs = (TExpr.RangeLocation, RangeCriteriaList)
  type CountIfsArgs = RangeCriteriaList
  type AverageIfArgs = (TExpr.RangeLocation, TExpr[Any], Option[TExpr.RangeLocation])
  type AverageIfsArgs = (TExpr.RangeLocation, RangeCriteriaList)
  type DateInt = (TExpr[LocalDate], TExpr[Int])
  type DatePairUnit = (TExpr[LocalDate], TExpr[LocalDate], TExpr[String])
  // GH-394: holiday args are RangeLocations (sheet-qualified ranges and defined names resolve
  // at evaluation through Evaluator.resolveRangeLocation, like every range-typed slot)
  type DatePairOptRange = (TExpr[LocalDate], TExpr[LocalDate], Option[TExpr.RangeLocation])
  type DateIntOptRange = (TExpr[LocalDate], TExpr[Int], Option[TExpr.RangeLocation])
  type DatePairOptBasis = (TExpr[LocalDate], TExpr[LocalDate], Option[TExpr[Int]])
  type IfArgs = (TExpr[Boolean], TExpr[Any], TExpr[Any])
  // GH-120 statistical functions over a range
  type RangeIntArgs = (TExpr.RangeLocation, TExpr[Int])
  type RankArgs = (TExpr[BigDecimal], TExpr.RangeLocation, Option[TExpr[Int]])
  type RangeNumArgs = (TExpr.RangeLocation, TExpr[BigDecimal])
  // GH-76 dynamic arrays (spill engine)
  type SequenceArgs =
    (TExpr[Int], Option[TExpr[Int]], Option[TExpr[BigDecimal]], Option[TExpr[BigDecimal]])
  type SortArgs = (TExpr.RangeLocation, Option[TExpr[Int]], Option[TExpr[Int]])
  type UniqueArgs = (TExpr.RangeLocation, Option[TExpr[Boolean]], Option[TExpr[Boolean]])
  type FilterArgs = (TExpr.RangeLocation, TExpr.RangeLocation, Option[TExpr[Any]])
  // GH-122 OFFSET: anchor ref + row/col offsets + optional height/width
  type OffsetArgs = (AnyExpr, TExpr[Int], TExpr[Int], Option[TExpr[Int]], Option[TExpr[Int]])
  // GH-274 INDIRECT: ref_text + optional a1 flag (FALSE = R1C1, documented-unsupported)
  type IndirectArgs = (TExpr[String], Option[TExpr[Boolean]])
  type IfErrorArgs = (TExpr[CellValue], TExpr[CellValue])
  type NoArgs = EmptyTuple
  type DateTripleInt = (TExpr[Int], TExpr[Int], TExpr[Int])
  type UnaryDate = TExpr[LocalDate]
  type AnyExpr = TExpr[Any]
  // GH-394: cashflow/date ranges are RangeLocations — `=XIRR(S!A1:B1, S!A2:B2)` and named
  // ranges resolve at evaluation; the CellRange-taking public builders wrap RangeLocation.Local
  type NpvArgs = (TExpr[BigDecimal], TExpr.RangeLocation)
  type IrrArgs = (TExpr.RangeLocation, Option[TExpr[BigDecimal]])
  type VlookupArgs = (TExpr[CellValue], TExpr.RangeLocation, TExpr[Int], Option[TExpr[Boolean]])
  // GH-197: Changed to accept both ranges AND array expressions
  type SumProductArgs = List[ArgSpec.SumProductArg]
  type XLookupArgs = (
    AnyExpr,
    TExpr.RangeLocation,
    TExpr.RangeLocation,
    Option[AnyExpr],
    Option[TExpr[Int]],
    Option[TExpr[Int]]
  )
  type IndexArgs = (TExpr.RangeLocation, TExpr[BigDecimal], Option[TExpr[BigDecimal]])
  type MatchArgs = (AnyExpr, TExpr.RangeLocation, Option[TExpr[BigDecimal]])
  type AddressArgs = (
    TExpr[BigDecimal],
    TExpr[BigDecimal],
    Option[TExpr[BigDecimal]],
    Option[TExpr[Boolean]],
    Option[TExpr[String]]
  )
  // GH-424 CELL: info_type + optional positional reference (never evaluated — its address is
  // the datum)
  type CellArgs = (TExpr[String], Option[AnyExpr])
  type XnpvArgs = (TExpr[BigDecimal], TExpr.RangeLocation, TExpr.RangeLocation)
  type XirrArgs = (TExpr.RangeLocation, TExpr.RangeLocation, Option[TExpr[BigDecimal]])
  type TvmArgs = (
    TExpr[BigDecimal],
    TExpr[BigDecimal],
    TExpr[BigDecimal],
    Option[TExpr[BigDecimal]],
    Option[TExpr[BigDecimal]]
  )
  type RateArgs = (
    TExpr[BigDecimal],
    TExpr[BigDecimal],
    TExpr[BigDecimal],
    Option[TExpr[BigDecimal]],
    Option[TExpr[BigDecimal]],
    Option[TExpr[BigDecimal]]
  )

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  protected def evalAny(ctx: EvalContext, expr: TExpr[?]): Either[EvalError, Any] =
    // Resolve PolyRef/SheetPolyRef to typed Ref before evaluation.
    // This fixes cell references used as criteria in SUMIFS, COUNTIF, etc.
    val resolved = expr match
      case _: TExpr.PolyRef | _: TExpr.SheetPolyRef => TExpr.asResolvedValueExpr(expr)
      case other => other
    ctx.evalExpr[Any](resolved.asInstanceOf[TExpr[Any]])

  /**
   * GH-333: evaluate an argument array-aware, materializing bare ranges.
   *
   * Unlike evalAny (whose scalar boundary collapses arrays and rejects bare RangeRefs), this yields
   * ArrayResults for range-shaped arguments and array-producing expressions — the IF
   * branch/condition convention, mirroring the evaluator's own evalMaybeArray operand handling.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  protected def evalMaybeArrayArg(ctx: EvalContext, expr: TExpr[?]): Either[EvalError, Any] =
    expr match
      case TExpr.RangeRef(range) =>
        Right(ArrayArithmetic.rangeToArray(range, ctx.sheet))
      case TExpr.SheetRange(sheetName, range) =>
        Evaluator
          .resolveRangeLocation(
            TExpr.RangeLocation.CrossSheet(sheetName, range),
            ctx.sheet,
            ctx.workbook
          )
          .map { case (targetSheet, _) => ArrayArithmetic.rangeToArray(range, targetSheet) }
      case _: TExpr.PolyRef | _: TExpr.SheetPolyRef =>
        ctx.evalArrayExpr(TExpr.asResolvedValueExpr(expr).asInstanceOf[TExpr[Any]])
      case other =>
        ctx.evalArrayExpr(other.asInstanceOf[TExpr[Any]])

  /** Normalize an evaluated value to an ArrayResult (scalars become 1x1). */
  protected def toCellArray(value: Any): ArrayResult = value match
    case arr: ArrayResult => arr
    case scalar => ArrayResult.single(ArrayArithmetic.anyToCellValue(scalar))

  protected def evalValue(ctx: EvalContext, expr: TExpr[?]): Either[EvalError, ExprValue] =
    evalAny(ctx, expr).map(ExprValue.from)

  protected def toCellValue(value: ExprValue): CellValue =
    value match
      case ExprValue.Cell(cv) => cv
      case ExprValue.Text(s) => CellValue.Text(s)
      case ExprValue.Number(n) => CellValue.Number(n)
      case ExprValue.Bool(b) => CellValue.Bool(b)
      case ExprValue.Date(d) => CellValue.DateTime(d.atStartOfDay())
      case ExprValue.DateTime(dt) => CellValue.DateTime(dt)
      case ExprValue.Opaque(other) => CellValue.Text(other.toString)

  /**
   * GH-307: total integer-argument extraction for ExprValue-evaluated arguments (EDATE/EOMONTH
   * months, WORKDAY days, YEARFRAC basis, XLOOKUP modes).
   *
   * Replaces the silent `toInt` (non-numeric → 0 → garbage results). Conventions come from the
   * shared ScalarCoercion Integer table: fractionals TRUNCATE toward zero (Excel truncates
   * months/days), numeric text parses ("3" → 3), booleans are TRUE = 1 / FALSE = 0, anything else
   * is a clean per-cell error naming the function (Excel: #VALUE!).
   */
  protected def toIntArg(fnName: String, value: ExprValue): Either[EvalError, Int] =
    val raw = value match
      case ExprValue.Number(n) => n
      case ExprValue.Text(s) => s
      case ExprValue.Bool(b) => b
      case ExprValue.Date(d) => d
      case ExprValue.DateTime(dt) => dt
      case ExprValue.Cell(cv) => cv
      case ExprValue.Opaque(other) => other
    ScalarCoercion.coerce(s"$fnName integer argument", raw, BindingCoercion.Integer).flatMap {
      case i: Int => Right(i)
      case other =>
        Left(EvalError.TypeMismatch(s"$fnName integer argument", "integer", s"$other"))
    }

  protected def coerceToNumeric(value: CellValue): BigDecimal =
    value match
      case CellValue.Number(n) => n
      case CellValue.Bool(b) => ArrayArithmetic.boolToNumeric(b)
      case CellValue.DateTime(dt) => BigDecimal(CellValue.dateTimeToExcelSerial(dt))
      case CellValue.Formula(_, Some(cached), _) => coerceToNumeric(cached)
      case _ => BigDecimal(0)

  /**
   * Extract numeric value from CellValue, handling formulas with cached results.
   *
   * GH-449: a DateTime cell yields its Excel serial — dates ARE numbers in Excel's aggregate
   * positions (=MIN/MAX over a date column are the earliest/latest date, =COUNT counts them,
   * =SUM sums the serials), mirroring TExprDecoders.decodeNumeric. Booleans stay skipped: Excel's
   * range aggregates ignore logicals, which is why this is not simply coerceToNumeric.
   */
  protected def extractNumericValue(value: CellValue): Option[BigDecimal] =
    value match
      case CellValue.Number(n) => Some(n)
      case CellValue.DateTime(dt) => Some(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case CellValue.Formula(_, Some(cached), _) => extractNumericValue(cached)
      case _ => None

  /** Extract the anchor ARef from a reference expression (cell ref, or the start of a range). */
  protected def extractARef(expr: TExpr[?]): Option[ARef] =
    expr match
      case TExpr.PolyRef(ref, _) => Some(ref)
      case TExpr.Ref(ref, _, _) => Some(ref)
      case TExpr.SheetPolyRef(_, ref, _) => Some(ref)
      case TExpr.SheetRef(_, ref, _, _) => Some(ref)
      case TExpr.RangeRef(range) => Some(range.start)
      case TExpr.SheetRange(_, range) => Some(range.start)
      case _ => None

  /**
   * GH-467: normalize a lookup value for MATCH/XLOOKUP comparison. Cell-ref lookup values arrive as
   * ExprValue.Cell and must be dereferenced to their underlying scalar before comparing (the
   * literal `"k2"` and a ref to a cell holding `"k2"` must behave identically); date/datetime
   * lookup values coerce to their Excel serial — dates ARE numbers in Excel's comparison plane
   * (mirroring GH-449's dateTimeToExcelSerial coercion at numeric boundaries).
   */
  protected def normalizeLookupValue(value: ExprValue): ExprValue =
    value match
      case ExprValue.Cell(CellValue.Number(n)) => ExprValue.Number(n)
      case ExprValue.Cell(CellValue.Text(s)) => ExprValue.Text(s)
      case ExprValue.Cell(CellValue.Bool(b)) => ExprValue.Bool(b)
      case ExprValue.Cell(CellValue.DateTime(dt)) =>
        ExprValue.Number(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case ExprValue.Cell(CellValue.Formula(_, Some(cached), _)) =>
        normalizeLookupValue(ExprValue.Cell(cached))
      case ExprValue.Date(d) =>
        ExprValue.Number(BigDecimal(CellValue.dateTimeToExcelSerial(d.atStartOfDay())))
      case ExprValue.DateTime(dt) =>
        ExprValue.Number(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case other => other

  /**
   * GH-467: exact equality for the MATCH/XLOOKUP lookup plane, over a lookup value already passed
   * through [[normalizeLookupValue]]. Total and blank-safe: a blank cell matches nothing (a
   * default-equal fallback here made blanks "match" every lookup value, shifting MATCH positions
   * over ranges with holes), mismatched types match nothing, and DateTime cells compare by their
   * Excel serial so date lookups find date columns.
   */
  protected def matchesLookupExact(cv: CellValue, lookup: ExprValue): Boolean =
    (cv, lookup) match
      case (CellValue.Number(n), ExprValue.Number(v)) => n == v
      case (CellValue.DateTime(dt), ExprValue.Number(v)) =>
        BigDecimal(CellValue.dateTimeToExcelSerial(dt)) == v
      case (CellValue.Text(s), ExprValue.Text(v)) => s.equalsIgnoreCase(v)
      case (CellValue.Bool(b), ExprValue.Bool(v)) => b == v
      case (CellValue.Error(e1), ExprValue.Cell(CellValue.Error(e2))) => e1 == e2
      case (CellValue.Formula(_, Some(cached), _), v) => matchesLookupExact(cached, v)
      case _ => false

  /** Extract text for matching, coercing numbers and booleans to strings. */
  protected def extractTextForMatch(cv: CellValue): Option[String] =
    cv match
      case CellValue.Text(s) => Some(s)
      case CellValue.Number(n) => Some(n.bigDecimal.stripTrailingZeros().toPlainString)
      case CellValue.Bool(b) => Some(if b then "TRUE" else "FALSE")
      case CellValue.Formula(_, Some(cached), _) => extractTextForMatch(cached)
      case _ => None

  /** Extract numeric for matching, parsing text as numbers. */
  protected def extractNumericForMatch(cv: CellValue): Option[BigDecimal] =
    cv match
      case CellValue.Number(n) => Some(n)
      case CellValue.Text(s) => scala.util.Try(BigDecimal(s.trim)).toOption
      case CellValue.Bool(b) => Some(ArrayArithmetic.boolToNumeric(b))
      case CellValue.Formula(_, Some(cached), _) => extractNumericForMatch(cached)
      case _ => None

  /** Validate that two ranges have the same dimensions. */
  protected def validateDimensions(
    range1: CellRange,
    range2: CellRange,
    fnName: String,
    range1Name: String,
    range2Name: String
  ): Either[EvalError, Unit] =
    if range1.width != range2.width || range1.height != range2.height then
      Left(
        EvalError.EvalFailed(
          s"$fnName: $range1Name and $range2Name must have same dimensions " +
            s"(${range1.height}×${range1.width} vs ${range2.height}×${range2.width})",
          Some(s"$fnName(${range1.toA1}, ..., ${range2.toA1})")
        )
      )
    else Right(())
