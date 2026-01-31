package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator, ArrayArithmetic}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.CellRange
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
  type DatePairOptRange = (TExpr[LocalDate], TExpr[LocalDate], Option[CellRange])
  type DateIntOptRange = (TExpr[LocalDate], TExpr[Int], Option[CellRange])
  type DatePairOptBasis = (TExpr[LocalDate], TExpr[LocalDate], Option[TExpr[Int]])
  type IfArgs = (TExpr[Boolean], TExpr[Any], TExpr[Any])
  type IfErrorArgs = (TExpr[CellValue], TExpr[CellValue])
  type NoArgs = EmptyTuple
  type DateTripleInt = (TExpr[Int], TExpr[Int], TExpr[Int])
  type UnaryDate = TExpr[LocalDate]
  type AnyExpr = TExpr[Any]
  type NpvArgs = (TExpr[BigDecimal], CellRange)
  type IrrArgs = (CellRange, Option[TExpr[BigDecimal]])
  type VlookupArgs = (TExpr[CellValue], TExpr.RangeLocation, TExpr[Int], Option[TExpr[Boolean]])
  // GH-197: Changed to accept both ranges AND array expressions
  type SumProductArgs = List[ArgSpec.SumProductArg]
  type XLookupArgs = (
    AnyExpr,
    CellRange,
    CellRange,
    Option[AnyExpr],
    Option[TExpr[Int]],
    Option[TExpr[Int]]
  )
  type IndexArgs = (CellRange, TExpr[BigDecimal], Option[TExpr[BigDecimal]])
  type MatchArgs = (AnyExpr, CellRange, Option[TExpr[BigDecimal]])
  type AddressArgs = (
    TExpr[BigDecimal],
    TExpr[BigDecimal],
    Option[TExpr[BigDecimal]],
    Option[TExpr[Boolean]],
    Option[TExpr[String]]
  )
  type XnpvArgs = (TExpr[BigDecimal], CellRange, CellRange)
  type XirrArgs = (CellRange, CellRange, Option[TExpr[BigDecimal]])
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

  protected def toInt(value: ExprValue): Int =
    value match
      case ExprValue.Number(n) => n.toInt
      case _ => 0

  protected def coerceToNumeric(value: CellValue): BigDecimal =
    value match
      case CellValue.Number(n) => n
      case CellValue.Bool(b) => ArrayArithmetic.boolToNumeric(b)
      case CellValue.Formula(_, Some(cached)) => coerceToNumeric(cached)
      case _ => BigDecimal(0)

  /** Extract numeric value from CellValue, handling formulas with cached results. */
  protected def extractNumericValue(value: CellValue): Option[BigDecimal] =
    value match
      case CellValue.Number(n) => Some(n)
      case CellValue.Formula(_, Some(CellValue.Number(n))) => Some(n)
      case _ => None

  /** Extract text for matching, coercing numbers and booleans to strings. */
  protected def extractTextForMatch(cv: CellValue): Option[String] =
    cv match
      case CellValue.Text(s) => Some(s)
      case CellValue.Number(n) => Some(n.bigDecimal.stripTrailingZeros().toPlainString)
      case CellValue.Bool(b) => Some(if b then "TRUE" else "FALSE")
      case CellValue.Formula(_, Some(cached)) => extractTextForMatch(cached)
      case _ => None

  /** Extract numeric for matching, parsing text as numbers. */
  protected def extractNumericForMatch(cv: CellValue): Option[BigDecimal] =
    cv match
      case CellValue.Number(n) => Some(n)
      case CellValue.Text(s) => scala.util.Try(BigDecimal(s.trim)).toOption
      case CellValue.Bool(b) => Some(ArrayArithmetic.boolToNumeric(b))
      case CellValue.Formula(_, Some(cached)) => extractNumericForMatch(cached)
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
