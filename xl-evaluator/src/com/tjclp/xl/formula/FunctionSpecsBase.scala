package com.tjclp.xl.formula

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
  type RangeCriteriaList = List[(CellRange, TExpr[Any])]
  type SumIfArgs = (CellRange, TExpr[Any], Option[CellRange])
  type CountIfArgs = (CellRange, TExpr[Any])
  type SumIfsArgs = (CellRange, RangeCriteriaList)
  type CountIfsArgs = RangeCriteriaList
  type AverageIfArgs = (CellRange, TExpr[Any], Option[CellRange])
  type AverageIfsArgs = (CellRange, RangeCriteriaList)
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
  type SumProductArgs = List[CellRange]
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
    ctx.evalExpr[Any](expr.asInstanceOf[TExpr[Any]])

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
      case CellValue.Bool(true) => BigDecimal(1)
      case CellValue.Bool(false) => BigDecimal(0)
      case CellValue.Formula(_, Some(cached)) => coerceToNumeric(cached)
      case _ => BigDecimal(0)
