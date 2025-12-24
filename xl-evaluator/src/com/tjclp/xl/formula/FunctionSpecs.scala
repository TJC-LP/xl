package com.tjclp.xl.formula

import com.tjclp.xl.CellRange
import com.tjclp.xl.cells.{CellError, CellValue}
import java.time.LocalDate
import java.time.temporal.ChronoUnit

object FunctionSpecs:
  private given numericExpr: ArgSpec[TExpr[BigDecimal]] = ArgSpec.expr[BigDecimal]
  private given stringExpr: ArgSpec[TExpr[String]] = ArgSpec.expr[String]
  private given intExpr: ArgSpec[TExpr[Int]] = ArgSpec.expr[Int]
  private given booleanExpr: ArgSpec[TExpr[Boolean]] = ArgSpec.expr[Boolean]
  private given cellValueExpr: ArgSpec[TExpr[CellValue]] = ArgSpec.expr[CellValue]
  private given dateExpr: ArgSpec[TExpr[LocalDate]] = ArgSpec.expr[LocalDate]
  private given rangeLocation: ArgSpec[TExpr.RangeLocation] = ArgSpec.rangeLocation
  private given cellRange: ArgSpec[CellRange] = ArgSpec.cellRange

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private given anyExpr: ArgSpec[TExpr[Any]] with
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
  type CriteriaExpr = TExpr[Any]
  type RangeCriteria = (CellRange, CriteriaExpr)
  type RangeCriteriaList = List[RangeCriteria]
  type SumIfArgs = (CellRange, CriteriaExpr, Option[CellRange])
  type CountIfArgs = (CellRange, CriteriaExpr)
  type SumIfsArgs = (CellRange, RangeCriteriaList)
  type CountIfsArgs = RangeCriteriaList
  type AverageIfArgs = (CellRange, CriteriaExpr, Option[CellRange])
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
  private def evalAny(ctx: EvalContext, expr: TExpr[?]): Either[EvalError, Any] =
    ctx.evalExpr[Any](expr.asInstanceOf[TExpr[Any]])

  private def toCellValue(value: Any): CellValue =
    value match
      case cv: CellValue => cv
      case s: String => CellValue.Text(s)
      case n: BigDecimal => CellValue.Number(n)
      case b: Boolean => CellValue.Bool(b)
      case n: Int => CellValue.Number(BigDecimal(n))
      case n: Long => CellValue.Number(BigDecimal(n))
      case n: Double => CellValue.Number(BigDecimal(n))
      case d: java.time.LocalDate => CellValue.DateTime(d.atStartOfDay())
      case dt: java.time.LocalDateTime => CellValue.DateTime(dt)
      case other => CellValue.Text(other.toString)

  private def toInt(value: Any): Int =
    value match
      case i: Int => i
      case bd: BigDecimal => bd.toInt
      case n: Number => n.intValue()
      case _ => 0

  private def roundToDigits(
    value: BigDecimal,
    numDigits: Int,
    mode: BigDecimal.RoundingMode.Value
  ): BigDecimal =
    if numDigits >= 0 then value.setScale(numDigits, mode)
    else
      val scale = math.pow(10, -numDigits).toLong
      val divided = value / scale
      val rounded = divided.setScale(0, mode)
      rounded * scale

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def countWorkingDays(
    earlier: LocalDate,
    later: LocalDate,
    holidays: Set[LocalDate]
  ): Int =
    var count = 0
    var current = earlier
    while !current.isAfter(later) do
      val dayOfWeek = current.getDayOfWeek
      if dayOfWeek != java.time.DayOfWeek.SATURDAY &&
        dayOfWeek != java.time.DayOfWeek.SUNDAY &&
        !holidays.contains(current)
      then count += 1
      current = current.plusDays(1)
    count

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def addWorkingDays(
    start: LocalDate,
    daysValue: Int,
    holidays: Set[LocalDate]
  ): LocalDate =
    var remaining = daysValue
    var current = start
    val direction = if remaining >= 0 then 1L else -1L

    while remaining != 0 do
      current = current.plusDays(direction)
      val dayOfWeek = current.getDayOfWeek
      if dayOfWeek != java.time.DayOfWeek.SATURDAY &&
        dayOfWeek != java.time.DayOfWeek.SUNDAY &&
        !holidays.contains(current)
      then remaining -= direction.toInt
    current

  val abs: FunctionSpec[BigDecimal] { type Args = UnaryNumeric } =
    FunctionSpec.simple[BigDecimal, UnaryNumeric]("ABS", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(_.abs)
    }

  val sqrt: FunctionSpec[BigDecimal] { type Args = UnaryNumeric } =
    FunctionSpec.simple[BigDecimal, UnaryNumeric]("SQRT", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).flatMap { value =>
        if value < 0 then
          Left(
            EvalError.EvalFailed(
              s"SQRT: cannot take square root of negative number ($value)",
              Some(s"SQRT($value)")
            )
          )
        else Right(BigDecimal(Math.sqrt(value.toDouble)))
      }
    }

  val round: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric]("ROUND", Arity.two) { (args, ctx) =>
      val (valueExpr, numDigitsExpr) = args
      for
        value <- ctx.evalExpr(valueExpr)
        numDigits <- ctx.evalExpr(numDigitsExpr)
      yield roundToDigits(value, numDigits.toInt, BigDecimal.RoundingMode.HALF_UP)
    }

  val roundUp: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "ROUNDUP",
      Arity.two
    ) { (args, ctx) =>
      val (valueExpr, numDigitsExpr) = args
      for
        value <- ctx.evalExpr(valueExpr)
        numDigits <- ctx.evalExpr(numDigitsExpr)
      yield
        val mode =
          if value >= 0 then BigDecimal.RoundingMode.CEILING
          else BigDecimal.RoundingMode.FLOOR
        roundToDigits(value, numDigits.toInt, mode)
    }

  val roundDown: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "ROUNDDOWN",
      Arity.two
    ) { (args, ctx) =>
      val (valueExpr, numDigitsExpr) = args
      for
        value <- ctx.evalExpr(valueExpr)
        numDigits <- ctx.evalExpr(numDigitsExpr)
      yield
        val mode =
          if value >= 0 then BigDecimal.RoundingMode.FLOOR
          else BigDecimal.RoundingMode.CEILING
        roundToDigits(value, numDigits.toInt, mode)
    }

  val mod: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric]("MOD", Arity.two) { (args, ctx) =>
      val (numberExpr, divisorExpr) = args
      for
        number <- ctx.evalExpr(numberExpr)
        divisor <- ctx.evalExpr(divisorExpr)
        result <-
          if divisor == 0 then
            Left(EvalError.EvalFailed("MOD: division by zero", Some(s"MOD($number, $divisor)")))
          else
            val quotient = (number / divisor).setScale(0, BigDecimal.RoundingMode.FLOOR)
            Right(number - divisor * quotient)
      yield result
    }

  val power: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "POWER",
      Arity.two
    ) { (args, ctx) =>
      val (numberExpr, powerExpr) = args
      for
        number <- ctx.evalExpr(numberExpr)
        power <- ctx.evalExpr(powerExpr)
      yield BigDecimal(Math.pow(number.toDouble, power.toDouble))
    }

  val log: FunctionSpec[BigDecimal] { type Args = BinaryNumericOpt } =
    FunctionSpec.simple[BigDecimal, BinaryNumericOpt](
      "LOG",
      Arity.Range(1, 2)
    ) { (args, ctx) =>
      val (numberExpr, baseExprOpt) = args
      val baseExpr = baseExprOpt.getOrElse(TExpr.Lit(BigDecimal(10)))
      for
        number <- ctx.evalExpr(numberExpr)
        base <- ctx.evalExpr(baseExpr)
        result <-
          if number <= 0 then
            Left(
              EvalError.EvalFailed(
                s"LOG: argument must be positive ($number)",
                Some(s"LOG($number, $base)")
              )
            )
          else if base <= 0 then
            Left(
              EvalError.EvalFailed(
                s"LOG: base must be positive ($base)",
                Some(s"LOG($number, $base)")
              )
            )
          else if base == 1 then
            Left(EvalError.EvalFailed("LOG: base cannot be 1", Some(s"LOG($number, $base)")))
          else Right(BigDecimal(Math.log(number.toDouble) / Math.log(base.toDouble)))
      yield result
    }

  val ln: FunctionSpec[BigDecimal] { type Args = UnaryNumeric } =
    FunctionSpec.simple[BigDecimal, UnaryNumeric]("LN", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).flatMap { value =>
        if value <= 0 then
          Left(
            EvalError.EvalFailed(s"LN: argument must be positive ($value)", Some(s"LN($value)"))
          )
        else Right(BigDecimal(Math.log(value.toDouble)))
      }
    }

  val exp: FunctionSpec[BigDecimal] { type Args = UnaryNumeric } =
    FunctionSpec.simple[BigDecimal, UnaryNumeric]("EXP", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map { value =>
        BigDecimal(Math.exp(value.toDouble))
      }
    }

  val floor: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric]("FLOOR", Arity.two) { (args, ctx) =>
      val (numberExpr, significanceExpr) = args
      for
        number <- ctx.evalExpr(numberExpr)
        significance <- ctx.evalExpr(significanceExpr)
        result <-
          if significance == 0 then
            Left(
              EvalError.EvalFailed(
                "FLOOR: significance cannot be zero",
                Some(s"FLOOR($number, $significance)")
              )
            )
          else if (number > 0 && significance < 0) || (number < 0 && significance > 0) then
            Left(
              EvalError.EvalFailed(
                "FLOOR: number and significance must have same sign",
                Some(s"FLOOR($number, $significance)")
              )
            )
          else
            val quotient = (number / significance).setScale(0, BigDecimal.RoundingMode.FLOOR)
            Right(quotient * significance)
      yield result
    }

  val ceiling: FunctionSpec[BigDecimal] { type Args = BinaryNumeric } =
    FunctionSpec.simple[BigDecimal, BinaryNumeric](
      "CEILING",
      Arity.two
    ) { (args, ctx) =>
      val (numberExpr, significanceExpr) = args
      for
        number <- ctx.evalExpr(numberExpr)
        significance <- ctx.evalExpr(significanceExpr)
        result <-
          if significance == 0 then
            Left(
              EvalError.EvalFailed(
                "CEILING: significance cannot be zero",
                Some(s"CEILING($number, $significance)")
              )
            )
          else if (number > 0 && significance < 0) || (number < 0 && significance > 0) then
            Left(
              EvalError.EvalFailed(
                "CEILING: number and significance must have same sign",
                Some(s"CEILING($number, $significance)")
              )
            )
          else
            val quotient =
              (number / significance).setScale(0, BigDecimal.RoundingMode.CEILING)
            Right(quotient * significance)
      yield result
    }

  val trunc: FunctionSpec[BigDecimal] { type Args = BinaryNumericOpt } =
    FunctionSpec.simple[BigDecimal, BinaryNumericOpt](
      "TRUNC",
      Arity.Range(1, 2)
    ) { (args, ctx) =>
      val (numberExpr, numDigitsExprOpt) = args
      val numDigitsExpr = numDigitsExprOpt.getOrElse(TExpr.Lit(BigDecimal(0)))
      for
        number <- ctx.evalExpr(numberExpr)
        numDigits <- ctx.evalExpr(numDigitsExpr)
      yield roundToDigits(number, numDigits.toInt, BigDecimal.RoundingMode.DOWN)
    }

  val sign: FunctionSpec[BigDecimal] { type Args = UnaryNumeric } =
    FunctionSpec.simple[BigDecimal, UnaryNumeric]("SIGN", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map { value =>
        if value > 0 then BigDecimal(1)
        else if value < 0 then BigDecimal(-1)
        else BigDecimal(0)
      }
    }

  val int: FunctionSpec[BigDecimal] { type Args = UnaryNumeric } =
    FunctionSpec.simple[BigDecimal, UnaryNumeric]("INT", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(_.setScale(0, BigDecimal.RoundingMode.FLOOR))
    }

  private def aggregateSpec(
    name: String
  ): FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    FunctionSpec.simple[BigDecimal, UnaryRange](name, Arity.one) { (location, ctx) =>
      ctx.evalExpr(TExpr.Aggregate(name, location))
    }

  private def toLocation(range: CellRange): TExpr.RangeLocation =
    TExpr.RangeLocation.Local(range)

  private def toConditions(
    pairs: RangeCriteriaList
  ): List[(TExpr.RangeLocation, TExpr[?])] =
    pairs.map { case (range, criteria) => (toLocation(range), criteria) }

  val sum: FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    aggregateSpec("SUM")

  val count: FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    aggregateSpec("COUNT")

  val counta: FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    aggregateSpec("COUNTA")

  val countblank: FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    aggregateSpec("COUNTBLANK")

  val average: FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    aggregateSpec("AVERAGE")

  val min: FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    aggregateSpec("MIN")

  val max: FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    aggregateSpec("MAX")

  val median: FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    aggregateSpec("MEDIAN")

  val stdev: FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    aggregateSpec("STDEV")

  val stdevp: FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    aggregateSpec("STDEVP")

  val variance: FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    aggregateSpec("VAR")

  val variancep: FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    aggregateSpec("VARP")

  val sumif: FunctionSpec[BigDecimal] { type Args = SumIfArgs } =
    FunctionSpec.simple[BigDecimal, SumIfArgs]("SUMIF", Arity.Range(2, 3)) { (args, ctx) =>
      val (range, criteria, sumRangeOpt) = args
      ctx.evalExpr(TExpr.SumIf(toLocation(range), criteria, sumRangeOpt.map(toLocation)))
    }

  val countif: FunctionSpec[BigDecimal] { type Args = CountIfArgs } =
    FunctionSpec.simple[BigDecimal, CountIfArgs]("COUNTIF", Arity.two) { (args, ctx) =>
      val (range, criteria) = args
      ctx.evalExpr(TExpr.CountIf(toLocation(range), criteria))
    }

  val sumifs: FunctionSpec[BigDecimal] { type Args = SumIfsArgs } =
    FunctionSpec.simple[BigDecimal, SumIfsArgs]("SUMIFS", Arity.AtLeast(3)) { (args, ctx) =>
      val (sumRange, conditions) = args
      ctx.evalExpr(TExpr.SumIfs(toLocation(sumRange), toConditions(conditions)))
    }

  val countifs: FunctionSpec[BigDecimal] { type Args = CountIfsArgs } =
    FunctionSpec.simple[BigDecimal, CountIfsArgs]("COUNTIFS", Arity.AtLeast(2)) {
      (conditions, ctx) =>
        ctx.evalExpr(TExpr.CountIfs(toConditions(conditions)))
    }

  val averageif: FunctionSpec[BigDecimal] { type Args = AverageIfArgs } =
    FunctionSpec.simple[BigDecimal, AverageIfArgs]("AVERAGEIF", Arity.Range(2, 3)) { (args, ctx) =>
      val (range, criteria, avgRangeOpt) = args
      ctx.evalExpr(TExpr.AverageIf(toLocation(range), criteria, avgRangeOpt.map(toLocation)))
    }

  val averageifs: FunctionSpec[BigDecimal] { type Args = AverageIfsArgs } =
    FunctionSpec.simple[BigDecimal, AverageIfsArgs]("AVERAGEIFS", Arity.AtLeast(3)) { (args, ctx) =>
      val (avgRange, conditions) = args
      ctx.evalExpr(TExpr.AverageIfs(toLocation(avgRange), toConditions(conditions)))
    }

  val and: FunctionSpec[Boolean] { type Args = BooleanList } =
    FunctionSpec.simple[Boolean, BooleanList]("AND", Arity.atLeastOne) { (args, ctx) =>
      @annotation.tailrec
      def loop(remaining: List[TExpr[Boolean]]): Either[EvalError, Boolean] =
        remaining match
          case Nil => Right(true)
          case head :: tail =>
            ctx.evalExpr(head) match
              case Left(err) => Left(err)
              case Right(value) =>
                if !value then Right(false)
                else loop(tail)
      loop(args)
    }

  val or: FunctionSpec[Boolean] { type Args = BooleanList } =
    FunctionSpec.simple[Boolean, BooleanList]("OR", Arity.atLeastOne) { (args, ctx) =>
      @annotation.tailrec
      def loop(remaining: List[TExpr[Boolean]]): Either[EvalError, Boolean] =
        remaining match
          case Nil => Right(false)
          case head :: tail =>
            ctx.evalExpr(head) match
              case Left(err) => Left(err)
              case Right(value) =>
                if value then Right(true)
                else loop(tail)
      loop(args)
    }

  val not: FunctionSpec[Boolean] { type Args = UnaryBoolean } =
    FunctionSpec.simple[Boolean, UnaryBoolean]("NOT", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(value => !value)
    }

  val ifFn: FunctionSpec[Any] { type Args = IfArgs } =
    FunctionSpec.simple[Any, IfArgs]("IF", Arity.three) { (args, ctx) =>
      val (condExpr, ifTrueExpr, ifFalseExpr) = args
      for
        cond <- ctx.evalExpr(condExpr)
        result <- if cond then evalAny(ctx, ifTrueExpr) else evalAny(ctx, ifFalseExpr)
      yield result
    }

  val iferror: FunctionSpec[CellValue] { type Args = IfErrorArgs } =
    FunctionSpec.simple[CellValue, IfErrorArgs]("IFERROR", Arity.two) { (args, ctx) =>
      val (valueExpr, valueIfErrorExpr) = args
      evalAny(ctx, valueExpr) match
        case Left(_) =>
          evalAny(ctx, valueIfErrorExpr).map(toCellValue)
        case Right(cv: CellValue) =>
          cv match
            case CellValue.Error(_) =>
              evalAny(ctx, valueIfErrorExpr).map(toCellValue)
            case _ => Right(cv)
        case Right(other) =>
          Right(toCellValue(other))
    }

  val iserror: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISERROR", Arity.one) { (expr, ctx) =>
      evalAny(ctx, expr) match
        case Left(_) => Right(true)
        case Right(cv: CellValue) =>
          cv match
            case CellValue.Error(_) => Right(true)
            case _ => Right(false)
        case Right(_) => Right(false)
    }

  val iserr: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISERR", Arity.one) { (expr, ctx) =>
      evalAny(ctx, expr) match
        case Left(_) => Right(true)
        case Right(cv: CellValue) =>
          cv match
            case CellValue.Error(err) => Right(err != CellError.NA)
            case _ => Right(false)
        case Right(_) => Right(false)
    }

  val isnumber: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISNUMBER", Arity.one) { (expr, ctx) =>
      evalAny(ctx, expr) match
        case Left(_) => Right(false)
        case Right(cv: CellValue) =>
          cv match
            case CellValue.Number(_) => Right(true)
            case CellValue.Formula(_, Some(CellValue.Number(_))) => Right(true)
            case _ => Right(false)
        case Right(_: BigDecimal) => Right(true)
        case Right(_: Int) => Right(true)
        case Right(_: Long) => Right(true)
        case Right(_: Double) => Right(true)
        case Right(_) => Right(false)
    }

  val istext: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISTEXT", Arity.one) { (expr, ctx) =>
      evalAny(ctx, expr) match
        case Left(_) => Right(false)
        case Right(cv: CellValue) =>
          cv match
            case CellValue.Text(_) => Right(true)
            case CellValue.Formula(_, Some(CellValue.Text(_))) => Right(true)
            case _ => Right(false)
        case Right(_: String) => Right(true)
        case Right(_) => Right(false)
    }

  val isblank: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISBLANK", Arity.one) { (expr, ctx) =>
      evalAny(ctx, expr) match
        case Left(_) => Right(false)
        case Right(cv: CellValue) =>
          cv match
            case CellValue.Empty => Right(true)
            case _ => Right(false)
        case Right(_) => Right(false)
    }

  val today: FunctionSpec[LocalDate] { type Args = NoArgs } =
    FunctionSpec.simple[LocalDate, NoArgs](
      "TODAY",
      Arity.none,
      flags = FunctionFlags(returnsDate = true)
    ) { (_, ctx) =>
      Right(ctx.clock.today())
    }

  val now: FunctionSpec[java.time.LocalDateTime] { type Args = NoArgs } =
    FunctionSpec.simple[java.time.LocalDateTime, NoArgs](
      "NOW",
      Arity.none,
      flags = FunctionFlags(returnsTime = true)
    ) { (_, ctx) =>
      Right(ctx.clock.now())
    }

  val pi: FunctionSpec[BigDecimal] { type Args = NoArgs } =
    FunctionSpec.simple[BigDecimal, NoArgs]("PI", Arity.none) { (_, _) =>
      Right(BigDecimal(Math.PI))
    }

  val date: FunctionSpec[LocalDate] { type Args = DateTripleInt } =
    FunctionSpec.simple[LocalDate, DateTripleInt](
      "DATE",
      Arity.three,
      flags = FunctionFlags(returnsDate = true)
    ) { (args, ctx) =>
      val (yearExpr, monthExpr, dayExpr) = args
      ctx.evalExpr(TExpr.Date(yearExpr, monthExpr, dayExpr))
    }

  val year: FunctionSpec[BigDecimal] { type Args = UnaryDate } =
    FunctionSpec.simple[BigDecimal, UnaryDate]("YEAR", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(TExpr.Year(expr))
    }

  val month: FunctionSpec[BigDecimal] { type Args = UnaryDate } =
    FunctionSpec.simple[BigDecimal, UnaryDate]("MONTH", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(TExpr.Month(expr))
    }

  val day: FunctionSpec[BigDecimal] { type Args = UnaryDate } =
    FunctionSpec.simple[BigDecimal, UnaryDate]("DAY", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(TExpr.Day(expr))
    }

  val eomonth: FunctionSpec[LocalDate] { type Args = DateInt } =
    FunctionSpec.simple[LocalDate, DateInt](
      "EOMONTH",
      Arity.two,
      flags = FunctionFlags(returnsDate = true)
    ) { (args, ctx) =>
      val (startDateExpr, monthsExpr) = args
      for
        date <- ctx.evalExpr(startDateExpr)
        monthsRaw <- evalAny(ctx, monthsExpr)
      yield
        val monthsValue = toInt(monthsRaw)
        val targetMonth = date.plusMonths(monthsValue.toLong)
        targetMonth.withDayOfMonth(targetMonth.lengthOfMonth)
    }

  val edate: FunctionSpec[LocalDate] { type Args = DateInt } =
    FunctionSpec.simple[LocalDate, DateInt](
      "EDATE",
      Arity.two,
      flags = FunctionFlags(returnsDate = true)
    ) { (args, ctx) =>
      val (startDateExpr, monthsExpr) = args
      for
        date <- ctx.evalExpr(startDateExpr)
        monthsRaw <- evalAny(ctx, monthsExpr)
      yield date.plusMonths(toInt(monthsRaw).toLong)
    }

  val datedif: FunctionSpec[BigDecimal] { type Args = DatePairUnit } =
    FunctionSpec.simple[BigDecimal, DatePairUnit]("DATEDIF", Arity.three) { (args, ctx) =>
      val (startDateExpr, endDateExpr, unitExpr) = args
      for
        start <- ctx.evalExpr(startDateExpr)
        end <- ctx.evalExpr(endDateExpr)
        unitStr <- ctx.evalExpr(unitExpr)
        result <- unitStr.toUpperCase match
          case "Y" =>
            Right(BigDecimal(ChronoUnit.YEARS.between(start, end)))
          case "M" =>
            Right(BigDecimal(ChronoUnit.MONTHS.between(start, end)))
          case "D" =>
            Right(BigDecimal(ChronoUnit.DAYS.between(start, end)))
          case "MD" =>
            val daysDiff = end.getDayOfMonth - start.getDayOfMonth
            val adjustedDays =
              if daysDiff >= 0 then daysDiff
              else
                val prevMonthLength = end.minusMonths(1).lengthOfMonth
                val effectiveStartDay = math.min(start.getDayOfMonth, prevMonthLength)
                prevMonthLength - effectiveStartDay + end.getDayOfMonth
            Right(BigDecimal(adjustedDays))
          case "YM" =>
            val monthsDiff = end.getMonthValue - start.getMonthValue
            val adjustedMonths = if monthsDiff < 0 then 12 + monthsDiff else monthsDiff
            val finalMonths =
              if end.getDayOfMonth < start.getDayOfMonth then (adjustedMonths - 1 + 12) % 12
              else adjustedMonths
            Right(BigDecimal(finalMonths))
          case "YD" =>
            val startAdjusted = start.withYear(end.getYear)
            val days =
              if startAdjusted.isAfter(end) then
                ChronoUnit.DAYS.between(startAdjusted.minusYears(1), end)
              else ChronoUnit.DAYS.between(startAdjusted, end)
            Right(BigDecimal(days))
          case other =>
            Left(
              EvalError.EvalFailed(
                s"DATEDIF: invalid unit '$other'. Valid units: Y, M, D, MD, YM, YD",
                Some("DATEDIF(start, end, unit)")
              )
            )
      yield result
    }

  val networkdays: FunctionSpec[BigDecimal] { type Args = DatePairOptRange } =
    FunctionSpec.simple[BigDecimal, DatePairOptRange]("NETWORKDAYS", Arity.Range(2, 3)) {
      (args, ctx) =>
        val (startDateExpr, endDateExpr, holidaysOpt) = args
        for
          start <- ctx.evalExpr(startDateExpr)
          end <- ctx.evalExpr(endDateExpr)
        yield
          val holidays: Set[LocalDate] = holidaysOpt
            .map { range =>
              range.cells
                .flatMap(ref => TExpr.decodeDate(ctx.sheet(ref)).toOption)
                .toSet
            }
            .getOrElse(Set.empty)

          val (earlier, later) = if start.isBefore(end) then (start, end) else (end, start)
          val count = countWorkingDays(earlier, later, holidays)
          BigDecimal(if start.isBefore(end) || start.isEqual(end) then count else -count)
    }

  val workday: FunctionSpec[LocalDate] { type Args = DateIntOptRange } =
    FunctionSpec.simple[LocalDate, DateIntOptRange](
      "WORKDAY",
      Arity.Range(2, 3),
      flags = FunctionFlags(returnsDate = true)
    ) { (args, ctx) =>
      val (startDateExpr, daysExpr, holidaysOpt) = args
      for
        start <- ctx.evalExpr(startDateExpr)
        daysRaw <- evalAny(ctx, daysExpr)
      yield
        val daysValue = toInt(daysRaw)
        val holidays: Set[LocalDate] = holidaysOpt
          .map { range =>
            range.cells
              .flatMap(ref => TExpr.decodeDate(ctx.sheet(ref)).toOption)
              .toSet
          }
          .getOrElse(Set.empty)
        addWorkingDays(start, daysValue, holidays)
    }

  val yearfrac: FunctionSpec[BigDecimal] { type Args = DatePairOptBasis } =
    FunctionSpec.simple[BigDecimal, DatePairOptBasis](
      "YEARFRAC",
      Arity.Range(2, 3),
      renderFn = Some { (args, printer) =>
        val (startDateExpr, endDateExpr, basisOpt) = args
        val rendered = basisOpt match
          case None =>
            List(printer.expr(startDateExpr), printer.expr(endDateExpr))
          case Some(TExpr.Lit(0)) =>
            List(printer.expr(startDateExpr), printer.expr(endDateExpr))
          case Some(basisExpr) =>
            List(
              printer.expr(startDateExpr),
              printer.expr(endDateExpr),
              printer.expr(basisExpr)
            )
        s"YEARFRAC(${rendered.mkString(", ")})"
      }
    ) { (args, ctx) =>
      val (startDateExpr, endDateExpr, basisOpt) = args
      val basisValueEither = basisOpt match
        case Some(expr) => evalAny(ctx, expr).map(toInt)
        case None => Right(0)
      for
        start <- ctx.evalExpr(startDateExpr)
        end <- ctx.evalExpr(endDateExpr)
        basisValue <- basisValueEither
        result <- basisValue match
          case 0 =>
            val d1 = math.min(30, start.getDayOfMonth)
            val d2 = if start.getDayOfMonth == 31 then 30 else end.getDayOfMonth
            val m1 = start.getMonthValue
            val m2 = end.getMonthValue
            val y1 = start.getYear
            val y2 = end.getYear
            val days = ((y2 - y1) * 360) + ((m2 - m1) * 30) + (d2 - d1)
            Right(BigDecimal(days) / BigDecimal(360))
          case 1 =>
            val daysBetween = ChronoUnit.DAYS.between(start, end)
            val year = start.getYear
            val daysInYear = if java.time.Year.isLeap(year) then 366 else 365
            Right(BigDecimal(daysBetween) / BigDecimal(daysInYear))
          case 2 =>
            val daysBetween = ChronoUnit.DAYS.between(start, end)
            Right(BigDecimal(daysBetween) / BigDecimal(360))
          case 3 =>
            val daysBetween = ChronoUnit.DAYS.between(start, end)
            Right(BigDecimal(daysBetween) / BigDecimal(365))
          case 4 =>
            val d1 = math.min(30, start.getDayOfMonth)
            val d2 = math.min(30, end.getDayOfMonth)
            val m1 = start.getMonthValue
            val m2 = end.getMonthValue
            val y1 = start.getYear
            val y2 = end.getYear
            val days = ((y2 - y1) * 360) + ((m2 - m1) * 30) + (d2 - d1)
            Right(BigDecimal(days) / BigDecimal(360))
          case other =>
            Left(
              EvalError.EvalFailed(
                s"YEARFRAC: invalid basis $other. Valid values: 0-4",
                Some("YEARFRAC(start, end, [basis])")
              )
            )
      yield result
    }

  val npv: FunctionSpec[BigDecimal] { type Args = NpvArgs } =
    FunctionSpec.simple[BigDecimal, NpvArgs]("NPV", Arity.two) { (args, ctx) =>
      val (rateExpr, range) = args
      ctx.evalExpr(TExpr.npv(rateExpr, range))
    }

  val irr: FunctionSpec[BigDecimal] { type Args = IrrArgs } =
    FunctionSpec.simple[BigDecimal, IrrArgs]("IRR", Arity.Range(1, 2)) { (args, ctx) =>
      val (range, guessOpt) = args
      ctx.evalExpr(TExpr.irr(range, guessOpt))
    }

  val xnpv: FunctionSpec[BigDecimal] { type Args = XnpvArgs } =
    FunctionSpec.simple[BigDecimal, XnpvArgs]("XNPV", Arity.three) { (args, ctx) =>
      val (rateExpr, valuesRange, datesRange) = args
      ctx.evalExpr(TExpr.xnpv(rateExpr, valuesRange, datesRange))
    }

  val xirr: FunctionSpec[BigDecimal] { type Args = XirrArgs } =
    FunctionSpec.simple[BigDecimal, XirrArgs]("XIRR", Arity.Range(2, 3)) { (args, ctx) =>
      val (valuesRange, datesRange, guessOpt) = args
      ctx.evalExpr(TExpr.xirr(valuesRange, datesRange, guessOpt))
    }

  val pmt: FunctionSpec[BigDecimal] { type Args = TvmArgs } =
    FunctionSpec.simple[BigDecimal, TvmArgs]("PMT", Arity.Range(3, 5)) { (args, ctx) =>
      val (rateExpr, nperExpr, pvExpr, fvOpt, typeOpt) = args
      ctx.evalExpr(TExpr.pmt(rateExpr, nperExpr, pvExpr, fvOpt, typeOpt))
    }

  val fv: FunctionSpec[BigDecimal] { type Args = TvmArgs } =
    FunctionSpec.simple[BigDecimal, TvmArgs]("FV", Arity.Range(3, 5)) { (args, ctx) =>
      val (rateExpr, nperExpr, pmtExpr, pvOpt, typeOpt) = args
      ctx.evalExpr(TExpr.fv(rateExpr, nperExpr, pmtExpr, pvOpt, typeOpt))
    }

  val pv: FunctionSpec[BigDecimal] { type Args = TvmArgs } =
    FunctionSpec.simple[BigDecimal, TvmArgs]("PV", Arity.Range(3, 5)) { (args, ctx) =>
      val (rateExpr, nperExpr, pmtExpr, fvOpt, typeOpt) = args
      ctx.evalExpr(TExpr.pv(rateExpr, nperExpr, pmtExpr, fvOpt, typeOpt))
    }

  val nper: FunctionSpec[BigDecimal] { type Args = TvmArgs } =
    FunctionSpec.simple[BigDecimal, TvmArgs]("NPER", Arity.Range(3, 5)) { (args, ctx) =>
      val (rateExpr, pmtExpr, pvExpr, fvOpt, typeOpt) = args
      ctx.evalExpr(TExpr.nper(rateExpr, pmtExpr, pvExpr, fvOpt, typeOpt))
    }

  val rate: FunctionSpec[BigDecimal] { type Args = RateArgs } =
    FunctionSpec.simple[BigDecimal, RateArgs]("RATE", Arity.Range(3, 6)) { (args, ctx) =>
      val (nperExpr, pmtExpr, pvExpr, fvOpt, typeOpt, guessOpt) = args
      ctx.evalExpr(TExpr.rate(nperExpr, pmtExpr, pvExpr, fvOpt, typeOpt, guessOpt))
    }

  val vlookup: FunctionSpec[CellValue] { type Args = VlookupArgs } =
    FunctionSpec.simple[CellValue, VlookupArgs]("VLOOKUP", Arity.Range(3, 4)) { (args, ctx) =>
      val (lookupExpr, table, colIndexExpr, rangeLookupOpt) = args
      val rangeLookupExpr = rangeLookupOpt.getOrElse(TExpr.Lit(true))
      ctx.evalExpr(TExpr.vlookupWithLocation(lookupExpr, table, colIndexExpr, rangeLookupExpr))
    }

  val sumproduct: FunctionSpec[BigDecimal] { type Args = SumProductArgs } =
    FunctionSpec.simple[BigDecimal, SumProductArgs]("SUMPRODUCT", Arity.atLeastOne) {
      (arrays, ctx) =>
        ctx.evalExpr(TExpr.sumProduct(arrays))
    }

  val xlookup: FunctionSpec[CellValue] { type Args = XLookupArgs } =
    FunctionSpec.simple[CellValue, XLookupArgs]("XLOOKUP", Arity.Range(3, 6)) { (args, ctx) =>
      val (lookupValue, lookupArray, returnArray, ifNotFoundOpt, matchModeOpt, searchModeOpt) =
        args
      val matchModeExpr = matchModeOpt.getOrElse(TExpr.Lit(0))
      val searchModeExpr = searchModeOpt.getOrElse(TExpr.Lit(1))
      ctx.evalExpr(
        TExpr.xlookup(
          lookupValue,
          lookupArray,
          returnArray,
          ifNotFoundOpt,
          matchModeExpr,
          searchModeExpr
        )
      )
    }

  val row: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr]("ROW", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(TExpr.row(expr))
    }

  val column: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr]("COLUMN", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(TExpr.column(expr))
    }

  val rows: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr](
      "ROWS",
      Arity.one,
      renderFn = Some { (expr, printer) =>
        val rendered = expr match
          case TExpr.FoldRange(range, _, _, _) => printer.cellRange(range)
          case TExpr.SheetFoldRange(sheet, range, _, _, _) =>
            printer.location(TExpr.RangeLocation.CrossSheet(sheet, range))
          case other => printer.expr(other)
        s"ROWS($rendered)"
      }
    ) { (expr, ctx) =>
      ctx.evalExpr(TExpr.rows(expr))
    }

  val columns: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr](
      "COLUMNS",
      Arity.one,
      renderFn = Some { (expr, printer) =>
        val rendered = expr match
          case TExpr.FoldRange(range, _, _, _) => printer.cellRange(range)
          case TExpr.SheetFoldRange(sheet, range, _, _, _) =>
            printer.location(TExpr.RangeLocation.CrossSheet(sheet, range))
          case other => printer.expr(other)
        s"COLUMNS($rendered)"
      }
    ) { (expr, ctx) =>
      ctx.evalExpr(TExpr.columns(expr))
    }

  val address: FunctionSpec[String] { type Args = AddressArgs } =
    FunctionSpec.simple[String, AddressArgs]("ADDRESS", Arity.Range(2, 5)) { (args, ctx) =>
      val (rowExpr, colExpr, absNumOpt, a1Opt, sheetOpt) = args
      val absNumExpr = absNumOpt.getOrElse(TExpr.Lit(BigDecimal(1)))
      val a1Expr = a1Opt.getOrElse(TExpr.Lit(true))
      ctx.evalExpr(TExpr.address(rowExpr, colExpr, absNumExpr, a1Expr, sheetOpt))
    }

  val index: FunctionSpec[CellValue] { type Args = IndexArgs } =
    FunctionSpec.simple[CellValue, IndexArgs]("INDEX", Arity.Range(2, 3)) { (args, ctx) =>
      val (array, rowNumExpr, colNumOpt) = args
      ctx.evalExpr(TExpr.index(array, rowNumExpr, colNumOpt))
    }

  val matchFn: FunctionSpec[BigDecimal] { type Args = MatchArgs } =
    FunctionSpec.simple[BigDecimal, MatchArgs]("MATCH", Arity.Range(2, 3)) { (args, ctx) =>
      val (lookupValue, lookupArray, matchTypeOpt) = args
      val matchTypeExpr = matchTypeOpt.getOrElse(TExpr.Lit(BigDecimal(1)))
      ctx.evalExpr(TExpr.matchExpr(lookupValue, lookupArray, matchTypeExpr))
    }

  val concatenate: FunctionSpec[String] { type Args = TextList } =
    FunctionSpec.simple[String, TextList]("CONCATENATE", Arity.atLeastOne) { (args, ctx) =>
      args.foldLeft[Either[EvalError, String]](Right("")) { (accEither, expr) =>
        for
          acc <- accEither
          value <- ctx.evalExpr(expr)
        yield acc + value
      }
    }

  val left: FunctionSpec[String] { type Args = BinaryTextInt } =
    FunctionSpec.simple[String, BinaryTextInt]("LEFT", Arity.two) { (args, ctx) =>
      val (textExpr, nExpr) = args
      for
        text <- ctx.evalExpr(textExpr)
        nValue <- ctx.evalExpr(nExpr)
        result <-
          if nValue < 0 then
            Left(EvalError.EvalFailed(s"LEFT: n must be non-negative, got $nValue"))
          else if nValue >= text.length then Right(text)
          else Right(text.take(nValue))
      yield result
    }

  val right: FunctionSpec[String] { type Args = BinaryTextInt } =
    FunctionSpec.simple[String, BinaryTextInt]("RIGHT", Arity.two) { (args, ctx) =>
      val (textExpr, nExpr) = args
      for
        text <- ctx.evalExpr(textExpr)
        nValue <- ctx.evalExpr(nExpr)
        result <-
          if nValue < 0 then
            Left(EvalError.EvalFailed(s"RIGHT: n must be non-negative, got $nValue"))
          else if nValue >= text.length then Right(text)
          else Right(text.takeRight(nValue))
      yield result
    }

  val len: FunctionSpec[BigDecimal] { type Args = UnaryText } =
    FunctionSpec.simple[BigDecimal, UnaryText]("LEN", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(text => BigDecimal(text.length))
    }

  val upper: FunctionSpec[String] { type Args = UnaryText } =
    FunctionSpec.simple[String, UnaryText]("UPPER", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(_.toUpperCase)
    }

  val lower: FunctionSpec[String] { type Args = UnaryText } =
    FunctionSpec.simple[String, UnaryText]("LOWER", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(_.toLowerCase)
    }
