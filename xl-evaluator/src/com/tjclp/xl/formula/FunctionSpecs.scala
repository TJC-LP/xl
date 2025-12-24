package com.tjclp.xl.formula

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
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

  private def coerceToNumeric(value: CellValue): BigDecimal =
    value match
      case CellValue.Number(n) => n
      case CellValue.Bool(true) => BigDecimal(1)
      case CellValue.Bool(false) => BigDecimal(0)
      case CellValue.Formula(_, Some(cached)) => coerceToNumeric(cached)
      case _ => BigDecimal(0)

  private def compareCellValues(cv: CellValue, value: Any): Int =
    (cv, value) match
      case (CellValue.Number(n1), n2: BigDecimal) => n1.compare(n2)
      case (CellValue.Number(n1), n2: Int) => n1.compare(BigDecimal(n2))
      case (CellValue.Number(n1), n2: Long) => n1.compare(BigDecimal(n2))
      case (CellValue.Number(n1), n2: Double) => n1.compare(BigDecimal(n2))
      case (CellValue.Text(s1), s2: String) => s1.compareToIgnoreCase(s2)
      case (CellValue.Bool(b1), b2: Boolean) => b1.compare(b2)
      case (CellValue.Number(n), CellValue.Number(n2)) => n.compare(n2)
      case (CellValue.Text(s), CellValue.Text(s2)) => s.compareToIgnoreCase(s2)
      case (CellValue.Bool(b), CellValue.Bool(b2)) => b.compare(b2)
      case _ => -2

  private def coerceToBigDecimal(value: Any): BigDecimal =
    value match
      case n: BigDecimal => n
      case n: Int => BigDecimal(n)
      case n: Long => BigDecimal(n)
      case n: Double => BigDecimal(n)
      case CellValue.Number(n) => n
      case CellValue.Bool(true) => BigDecimal(1)
      case CellValue.Bool(false) => BigDecimal(0)
      case _ => BigDecimal(0)

  private def performXLookup(
    lookupValue: Any,
    lookupArray: CellRange,
    returnArray: CellRange,
    ifNotFoundOpt: Option[TExpr[?]],
    matchMode: Int,
    searchMode: Int,
    ctx: EvalContext
  ): Either[EvalError, CellValue] =
    val lookupCells = lookupArray.cells.toList
    val returnCells = returnArray.cells.toList

    val indices: List[Int] = searchMode match
      case -1 => lookupCells.indices.reverse.toList
      case _ => lookupCells.indices.toList

    val matchingIndexOpt: Option[Int] = matchMode match
      case 0 =>
        indices.find { idx =>
          val cellValue = ctx.sheet(lookupCells(idx)).value
          matchesExactForXLookup(cellValue, lookupValue)
        }
      case 2 =>
        lookupValue match
          case pattern: String =>
            val criterion = CriteriaMatcher.parse(pattern)
            indices.find { idx =>
              CriteriaMatcher.matches(ctx.sheet(lookupCells(idx)).value, criterion)
            }
          case _ => None
      case -1 =>
        findNextSmaller(lookupValue, lookupCells, ctx.sheet, indices)
      case 1 =>
        findNextLarger(lookupValue, lookupCells, ctx.sheet, indices)
      case _ => None

    matchingIndexOpt match
      case Some(idx) =>
        Right(ctx.sheet(returnCells(idx)).value)
      case None =>
        ifNotFoundOpt match
          case Some(expr) => evalAny(ctx, expr).map(toCellValue)
          case None => Right(CellValue.Error(CellError.NA))

  private def matchesExactForXLookup(cellValue: CellValue, lookupValue: Any): Boolean =
    CriteriaMatcher.matches(cellValue, CriteriaMatcher.Exact(lookupValue))

  private def findNextSmaller(
    lookupValue: Any,
    lookupCells: List[ARef],
    sheet: Sheet,
    indices: List[Int]
  ): Option[Int] =
    lookupValue match
      case targetNum: BigDecimal =>
        val candidates = indices
          .flatMap { idx =>
            extractNumericValue(sheet(lookupCells(idx)).value).map(n => (idx, n))
          }
          .filter(_._2 <= targetNum)
        candidates.sortBy(-_._2).headOption.map(_._1)
      case _ => None

  private def findNextLarger(
    lookupValue: Any,
    lookupCells: List[ARef],
    sheet: Sheet,
    indices: List[Int]
  ): Option[Int] =
    lookupValue match
      case targetNum: BigDecimal =>
        val candidates = indices
          .flatMap { idx =>
            extractNumericValue(sheet(lookupCells(idx)).value).map(n => (idx, n))
          }
          .filter(_._2 >= targetNum)
        candidates.sortBy(_._2).headOption.map(_._1)
      case _ => None

  private def extractNumericValue(value: CellValue): Option[BigDecimal] =
    value match
      case CellValue.Number(n) => Some(n)
      case CellValue.Formula(_, Some(CellValue.Number(n))) => Some(n)
      case _ => None

  private def extractARef(expr: TExpr[?]): Option[ARef] = expr match
    case TExpr.PolyRef(ref, _) => Some(ref)
    case TExpr.Ref(ref, _, _) => Some(ref)
    case TExpr.SheetPolyRef(_, ref, _) => Some(ref)
    case TExpr.SheetRef(_, ref, _, _) => Some(ref)
    case TExpr.RangeRef(range) => Some(range.start)
    case TExpr.SheetRange(_, range) => Some(range.start)
    case _ => None

  private def extractCellRange(expr: TExpr[?]): Option[CellRange] = expr match
    case TExpr.RangeRef(range) => Some(range)
    case TExpr.SheetRange(_, range) => Some(range)
    case _ => None

  @annotation.tailrec
  private def columnToLetter(col: Int, acc: String = ""): String =
    if col < 0 then acc
    else if acc.isEmpty && col <= 25 then ('A' + col).toChar.toString
    else
      val remainder = col % 26
      val quotient = col / 26 - 1
      val letter = ('A' + remainder).toChar
      if quotient < 0 then letter.toString + acc
      else columnToLetter(quotient, letter.toString + acc)

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
      evalAny(ctx, criteria).flatMap { criteriaValue =>
        val criterion = CriteriaMatcher.parse(criteriaValue)
        val effectiveRange = sumRangeOpt.getOrElse(range)
        val rangeRefsList = range.cells.toList
        val sumRefsList = effectiveRange.cells.toList

        if range.width != effectiveRange.width || range.height != effectiveRange.height then
          Left(
            EvalError.EvalFailed(
              s"SUMIF: range and sum_range must have same dimensions (${range.height}×${range.width} vs ${effectiveRange.height}×${effectiveRange.width})",
              Some(s"SUMIF(${range.toA1}, ..., ${effectiveRange.toA1})")
            )
          )
        else
          val pairs = rangeRefsList.zip(sumRefsList)
          val sum = pairs.foldLeft(BigDecimal(0)) { case (acc, (testRef, sumRef)) =>
            val testCell = ctx.sheet(testRef)
            if CriteriaMatcher.matches(testCell.value, criterion) then
              ctx.sheet(sumRef).value match
                case CellValue.Number(n) => acc + n
                case _ => acc
            else acc
          }
          Right(sum)
      }
    }

  val countif: FunctionSpec[BigDecimal] { type Args = CountIfArgs } =
    FunctionSpec.simple[BigDecimal, CountIfArgs]("COUNTIF", Arity.two) { (args, ctx) =>
      val (range, criteria) = args
      evalAny(ctx, criteria).map { criteriaValue =>
        val criterion = CriteriaMatcher.parse(criteriaValue)
        val count = range.cells.count { ref =>
          CriteriaMatcher.matches(ctx.sheet(ref).value, criterion)
        }
        BigDecimal(count)
      }
    }

  val sumifs: FunctionSpec[BigDecimal] { type Args = SumIfsArgs } =
    FunctionSpec.simple[BigDecimal, SumIfsArgs]("SUMIFS", Arity.AtLeast(3)) { (args, ctx) =>
      val (sumRange, conditions) = args
      val criteriaEithers = conditions.map { case (_, criteriaExpr) =>
        evalAny(ctx, criteriaExpr)
      }

      criteriaEithers
        .foldLeft[Either[EvalError, List[Any]]](Right(List.empty)) { (acc, either) =>
          acc.flatMap(list => either.map(v => v :: list))
        }
        .map(_.reverse)
        .flatMap { criteriaValues =>
          val parsedConditions = conditions
            .zip(criteriaValues)
            .map { case ((range, _), criteriaValue) =>
              (range, CriteriaMatcher.parse(criteriaValue))
            }

          val sumRefsList = sumRange.cells.toList
          val dimensionError = parsedConditions.collectFirst {
            case (range, _) if range.width != sumRange.width || range.height != sumRange.height =>
              EvalError.EvalFailed(
                s"SUMIFS: all ranges must have same dimensions (sum_range is ${sumRange.height}×${sumRange.width}, criteria_range is ${range.height}×${range.width})",
                Some(s"SUMIFS(${sumRange.toA1}, ${range.toA1}, ...)")
              )
          }

          dimensionError match
            case Some(err) => Left(err)
            case None =>
              val sum = sumRefsList.indices.foldLeft(BigDecimal(0)) { (acc, idx) =>
                val allMatch = parsedConditions.forall { case (criteriaRange, criterion) =>
                  val testRef = criteriaRange.cells.toList(idx)
                  CriteriaMatcher.matches(ctx.sheet(testRef).value, criterion)
                }
                if allMatch then
                  ctx.sheet(sumRefsList(idx)).value match
                    case CellValue.Number(n) => acc + n
                    case _ => acc
                else acc
              }
              Right(sum)
        }
    }

  val countifs: FunctionSpec[BigDecimal] { type Args = CountIfsArgs } =
    FunctionSpec.simple[BigDecimal, CountIfsArgs]("COUNTIFS", Arity.AtLeast(2)) {
      (conditions, ctx) =>
        val criteriaEithers = conditions.map { case (_, criteriaExpr) =>
          evalAny(ctx, criteriaExpr)
        }

        criteriaEithers
          .foldLeft[Either[EvalError, List[Any]]](Right(List.empty)) { (acc, either) =>
            acc.flatMap(list => either.map(v => v :: list))
          }
          .map(_.reverse)
          .flatMap { criteriaValues =>
            val parsedConditions = conditions
              .zip(criteriaValues)
              .map { case ((range, _), criteriaValue) =>
                (range, CriteriaMatcher.parse(criteriaValue))
              }

            parsedConditions match
              case (firstRange, _) :: _ =>
                val refCount = firstRange.cells.toList.length

                val dimensionError = parsedConditions.collectFirst {
                  case (range, _)
                      if range.width != firstRange.width || range.height != firstRange.height =>
                    EvalError.EvalFailed(
                      s"COUNTIFS: all ranges must have same dimensions (first is ${firstRange.height}×${firstRange.width}, this is ${range.height}×${range.width})",
                      Some(s"COUNTIFS(${firstRange.toA1}, ..., ${range.toA1}, ...)")
                    )
                }

                dimensionError match
                  case Some(err) => Left(err)
                  case None =>
                    val count = (0 until refCount).count { idx =>
                      parsedConditions.forall { case (criteriaRange, criterion) =>
                        val testRef = criteriaRange.cells.toList(idx)
                        CriteriaMatcher.matches(ctx.sheet(testRef).value, criterion)
                      }
                    }
                    Right(BigDecimal(count))
              case Nil =>
                Right(BigDecimal(0))
          }
    }

  val averageif: FunctionSpec[BigDecimal] { type Args = AverageIfArgs } =
    FunctionSpec.simple[BigDecimal, AverageIfArgs]("AVERAGEIF", Arity.Range(2, 3)) { (args, ctx) =>
      val (range, criteria, avgRangeOpt) = args
      evalAny(ctx, criteria).flatMap { criteriaValue =>
        val criterion = CriteriaMatcher.parse(criteriaValue)
        val effectiveRange = avgRangeOpt.getOrElse(range)
        val rangeRefsList = range.cells.toList
        val avgRefsList = effectiveRange.cells.toList

        if range.width != effectiveRange.width || range.height != effectiveRange.height then
          Left(
            EvalError.EvalFailed(
              s"AVERAGEIF: range and average_range must have same dimensions (${range.height}×${range.width} vs ${effectiveRange.height}×${effectiveRange.width})",
              Some(s"AVERAGEIF(${range.toA1}, ..., ${effectiveRange.toA1})")
            )
          )
        else
          val pairs = rangeRefsList.zip(avgRefsList)
          val (sum, count) = pairs.foldLeft((BigDecimal(0), 0)) {
            case ((accSum, accCount), (testRef, avgRef)) =>
              val testCell = ctx.sheet(testRef)
              if CriteriaMatcher.matches(testCell.value, criterion) then
                ctx.sheet(avgRef).value match
                  case CellValue.Number(n) => (accSum + n, accCount + 1)
                  case _ => (accSum, accCount)
              else (accSum, accCount)
          }
          if count == 0 then
            Left(EvalError.DivByZero("AVERAGEIF sum", "0 (no matching numeric cells)"))
          else Right(sum / count)
      }
    }

  val averageifs: FunctionSpec[BigDecimal] { type Args = AverageIfsArgs } =
    FunctionSpec.simple[BigDecimal, AverageIfsArgs]("AVERAGEIFS", Arity.AtLeast(3)) { (args, ctx) =>
      val (avgRange, conditions) = args
      val criteriaEithers = conditions.map { case (_, criteriaExpr) =>
        evalAny(ctx, criteriaExpr)
      }

      criteriaEithers
        .foldLeft[Either[EvalError, List[Any]]](Right(List.empty)) { (acc, either) =>
          acc.flatMap(list => either.map(v => v :: list))
        }
        .map(_.reverse)
        .flatMap { criteriaValues =>
          val parsedConditions = conditions
            .zip(criteriaValues)
            .map { case ((range, _), criteriaValue) =>
              (range, CriteriaMatcher.parse(criteriaValue))
            }

          val avgRefsList = avgRange.cells.toList
          val dimensionError = parsedConditions.collectFirst {
            case (range, _) if range.width != avgRange.width || range.height != avgRange.height =>
              EvalError.EvalFailed(
                s"AVERAGEIFS: all ranges must have same dimensions (average_range is ${avgRange.height}×${avgRange.width}, criteria_range is ${range.height}×${range.width})",
                Some(s"AVERAGEIFS(${avgRange.toA1}, ${range.toA1}, ...)")
              )
          }

          dimensionError match
            case Some(err) => Left(err)
            case None =>
              val (sum, count) = avgRefsList.indices.foldLeft((BigDecimal(0), 0)) { (acc, idx) =>
                val (accSum, accCount) = acc
                val allMatch = parsedConditions.forall { case (criteriaRange, criterion) =>
                  val testRef = criteriaRange.cells.toList(idx)
                  CriteriaMatcher.matches(ctx.sheet(testRef).value, criterion)
                }
                if allMatch then
                  ctx.sheet(avgRefsList(idx)).value match
                    case CellValue.Number(n) => (accSum + n, accCount + 1)
                    case _ => (accSum, accCount)
                else (accSum, accCount)
              }
              if count == 0 then
                Left(EvalError.DivByZero("AVERAGEIFS sum", "0 (no matching numeric cells)"))
              else Right(sum / count)
        }
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
      for
        y <- ctx.evalExpr(yearExpr)
        m <- ctx.evalExpr(monthExpr)
        d <- ctx.evalExpr(dayExpr)
        result <- scala.util.Try(LocalDate.of(y, m, d)).toEither.left.map { ex =>
          EvalError.EvalFailed(
            s"DATE: invalid date components (year=$y, month=$m, day=$d): ${ex.getMessage}"
          )
        }
      yield result
    }

  val year: FunctionSpec[BigDecimal] { type Args = UnaryDate } =
    FunctionSpec.simple[BigDecimal, UnaryDate]("YEAR", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(date => BigDecimal(date.getYear))
    }

  val month: FunctionSpec[BigDecimal] { type Args = UnaryDate } =
    FunctionSpec.simple[BigDecimal, UnaryDate]("MONTH", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(date => BigDecimal(date.getMonthValue))
    }

  val day: FunctionSpec[BigDecimal] { type Args = UnaryDate } =
    FunctionSpec.simple[BigDecimal, UnaryDate]("DAY", Arity.one) { (expr, ctx) =>
      ctx.evalExpr(expr).map(date => BigDecimal(date.getDayOfMonth))
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
      ctx.evalExpr(rateExpr).flatMap { rate =>
        val onePlusR = BigDecimal(1) + rate
        if onePlusR == BigDecimal(0) then
          Left(
            EvalError.EvalFailed(
              "NPV: rate = -1 would require division by zero",
              Some("NPV(rate, values)")
            )
          )
        else
          val cashFlows: List[BigDecimal] =
            range.cells
              .map(ref => ctx.sheet(ref))
              .flatMap(cell => TExpr.decodeNumeric(cell).toOption)
              .toList

          val npv =
            cashFlows.zipWithIndex.foldLeft(BigDecimal(0)) { case (acc, (cf, idx)) =>
              val period = idx + 1
              acc + cf / onePlusR.pow(period)
            }
          Right(npv)
      }
    }

  val irr: FunctionSpec[BigDecimal] { type Args = IrrArgs } =
    FunctionSpec.simple[BigDecimal, IrrArgs]("IRR", Arity.Range(1, 2)) { (args, ctx) =>
      val (range, guessOpt) = args
      val cashFlows: List[BigDecimal] =
        range.cells
          .map(ref => ctx.sheet(ref))
          .flatMap(cell => TExpr.decodeNumeric(cell).toOption)
          .toList

      if cashFlows.isEmpty || !cashFlows.exists(_ < 0) || !cashFlows.exists(_ > 0) then
        Left(
          EvalError.EvalFailed(
            "IRR requires at least one positive and one negative cash flow",
            Some("IRR(values[, guess])")
          )
        )
      else
        val guessEither: Either[EvalError, BigDecimal] =
          guessOpt match
            case Some(guessExpr) => ctx.evalExpr(guessExpr)
            case None => Right(BigDecimal("0.1"))

        guessEither.flatMap { guess0 =>
          val maxIter = 50
          val tolerance = BigDecimal("1e-7")
          val one = BigDecimal(1)

          def npvAt(rate: BigDecimal): BigDecimal =
            val onePlusR = one + rate
            cashFlows.zipWithIndex.foldLeft(BigDecimal(0)) { case (acc, (cf, idx)) =>
              if idx == 0 then acc + cf
              else acc + cf / onePlusR.pow(idx)
            }

          def dNpvAt(rate: BigDecimal): BigDecimal =
            val onePlusR = one + rate
            cashFlows.zipWithIndex.foldLeft(BigDecimal(0)) { case (acc, (cf, idx)) =>
              if idx == 0 then acc
              else acc - (idx * cf) / onePlusR.pow(idx + 1)
            }

          @annotation.tailrec
          def loop(iter: Int, r: BigDecimal): Either[EvalError, BigDecimal] =
            if iter >= maxIter then
              Left(
                EvalError.EvalFailed(
                  s"IRR did not converge after $maxIter iterations",
                  Some("IRR(values[, guess])")
                )
              )
            else
              val f = npvAt(r)
              val df = dNpvAt(r)
              if df == BigDecimal(0) then
                Left(
                  EvalError.EvalFailed(
                    "IRR derivative is zero; cannot continue iteration",
                    Some("IRR(values[, guess])")
                  )
                )
              else
                val next = r - f / df
                if (next - r).abs <= tolerance then Right(next)
                else loop(iter + 1, next)

          loop(0, guess0)
        }
    }

  val xnpv: FunctionSpec[BigDecimal] { type Args = XnpvArgs } =
    FunctionSpec.simple[BigDecimal, XnpvArgs]("XNPV", Arity.three) { (args, ctx) =>
      val (rateExpr, valuesRange, datesRange) = args
      for
        rate <- ctx.evalExpr(rateExpr)
        result <- {
          val values: List[BigDecimal] =
            valuesRange.cells
              .map(ref => ctx.sheet(ref))
              .flatMap(cell => TExpr.decodeNumeric(cell).toOption)
              .toList

          val dates: List[LocalDate] =
            datesRange.cells
              .map(ref => ctx.sheet(ref))
              .flatMap(cell => TExpr.decodeDate(cell).toOption)
              .toList

          if values.isEmpty || dates.isEmpty then
            Left(
              EvalError.EvalFailed(
                "XNPV requires non-empty values and dates ranges",
                Some("XNPV(rate, values, dates)")
              )
            )
          else if values.length != dates.length then
            Left(
              EvalError.EvalFailed(
                s"XNPV: values (${values.length}) and dates (${dates.length}) must have same length",
                Some("XNPV(rate, values, dates)")
              )
            )
          else
            dates match
              case date0 :: _ =>
                val onePlusR = BigDecimal(1) + rate
                val npv = values.zip(dates).foldLeft(BigDecimal(0)) { case (acc, (value, date)) =>
                  val daysDiff = ChronoUnit.DAYS.between(date0, date)
                  val yearFraction = BigDecimal(daysDiff) / BigDecimal(365)
                  val discountFactor = math.pow(onePlusR.toDouble, yearFraction.toDouble)
                  acc + value / BigDecimal(discountFactor)
                }
                Right(npv)
              case Nil =>
                Left(EvalError.EvalFailed("XNPV: dates cannot be empty", None))
        }
      yield result
    }

  val xirr: FunctionSpec[BigDecimal] { type Args = XirrArgs } =
    FunctionSpec.simple[BigDecimal, XirrArgs]("XIRR", Arity.Range(2, 3)) { (args, ctx) =>
      val (valuesRange, datesRange, guessOpt) = args
      val values: List[BigDecimal] =
        valuesRange.cells
          .map(ref => ctx.sheet(ref))
          .flatMap(cell => TExpr.decodeNumeric(cell).toOption)
          .toList

      val dates: List[LocalDate] =
        datesRange.cells
          .map(ref => ctx.sheet(ref))
          .flatMap(cell => TExpr.decodeDate(cell).toOption)
          .toList

      if values.isEmpty || dates.isEmpty then
        Left(
          EvalError.EvalFailed(
            "XIRR requires non-empty values and dates ranges",
            Some("XIRR(values, dates[, guess])")
          )
        )
      else if values.length != dates.length then
        Left(
          EvalError.EvalFailed(
            s"XIRR: values (${values.length}) and dates (${dates.length}) must have same length",
            Some("XIRR(values, dates[, guess])")
          )
        )
      else if !values.exists(_ < 0) || !values.exists(_ > 0) then
        Left(
          EvalError.EvalFailed(
            "XIRR requires at least one positive and one negative cash flow",
            Some("XIRR(values, dates[, guess])")
          )
        )
      else
        val guessEither: Either[EvalError, BigDecimal] =
          guessOpt match
            case Some(guessExpr) => ctx.evalExpr(guessExpr)
            case None => Right(BigDecimal("0.1"))

        guessEither.flatMap { guess0 =>
          val maxIter = 100
          val tolerance = BigDecimal("1e-7")
          dates match
            case date0 :: _ =>
              val yearFractions: List[BigDecimal] = dates.map { date =>
                val daysDiff = ChronoUnit.DAYS.between(date0, date)
                BigDecimal(daysDiff) / BigDecimal(365)
              }

              def xnpvAt(rate: BigDecimal): BigDecimal =
                val onePlusR = BigDecimal(1) + rate
                values.zip(yearFractions).foldLeft(BigDecimal(0)) { case (acc, (cf, yf)) =>
                  val discountFactor = math.pow(onePlusR.toDouble, yf.toDouble)
                  acc + cf / BigDecimal(discountFactor)
                }

              def dXnpvAt(rate: BigDecimal): BigDecimal =
                val onePlusR = BigDecimal(1) + rate
                values.zip(yearFractions).foldLeft(BigDecimal(0)) { case (acc, (cf, yf)) =>
                  val discountFactor = math.pow(onePlusR.toDouble, (yf + 1).toDouble)
                  acc - (yf * cf) / BigDecimal(discountFactor)
                }

              @annotation.tailrec
              def loop(iter: Int, r: BigDecimal): Either[EvalError, BigDecimal] =
                if iter >= maxIter then
                  Left(
                    EvalError.EvalFailed(
                      s"XIRR did not converge after $maxIter iterations",
                      Some("XIRR(values, dates[, guess])")
                    )
                  )
                else
                  val f = xnpvAt(r)
                  val df = dXnpvAt(r)
                  if df.abs < BigDecimal("1e-10") then
                    Left(
                      EvalError.EvalFailed(
                        "XIRR derivative is near zero; cannot continue iteration",
                        Some("XIRR(values, dates[, guess])")
                      )
                    )
                  else
                    val next = r - f / df
                    if (next - r).abs <= tolerance then Right(next)
                    else loop(iter + 1, next)

              loop(0, guess0)
            case Nil =>
              Left(EvalError.EvalFailed("XIRR: dates cannot be empty", None))
        }
    }

  val pmt: FunctionSpec[BigDecimal] { type Args = TvmArgs } =
    FunctionSpec.simple[BigDecimal, TvmArgs]("PMT", Arity.Range(3, 5)) { (args, ctx) =>
      val (rateExpr, nperExpr, pvExpr, fvOpt, typeOpt) = args
      for
        rate <- ctx.evalExpr(rateExpr).map(_.toDouble)
        nper <- ctx.evalExpr(nperExpr).map(_.toDouble)
        pv <- ctx.evalExpr(pvExpr).map(_.toDouble)
        fv <- fvOpt match
          case Some(expr) => ctx.evalExpr(expr).map(_.toDouble)
          case None => Right(0.0)
        pmtType <- typeOpt match
          case Some(expr) => ctx.evalExpr(expr).map(v => if v.toInt != 0 then 1 else 0)
          case None => Right(0)
      yield
        if math.abs(rate) < 1e-10 then
          if nper == 0.0 then BigDecimal(Double.NaN)
          else BigDecimal(-(pv + fv) / nper)
        else
          val pvif = math.pow(1.0 + rate, nper)
          BigDecimal(-rate * (pv * pvif + fv) / ((1.0 + rate * pmtType) * (pvif - 1.0)))
    }

  val fv: FunctionSpec[BigDecimal] { type Args = TvmArgs } =
    FunctionSpec.simple[BigDecimal, TvmArgs]("FV", Arity.Range(3, 5)) { (args, ctx) =>
      val (rateExpr, nperExpr, pmtExpr, pvOpt, typeOpt) = args
      for
        rate <- ctx.evalExpr(rateExpr).map(_.toDouble)
        nper <- ctx.evalExpr(nperExpr).map(_.toDouble)
        pmt <- ctx.evalExpr(pmtExpr).map(_.toDouble)
        pv <- pvOpt match
          case Some(expr) => ctx.evalExpr(expr).map(_.toDouble)
          case None => Right(0.0)
        pmtType <- typeOpt match
          case Some(expr) => ctx.evalExpr(expr).map(v => if v.toInt != 0 then 1 else 0)
          case None => Right(0)
      yield
        if math.abs(rate) < 1e-10 then BigDecimal(-pv - pmt * nper)
        else
          val pvif = math.pow(1.0 + rate, nper)
          val fvifa = (pvif - 1.0) / rate
          BigDecimal(-pv * pvif - pmt * (1.0 + rate * pmtType) * fvifa)
    }

  val pv: FunctionSpec[BigDecimal] { type Args = TvmArgs } =
    FunctionSpec.simple[BigDecimal, TvmArgs]("PV", Arity.Range(3, 5)) { (args, ctx) =>
      val (rateExpr, nperExpr, pmtExpr, fvOpt, typeOpt) = args
      for
        rate <- ctx.evalExpr(rateExpr).map(_.toDouble)
        nper <- ctx.evalExpr(nperExpr).map(_.toDouble)
        pmt <- ctx.evalExpr(pmtExpr).map(_.toDouble)
        fv <- fvOpt match
          case Some(expr) => ctx.evalExpr(expr).map(_.toDouble)
          case None => Right(0.0)
        pmtType <- typeOpt match
          case Some(expr) => ctx.evalExpr(expr).map(v => if v.toInt != 0 then 1 else 0)
          case None => Right(0)
      yield
        if math.abs(rate) < 1e-10 then BigDecimal(-fv - pmt * nper)
        else
          val pvif = math.pow(1.0 + rate, nper)
          val fvifa = (pvif - 1.0) / rate
          BigDecimal((-fv - pmt * (1.0 + rate * pmtType) * fvifa) / pvif)
    }

  val nper: FunctionSpec[BigDecimal] { type Args = TvmArgs } =
    FunctionSpec.simple[BigDecimal, TvmArgs]("NPER", Arity.Range(3, 5)) { (args, ctx) =>
      val (rateExpr, pmtExpr, pvExpr, fvOpt, typeOpt) = args
      for
        rate <- ctx.evalExpr(rateExpr).map(_.toDouble)
        pmt <- ctx.evalExpr(pmtExpr).map(_.toDouble)
        pv <- ctx.evalExpr(pvExpr).map(_.toDouble)
        fv <- fvOpt match
          case Some(expr) => ctx.evalExpr(expr).map(_.toDouble)
          case None => Right(0.0)
        pmtType <- typeOpt match
          case Some(expr) => ctx.evalExpr(expr).map(v => if v.toInt != 0 then 1 else 0)
          case None => Right(0)
      yield
        if math.abs(rate) < 1e-10 then
          if pmt == 0.0 then BigDecimal(Double.NaN)
          else BigDecimal(-(pv + fv) / pmt)
        else
          val ratep1 = 1.0 + rate
          val numerator = -fv * rate + pmt * (1.0 + rate * pmtType)
          val denominator = pv * rate + pmt * (1.0 + rate * pmtType)
          BigDecimal(math.log(numerator / denominator) / math.log(ratep1))
    }

  val rate: FunctionSpec[BigDecimal] { type Args = RateArgs } =
    FunctionSpec.simple[BigDecimal, RateArgs]("RATE", Arity.Range(3, 6)) { (args, ctx) =>
      val (nperExpr, pmtExpr, pvExpr, fvOpt, typeOpt, guessOpt) = args
      for
        nper <- ctx.evalExpr(nperExpr).map(_.toDouble)
        pmt <- ctx.evalExpr(pmtExpr).map(_.toDouble)
        pv <- ctx.evalExpr(pvExpr).map(_.toDouble)
        fv <- fvOpt match
          case Some(expr) => ctx.evalExpr(expr).map(_.toDouble)
          case None => Right(0.0)
        pmtType <- typeOpt match
          case Some(expr) => ctx.evalExpr(expr).map(v => if v.toInt != 0 then 1 else 0)
          case None => Right(0)
        guess <- guessOpt match
          case Some(expr) => ctx.evalExpr(expr).map(_.toDouble)
          case None => Right(0.1)
        result <- {
          val maxIter = 100
          val tolerance = 1e-7

          def f(rate: Double): Double =
            if math.abs(rate) < 1e-10 then pv + pmt * nper + fv
            else
              val pvif = math.pow(1.0 + rate, nper)
              val fvifa = (pvif - 1.0) / rate
              pv * pvif + pmt * (1.0 + rate * pmtType) * fvifa + fv

          def df(rate: Double): Double =
            if math.abs(rate) < 1e-10 then pv * nper + pmt * nper * (nper - 1.0) / 2.0
            else
              val pvif = math.pow(1.0 + rate, nper)
              val dpvif = nper * math.pow(1.0 + rate, nper - 1.0)
              val fvifa = (pvif - 1.0) / rate
              val dfvifa = (dpvif * rate - (pvif - 1.0)) / (rate * rate)
              pv * dpvif + pmt * pmtType * fvifa + pmt * (1.0 + rate * pmtType) * dfvifa

          @annotation.tailrec
          def loop(iter: Int, r: Double): Either[EvalError, BigDecimal] =
            if iter >= maxIter then
              Left(
                EvalError.EvalFailed(
                  s"RATE did not converge after $maxIter iterations",
                  Some("RATE(nper, pmt, pv, [fv], [type], [guess])")
                )
              )
            else
              val fVal = f(r)
              val dfVal = df(r)
              if math.abs(dfVal) < 1e-14 then
                Left(
                  EvalError.EvalFailed(
                    "RATE derivative is zero; cannot continue iteration",
                    Some("RATE(nper, pmt, pv, [fv], [type], [guess])")
                  )
                )
              else
                val next = r - fVal / dfVal
                if math.abs(next - r) <= tolerance then Right(BigDecimal(next))
                else loop(iter + 1, next)

          loop(0, guess)
        }
      yield result
    }

  val vlookup: FunctionSpec[CellValue] { type Args = VlookupArgs } =
    FunctionSpec.simple[CellValue, VlookupArgs]("VLOOKUP", Arity.Range(3, 4)) { (args, ctx) =>
      val (lookupExpr, table, colIndexExpr, rangeLookupOpt) = args
      val rangeLookupExpr = rangeLookupOpt.getOrElse(TExpr.Lit(true))
      for
        lookupValue <- evalAny(ctx, lookupExpr)
        colIndex <- ctx.evalExpr(colIndexExpr)
        rangeMatch <- ctx.evalExpr(rangeLookupExpr)
        targetSheet <- Evaluator.resolveRangeLocation(table, ctx.sheet, ctx.workbook)
        result <-
          if colIndex < 1 || colIndex > table.range.width then
            Left(
              EvalError.EvalFailed(
                s"VLOOKUP: col_index_num $colIndex is outside 1..${table.range.width}",
                Some(s"VLOOKUP(…, ${table.range.toA1})")
              )
            )
          else
            val rowIndices = 0 until table.range.height
            val keyCol0 = table.range.colStart.index0
            val rowStart0 = table.range.rowStart.index0
            val resultCol0 = keyCol0 + (colIndex - 1)

            def extractTextForMatch(cv: CellValue): Option[String] = cv match
              case CellValue.Text(s) => Some(s)
              case CellValue.Number(n) => Some(n.bigDecimal.stripTrailingZeros().toPlainString)
              case CellValue.Bool(b) => Some(if b then "TRUE" else "FALSE")
              case CellValue.Formula(_, Some(cached)) => extractTextForMatch(cached)
              case _ => None

            def extractNumericForMatch(cv: CellValue): Option[BigDecimal] = cv match
              case CellValue.Number(n) => Some(n)
              case CellValue.Text(s) => scala.util.Try(BigDecimal(s.trim)).toOption
              case CellValue.Bool(b) => Some(if b then BigDecimal(1) else BigDecimal(0))
              case CellValue.Formula(_, Some(cached)) => extractNumericForMatch(cached)
              case _ => None

            val normalizedLookup: Any = lookupValue match
              case cv: CellValue =>
                cv match
                  case CellValue.Number(n) => n
                  case CellValue.Text(s) => s
                  case CellValue.Bool(b) => b
                  case CellValue.Formula(_, Some(cached)) =>
                    cached match
                      case CellValue.Number(n) => n
                      case CellValue.Text(s) => s
                      case CellValue.Bool(b) => b
                      case other => other
                  case other => other
              case other => other

            val isTextLookup = normalizedLookup match
              case _: String => true
              case _: BigDecimal => false
              case _: Int => false
              case _: Boolean => false
              case _ => true

            val chosenRowOpt: Option[Int] =
              if rangeMatch then
                val numericLookup: Option[BigDecimal] = normalizedLookup match
                  case n: BigDecimal => Some(n)
                  case i: Int => Some(BigDecimal(i))
                  case s: String => scala.util.Try(BigDecimal(s.trim)).toOption
                  case _ => None

                numericLookup.flatMap { lookup =>
                  val keyedRows: List[(Int, BigDecimal)] =
                    rowIndices.toList.flatMap { i =>
                      val keyRef = ARef.from0(keyCol0, rowStart0 + i)
                      extractNumericForMatch(targetSheet(keyRef).value).map(k => (i, k))
                    }
                  keyedRows
                    .filter(_._2 <= lookup)
                    .sortBy(_._2)
                    .lastOption
                    .map(_._1)
                }
              else if isTextLookup then
                val lookupText = normalizedLookup.toString.toLowerCase
                rowIndices.find { i =>
                  val keyRef = ARef.from0(keyCol0, rowStart0 + i)
                  extractTextForMatch(targetSheet(keyRef).value)
                    .exists(_.toLowerCase == lookupText)
                }
              else
                val numericLookup: Option[BigDecimal] = normalizedLookup match
                  case n: BigDecimal => Some(n)
                  case i: Int => Some(BigDecimal(i))
                  case _ => None
                numericLookup.flatMap { lookup =>
                  rowIndices.find { i =>
                    val keyRef = ARef.from0(keyCol0, rowStart0 + i)
                    extractNumericForMatch(targetSheet(keyRef).value).contains(lookup)
                  }
                }

            chosenRowOpt match
              case Some(rowIndex) =>
                val resultRef = ARef.from0(resultCol0, rowStart0 + rowIndex)
                Right(targetSheet(resultRef).value)
              case None =>
                Left(
                  EvalError.EvalFailed(
                    if rangeMatch then "VLOOKUP approximate match not found"
                    else "VLOOKUP exact match not found",
                    Some(
                      s"VLOOKUP($normalizedLookup, ${table.range.toA1}, $colIndex, $rangeMatch)"
                    )
                  )
                )
      yield result
    }

  val sumproduct: FunctionSpec[BigDecimal] { type Args = SumProductArgs } =
    FunctionSpec.simple[BigDecimal, SumProductArgs]("SUMPRODUCT", Arity.atLeastOne) {
      (arrays, ctx) =>
        arrays match
          case Nil => Right(BigDecimal(0))
          case first :: rest =>
            val firstWidth = first.width
            val firstHeight = first.height

            val dimensionError = rest.collectFirst {
              case range if range.width != firstWidth || range.height != firstHeight =>
                EvalError.EvalFailed(
                  s"SUMPRODUCT: all arrays must have same dimensions (first is ${firstHeight}×${firstWidth}, got ${range.height}×${range.width})",
                  Some(s"SUMPRODUCT(${first.toA1}, ${range.toA1}, ...)")
                )
            }

            dimensionError match
              case Some(err) => Left(err)
              case None =>
                val cellLists = arrays.map(_.cells.toList)
                val cellCount = cellLists.headOption.map(_.length).getOrElse(0)

                val sum = (0 until cellCount).foldLeft(BigDecimal(0)) { (acc, idx) =>
                  val values = cellLists.map { cells =>
                    val ref = cells(idx)
                    coerceToNumeric(ctx.sheet(ref).value)
                  }
                  val product = values.foldLeft(BigDecimal(1))(_ * _)
                  acc + product
                }

                Right(sum)
    }

  val xlookup: FunctionSpec[CellValue] { type Args = XLookupArgs } =
    FunctionSpec.simple[CellValue, XLookupArgs]("XLOOKUP", Arity.Range(3, 6)) { (args, ctx) =>
      val (lookupValue, lookupArray, returnArray, ifNotFoundOpt, matchModeOpt, searchModeOpt) =
        args
      val matchModeExpr = matchModeOpt.getOrElse(TExpr.Lit(0))
      val searchModeExpr = searchModeOpt.getOrElse(TExpr.Lit(1))
      if lookupArray.width != returnArray.width || lookupArray.height != returnArray.height then
        Left(
          EvalError.EvalFailed(
            s"XLOOKUP: lookup_array and return_array must have same dimensions (${lookupArray.height}×${lookupArray.width} vs ${returnArray.height}×${returnArray.width})",
            Some(s"XLOOKUP(..., ${lookupArray.toA1}, ${returnArray.toA1}, ...)")
          )
        )
      else
        for
          lookupValueEval <- evalAny(ctx, lookupValue)
          matchModeRaw <- evalAny(ctx, matchModeExpr)
          searchModeRaw <- evalAny(ctx, searchModeExpr)
          matchMode = toInt(matchModeRaw)
          searchMode = toInt(searchModeRaw)
          result <- performXLookup(
            lookupValueEval,
            lookupArray,
            returnArray,
            ifNotFoundOpt,
            matchMode,
            searchMode,
            ctx
          )
        yield result
    }

  val row: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr]("ROW", Arity.one) { (expr, ctx) =>
      extractARef(expr) match
        case Some(aref) => Right(BigDecimal(aref.row.index0 + 1))
        case None =>
          Left(
            EvalError.EvalFailed(
              "ROW requires a cell reference",
              Some(s"ROW($expr)")
            )
          )
    }

  val column: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr]("COLUMN", Arity.one) { (expr, ctx) =>
      extractARef(expr) match
        case Some(aref) => Right(BigDecimal(aref.col.index0 + 1))
        case None =>
          Left(
            EvalError.EvalFailed(
              "COLUMN requires a cell reference",
              Some(s"COLUMN($expr)")
            )
          )
    }

  val rows: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr]("ROWS", Arity.one) { (expr, ctx) =>
      extractCellRange(expr) match
        case Some(range) =>
          val rowCount = range.rowEnd.index0 - range.rowStart.index0 + 1
          Right(BigDecimal(rowCount))
        case None =>
          Left(
            EvalError.EvalFailed(
              "ROWS requires a range argument",
              Some(s"ROWS($expr)")
            )
          )
    }

  val columns: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr]("COLUMNS", Arity.one) { (expr, ctx) =>
      extractCellRange(expr) match
        case Some(range) =>
          val colCount = range.colEnd.index0 - range.colStart.index0 + 1
          Right(BigDecimal(colCount))
        case None =>
          Left(
            EvalError.EvalFailed(
              "COLUMNS requires a range argument",
              Some(s"COLUMNS($expr)")
            )
          )
    }

  val address: FunctionSpec[String] { type Args = AddressArgs } =
    FunctionSpec.simple[String, AddressArgs]("ADDRESS", Arity.Range(2, 5)) { (args, ctx) =>
      val (rowExpr, colExpr, absNumOpt, a1Opt, sheetOpt) = args
      val absNumExpr = absNumOpt.getOrElse(TExpr.Lit(BigDecimal(1)))
      val a1Expr = a1Opt.getOrElse(TExpr.Lit(true))
      for
        row <- ctx.evalExpr(rowExpr)
        col <- ctx.evalExpr(colExpr)
        absNum <- ctx.evalExpr(absNumExpr)
        a1Style <- ctx.evalExpr(a1Expr)
        sheetName <- sheetOpt match
          case Some(expr) => ctx.evalExpr(expr).map(Some(_))
          case None => Right(None)
      yield
        val rowInt = row.toInt
        val colInt = col.toInt
        val absType = absNum.toInt

        if rowInt < 1 || colInt < 1 then "#VALUE!"
        else if a1Style then
          val colLetter = columnToLetter(colInt - 1)
          val (colPrefix, rowPrefix) = absType match
            case 1 => ("$", "$")
            case 2 => ("", "$")
            case 3 => ("$", "")
            case _ => ("", "")
          val refStr = s"$colPrefix$colLetter$rowPrefix$rowInt"
          sheetName match
            case Some(sn) => s"$sn!$refStr"
            case None => refStr
        else
          val rowPart = absType match
            case 1 | 2 => s"R$rowInt"
            case _ => s"R[$rowInt]"
          val colPart = absType match
            case 1 | 3 => s"C$colInt"
            case _ => s"C[$colInt]"
          val refStr = s"$rowPart$colPart"
          sheetName match
            case Some(sn) => s"$sn!$refStr"
            case None => refStr
    }

  val index: FunctionSpec[CellValue] { type Args = IndexArgs } =
    FunctionSpec.simple[CellValue, IndexArgs]("INDEX", Arity.Range(2, 3)) { (args, ctx) =>
      val (array, rowNumExpr, colNumOpt) = args
      for
        rowNum <- ctx.evalExpr(rowNumExpr)
        colNum <- colNumOpt match
          case Some(expr) => ctx.evalExpr(expr).map(Some(_))
          case None => Right(None)
        result <- {
          val rowIdx = rowNum.toInt - 1
          val colIdx = colNum.map(_.toInt - 1).getOrElse(0)
          val startCol = array.colStart.index0
          val startRow = array.rowStart.index0
          val numCols = array.colEnd.index0 - startCol + 1
          val numRows = array.rowEnd.index0 - startRow + 1

          if rowIdx < 0 || rowIdx >= numRows then
            Left(
              EvalError.EvalFailed(
                s"INDEX: row_num ${rowNum.toInt} is out of bounds (array has $numRows rows, valid range: 1-$numRows) (#REF!)",
                Some(s"INDEX(${array.toA1}, $rowNum${colNum.map(c => s", $c").getOrElse("")})")
              )
            )
          else if colIdx < 0 || colIdx >= numCols then
            Left(
              EvalError.EvalFailed(
                s"INDEX: col_num ${colNum.map(_.toInt).getOrElse(1)} is out of bounds (array has $numCols columns, valid range: 1-$numCols) (#REF!)",
                Some(s"INDEX(${array.toA1}, $rowNum${colNum.map(c => s", $c").getOrElse("")})")
              )
            )
          else
            val targetRef = ARef.from0(startCol + colIdx, startRow + rowIdx)
            Right(ctx.sheet(targetRef).value)
        }
      yield result
    }

  val matchFn: FunctionSpec[BigDecimal] { type Args = MatchArgs } =
    FunctionSpec.simple[BigDecimal, MatchArgs]("MATCH", Arity.Range(2, 3)) { (args, ctx) =>
      val (lookupValue, lookupArray, matchTypeOpt) = args
      val matchTypeExpr = matchTypeOpt.getOrElse(TExpr.Lit(BigDecimal(1)))
      for
        lookupValueEval <- evalAny(ctx, lookupValue)
        matchType <- ctx.evalExpr(matchTypeExpr)
        result <- {
          val matchTypeInt = matchType.toInt
          val cells: List[(Int, CellValue)] =
            lookupArray.cells.toList.zipWithIndex.map { case (ref, idx) =>
              (idx + 1, ctx.sheet(ref).value)
            }

          val positionOpt: Option[Int] = matchTypeInt match
            case 0 =>
              cells
                .find { case (_, cv) =>
                  compareCellValues(cv, lookupValueEval) == 0
                }
                .map(_._1)
            case 1 =>
              val numericLookup = coerceToBigDecimal(lookupValueEval)
              val candidates = cells.flatMap { case (pos, cv) =>
                val numericCv = coerceToNumeric(cv)
                if numericCv <= numericLookup then Some((pos, numericCv))
                else None
              }
              candidates.maxByOption(_._2).map(_._1)
            case -1 =>
              val numericLookup = coerceToBigDecimal(lookupValueEval)
              val candidates = cells.flatMap { case (pos, cv) =>
                val numericCv = coerceToNumeric(cv)
                if numericCv >= numericLookup then Some((pos, numericCv))
                else None
              }
              candidates.minByOption(_._2).map(_._1)
            case _ =>
              None

          positionOpt match
            case Some(pos) => Right(BigDecimal(pos))
            case None =>
              Left(
                EvalError.EvalFailed(
                  "MATCH: no match found for lookup value (#N/A)",
                  Some("MATCH(lookup_value, lookup_array, [match_type])")
                )
              )
        }
      yield result
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
