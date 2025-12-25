package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator, CriteriaMatcher}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.CellValue

trait FunctionSpecsAggregate extends FunctionSpecsBase:
  private def aggregateSpec(
    name: String
  ): FunctionSpec[BigDecimal] { type Args = UnaryRange } =
    FunctionSpec.simple[BigDecimal, UnaryRange](name, Arity.one) { (location, ctx) =>
      ctx.evalExpr(TExpr.Aggregate(name, location))
    }

  private def evalCriteriaValues(
    ctx: EvalContext,
    conditions: RangeCriteriaList
  ): Either[EvalError, List[ExprValue]] =
    conditions
      .map { case (_, criteriaExpr) => evalValue(ctx, criteriaExpr) }
      .foldLeft[Either[EvalError, List[ExprValue]]](Right(List.empty)) { (acc, either) =>
        acc.flatMap(list => either.map(v => v :: list))
      }
      .map(_.reverse)

  private def parseConditions(
    conditions: RangeCriteriaList,
    criteriaValues: List[ExprValue]
  ): List[(CellRange, CriteriaMatcher.Criterion)] =
    conditions
      .zip(criteriaValues)
      .map { case ((range, _), criteriaValue) =>
        (range, CriteriaMatcher.parse(criteriaValue))
      }

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
      evalValue(ctx, criteria).flatMap { criteriaValue =>
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
      evalValue(ctx, criteria).map { criteriaValue =>
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
      evalCriteriaValues(ctx, conditions)
        .flatMap { criteriaValues =>
          val parsedConditions = parseConditions(conditions, criteriaValues)

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
        evalCriteriaValues(ctx, conditions)
          .flatMap { criteriaValues =>
            val parsedConditions = parseConditions(conditions, criteriaValues)

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
      evalValue(ctx, criteria).flatMap { criteriaValue =>
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
      evalCriteriaValues(ctx, conditions)
        .flatMap { criteriaValues =>
          val parsedConditions = parseConditions(conditions, criteriaValues)

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
