package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator, CriteriaMatcher, Aggregator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.CellValue

trait FunctionSpecsAggregate extends FunctionSpecsBase:
  import ArgSpec.NumericArg

  // Import the ArgSpec for variadic numeric args
  protected given variadicNumeric: ArgSpec[List[NumericArg]] = ArgSpec.list[NumericArg]

  // GH-187: Helper to extract numeric value, evaluating uncached formulas if needed.
  // Moved to class level so it can be reused by conditional aggregates (SUMIF, SUMIFS, etc.)
  private def extractOrEvalNumeric(
    cellValue: CellValue,
    targetSheet: com.tjclp.xl.sheets.Sheet,
    ctx: EvalContext
  ): Either[EvalError, Option[BigDecimal]] =
    cellValue match
      case CellValue.Formula(formulaStr, None) =>
        // Recursively evaluate uncached formula
        Evaluator
          .evalCrossSheetFormula(formulaStr, targetSheet, ctx.clock, ctx.workbook)
          .map {
            case CellValue.Number(n) => Some(n)
            case _ => None // Non-numeric result, skip
          }
      case _ =>
        // Fall back to standard extraction
        Right(extractNumericValue(cellValue))

  // GH-187: Helper to coerce cell value to numeric, evaluating uncached formulas if needed.
  // Used by SUMPRODUCT which needs BigDecimal(0) for non-numeric values instead of None.
  private def coerceToNumericWithEval(
    cellValue: CellValue,
    targetSheet: com.tjclp.xl.sheets.Sheet,
    ctx: EvalContext
  ): Either[EvalError, BigDecimal] =
    cellValue match
      case CellValue.Formula(formulaStr, None) =>
        // Recursively evaluate uncached formula
        Evaluator
          .evalCrossSheetFormula(formulaStr, targetSheet, ctx.clock, ctx.workbook)
          .map(coerceToNumeric)
      case _ =>
        Right(coerceToNumeric(cellValue))

  // GH-187: Helper to evaluate cell value for criteria matching.
  // Evaluates uncached formulas before matching, returning the resolved CellValue.
  private def evalCellValueForMatch(
    cellValue: CellValue,
    targetSheet: com.tjclp.xl.sheets.Sheet,
    ctx: EvalContext
  ): Either[EvalError, CellValue] =
    cellValue match
      case CellValue.Formula(formulaStr, None) =>
        // Recursively evaluate uncached formula
        Evaluator
          .evalCrossSheetFormula(formulaStr, targetSheet, ctx.clock, ctx.workbook)
      case CellValue.Formula(_, Some(cached)) =>
        // Use cached value
        Right(cached)
      case other =>
        Right(other)

  /**
   * Create a variadic aggregate function spec.
   *
   * Supports Excel-compatible syntax: SUM(1,2,3) / SUM(A1:A5,B1:B5) / SUM(A1,5,B1:B3)
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def variadicAggregateSpec(
    name: String
  ): FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    FunctionSpec.simple[BigDecimal, List[NumericArg]](name, Arity.atLeastOne) { (args, ctx) =>
      Aggregator.lookup(name) match
        case None =>
          Left(EvalError.EvalFailed(s"Unknown aggregator: $name", None))
        case Some(agg) =>
          evalVariadicAggregate(agg, args, ctx)
    }

  /** Helper to evaluate variadic aggregates with proper type handling. */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def evalVariadicAggregate[A](
    agg: Aggregator[A],
    args: List[NumericArg],
    ctx: EvalContext
  ): Either[EvalError, BigDecimal] =
    // Collect all numeric values from both ranges and individual expressions
    val valuesResult: Either[EvalError, Vector[BigDecimal]] =
      args.foldLeft[Either[EvalError, Vector[BigDecimal]]](Right(Vector.empty)) {
        case (Left(err), _) => Left(err)
        case (Right(acc), Left(location)) =>
          // Range argument - extract all numeric values from cells
          Evaluator.resolveRangeLocation(location, ctx.sheet, ctx.workbook).flatMap { targetSheet =>
            // GH-187: Use fold to handle potential errors from formula evaluation
            location.range.cells.foldLeft[Either[EvalError, Vector[BigDecimal]]](Right(acc)) {
              case (Left(err), _) => Left(err)
              case (Right(values), cellRef) =>
                val cellValue = targetSheet(cellRef).value
                // Handle different cell types for aggregation
                if agg.countsNonEmpty then
                  // COUNTA mode: count any non-empty cell
                  cellValue match
                    case CellValue.Empty => Right(values)
                    case _ => Right(values :+ BigDecimal(1))
                else if agg.countsEmpty then
                  // COUNTBLANK mode: count only empty cells
                  cellValue match
                    case CellValue.Empty => Right(values :+ BigDecimal(1))
                    case _ => Right(values)
                else
                  // Standard numeric mode: extract or evaluate formulas
                  extractOrEvalNumeric(cellValue, targetSheet, ctx).map {
                    case Some(n) => values :+ n
                    case None => values
                  }
            }
          }
        case (Right(acc), Right(expr)) =>
          // Individual numeric expression - evaluate it
          ctx.evalExpr(expr).map { value =>
            if agg.countsNonEmpty || agg.countsEmpty then
              // For COUNT/COUNTA/COUNTBLANK, individual values always count
              acc :+ BigDecimal(1)
            else acc :+ value
          }
      }

    valuesResult.flatMap { values =>
      // Apply the aggregator to all collected values
      val result = values.foldLeft(agg.empty)((acc, v) => agg.combine(acc, v))
      agg.finalizeWithError(result)
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

  val sum: FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    variadicAggregateSpec("SUM")

  val count: FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    variadicAggregateSpec("COUNT")

  val counta: FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    variadicAggregateSpec("COUNTA")

  val countblank: FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    variadicAggregateSpec("COUNTBLANK")

  val average: FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    variadicAggregateSpec("AVERAGE")

  val min: FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    variadicAggregateSpec("MIN")

  val max: FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    variadicAggregateSpec("MAX")

  val median: FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    variadicAggregateSpec("MEDIAN")

  val stdev: FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    variadicAggregateSpec("STDEV")

  val stdevp: FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    variadicAggregateSpec("STDEVP")

  val variance: FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    variadicAggregateSpec("VAR")

  val variancep: FunctionSpec[BigDecimal] { type Args = List[NumericArg] } =
    variadicAggregateSpec("VARP")

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
          // GH-187: Use fold to handle uncached formula evaluation in both test and sum ranges
          pairs.foldLeft[Either[EvalError, BigDecimal]](Right(BigDecimal(0))) {
            case (Left(err), _) => Left(err)
            case (Right(acc), (testRef, sumRef)) =>
              // Evaluate test cell value (may be uncached formula)
              evalCellValueForMatch(ctx.sheet(testRef).value, ctx.sheet, ctx).flatMap { testValue =>
                if CriteriaMatcher.matches(testValue, criterion) then
                  extractOrEvalNumeric(ctx.sheet(sumRef).value, ctx.sheet, ctx).map {
                    case Some(n) => acc + n
                    case None => acc
                  }
                else Right(acc)
              }
          }
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
              // GH-187: Use fold to handle uncached formula evaluation in both test and sum ranges
              sumRefsList.indices.foldLeft[Either[EvalError, BigDecimal]](Right(BigDecimal(0))) {
                case (Left(err), _) => Left(err)
                case (Right(acc), idx) =>
                  // Check all conditions, evaluating uncached formulas in test cells
                  val matchResult =
                    parsedConditions.foldLeft[Either[EvalError, Boolean]](Right(true)) {
                      case (Left(err), _) => Left(err)
                      case (Right(false), _) => Right(false) // Short-circuit if already failed
                      case (Right(true), (criteriaRange, criterion)) =>
                        val testRef = criteriaRange.cells.toList(idx)
                        evalCellValueForMatch(ctx.sheet(testRef).value, ctx.sheet, ctx).map {
                          testValue =>
                            CriteriaMatcher.matches(testValue, criterion)
                        }
                    }
                  matchResult.flatMap { allMatch =>
                    if allMatch then
                      extractOrEvalNumeric(ctx.sheet(sumRefsList(idx)).value, ctx.sheet, ctx).map {
                        case Some(n) => acc + n
                        case None => acc
                      }
                    else Right(acc)
                  }
              }
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
          // GH-187: Use fold to handle uncached formula evaluation in both test and average ranges
          pairs
            .foldLeft[Either[EvalError, (BigDecimal, Int)]](Right((BigDecimal(0), 0))) {
              case (Left(err), _) => Left(err)
              case (Right((accSum, accCount)), (testRef, avgRef)) =>
                // Evaluate test cell value (may be uncached formula)
                evalCellValueForMatch(ctx.sheet(testRef).value, ctx.sheet, ctx).flatMap {
                  testValue =>
                    if CriteriaMatcher.matches(testValue, criterion) then
                      extractOrEvalNumeric(ctx.sheet(avgRef).value, ctx.sheet, ctx).map {
                        case Some(n) => (accSum + n, accCount + 1)
                        case None => (accSum, accCount)
                      }
                    else Right((accSum, accCount))
                }
            }
            .flatMap { case (sum, count) =>
              if count == 0 then
                Left(EvalError.DivByZero("AVERAGEIF sum", "0 (no matching numeric cells)"))
              else Right(sum / count)
            }
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
              // GH-187: Use fold to handle uncached formula evaluation in both test and average ranges
              avgRefsList.indices
                .foldLeft[Either[EvalError, (BigDecimal, Int)]](Right((BigDecimal(0), 0))) {
                  case (Left(err), _) => Left(err)
                  case (Right((accSum, accCount)), idx) =>
                    // Check all conditions, evaluating uncached formulas in test cells
                    val matchResult =
                      parsedConditions.foldLeft[Either[EvalError, Boolean]](Right(true)) {
                        case (Left(err), _) => Left(err)
                        case (Right(false), _) => Right(false) // Short-circuit if already failed
                        case (Right(true), (criteriaRange, criterion)) =>
                          val testRef = criteriaRange.cells.toList(idx)
                          evalCellValueForMatch(ctx.sheet(testRef).value, ctx.sheet, ctx).map {
                            testValue =>
                              CriteriaMatcher.matches(testValue, criterion)
                          }
                      }
                    matchResult.flatMap { allMatch =>
                      if allMatch then
                        extractOrEvalNumeric(ctx.sheet(avgRefsList(idx)).value, ctx.sheet, ctx)
                          .map {
                            case Some(n) => (accSum + n, accCount + 1)
                            case None => (accSum, accCount)
                          }
                      else Right((accSum, accCount))
                    }
                }
                .flatMap { case (sum, count) =>
                  if count == 0 then
                    Left(EvalError.DivByZero("AVERAGEIFS sum", "0 (no matching numeric cells)"))
                  else Right(sum / count)
                }
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

                // GH-187: Use fold to handle uncached formula evaluation
                (0 until cellCount).foldLeft[Either[EvalError, BigDecimal]](Right(BigDecimal(0))) {
                  case (Left(err), _) => Left(err)
                  case (Right(acc), idx) =>
                    // Evaluate each cell in the row, collecting values
                    cellLists
                      .foldLeft[Either[EvalError, List[BigDecimal]]](Right(List.empty)) {
                        case (Left(err), _) => Left(err)
                        case (Right(values), cells) =>
                          val ref = cells(idx)
                          coerceToNumericWithEval(ctx.sheet(ref).value, ctx.sheet, ctx)
                            .map(n => values :+ n)
                      }
                      .map { values =>
                        val product = values.foldLeft(BigDecimal(1))(_ * _)
                        acc + product
                      }
                }
    }
