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

  /**
   * GH-192: Constrain range to target sheet's used area for full-column/row optimization.
   *
   * When users write formulas like `=SUMIFS(C:C, A:A, "x")`, full columns contain 1M+ cells. This
   * method constrains such ranges to the actual used area of the target sheet, reducing iteration
   * from millions to just the populated rows.
   *
   * @param range
   *   The range to constrain (may be full column like A:A or full row like 1:1)
   * @param sheet
   *   The target sheet whose usedRange provides the bounds
   * @return
   *   The constrained range, or empty range if no intersection
   */
  private def constrainToUsedRange(range: CellRange, sheet: com.tjclp.xl.sheets.Sheet): CellRange =
    if range.isFullColumn || range.isFullRow then
      sheet.usedRange.flatMap(range.intersect).getOrElse(CellRange.empty)
    else range

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
          .evalCrossSheetFormula(formulaStr, targetSheet, ctx.clock, ctx.workbook, ctx.depth + 1)
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
          .evalCrossSheetFormula(formulaStr, targetSheet, ctx.clock, ctx.workbook, ctx.depth + 1)
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
          .evalCrossSheetFormula(formulaStr, targetSheet, ctx.clock, ctx.workbook, ctx.depth + 1)
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
            // GH-192: Constrain range to used area for full-column optimization
            val constrainedRange = constrainToUsedRange(location.range, targetSheet)
            // GH-192: Use iterator-based folding (no .toList) for memory efficiency
            constrainedRange.cells.foldLeft[Either[EvalError, Vector[BigDecimal]]](Right(acc)) {
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
  ): List[(TExpr.RangeLocation, CriteriaMatcher.Criterion)] =
    conditions
      .zip(criteriaValues)
      .map { case ((location, _), criteriaValue) =>
        (location, CriteriaMatcher.parse(criteriaValue))
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
      val (rangeLocation, criteria, sumRangeLocationOpt) = args
      evalValue(ctx, criteria).flatMap { criteriaValue =>
        val criterion = CriteriaMatcher.parse(criteriaValue)
        val effectiveLocation = sumRangeLocationOpt.getOrElse(rangeLocation)

        // GH-192: Resolve target sheets for cross-sheet support BEFORE constraining
        for
          criteriaSheet <- Evaluator.resolveRangeLocation(rangeLocation, ctx.sheet, ctx.workbook)
          sumSheet <- Evaluator.resolveRangeLocation(effectiveLocation, ctx.sheet, ctx.workbook)
          result <- {
            // GH-192: Constrain ranges to used area for full-column optimization
            val criteriaRange = constrainToUsedRange(rangeLocation.range, criteriaSheet)
            val sumRange = constrainToUsedRange(effectiveLocation.range, sumSheet)

            // Validate dimensions AFTER constraining (full columns always match before)
            if criteriaRange.width != sumRange.width || criteriaRange.height != sumRange.height then
              Left(
                EvalError.EvalFailed(
                  s"SUMIF: range and sum_range must have same dimensions (${criteriaRange.height}×${criteriaRange.width} vs ${sumRange.height}×${sumRange.width})",
                  Some(s"SUMIF(${rangeLocation.toA1}, ..., ${effectiveLocation.toA1})")
                )
              )
            else
              // GH-192: Use iterator-based folding (no .toList) for memory efficiency
              criteriaRange.cells
                .zip(sumRange.cells)
                .foldLeft[Either[EvalError, BigDecimal]](Right(BigDecimal(0))) {
                  case (Left(err), _) => Left(err)
                  case (Right(acc), (testRef, sumRef)) =>
                    // Evaluate test cell value (may be uncached formula)
                    evalCellValueForMatch(criteriaSheet(testRef).value, criteriaSheet, ctx)
                      .flatMap { testValue =>
                        if CriteriaMatcher.matches(testValue, criterion) then
                          extractOrEvalNumeric(sumSheet(sumRef).value, sumSheet, ctx).map {
                            case Some(n) => acc + n
                            case None => acc
                          }
                        else Right(acc)
                      }
                }
          }
        yield result
      }
    }

  val countif: FunctionSpec[BigDecimal] { type Args = CountIfArgs } =
    FunctionSpec.simple[BigDecimal, CountIfArgs]("COUNTIF", Arity.two) { (args, ctx) =>
      val (rangeLocation, criteria) = args
      evalValue(ctx, criteria).flatMap { criteriaValue =>
        val criterion = CriteriaMatcher.parse(criteriaValue)
        // GH-192: Resolve target sheet for cross-sheet support
        Evaluator.resolveRangeLocation(rangeLocation, ctx.sheet, ctx.workbook).flatMap {
          criteriaSheet =>
            // GH-192: Constrain range to used area for full-column optimization
            val constrainedRange = constrainToUsedRange(rangeLocation.range, criteriaSheet)
            // GH-192: Use iterator-based folding (no .toList) for memory efficiency
            constrainedRange.cells
              .foldLeft[Either[EvalError, Int]](Right(0)) {
                case (Left(err), _) => Left(err)
                case (Right(count), ref) =>
                  evalCellValueForMatch(criteriaSheet(ref).value, criteriaSheet, ctx).map {
                    testValue =>
                      if CriteriaMatcher.matches(testValue, criterion) then count + 1 else count
                  }
              }
              .map(BigDecimal(_))
        }
      }
    }

  val sumifs: FunctionSpec[BigDecimal] { type Args = SumIfsArgs } =
    FunctionSpec.simple[BigDecimal, SumIfsArgs]("SUMIFS", Arity.AtLeast(3)) { (args, ctx) =>
      val (sumRangeLocation, conditions) = args
      evalCriteriaValues(ctx, conditions)
        .flatMap { criteriaValues =>
          val parsedConditions = parseConditions(conditions, criteriaValues)

          // GH-192: Resolve sum range and all criteria ranges to their target sheets FIRST
          Evaluator.resolveRangeLocation(sumRangeLocation, ctx.sheet, ctx.workbook).flatMap {
            sumSheet =>
              // Resolve all criteria ranges upfront
              val resolvedConditions: Either[
                EvalError,
                List[(com.tjclp.xl.sheets.Sheet, TExpr.RangeLocation, CriteriaMatcher.Criterion)]
              ] =
                parsedConditions.foldLeft[Either[
                  EvalError,
                  List[(com.tjclp.xl.sheets.Sheet, TExpr.RangeLocation, CriteriaMatcher.Criterion)]
                ]](Right(List.empty)) {
                  case (Left(err), _) => Left(err)
                  case (Right(acc), (loc, criterion)) =>
                    Evaluator.resolveRangeLocation(loc, ctx.sheet, ctx.workbook).map { sheet =>
                      acc :+ (sheet, loc, criterion)
                    }
                }

              resolvedConditions.flatMap { resolved =>
                // GH-192: Constrain all ranges to used area for full-column optimization
                val constrainedSumRange = constrainToUsedRange(sumRangeLocation.range, sumSheet)
                val constrainedConditions = resolved.map { case (sheet, loc, criterion) =>
                  (sheet, constrainToUsedRange(loc.range, sheet), criterion)
                }

                // Validate dimensions AFTER constraining
                val dimensionError = constrainedConditions.collectFirst {
                  case (_, criteriaRange, _)
                      if criteriaRange.width != constrainedSumRange.width ||
                        criteriaRange.height != constrainedSumRange.height =>
                    EvalError.EvalFailed(
                      s"SUMIFS: all ranges must have same dimensions (sum_range is ${constrainedSumRange.height}×${constrainedSumRange.width}, criteria_range is ${criteriaRange.height}×${criteriaRange.width})",
                      Some(s"SUMIFS(${sumRangeLocation.toA1}, ...)")
                    )
                }

                dimensionError match
                  case Some(err) => Left(err)
                  case None =>
                    // GH-192: Use iterator-based folding with index tracking
                    val sumCells = constrainedSumRange.cells.toVector
                    val criteriaCells =
                      constrainedConditions.map { case (sheet, range, criterion) =>
                        (sheet, range.cells.toVector, criterion)
                      }

                    sumCells.indices.foldLeft[Either[EvalError, BigDecimal]](
                      Right(BigDecimal(0))
                    ) {
                      case (Left(err), _) => Left(err)
                      case (Right(acc), idx) =>
                        // Check all conditions
                        val matchResult =
                          criteriaCells.foldLeft[Either[EvalError, Boolean]](Right(true)) {
                            case (Left(err), _) => Left(err)
                            case (Right(false), _) => Right(false) // Short-circuit
                            case (Right(true), (criteriaSheet, cells, criterion)) =>
                              val testRef = cells(idx)
                              evalCellValueForMatch(
                                criteriaSheet(testRef).value,
                                criteriaSheet,
                                ctx
                              ).map { testValue =>
                                CriteriaMatcher.matches(testValue, criterion)
                              }
                          }
                        matchResult.flatMap { allMatch =>
                          if allMatch then
                            extractOrEvalNumeric(sumSheet(sumCells(idx)).value, sumSheet, ctx)
                              .map {
                                case Some(n) => acc + n
                                case None => acc
                              }
                          else Right(acc)
                        }
                    }
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
              case Nil => Right(BigDecimal(0))
              case _ =>
                // GH-192: Resolve all criteria ranges to their target sheets FIRST
                val resolvedConditions: Either[
                  EvalError,
                  List[(com.tjclp.xl.sheets.Sheet, TExpr.RangeLocation, CriteriaMatcher.Criterion)]
                ] =
                  parsedConditions.foldLeft[Either[
                    EvalError,
                    List[
                      (com.tjclp.xl.sheets.Sheet, TExpr.RangeLocation, CriteriaMatcher.Criterion)
                    ]
                  ]](Right(List.empty)) {
                    case (Left(err), _) => Left(err)
                    case (Right(acc), (loc, criterion)) =>
                      Evaluator.resolveRangeLocation(loc, ctx.sheet, ctx.workbook).map { sheet =>
                        acc :+ (sheet, loc, criterion)
                      }
                  }

                resolvedConditions.flatMap { resolved =>
                  // GH-192: Constrain all ranges to used area for full-column optimization
                  val constrainedConditions = resolved.map { case (sheet, loc, criterion) =>
                    (sheet, constrainToUsedRange(loc.range, sheet), criterion)
                  }

                  // Validate dimensions AFTER constraining
                  constrainedConditions match
                    case (_, firstRange, _) :: rest =>
                      val dimensionError = rest.collectFirst {
                        case (_, range, _)
                            if range.width != firstRange.width || range.height != firstRange.height =>
                          EvalError.EvalFailed(
                            s"COUNTIFS: all ranges must have same dimensions (first is ${firstRange.height}×${firstRange.width}, this is ${range.height}×${range.width})",
                            Some(s"COUNTIFS(...)")
                          )
                      }

                      dimensionError match
                        case Some(err) => Left(err)
                        case None =>
                          // GH-192: Use iterator-based folding with index tracking
                          val criteriaCells =
                            constrainedConditions.map { case (sheet, range, criterion) =>
                              (sheet, range.cells.toVector, criterion)
                            }
                          val refCount =
                            criteriaCells.headOption.map(_._2.length).getOrElse(0)

                          (0 until refCount)
                            .foldLeft[Either[EvalError, Int]](Right(0)) {
                              case (Left(err), _) => Left(err)
                              case (Right(count), idx) =>
                                // Check all conditions
                                val matchResult =
                                  criteriaCells.foldLeft[Either[EvalError, Boolean]](Right(true)) {
                                    case (Left(err), _) => Left(err)
                                    case (Right(false), _) => Right(false) // Short-circuit
                                    case (Right(true), (criteriaSheet, cells, criterion)) =>
                                      val testRef = cells(idx)
                                      evalCellValueForMatch(
                                        criteriaSheet(testRef).value,
                                        criteriaSheet,
                                        ctx
                                      )
                                        .map { testValue =>
                                          CriteriaMatcher.matches(testValue, criterion)
                                        }
                                  }
                                matchResult.map { allMatch =>
                                  if allMatch then count + 1 else count
                                }
                            }
                            .map(BigDecimal(_))
                    case Nil => Right(BigDecimal(0))
                }
          }
    }

  val averageif: FunctionSpec[BigDecimal] { type Args = AverageIfArgs } =
    FunctionSpec.simple[BigDecimal, AverageIfArgs]("AVERAGEIF", Arity.Range(2, 3)) { (args, ctx) =>
      val (rangeLocation, criteria, avgRangeLocationOpt) = args
      evalValue(ctx, criteria).flatMap { criteriaValue =>
        val criterion = CriteriaMatcher.parse(criteriaValue)
        val effectiveLocation = avgRangeLocationOpt.getOrElse(rangeLocation)

        // GH-192: Resolve target sheets for cross-sheet support BEFORE constraining
        for
          criteriaSheet <- Evaluator.resolveRangeLocation(rangeLocation, ctx.sheet, ctx.workbook)
          avgSheet <- Evaluator.resolveRangeLocation(effectiveLocation, ctx.sheet, ctx.workbook)
          result <- {
            // GH-192: Constrain ranges to used area for full-column optimization
            val criteriaRange = constrainToUsedRange(rangeLocation.range, criteriaSheet)
            val avgRange = constrainToUsedRange(effectiveLocation.range, avgSheet)

            // Validate dimensions AFTER constraining
            if criteriaRange.width != avgRange.width || criteriaRange.height != avgRange.height then
              Left(
                EvalError.EvalFailed(
                  s"AVERAGEIF: range and average_range must have same dimensions (${criteriaRange.height}×${criteriaRange.width} vs ${avgRange.height}×${avgRange.width})",
                  Some(s"AVERAGEIF(${rangeLocation.toA1}, ..., ${effectiveLocation.toA1})")
                )
              )
            else
              // GH-192: Use iterator-based folding (no .toList) for memory efficiency
              criteriaRange.cells
                .zip(avgRange.cells)
                .foldLeft[Either[EvalError, (BigDecimal, Int)]](Right((BigDecimal(0), 0))) {
                  case (Left(err), _) => Left(err)
                  case (Right((accSum, accCount)), (testRef, avgRef)) =>
                    // Evaluate test cell value (may be uncached formula)
                    evalCellValueForMatch(criteriaSheet(testRef).value, criteriaSheet, ctx)
                      .flatMap { testValue =>
                        if CriteriaMatcher.matches(testValue, criterion) then
                          extractOrEvalNumeric(avgSheet(avgRef).value, avgSheet, ctx).map {
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
        yield result
      }
    }

  val averageifs: FunctionSpec[BigDecimal] { type Args = AverageIfsArgs } =
    FunctionSpec.simple[BigDecimal, AverageIfsArgs]("AVERAGEIFS", Arity.AtLeast(3)) { (args, ctx) =>
      val (avgRangeLocation, conditions) = args
      evalCriteriaValues(ctx, conditions)
        .flatMap { criteriaValues =>
          val parsedConditions = parseConditions(conditions, criteriaValues)

          // GH-192: Resolve average range and all criteria ranges to their target sheets FIRST
          Evaluator.resolveRangeLocation(avgRangeLocation, ctx.sheet, ctx.workbook).flatMap {
            avgSheet =>
              // Resolve all criteria ranges upfront
              val resolvedConditions: Either[
                EvalError,
                List[(com.tjclp.xl.sheets.Sheet, TExpr.RangeLocation, CriteriaMatcher.Criterion)]
              ] =
                parsedConditions.foldLeft[Either[
                  EvalError,
                  List[(com.tjclp.xl.sheets.Sheet, TExpr.RangeLocation, CriteriaMatcher.Criterion)]
                ]](Right(List.empty)) {
                  case (Left(err), _) => Left(err)
                  case (Right(acc), (loc, criterion)) =>
                    Evaluator.resolveRangeLocation(loc, ctx.sheet, ctx.workbook).map { sheet =>
                      acc :+ (sheet, loc, criterion)
                    }
                }

              resolvedConditions.flatMap { resolved =>
                // GH-192: Constrain all ranges to used area for full-column optimization
                val constrainedAvgRange = constrainToUsedRange(avgRangeLocation.range, avgSheet)
                val constrainedConditions = resolved.map { case (sheet, loc, criterion) =>
                  (sheet, constrainToUsedRange(loc.range, sheet), criterion)
                }

                // Validate dimensions AFTER constraining
                val dimensionError = constrainedConditions.collectFirst {
                  case (_, criteriaRange, _)
                      if criteriaRange.width != constrainedAvgRange.width ||
                        criteriaRange.height != constrainedAvgRange.height =>
                    EvalError.EvalFailed(
                      s"AVERAGEIFS: all ranges must have same dimensions (average_range is ${constrainedAvgRange.height}×${constrainedAvgRange.width}, criteria_range is ${criteriaRange.height}×${criteriaRange.width})",
                      Some(s"AVERAGEIFS(${avgRangeLocation.toA1}, ...)")
                    )
                }

                dimensionError match
                  case Some(err) => Left(err)
                  case None =>
                    // GH-192: Use iterator-based folding with index tracking
                    val avgCells = constrainedAvgRange.cells.toVector
                    val criteriaCells =
                      constrainedConditions.map { case (sheet, range, criterion) =>
                        (sheet, range.cells.toVector, criterion)
                      }

                    avgCells.indices
                      .foldLeft[Either[EvalError, (BigDecimal, Int)]](Right((BigDecimal(0), 0))) {
                        case (Left(err), _) => Left(err)
                        case (Right((accSum, accCount)), idx) =>
                          // Check all conditions
                          val matchResult =
                            criteriaCells.foldLeft[Either[EvalError, Boolean]](Right(true)) {
                              case (Left(err), _) => Left(err)
                              case (Right(false), _) => Right(false) // Short-circuit
                              case (Right(true), (criteriaSheet, cells, criterion)) =>
                                val testRef = cells(idx)
                                evalCellValueForMatch(
                                  criteriaSheet(testRef).value,
                                  criteriaSheet,
                                  ctx
                                ).map { testValue =>
                                  CriteriaMatcher.matches(testValue, criterion)
                                }
                            }
                          matchResult.flatMap { allMatch =>
                            if allMatch then
                              extractOrEvalNumeric(avgSheet(avgCells(idx)).value, avgSheet, ctx)
                                .map {
                                  case Some(n) => (accSum + n, accCount + 1)
                                  case None => (accSum, accCount)
                                }
                            else Right((accSum, accCount))
                          }
                      }
                      .flatMap { case (sum, count) =>
                        if count == 0 then
                          Left(
                            EvalError.DivByZero("AVERAGEIFS sum", "0 (no matching numeric cells)")
                          )
                        else Right(sum / count)
                      }
              }
          }
        }
    }

  val sumproduct: FunctionSpec[BigDecimal] { type Args = SumProductArgs } =
    FunctionSpec.simple[BigDecimal, SumProductArgs]("SUMPRODUCT", Arity.atLeastOne) {
      (arrayLocations, ctx) =>
        arrayLocations match
          case Nil => Right(BigDecimal(0))
          case _ =>
            // GH-192: Resolve all arrays to their target sheets FIRST
            val resolvedArrays: Either[
              EvalError,
              List[(com.tjclp.xl.sheets.Sheet, TExpr.RangeLocation)]
            ] =
              arrayLocations.foldLeft[Either[
                EvalError,
                List[(com.tjclp.xl.sheets.Sheet, TExpr.RangeLocation)]
              ]](Right(List.empty)) {
                case (Left(err), _) => Left(err)
                case (Right(acc), loc) =>
                  Evaluator.resolveRangeLocation(loc, ctx.sheet, ctx.workbook).map { sheet =>
                    acc :+ (sheet, loc)
                  }
              }

            resolvedArrays.flatMap { resolved =>
              // GH-192: Constrain all ranges to used area for full-column optimization
              val constrainedArrays = resolved.map { case (sheet, loc) =>
                (sheet, constrainToUsedRange(loc.range, sheet))
              }

              // Validate dimensions AFTER constraining
              constrainedArrays match
                case (_, firstRange) :: rest =>
                  val dimensionError = rest.collectFirst {
                    case (_, range)
                        if range.width != firstRange.width || range.height != firstRange.height =>
                      EvalError.EvalFailed(
                        s"SUMPRODUCT: all arrays must have same dimensions (first is ${firstRange.height}×${firstRange.width}, got ${range.height}×${range.width})",
                        Some(s"SUMPRODUCT(...)")
                      )
                  }

                  dimensionError match
                    case Some(err) => Left(err)
                    case None =>
                      // GH-192: Use iterator-based folding with index tracking
                      val arrayCells = constrainedArrays.map { case (sheet, range) =>
                        (sheet, range.cells.toVector)
                      }
                      val cellCount = arrayCells.headOption.map(_._2.length).getOrElse(0)

                      (0 until cellCount).foldLeft[Either[EvalError, BigDecimal]](
                        Right(BigDecimal(0))
                      ) {
                        case (Left(err), _) => Left(err)
                        case (Right(acc), idx) =>
                          // Evaluate each cell in the row, collecting values
                          arrayCells
                            .foldLeft[Either[EvalError, List[BigDecimal]]](Right(List.empty)) {
                              case (Left(err), _) => Left(err)
                              case (Right(values), (sheet, cells)) =>
                                val ref = cells(idx)
                                coerceToNumericWithEval(sheet(ref).value, sheet, ctx)
                                  .map(n => values :+ n)
                            }
                            .map { values =>
                              val product = values.foldLeft(BigDecimal(1))(_ * _)
                              acc + product
                            }
                      }
                case Nil => Right(BigDecimal(0))
            }
    }
