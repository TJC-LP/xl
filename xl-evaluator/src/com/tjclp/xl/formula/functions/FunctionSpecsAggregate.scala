package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator, CriteriaMatcher, Aggregator, ArrayResult}
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.CellValue

trait FunctionSpecsAggregate extends FunctionSpecsBase:
  import ArgSpec.NumericArg

  // Import the ArgSpec for variadic numeric args
  protected given variadicNumeric: ArgSpec[List[NumericArg]] = ArgSpec.list[NumericArg]

  private type RowBounds = (Row, Row)
  private type ColBounds = (Column, Column)

  /**
   * GH-192: Compute shared bounds across all involved sheets for full-column/row optimization.
   *
   * We still enforce Excel's dimension rules on the original ranges, but we constrain full
   * column/row ranges to the union of used ranges across all involved sheets. This prevents
   * mismatched lengths after constraining while preserving performance.
   */
  private def computeBounds(
    ranges: List[(CellRange, com.tjclp.xl.sheets.Sheet)]
  ): (Option[RowBounds], Option[ColBounds]) =
    val usedRanges = ranges.flatMap(_._2.usedRange)
    val rowBounds =
      if ranges.exists(_._1.isFullColumn) then
        if usedRanges.isEmpty then None
        else
          val minRow =
            usedRanges.foldLeft(Int.MaxValue)((acc, r) => Math.min(acc, r.rowStart.index0))
          val maxRow = usedRanges.foldLeft(Int.MinValue)((acc, r) => Math.max(acc, r.rowEnd.index0))
          Some((Row.from0(minRow), Row.from0(maxRow)))
      else None
    val colBounds =
      if ranges.exists(_._1.isFullRow) then
        if usedRanges.isEmpty then None
        else
          val minCol =
            usedRanges.foldLeft(Int.MaxValue)((acc, r) => Math.min(acc, r.colStart.index0))
          val maxCol = usedRanges.foldLeft(Int.MinValue)((acc, r) => Math.max(acc, r.colEnd.index0))
          Some((Column.from0(minCol), Column.from0(maxCol)))
      else None
    (rowBounds, colBounds)

  /**
   * GH-192: Constrain full-column/row ranges to shared bounds.
   *
   * If all sheets are empty (no usedRange), full-column/row ranges collapse to CellRange.empty.
   */
  private def constrainRange(
    range: CellRange,
    bounds: (Option[RowBounds], Option[ColBounds])
  ): CellRange =
    val (rowBounds, colBounds) = bounds
    if range.isFullColumn && rowBounds.isEmpty then CellRange.empty
    else if range.isFullRow && colBounds.isEmpty then CellRange.empty
    else
      val rowStart =
        if range.isFullColumn then rowBounds.map(_._1).getOrElse(range.rowStart) else range.rowStart
      val rowEnd =
        if range.isFullColumn then rowBounds.map(_._2).getOrElse(range.rowEnd) else range.rowEnd
      val colStart =
        if range.isFullRow then colBounds.map(_._1).getOrElse(range.colStart) else range.colStart
      val colEnd =
        if range.isFullRow then colBounds.map(_._2).getOrElse(range.colEnd) else range.colEnd

      if rowStart.index0 > rowEnd.index0 || colStart.index0 > colEnd.index0 then CellRange.empty
      else
        new CellRange(
          ARef(colStart, rowStart),
          ARef(colEnd, rowEnd),
          range.startAnchor,
          range.endAnchor
        )

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
            // GH-192: Constrain full-column/row ranges to used area for performance
            val bounds = computeBounds(List((location.range, targetSheet)))
            val constrainedRange = constrainRange(location.range, bounds)
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

        // Validate dimensions using original ranges (Excel semantics)
        if rangeLocation.range.width != effectiveLocation.range.width ||
          rangeLocation.range.height != effectiveLocation.range.height
        then
          Left(
            EvalError.EvalFailed(
              s"SUMIF: range and sum_range must have same dimensions (${rangeLocation.range.height}×${rangeLocation.range.width} vs ${effectiveLocation.range.height}×${effectiveLocation.range.width})",
              Some(s"SUMIF(${rangeLocation.toA1}, ..., ${effectiveLocation.toA1})")
            )
          )
        else
          // GH-192: Resolve target sheets for cross-sheet support BEFORE constraining
          for
            criteriaSheet <- Evaluator.resolveRangeLocation(rangeLocation, ctx.sheet, ctx.workbook)
            sumSheet <- Evaluator.resolveRangeLocation(effectiveLocation, ctx.sheet, ctx.workbook)
            result <- {
              val bounds = computeBounds(
                List(
                  (rangeLocation.range, criteriaSheet),
                  (effectiveLocation.range, sumSheet)
                )
              )
              // GH-192: Constrain full-column/row ranges to shared bounds
              val criteriaRange = constrainRange(rangeLocation.range, bounds)
              val sumRange = constrainRange(effectiveLocation.range, bounds)

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
            // GH-192: Constrain full-column/row ranges to used area for performance
            val bounds = computeBounds(List((rangeLocation.range, criteriaSheet)))
            val constrainedRange = constrainRange(rangeLocation.range, bounds)
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

          val dimensionError = parsedConditions.collectFirst {
            case (loc, _)
                if loc.range.width != sumRangeLocation.range.width ||
                  loc.range.height != sumRangeLocation.range.height =>
              EvalError.EvalFailed(
                s"SUMIFS: all ranges must have same dimensions (sum_range is ${sumRangeLocation.range.height}×${sumRangeLocation.range.width}, criteria_range is ${loc.range.height}×${loc.range.width})",
                Some(s"SUMIFS(${sumRangeLocation.toA1}, ...)")
              )
          }

          dimensionError match
            case Some(err) => Left(err)
            case None =>
              // GH-192: Resolve sum range and all criteria ranges to their target sheets FIRST
              Evaluator.resolveRangeLocation(sumRangeLocation, ctx.sheet, ctx.workbook).flatMap {
                sumSheet =>
                  // Resolve all criteria ranges upfront
                  val resolvedConditions: Either[
                    EvalError,
                    List[
                      (com.tjclp.xl.sheets.Sheet, TExpr.RangeLocation, CriteriaMatcher.Criterion)
                    ]
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
                    val bounds = computeBounds(
                      (sumRangeLocation.range, sumSheet) ::
                        resolved.map { case (sheet, loc, _) => (loc.range, sheet) }
                    )
                    // GH-192: Constrain full-column/row ranges to shared bounds
                    val constrainedSumRange = constrainRange(sumRangeLocation.range, bounds)
                    val constrainedConditions = resolved.map { case (sheet, loc, criterion) =>
                      (sheet, constrainRange(loc.range, bounds), criterion)
                    }

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
              case (firstLoc, _) :: rest =>
                val dimensionError = rest.collectFirst {
                  case (loc, _)
                      if loc.range.width != firstLoc.range.width ||
                        loc.range.height != firstLoc.range.height =>
                    EvalError.EvalFailed(
                      s"COUNTIFS: all ranges must have same dimensions (first is ${firstLoc.range.height}×${firstLoc.range.width}, this is ${loc.range.height}×${loc.range.width})",
                      Some(s"COUNTIFS(...)")
                    )
                }

                dimensionError match
                  case Some(err) => Left(err)
                  case None =>
                    // GH-192: Resolve all criteria ranges to their target sheets FIRST
                    val resolvedConditions: Either[
                      EvalError,
                      List[
                        (com.tjclp.xl.sheets.Sheet, TExpr.RangeLocation, CriteriaMatcher.Criterion)
                      ]
                    ] =
                      parsedConditions.foldLeft[Either[
                        EvalError,
                        List[
                          (
                            com.tjclp.xl.sheets.Sheet,
                            TExpr.RangeLocation,
                            CriteriaMatcher.Criterion
                          )
                        ]
                      ]](Right(List.empty)) {
                        case (Left(err), _) => Left(err)
                        case (Right(acc), (loc, criterion)) =>
                          Evaluator.resolveRangeLocation(loc, ctx.sheet, ctx.workbook).map {
                            sheet =>
                              acc :+ (sheet, loc, criterion)
                          }
                      }

                    resolvedConditions.flatMap { resolved =>
                      val bounds = computeBounds(resolved.map { case (sheet, loc, _) =>
                        (loc.range, sheet)
                      })
                      // GH-192: Constrain full-column/row ranges to shared bounds
                      val constrainedConditions = resolved.map { case (sheet, loc, criterion) =>
                        (sheet, constrainRange(loc.range, bounds), criterion)
                      }

                      // GH-192: Use iterator-based folding with index tracking
                      val criteriaCells =
                        constrainedConditions.map { case (sheet, range, criterion) =>
                          (sheet, range.cells.toVector, criterion)
                        }
                      val refCount = criteriaCells.headOption.map(_._2.length).getOrElse(0)

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
                    }
          }
    }

  val averageif: FunctionSpec[BigDecimal] { type Args = AverageIfArgs } =
    FunctionSpec.simple[BigDecimal, AverageIfArgs]("AVERAGEIF", Arity.Range(2, 3)) { (args, ctx) =>
      val (rangeLocation, criteria, avgRangeLocationOpt) = args
      evalValue(ctx, criteria).flatMap { criteriaValue =>
        val criterion = CriteriaMatcher.parse(criteriaValue)
        val effectiveLocation = avgRangeLocationOpt.getOrElse(rangeLocation)

        // Validate dimensions using original ranges (Excel semantics)
        if rangeLocation.range.width != effectiveLocation.range.width ||
          rangeLocation.range.height != effectiveLocation.range.height
        then
          Left(
            EvalError.EvalFailed(
              s"AVERAGEIF: range and average_range must have same dimensions (${rangeLocation.range.height}×${rangeLocation.range.width} vs ${effectiveLocation.range.height}×${effectiveLocation.range.width})",
              Some(s"AVERAGEIF(${rangeLocation.toA1}, ..., ${effectiveLocation.toA1})")
            )
          )
        else
          // GH-192: Resolve target sheets for cross-sheet support BEFORE constraining
          for
            criteriaSheet <- Evaluator.resolveRangeLocation(rangeLocation, ctx.sheet, ctx.workbook)
            avgSheet <- Evaluator.resolveRangeLocation(effectiveLocation, ctx.sheet, ctx.workbook)
            result <- {
              val bounds = computeBounds(
                List(
                  (rangeLocation.range, criteriaSheet),
                  (effectiveLocation.range, avgSheet)
                )
              )
              // GH-192: Constrain full-column/row ranges to shared bounds
              val criteriaRange = constrainRange(rangeLocation.range, bounds)
              val avgRange = constrainRange(effectiveLocation.range, bounds)

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

          val dimensionError = parsedConditions.collectFirst {
            case (loc, _)
                if loc.range.width != avgRangeLocation.range.width ||
                  loc.range.height != avgRangeLocation.range.height =>
              EvalError.EvalFailed(
                s"AVERAGEIFS: all ranges must have same dimensions (average_range is ${avgRangeLocation.range.height}×${avgRangeLocation.range.width}, criteria_range is ${loc.range.height}×${loc.range.width})",
                Some(s"AVERAGEIFS(${avgRangeLocation.toA1}, ...)")
              )
          }

          dimensionError match
            case Some(err) => Left(err)
            case None =>
              // GH-192: Resolve average range and all criteria ranges to their target sheets FIRST
              Evaluator.resolveRangeLocation(avgRangeLocation, ctx.sheet, ctx.workbook).flatMap {
                avgSheet =>
                  // Resolve all criteria ranges upfront
                  val resolvedConditions: Either[
                    EvalError,
                    List[
                      (com.tjclp.xl.sheets.Sheet, TExpr.RangeLocation, CriteriaMatcher.Criterion)
                    ]
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
                    val bounds = computeBounds(
                      (avgRangeLocation.range, avgSheet) ::
                        resolved.map { case (sheet, loc, _) => (loc.range, sheet) }
                    )
                    // GH-192: Constrain full-column/row ranges to shared bounds
                    val constrainedAvgRange = constrainRange(avgRangeLocation.range, bounds)
                    val constrainedConditions = resolved.map { case (sheet, loc, criterion) =>
                      (sheet, constrainRange(loc.range, bounds), criterion)
                    }

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

  /**
   * GH-197: Resolved array for SUMPRODUCT.
   *
   * Represents a resolved argument as either:
   *   - A range with its sheet reference for cell-by-cell access
   *   - A pre-computed numeric matrix from an evaluated expression
   */
  private sealed trait ResolvedArray:
    def rows: Int
    def cols: Int
    def valueAt(row: Int, col: Int, ctx: EvalContext): Either[EvalError, BigDecimal]

  private final case class RangeArray(
    sheet: com.tjclp.xl.sheets.Sheet,
    range: CellRange,
    ctx: EvalContext
  ) extends ResolvedArray:
    def rows: Int = range.height
    def cols: Int = range.width
    def valueAt(row: Int, col: Int, ctx: EvalContext): Either[EvalError, BigDecimal] =
      val ref = ARef.from0(range.colStart.index0 + col, range.rowStart.index0 + row)
      coerceToNumericWithEval(sheet(ref).value, sheet, ctx)

  private final case class MatrixArray(matrix: Vector[Vector[BigDecimal]]) extends ResolvedArray:
    def rows: Int = matrix.length
    def cols: Int = matrix.headOption.map(_.length).getOrElse(0)
    def valueAt(row: Int, col: Int, ctx: EvalContext): Either[EvalError, BigDecimal] =
      Right(matrix(row)(col))

  val sumproduct: FunctionSpec[BigDecimal] { type Args = SumProductArgs } =
    FunctionSpec.simple[BigDecimal, SumProductArgs]("SUMPRODUCT", Arity.atLeastOne) { (args, ctx) =>
      import com.tjclp.xl.formula.ast.TExpr

      args match
        case Nil => Right(BigDecimal(0))
        case _ =>
          // GH-197: First collect ALL ranges from both locations and expressions for bounds calc
          val allRangesWithSheets: Either[EvalError, List[(CellRange, com.tjclp.xl.sheets.Sheet)]] =
            args.foldLeft[Either[EvalError, List[(CellRange, com.tjclp.xl.sheets.Sheet)]]](
              Right(List.empty)
            ) {
              case (Left(err), _) => Left(err)
              case (Right(acc), Left(loc)) =>
                // Range location argument
                Evaluator.resolveRangeLocation(loc, ctx.sheet, ctx.workbook).map { sheet =>
                  acc :+ (loc.range, sheet)
                }
              case (Right(acc), Right(expr)) =>
                // Expression argument - collect ranges from AST
                val exprRanges = TExpr.collectRanges(expr)
                exprRanges
                  .foldLeft[Either[EvalError, List[(CellRange, com.tjclp.xl.sheets.Sheet)]]](
                    Right(acc)
                  ) {
                    case (Left(err), _) => Left(err)
                    case (Right(innerAcc), (sheetOpt, range)) =>
                      sheetOpt match
                        case Some(sheetName) =>
                          // Cross-sheet range - resolve sheet
                          ctx.workbook match
                            case Some(wb) =>
                              wb(sheetName) match
                                case Right(targetSheet) => Right(innerAcc :+ (range, targetSheet))
                                case Left(_) =>
                                  Left(
                                    EvalError.EvalFailed(
                                      s"Sheet '$sheetName' not found",
                                      Some("SUMPRODUCT")
                                    )
                                  )
                            case None =>
                              Left(
                                EvalError.EvalFailed(
                                  "Cross-sheet reference requires workbook context",
                                  Some("SUMPRODUCT")
                                )
                              )
                        case None =>
                          // Local range - use current sheet
                          Right(innerAcc :+ (range, ctx.sheet))
                  }
            }

          allRangesWithSheets.flatMap { rangesWithSheets =>
            // GH-192/197: Compute shared bounds across ALL ranges (locations + expressions)
            val bounds: (Option[RowBounds], Option[ColBounds]) =
              if rangesWithSheets.nonEmpty then computeBounds(rangesWithSheets)
              else (None, None)

            // Helper to constrain ranges in expressions
            def constrainExprRanges(expr: TExpr[Any]): TExpr[Any] =
              TExpr.transformRanges(
                expr,
                { (sheetOpt, range) =>
                  constrainRange(range, bounds)
                }
              )

            // GH-197: Resolve each argument to a ResolvedArray with bounded ranges
            val resolvedResult: Either[EvalError, List[ResolvedArray]] =
              args.foldLeft[Either[EvalError, List[ResolvedArray]]](Right(List.empty)) {
                case (Left(err), _) => Left(err)
                case (Right(acc), Left(loc)) =>
                  // Range location - resolve to sheet and constrain
                  Evaluator.resolveRangeLocation(loc, ctx.sheet, ctx.workbook).map { sheet =>
                    acc :+ RangeArray(sheet, constrainRange(loc.range, bounds), ctx)
                  }
                case (Right(acc), Right(expr)) =>
                  // GH-197: Expression - constrain ranges, then evaluate with array support
                  val boundedExpr = constrainExprRanges(expr)
                  ctx.evalArrayExpr(boundedExpr).flatMap {
                    case ar: ArrayResult =>
                      // Convert ArrayResult to numeric matrix with boolean coercion
                      val matrix = (0 until ar.rows).map { row =>
                        (0 until ar.cols).map { col =>
                          coerceToNumeric(ar.values(row)(col))
                        }.toVector
                      }.toVector
                      Right(acc :+ MatrixArray(matrix))
                    case bd: BigDecimal =>
                      // Scalar treated as 1x1 matrix
                      Right(acc :+ MatrixArray(Vector(Vector(bd))))
                    case b: Boolean =>
                      // Boolean coerced to 1/0
                      val n = if b then BigDecimal(1) else BigDecimal(0)
                      Right(acc :+ MatrixArray(Vector(Vector(n))))
                    case cv: CellValue =>
                      // CellValue coerced to numeric
                      Right(acc :+ MatrixArray(Vector(Vector(coerceToNumeric(cv)))))
                    case other =>
                      Left(EvalError.TypeMismatch("SUMPRODUCT", "array or number", other.toString))
                  }
              }

            resolvedResult.flatMap { resolved =>
              resolved match
                case Nil => Right(BigDecimal(0))
                case first :: rest =>
                  // Validate dimensions match
                  val dimensionError = rest.collectFirst {
                    case arr if arr.rows != first.rows || arr.cols != first.cols =>
                      EvalError.EvalFailed(
                        s"SUMPRODUCT: all arrays must have same dimensions (first is ${first.rows}×${first.cols}, got ${arr.rows}×${arr.cols})",
                        Some("SUMPRODUCT(...)")
                      )
                  }

                  dimensionError match
                    case Some(err) => Left(err)
                    case None =>
                      // Compute final dimensions
                      val finalRows = resolved.headOption.map(_.rows).getOrElse(0)
                      val finalCols = resolved.headOption.map(_.cols).getOrElse(0)

                      // Element-wise multiplication and sum
                      (0 until finalRows).foldLeft[Either[EvalError, BigDecimal]](
                        Right(BigDecimal(0))
                      ) {
                        case (Left(err), _) => Left(err)
                        case (Right(acc), row) =>
                          (0 until finalCols).foldLeft[Either[EvalError, BigDecimal]](Right(acc)) {
                            case (Left(err), _) => Left(err)
                            case (Right(rowAcc), col) =>
                              // Get values from all arrays at this position
                              resolved
                                .foldLeft[Either[EvalError, BigDecimal]](Right(BigDecimal(1))) {
                                  case (Left(err), _) => Left(err)
                                  case (Right(product), arr) =>
                                    arr.valueAt(row, col, ctx).map(v => product * v)
                                }
                                .map(product => rowAcc + product)
                          }
                      }
            }
          }
    }
