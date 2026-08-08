package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{
  EvalError,
  Evaluator,
  CriteriaMatcher,
  Aggregator,
  ArrayResult,
  ArrayArithmetic
}
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{CellError, CellValue}

trait FunctionSpecsAggregate extends FunctionSpecsBase:
  import ArgSpec.NumericArg

  // Import the ArgSpec for variadic numeric args
  protected given variadicNumeric: ArgSpec[List[NumericArg]] = ArgSpec.list[NumericArg]

  private type RowBounds = (Row, Row)
  private type ColBounds = (Column, Column)
  private val aggregateCountUnit = BigDecimal(1)

  /**
   * GH-337/GH-344: the propagated failure for a carried error element consumed by an aggregate —
   * the element's Excel error VALUE (surfacing as `Right(CellValue.Error(code))` at the result
   * boundary, Excel-exact), IFERROR-catchable, with a context naming the function.
   */
  private def propagatedElementError(fnName: String, err: CellError): EvalError =
    EvalError.ErrorValue(err, Some(s"$fnName: array element contains ${err.toExcel} error"))

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
    val hasFullColumn = ranges.exists(_._1.isFullColumn)
    val hasFullRow = ranges.exists(_._1.isFullRow)
    if !hasFullColumn && !hasFullRow then (None, None)
    else
      // Computing usedRange scans every populated cell. Bounded ranges need no shared bounds, so
      // keep that scan entirely off their hot path.
      val usedRanges = ranges.flatMap(_._2.usedRange)
      val rowBounds =
        if hasFullColumn then
          if usedRanges.isEmpty then None
          else
            val minRow =
              usedRanges.foldLeft(Int.MaxValue)((acc, r) => Math.min(acc, r.rowStart.index0))
            val maxRow =
              usedRanges.foldLeft(Int.MinValue)((acc, r) => Math.max(acc, r.rowEnd.index0))
            Some((Row.from0(minRow), Row.from0(maxRow)))
        else None
      val colBounds =
        if hasFullRow then
          if usedRanges.isEmpty then None
          else
            val minCol =
              usedRanges.foldLeft(Int.MaxValue)((acc, r) => Math.min(acc, r.colStart.index0))
            val maxCol =
              usedRanges.foldLeft(Int.MinValue)((acc, r) => Math.max(acc, r.colEnd.index0))
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

  /**
   * GH-344 item 6: resolve-then-police for raw-range numeric extraction — resolve the cell to its
   * effective CellValue FIRST (recursively evaluating uncached formulas via
   * [[evalCellValueForMatch]], which also fixes the pre-GH-344 swallow where an uncached formula
   * recursing to an error mapped to None and vanished), THEN police carried errors: propagate as
   * the element's Excel error VALUE when the caller's policy says so, or skip (COUNT — an error is
   * not a number). Non-error values keep the numeric extract-or-skip semantics.
   */
  private def resolveNumericPolicing(
    cellValue: CellValue,
    targetSheet: com.tjclp.xl.sheets.Sheet,
    ctx: EvalContext,
    fnName: String,
    propagateErrors: Boolean
  ): Either[EvalError, Option[BigDecimal]] =
    evalCellValueForMatch(cellValue, targetSheet, ctx).flatMap { resolved =>
      ArrayArithmetic.carriedError(resolved) match
        case Some(err) if propagateErrors => Left(propagatedElementError(fnName, err))
        case Some(_) => Right(None) // COUNT: errors are not numbers
        case None => Right(extractNumericValue(resolved))
    }

  // GH-187: Helper to coerce cell value to numeric, evaluating uncached formulas if needed.
  // Used by SUMPRODUCT which needs BigDecimal(0) for non-numeric values instead of None.
  // GH-344 item 6: carried error cells propagate as the error VALUE (the raw-range coerce-to-0
  // leniency produced wrong sums); non-error non-numerics keep coercing to 0.
  private def coerceToNumericWithEval(
    cellValue: CellValue,
    targetSheet: com.tjclp.xl.sheets.Sheet,
    ctx: EvalContext
  ): Either[EvalError, BigDecimal] =
    evalCellValueForMatch(cellValue, targetSheet, ctx).flatMap { resolved =>
      ArrayArithmetic.carriedError(resolved) match
        case Some(err) => Left(propagatedElementError("SUMPRODUCT", err))
        case None => Right(coerceToNumeric(resolved))
    }

  // GH-187: Helper to evaluate cell value for criteria matching.
  // Evaluates uncached formulas before matching, returning the resolved CellValue.
  private def evalCellValueForMatch(
    cellValue: CellValue,
    targetSheet: com.tjclp.xl.sheets.Sheet,
    ctx: EvalContext
  ): Either[EvalError, CellValue] =
    cellValue match
      case CellValue.Formula(formulaStr, None, _) =>
        // Recursively evaluate uncached formula (GH-346: memoized once per pass)
        Evaluator
          .evalCrossSheetFormula(
            formulaStr,
            targetSheet,
            ctx.clock,
            ctx.workbook,
            ctx.depth + 1,
            ctx.rng,
            ctx.memo.getOrElse(new Evaluator.EvalMemo),
            aggregateMemo = ctx.aggregateMemo
          )
      case CellValue.Formula(_, Some(cached), _) =>
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
    FunctionSpec.simple[BigDecimal, List[NumericArg]](
      name,
      Arity.atLeastOne,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      Aggregator.lookup(name) match
        case None =>
          Left(EvalError.EvalFailed(s"Unknown aggregator: $name", None))
        case Some(agg) =>
          evalVariadicAggregate(agg, args, ctx)
    }

  /**
   * GH-395: the per-CELL triage shared by the raw-range fold and direct single-cell reference
   * arguments — COUNTA counts non-empty, COUNTBLANK counts empty, numeric mode resolves the
   * effective value (recursively evaluating uncached formulas), polices carried errors per the
   * aggregator's policy (GH-344 item 6), then extracts-or-skips. Factored so =COUNT(A1, A2) with a
   * blank A2 triages A2 exactly like the 1×1 range =COUNT(A2:A2) would (Excel ignores it) instead
   * of coercing it to a genuine 0 through the scalar decode boundary.
   */
  private def triageCellForAggregate[A](
    agg: Aggregator[A],
    cellValue: CellValue,
    targetSheet: com.tjclp.xl.sheets.Sheet,
    ctx: EvalContext,
    acc: A
  ): Either[EvalError, A] =
    if agg.countsNonEmpty then
      // COUNTA mode: count any non-empty cell
      cellValue match
        case CellValue.Empty => Right(acc)
        case _ => Right(agg.combine(acc, aggregateCountUnit))
    else if agg.countsEmpty then
      // COUNTBLANK mode: count only empty cells
      cellValue match
        case CellValue.Empty => Right(agg.combine(acc, aggregateCountUnit))
        case _ => Right(acc)
    else
      // Standard numeric mode: resolve the effective value, police carried errors
      // per the aggregator's policy (GH-344 item 6), then extract-or-skip
      resolveNumericPolicing(
        cellValue,
        targetSheet,
        ctx,
        agg.name,
        agg.propagatesErrors
      )
        .map {
          case Some(n) => agg.combine(acc, n)
          case None => acc
        }

  /** Resolve a raw range and apply the full-row/full-column used-area constraint once. */
  private def resolveConstrainedRange(
    location: TExpr.RangeLocation,
    ctx: EvalContext
  ): Either[EvalError, (com.tjclp.xl.sheets.Sheet, CellRange)] =
    Evaluator.resolveRangeLocation(location, ctx.sheet, ctx.workbook).map {
      case (targetSheet, range) =>
        val bounds = computeBounds(List((range, targetSheet)))
        (targetSheet, constrainRange(range, bounds))
    }

  /** Stream one already-constrained raw range into the supplied accumulator. */
  private def foldRawRange[A](
    agg: Aggregator[A],
    targetSheet: com.tjclp.xl.sheets.Sheet,
    range: CellRange,
    ctx: EvalContext,
    initial: A
  ): Either[EvalError, A] =
    range.cells.foldLeft[Either[EvalError, A]](Right(initial)) {
      case (Left(err), _) => Left(err)
      case (Right(current), cellRef) =>
        triageCellForAggregate(agg, targetSheet(cellRef).value, targetSheet, ctx, current)
    }

  /** Helper to evaluate variadic aggregates with proper type handling. */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def evalVariadicAggregate[A](
    agg: Aggregator[A],
    args: List[NumericArg],
    ctx: EvalContext
  ): Either[EvalError, BigDecimal] =
    // Stream values into the aggregator's own accumulator. Aggregators that need to retain their
    // inputs (for example MEDIAN) do so in A; ordinary aggregates avoid an intermediate Vector.
    def foldAllArgs: Either[EvalError, A] =
      args.foldLeft[Either[EvalError, A]](Right(agg.empty)) {
        case (Left(err), _) => Left(err)
        case (Right(acc), Left(location)) =>
          resolveConstrainedRange(location, ctx).flatMap { case (targetSheet, constrainedRange) =>
            foldRawRange(agg, targetSheet, constrainedRange, ctx, acc)
          }
        // GH-395: direct single-cell references triage per-cell like a 1×1 range. In NumericArg
        // position Ref/SheetRef can ONLY arise from direct refs (asNumericExpr rewrites
        // PolyRef → Ref and SheetPolyRef → SheetRef; no other NumericArg source produces these
        // shapes), so matching them BEFORE the scalar evaluation keeps Excel's semantics: a
        // blank direct arg is ignored by COUNT/AVERAGE (not coerced to a genuine 0 by the
        // scalar decode), a text cell skips like the range fold, and COUNTBLANK still counts
        // the blank.
        case (Right(acc), Right(TExpr.Ref(at, _, _))) =>
          triageCellForAggregate(agg, ctx.sheet(at).value, ctx.sheet, ctx, acc)
        case (Right(acc), Right(TExpr.SheetRef(sheetName, at, _, _))) =>
          Evaluator
            .resolveRangeLocation(
              TExpr.RangeLocation.CrossSheet(sheetName, CellRange(at, at)),
              ctx.sheet,
              ctx.workbook
            )
            .flatMap { case (targetSheet, _) =>
              triageCellForAggregate(agg, targetSheet(at).value, targetSheet, ctx, acc)
            }
        case (Right(acc), Right(expr)) =>
          // GH-122: evaluate array-aware so a range-returning call (e.g. OFFSET) flattens into the
          // aggregate exactly like a literal range would; scalars keep their existing behavior.
          // GH-337: carried error ELEMENTS in the evaluated array fail loudly (per the
          // aggregator's propagatesErrors policy) instead of silently skipping — an error the
          // broadcast carried must not vanish into a wrong number. Raw-range arguments (the
          // Left(location) branch above) keep their pre-existing skip semantics.
          ctx.evalArrayExpr(expr.asInstanceOf[TExpr[Any]]).flatMap {
            case ar: ArrayResult =>
              ar.values.iterator.flatten.foldLeft[Either[EvalError, A]](Right(acc)) {
                case (Left(err), _) => Left(err)
                case (Right(current), cellValue) =>
                  if agg.countsNonEmpty then
                    cellValue match
                      case CellValue.Empty => Right(current)
                      case _ => Right(agg.combine(current, aggregateCountUnit))
                  else if agg.countsEmpty then
                    cellValue match
                      case CellValue.Empty => Right(agg.combine(current, aggregateCountUnit))
                      case _ => Right(current)
                  else
                    ArrayArithmetic.carriedError(cellValue) match
                      case Some(err) if agg.propagatesErrors =>
                        Left(propagatedElementError(agg.name, err))
                      case Some(_) => Right(current) // COUNT: errors are not numbers
                      case None =>
                        extractNumericValue(cellValue) match
                          case Some(n) => Right(agg.combine(current, n))
                          case None => Right(current)
              }
            case value: BigDecimal =>
              // GH-395: a numeric scalar (literal, arithmetic result) is never blank —
              // COUNTBLANK must not count it (=COUNTBLANK(5) is 0); COUNTA counts it
              Right(
                if agg.countsNonEmpty then agg.combine(acc, aggregateCountUnit)
                else if agg.countsEmpty then acc
                else agg.combine(acc, value)
              )
            case other =>
              // GH-395: non-numeric scalars triage on their CellValue shape — only a genuine
              // CellValue.Empty counts for COUNTBLANK (and is NOT counted by COUNTA); the
              // numeric-mode error policing is unchanged
              val cellValue = ArrayArithmetic.anyToCellValue(other)
              if agg.countsNonEmpty then
                cellValue match
                  case CellValue.Empty => Right(acc)
                  case _ => Right(agg.combine(acc, aggregateCountUnit))
              else if agg.countsEmpty then
                cellValue match
                  case CellValue.Empty => Right(agg.combine(acc, aggregateCountUnit))
                  case _ => Right(acc)
              else
                ArrayArithmetic.carriedError(cellValue) match
                  case Some(err) if agg.propagatesErrors =>
                    Left(propagatedElementError(agg.name, err))
                  case Some(_) => Right(acc)
                  case None =>
                    extractNumericValue(cellValue) match
                      case Some(n) => Right(agg.combine(acc, n))
                      case None => Right(acc)
          }
      }

    // The common workbook-recalc hot shape is one literal raw range. Its finalized immutable
    // result can be shared across formulas in this calculation generation when the target range
    // contains no uncached formula cells. Mixed/scalar/array arguments keep the exact streaming
    // fold above and are deliberately outside the cache's narrow safety proof.
    args match
      case List(Left(location)) =>
        Evaluator.resolveRangeLocation(location, ctx.sheet, ctx.workbook).flatMap {
          case (targetSheet, rawRange) =>
            // Full-row/column aggregate semantics depend on the CURRENT sheet used bounds. Those
            // bounds may shrink when an earlier formula evaluates to Empty even though the formula
            // lies outside the referenced row/column, so resolve them before keying. Ordinary
            // bounded ranges skip usedRange entirely and use their declared geometry directly.
            val effectiveRange =
              if rawRange.isFullColumn || rawRange.isFullRow then
                constrainRange(rawRange, computeBounds(List((rawRange, targetSheet))))
              else rawRange

            def compute: Either[EvalError, BigDecimal] =
              foldRawRange(agg, targetSheet, effectiveRange, ctx, agg.empty)
                .flatMap(agg.finalizeWithError)

            ctx.aggregateMemo match
              case Some(memo) =>
                memo.getOrCompute(
                  targetSheet,
                  effectiveRange,
                  agg.name,
                  Evaluator.AggregateMemoMode.FunctionCall
                )(compute)
              case None => compute
        }
      case _ => foldAllArgs.flatMap(agg.finalizeWithError)

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

  /**
   * GH-394: resolve every (location, criterion) pair to its (sheet, range, criterion) triple
   * through the single resolution boundary — Name locations have no static range, so all dimension
   * math happens on the resolved shapes downstream.
   */
  private def resolveConditions(
    parsedConditions: List[(TExpr.RangeLocation, CriteriaMatcher.Criterion)],
    ctx: EvalContext
  ): Either[EvalError, List[(com.tjclp.xl.sheets.Sheet, CellRange, CriteriaMatcher.Criterion)]] =
    parsedConditions.foldLeft[Either[
      EvalError,
      List[(com.tjclp.xl.sheets.Sheet, CellRange, CriteriaMatcher.Criterion)]
    ]](Right(List.empty)) {
      case (Left(err), _) => Left(err)
      case (Right(acc), (loc, criterion)) =>
        Evaluator.resolveRangeLocation(loc, ctx.sheet, ctx.workbook).map { case (sheet, range) =>
          acc :+ (sheet, range, criterion)
        }
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

  // ===== GH-120: statistical functions over a single range =====

  /**
   * Collect numeric values from one range (standard numeric mode), for LARGE/SMALL/RANK/etc. GH-344
   * item 6: carried error cells always propagate here — every order statistic returns the error
   * VALUE when its data contains one (Excel).
   */
  private def collectRangeNumerics(
    location: TExpr.RangeLocation,
    ctx: EvalContext,
    fnName: String
  ): Either[EvalError, Vector[BigDecimal]] =
    Evaluator.resolveRangeLocation(location, ctx.sheet, ctx.workbook).flatMap {
      case (targetSheet, range) =>
        val bounds = computeBounds(List((range, targetSheet)))
        val constrainedRange = constrainRange(range, bounds)
        constrainedRange.cells
          .foldLeft[Either[EvalError, Vector[BigDecimal]]](Right(Vector.empty)) {
            case (Left(err), _) => Left(err)
            case (Right(values), cellRef) =>
              resolveNumericPolicing(
                targetSheet(cellRef).value,
                targetSheet,
                ctx,
                fnName,
                propagateErrors = true
              ).map {
                case Some(n) => values :+ n
                case None => values
              }
          }
    }

  /** LARGE(range, k) — k-th largest value (1-based). */
  val large: FunctionSpec[BigDecimal] { type Args = RangeIntArgs } =
    FunctionSpec.simple[BigDecimal, RangeIntArgs](
      "LARGE",
      Arity.Exact(2),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (loc, kExpr) = args
      for
        nums <- collectRangeNumerics(loc, ctx, "LARGE")
        k <- ctx.evalExpr(kExpr)
        result <-
          val desc = nums.sorted.reverse
          if k >= 1 && k <= desc.length then Right(desc(k - 1))
          // GH-344: Excel #NUM! (op-level classification)
          else Left(EvalError.ErrorValue(CellError.Num, Some(s"LARGE: k=$k out of range")))
      yield result
    }

  /** SMALL(range, k) — k-th smallest value (1-based). */
  val small: FunctionSpec[BigDecimal] { type Args = RangeIntArgs } =
    FunctionSpec.simple[BigDecimal, RangeIntArgs](
      "SMALL",
      Arity.Exact(2),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (loc, kExpr) = args
      for
        nums <- collectRangeNumerics(loc, ctx, "SMALL")
        k <- ctx.evalExpr(kExpr)
        result <-
          val asc = nums.sorted
          if k >= 1 && k <= asc.length then Right(asc(k - 1))
          // GH-344: Excel #NUM! (op-level classification)
          else Left(EvalError.ErrorValue(CellError.Num, Some(s"SMALL: k=$k out of range")))
      yield result
    }

  /** RANK(number, ref, [order]) — rank of number in ref; order 0/omitted = descending. */
  val rank: FunctionSpec[BigDecimal] { type Args = RankArgs } =
    FunctionSpec.simple[BigDecimal, RankArgs](
      "RANK",
      Arity.Range(2, 3),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (numExpr, loc, orderOpt) = args
      for
        num <- ctx.evalExpr(numExpr)
        nums <- collectRangeNumerics(loc, ctx, "RANK")
        order <- orderOpt match
          case Some(e) => ctx.evalExpr(e)
          case None => Right(0)
        result <-
          if !nums.contains(num) then
            // GH-344: Excel #N/A (op-level classification)
            Left(EvalError.ErrorValue(CellError.NA, Some(s"RANK: $num not found in range")))
          else if order == 0 then Right(BigDecimal(nums.count(_ > num) + 1))
          else Right(BigDecimal(nums.count(_ < num) + 1))
      yield result
    }

  /** Excel PERCENTILE.INC linear interpolation over a sorted-ascending vector; None if invalid. */
  private def percentileInc(sortedAsc: Vector[BigDecimal], p: BigDecimal): Option[BigDecimal] =
    if sortedAsc.isEmpty || p < 0 || p > 1 then None
    else
      val n = sortedAsc.length
      if n == 1 then Some(sortedAsc(0))
      else
        val rank = p * BigDecimal(n - 1)
        val lo = rank.toInt
        if lo >= n - 1 then Some(sortedAsc(n - 1))
        else
          val frac = rank - BigDecimal(lo)
          Some(sortedAsc(lo) + frac * (sortedAsc(lo + 1) - sortedAsc(lo)))

  /** PERCENTILE(range, p) — p in [0,1], inclusive linear interpolation. */
  val percentile: FunctionSpec[BigDecimal] { type Args = RangeNumArgs } =
    FunctionSpec.simple[BigDecimal, RangeNumArgs](
      "PERCENTILE",
      Arity.Exact(2),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (loc, pExpr) = args
      for
        nums <- collectRangeNumerics(loc, ctx, "PERCENTILE")
        p <- ctx.evalExpr(pExpr)
        result <- percentileInc(nums.sorted, p) match
          case Some(v) => Right(v)
          case None =>
            // GH-344: Excel #NUM! (op-level classification)
            Left(
              EvalError.ErrorValue(CellError.Num, Some(s"PERCENTILE: invalid p=$p or empty range"))
            )
      yield result
    }

  /** QUARTILE(range, quart) — quart 0..4 maps to PERCENTILE p = quart/4. */
  val quartile: FunctionSpec[BigDecimal] { type Args = RangeIntArgs } =
    FunctionSpec.simple[BigDecimal, RangeIntArgs](
      "QUARTILE",
      Arity.Exact(2),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (loc, qExpr) = args
      for
        nums <- collectRangeNumerics(loc, ctx, "QUARTILE")
        q <- ctx.evalExpr(qExpr)
        result <-
          if q < 0 || q > 4 then
            // GH-344: Excel #NUM! (op-level classification)
            Left(EvalError.ErrorValue(CellError.Num, Some(s"QUARTILE: quart=$q must be 0-4")))
          else
            percentileInc(nums.sorted, BigDecimal(q) / 4) match
              case Some(v) => Right(v)
              // GH-344: Excel #NUM! (op-level classification)
              case None => Left(EvalError.ErrorValue(CellError.Num, Some("QUARTILE: empty range")))
      yield result
    }

  val sumif: FunctionSpec[BigDecimal] { type Args = SumIfArgs } =
    FunctionSpec.simple[BigDecimal, SumIfArgs](
      "SUMIF",
      Arity.Range(2, 3),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (rangeLocation, criteria, sumRangeLocationOpt) = args
      evalValue(ctx, criteria).flatMap { criteriaValue =>
        val criterion = CriteriaMatcher.parse(criteriaValue)
        val effectiveLocation = sumRangeLocationOpt.getOrElse(rangeLocation)

        // GH-192: Resolve target sheets for cross-sheet support BEFORE constraining.
        // GH-394: resolution now also yields the ranges (Name locations have no static range),
        // so the dimension validation moved after it — same Excel semantics, resolved shapes.
        for
          resolvedCriteria <- Evaluator.resolveRangeLocation(rangeLocation, ctx.sheet, ctx.workbook)
          resolvedSum <- Evaluator.resolveRangeLocation(effectiveLocation, ctx.sheet, ctx.workbook)
          (criteriaSheet, criteriaRange0) = resolvedCriteria
          (sumSheet, sumRange0) = resolvedSum
          _ <-
            if criteriaRange0.width != sumRange0.width ||
              criteriaRange0.height != sumRange0.height
            then
              Left(
                EvalError.EvalFailed(
                  s"SUMIF: range and sum_range must have same dimensions (${criteriaRange0.height}×${criteriaRange0.width} vs ${sumRange0.height}×${sumRange0.width})",
                  Some(s"SUMIF(${rangeLocation.toA1}, ..., ${effectiveLocation.toA1})")
                )
              )
            else Right(())
          result <- {
            val bounds = computeBounds(
              List(
                (criteriaRange0, criteriaSheet),
                (sumRange0, sumSheet)
              )
            )
            // GH-192: Constrain full-column/row ranges to shared bounds
            val criteriaRange = constrainRange(criteriaRange0, bounds)
            val sumRange = constrainRange(sumRange0, bounds)

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
                        resolveNumericPolicing(
                          sumSheet(sumRef).value,
                          sumSheet,
                          ctx,
                          "SUMIF",
                          propagateErrors = true
                        ).map {
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
    FunctionSpec.simple[BigDecimal, CountIfArgs](
      "COUNTIF",
      Arity.two,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (rangeLocation, criteria) = args
      evalValue(ctx, criteria).flatMap { criteriaValue =>
        val criterion = CriteriaMatcher.parse(criteriaValue)
        // GH-192: Resolve target sheet for cross-sheet support
        Evaluator.resolveRangeLocation(rangeLocation, ctx.sheet, ctx.workbook).flatMap {
          case (criteriaSheet, criteriaRange0) =>
            // GH-192: Constrain full-column/row ranges to used area for performance
            val bounds = computeBounds(List((criteriaRange0, criteriaSheet)))
            val constrainedRange = constrainRange(criteriaRange0, bounds)
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
    FunctionSpec.simple[BigDecimal, SumIfsArgs](
      "SUMIFS",
      Arity.AtLeast(3),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (sumRangeLocation, conditions) = args
      evalCriteriaValues(ctx, conditions)
        .flatMap { criteriaValues =>
          val parsedConditions = parseConditions(conditions, criteriaValues)

          // GH-192: Resolve sum range and all criteria ranges to their target sheets FIRST.
          // GH-394: resolution now also yields the ranges (Name locations have no static
          // range), so the dimension validation runs on the resolved shapes.
          for
            resolvedSum <- Evaluator.resolveRangeLocation(
              sumRangeLocation,
              ctx.sheet,
              ctx.workbook
            )
            (sumSheet, sumRange0) = resolvedSum
            resolved <- resolveConditions(parsedConditions, ctx)
            _ <- resolved
              .collectFirst {
                case (_, range, _)
                    if range.width != sumRange0.width || range.height != sumRange0.height =>
                  EvalError.EvalFailed(
                    s"SUMIFS: all ranges must have same dimensions (sum_range is ${sumRange0.height}×${sumRange0.width}, criteria_range is ${range.height}×${range.width})",
                    Some(s"SUMIFS(${sumRangeLocation.toA1}, ...)")
                  )
              }
              .map(Left(_))
              .getOrElse(Right(()))
            result <- {
              val bounds = computeBounds(
                (sumRange0, sumSheet) ::
                  resolved.map { case (sheet, range, _) => (range, sheet) }
              )
              // GH-192: Constrain full-column/row ranges to shared bounds
              val constrainedSumRange = constrainRange(sumRange0, bounds)
              val constrainedConditions = resolved.map { case (sheet, range, criterion) =>
                (sheet, constrainRange(range, bounds), criterion)
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
                      resolveNumericPolicing(
                        sumSheet(sumCells(idx)).value,
                        sumSheet,
                        ctx,
                        "SUMIFS",
                        propagateErrors = true
                      )
                        .map {
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

  /**
   * GH-76: Collect the numeric values from `valueRangeLocation` at positions where ALL criteria
   * match. Mirrors the SUMIFS resolve/constrain/iterate pipeline but accumulates the matched values
   * (rather than summing) so MAXIFS/MINIFS can reduce them. Empty result = no matches.
   */
  private def collectIfsValues(
    valueRangeLocation: TExpr.RangeLocation,
    conditions: RangeCriteriaList,
    fnName: String,
    ctx: EvalContext
  ): Either[EvalError, Vector[BigDecimal]] =
    evalCriteriaValues(ctx, conditions).flatMap { criteriaValues =>
      val parsedConditions = parseConditions(conditions, criteriaValues)
      // GH-394: resolve first (Name locations have no static range), then validate dimensions
      // on the resolved shapes
      for
        resolvedValue <- Evaluator.resolveRangeLocation(valueRangeLocation, ctx.sheet, ctx.workbook)
        (valueSheet, valueRange0) = resolvedValue
        resolved <- resolveConditions(parsedConditions, ctx)
        _ <- resolved
          .collectFirst {
            case (_, range, _)
                if range.width != valueRange0.width || range.height != valueRange0.height =>
              EvalError.EvalFailed(
                s"$fnName: all ranges must have same dimensions (value_range is ${valueRange0.height}×${valueRange0.width}, criteria_range is ${range.height}×${range.width})",
                Some(s"$fnName(${valueRangeLocation.toA1}, ...)")
              )
          }
          .map(Left(_))
          .getOrElse(Right(()))
        result <- {
          val bounds = computeBounds(
            (valueRange0, valueSheet) ::
              resolved.map { case (sheet, range, _) => (range, sheet) }
          )
          val constrainedValueRange = constrainRange(valueRange0, bounds)
          val constrainedConditions = resolved.map { case (sheet, range, criterion) =>
            (sheet, constrainRange(range, bounds), criterion)
          }
          val valueCells = constrainedValueRange.cells.toVector
          val criteriaCells = constrainedConditions.map { case (sheet, range, criterion) =>
            (sheet, range.cells.toVector, criterion)
          }
          valueCells.indices.foldLeft[Either[EvalError, Vector[BigDecimal]]](
            Right(Vector.empty)
          ) {
            case (Left(err), _) => Left(err)
            case (Right(acc), idx) =>
              val matchResult =
                criteriaCells.foldLeft[Either[EvalError, Boolean]](Right(true)) {
                  case (Left(err), _) => Left(err)
                  case (Right(false), _) => Right(false)
                  case (Right(true), (criteriaSheet, cells, criterion)) =>
                    evalCellValueForMatch(
                      criteriaSheet(cells(idx)).value,
                      criteriaSheet,
                      ctx
                    ).map(tv => CriteriaMatcher.matches(tv, criterion))
                }
              matchResult.flatMap { allMatch =>
                if allMatch then
                  resolveNumericPolicing(
                    valueSheet(valueCells(idx)).value,
                    valueSheet,
                    ctx,
                    fnName,
                    propagateErrors = true
                  )
                    .map {
                      case Some(n) => acc :+ n
                      case None => acc
                    }
                else Right(acc)
              }
          }
        }
      yield result
    }

  /** MAXIFS(max_range, criteria_range1, criteria1, ...) — max over matching cells, 0 if none. */
  val maxifs: FunctionSpec[BigDecimal] { type Args = SumIfsArgs } =
    FunctionSpec.simple[BigDecimal, SumIfsArgs](
      "MAXIFS",
      Arity.AtLeast(3),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (valueRange, conditions) = args
      collectIfsValues(valueRange, conditions, "MAXIFS", ctx).map { vs =>
        if vs.isEmpty then BigDecimal(0) else vs.max
      }
    }

  /** MINIFS(min_range, criteria_range1, criteria1, ...) — min over matching cells, 0 if none. */
  val minifs: FunctionSpec[BigDecimal] { type Args = SumIfsArgs } =
    FunctionSpec.simple[BigDecimal, SumIfsArgs](
      "MINIFS",
      Arity.AtLeast(3),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (valueRange, conditions) = args
      collectIfsValues(valueRange, conditions, "MINIFS", ctx).map { vs =>
        if vs.isEmpty then BigDecimal(0) else vs.min
      }
    }

  val countifs: FunctionSpec[BigDecimal] { type Args = CountIfsArgs } =
    FunctionSpec.simple[BigDecimal, CountIfsArgs](
      "COUNTIFS",
      Arity.AtLeast(2),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (conditions, ctx) =>
      evalCriteriaValues(ctx, conditions)
        .flatMap { criteriaValues =>
          val parsedConditions = parseConditions(conditions, criteriaValues)

          parsedConditions match
            case Nil => Right(BigDecimal(0))
            case _ =>
              // GH-192: Resolve all criteria ranges to their target sheets FIRST.
              // GH-394: resolution now also yields the ranges (Name locations have no static
              // range), so the dimension validation runs on the resolved shapes.
              resolveConditions(parsedConditions, ctx)
                .flatMap { resolved =>
                  val dimensionError = resolved.headOption.flatMap { case (_, firstRange, _) =>
                    resolved.collectFirst {
                      case (_, range, _)
                          if range.width != firstRange.width ||
                            range.height != firstRange.height =>
                        EvalError.EvalFailed(
                          s"COUNTIFS: all ranges must have same dimensions (first is ${firstRange.height}×${firstRange.width}, this is ${range.height}×${range.width})",
                          Some(s"COUNTIFS(...)")
                        )
                    }
                  }
                  dimensionError match
                    case Some(err) => Left(err)
                    case None => Right(resolved)
                }
                .flatMap { resolved =>
                  val bounds = computeBounds(resolved.map { case (sheet, range, _) =>
                    (range, sheet)
                  })
                  // GH-192: Constrain full-column/row ranges to shared bounds
                  val constrainedConditions = resolved.map { case (sheet, range, criterion) =>
                    (sheet, constrainRange(range, bounds), criterion)
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
    FunctionSpec.simple[BigDecimal, AverageIfArgs](
      "AVERAGEIF",
      Arity.Range(2, 3),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (rangeLocation, criteria, avgRangeLocationOpt) = args
      evalValue(ctx, criteria).flatMap { criteriaValue =>
        val criterion = CriteriaMatcher.parse(criteriaValue)
        val effectiveLocation = avgRangeLocationOpt.getOrElse(rangeLocation)

        // GH-192: Resolve target sheets for cross-sheet support BEFORE constraining.
        // GH-394: resolution now also yields the ranges (Name locations have no static range),
        // so the dimension validation moved after it — same Excel semantics, resolved shapes.
        for
          resolvedCriteria <- Evaluator.resolveRangeLocation(rangeLocation, ctx.sheet, ctx.workbook)
          resolvedAvg <- Evaluator.resolveRangeLocation(effectiveLocation, ctx.sheet, ctx.workbook)
          (criteriaSheet, criteriaRange0) = resolvedCriteria
          (avgSheet, avgRange0) = resolvedAvg
          _ <-
            if criteriaRange0.width != avgRange0.width ||
              criteriaRange0.height != avgRange0.height
            then
              Left(
                EvalError.EvalFailed(
                  s"AVERAGEIF: range and average_range must have same dimensions (${criteriaRange0.height}×${criteriaRange0.width} vs ${avgRange0.height}×${avgRange0.width})",
                  Some(s"AVERAGEIF(${rangeLocation.toA1}, ..., ${effectiveLocation.toA1})")
                )
              )
            else Right(())
          result <- {
            val bounds = computeBounds(
              List(
                (criteriaRange0, criteriaSheet),
                (avgRange0, avgSheet)
              )
            )
            // GH-192: Constrain full-column/row ranges to shared bounds
            val criteriaRange = constrainRange(criteriaRange0, bounds)
            val avgRange = constrainRange(avgRange0, bounds)

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
                        resolveNumericPolicing(
                          avgSheet(avgRef).value,
                          avgSheet,
                          ctx,
                          "AVERAGEIF",
                          propagateErrors = true
                        ).map {
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
    FunctionSpec.simple[BigDecimal, AverageIfsArgs](
      "AVERAGEIFS",
      Arity.AtLeast(3),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (avgRangeLocation, conditions) = args
      evalCriteriaValues(ctx, conditions)
        .flatMap { criteriaValues =>
          val parsedConditions = parseConditions(conditions, criteriaValues)

          // GH-192: Resolve average range and all criteria ranges to their target sheets FIRST.
          // GH-394: resolution now also yields the ranges (Name locations have no static
          // range), so the dimension validation runs on the resolved shapes.
          for
            resolvedAvg <- Evaluator.resolveRangeLocation(
              avgRangeLocation,
              ctx.sheet,
              ctx.workbook
            )
            (avgSheet, avgRange0) = resolvedAvg
            resolved <- resolveConditions(parsedConditions, ctx)
            _ <- resolved
              .collectFirst {
                case (_, range, _)
                    if range.width != avgRange0.width || range.height != avgRange0.height =>
                  EvalError.EvalFailed(
                    s"AVERAGEIFS: all ranges must have same dimensions (average_range is ${avgRange0.height}×${avgRange0.width}, criteria_range is ${range.height}×${range.width})",
                    Some(s"AVERAGEIFS(${avgRangeLocation.toA1}, ...)")
                  )
              }
              .map(Left(_))
              .getOrElse(Right(()))
            result <- {
              val bounds = computeBounds(
                (avgRange0, avgSheet) ::
                  resolved.map { case (sheet, range, _) => (range, sheet) }
              )
              // GH-192: Constrain full-column/row ranges to shared bounds
              val constrainedAvgRange = constrainRange(avgRange0, bounds)
              val constrainedConditions = resolved.map { case (sheet, range, criterion) =>
                (sheet, constrainRange(range, bounds), criterion)
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
                        resolveNumericPolicing(
                          avgSheet(avgCells(idx)).value,
                          avgSheet,
                          ctx,
                          "AVERAGEIFS",
                          propagateErrors = true
                        )
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
          yield result
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
    range: CellRange
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
      if row < 0 || row >= rows || col < 0 || col >= cols then
        Left(
          EvalError.EvalFailed(
            s"SUMPRODUCT: index out of bounds ($row, $col) for ${rows}x$cols matrix",
            None
          )
        )
      else Right(matrix(row)(col))

  val sumproduct: FunctionSpec[BigDecimal] { type Args = SumProductArgs } =
    FunctionSpec.simple[BigDecimal, SumProductArgs](
      "SUMPRODUCT",
      Arity.atLeastOne,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
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
                Evaluator.resolveRangeLocation(loc, ctx.sheet, ctx.workbook).map {
                  case (sheet, range) => acc :+ (range, sheet)
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
                  Evaluator.resolveRangeLocation(loc, ctx.sheet, ctx.workbook).map {
                    case (sheet, range) => acc :+ RangeArray(sheet, constrainRange(range, bounds))
                  }
                case (Right(acc), Right(expr)) =>
                  // GH-197: Expression - constrain ranges, then evaluate with array support
                  val boundedExpr = constrainExprRanges(expr)
                  ctx.evalArrayExpr(boundedExpr).flatMap {
                    case ar: ArrayResult =>
                      // Convert ArrayResult to numeric matrix with boolean coercion.
                      // GH-337: a carried error element fails loudly naming the code — the
                      // pre-carriage coerceToNumeric would silently turn it into 0 and produce
                      // a wrong sum for the exact issue-cited =SUMPRODUCT((A1:A3<5)*1) case.
                      val matrix = ar.values.flatten
                        .collectFirst(Function.unlift(ArrayArithmetic.carriedError)) match
                        case Some(err) => Left(propagatedElementError("SUMPRODUCT", err))
                        case None =>
                          Right(
                            (0 until ar.rows).map { row =>
                              (0 until ar.cols).map { col =>
                                coerceToNumeric(ar.values(row)(col))
                              }.toVector
                            }.toVector
                          )
                      matrix.map(m => acc :+ MatrixArray(m))
                    case bd: BigDecimal =>
                      // Scalar treated as 1x1 matrix
                      Right(acc :+ MatrixArray(Vector(Vector(bd))))
                    case b: Boolean =>
                      // Boolean coerced to 1/0
                      val n = if b then BigDecimal(1) else BigDecimal(0)
                      Right(acc :+ MatrixArray(Vector(Vector(n))))
                    case cv: CellValue =>
                      // CellValue coerced to numeric; GH-337: carried errors fail loudly
                      ArrayArithmetic.carriedError(cv) match
                        case Some(err) => Left(propagatedElementError("SUMPRODUCT", err))
                        case None => Right(acc :+ MatrixArray(Vector(Vector(coerceToNumeric(cv)))))
                    case other =>
                      Left(EvalError.TypeMismatch("SUMPRODUCT", "array or number", other.toString))
                  }
              }

            resolvedResult.flatMap { resolved =>
              resolved match
                case Nil => Right(BigDecimal(0))
                case first :: rest =>
                  // Validate dimensions match
                  // GH-344 4b: SUMPRODUCT keeps exact-dimension enforcement — Excel raises
                  // #VALUE! here, never the broadcast #N/A padding
                  val dimensionError = rest.collectFirst {
                    case arr if arr.rows != first.rows || arr.cols != first.cols =>
                      EvalError.ErrorValue(
                        CellError.Value,
                        Some(
                          s"SUMPRODUCT: all arrays must have same dimensions (first is ${first.rows}×${first.cols}, got ${arr.rows}×${arr.cols})"
                        )
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
