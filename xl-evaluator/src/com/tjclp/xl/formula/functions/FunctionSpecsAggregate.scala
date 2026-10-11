package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{
  EvalError,
  Evaluator,
  CriteriaMatcher,
  Aggregator,
  ArrayResult,
  ArrayArithmetic,
  RangeOperand,
  ScalarCoercion
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
    targetSheet: com.tjclp.xl.sheets.Sheet,
    at: ARef,
    ctx: EvalContext,
    fnName: String,
    propagateErrors: Boolean
  ): Either[EvalError, Option[BigDecimal]] =
    evalCellValueForMatch(targetSheet, at, ctx).flatMap { resolved =>
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
    targetSheet: com.tjclp.xl.sheets.Sheet,
    at: ARef,
    ctx: EvalContext
  ): Either[EvalError, BigDecimal] =
    evalCellValueForMatch(targetSheet, at, ctx).flatMap { resolved =>
      ArrayArithmetic.carriedError(resolved) match
        case Some(err) => Left(propagatedElementError("SUMPRODUCT", err))
        case None => Right(coerceToNumeric(resolved))
    }

  // GH-187: Helper to evaluate cell value for criteria matching.
  // Evaluates uncached formulas before matching, returning the resolved CellValue.
  private def evalCellValueForMatch(
    targetSheet: com.tjclp.xl.sheets.Sheet,
    at: ARef,
    ctx: EvalContext
  ): Either[EvalError, CellValue] =
    targetSheet(at).value match
      case CellValue.Formula(formulaStr, None, kind) =>
        // Recursively evaluate uncached formula (GH-346: memoized once per pass), at its own
        // position and in its own mode, exactly as a direct reference to the cell evaluates it
        Evaluator
          .evalCrossSheetFormula(
            formulaStr,
            targetSheet,
            ctx.clock,
            ctx.workbook,
            ctx.depth + 1,
            ctx.rng,
            ctx.memo.getOrElse(new Evaluator.EvalMemo),
            ctx.workbookPath,
            ctx.aggregateMemo,
            currentCell = Some(at),
            arrayRecord = Evaluator.isArrayRecord(kind)
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
    targetSheet: com.tjclp.xl.sheets.Sheet,
    at: ARef,
    ctx: EvalContext,
    acc: A
  ): Either[EvalError, A] =
    val cellValue = targetSheet(at).value
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
        targetSheet,
        at,
        ctx,
        agg.name,
        agg.propagatesErrors
      )
        .map {
          case Some(n) => agg.combine(acc, n)
          case None => acc
        }

  /**
   * GH-711: the cells of a whole-column or whole-row `range` that the used-area constraint dropped
   * (`constrained` is `range` after [[constrainRange]]). They lie past every non-empty cell, so all
   * are blank: COUNTBLANK counts them, as Excel does (`COUNTBLANK(B:B)` with B1:B3 filled is
   * 1048573), and every other aggregate skips them. A whole sheet holds 2^34 cells, past `Int`, so
   * the count rides beside the accumulator and joins the finalized result.
   */
  private def blanksTrimmed(
    agg: Aggregator[?],
    range: CellRange,
    constrained: CellRange
  ): BigDecimal =
    if agg.countsEmpty && (range.isFullColumn || range.isFullRow) then
      BigDecimal(range.cellCount - constrained.cellCount)
    else BigDecimal(0)

  /**
   * Fold one reference, a whole column or row within the sheet's used area, into the accumulator;
   * the second component is [[blanksTrimmed]].
   */
  private def foldReference[A](
    agg: Aggregator[A],
    targetSheet: com.tjclp.xl.sheets.Sheet,
    range: CellRange,
    ctx: EvalContext,
    acc: A
  ): Either[EvalError, (A, BigDecimal)] =
    val constrained = constrainRange(range, computeBounds(List((range, targetSheet))))
    foldRawRange(agg, targetSheet, constrained, ctx, acc)
      .map((_, blanksTrimmed(agg, range, constrained)))
