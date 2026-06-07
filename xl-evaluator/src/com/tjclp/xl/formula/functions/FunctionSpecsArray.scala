package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.Arity

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{CellValue, CellError}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

/**
 * Array function specifications (TRANSPOSE, etc.).
 *
 * These functions return ArrayResult which can be "spilled" into adjacent cells when evaluated as
 * array formulas.
 */
trait FunctionSpecsArray extends FunctionSpecsBase:

  /**
   * TRANSPOSE(array)
   *
   * Transposes a range, swapping rows and columns.
   *
   * Example:
   *   - TRANSPOSE(A1:C2) with 2 rows × 3 cols returns 3 rows × 2 cols
   *   - Input: [[1,2,3],[4,5,6]] → Output: [[1,4],[2,5],[3,6]]
   */
  val transpose: FunctionSpec[ArrayResult] { type Args = UnaryRange } =
    FunctionSpec.simple[ArrayResult, UnaryRange]("TRANSPOSE", Arity.one) { (location, ctx) =>
      Evaluator.resolveRangeLocation(location, ctx.sheet, ctx.workbook) match
        case Left(err) => Left(err)
        case Right(targetSheet) =>
          val range = location.range
          val values = extractRangeAsMatrix(range, targetSheet)
          Right(ArrayResult(values.transpose))
    }

  /**
   * OFFSET(reference, rows, cols, [height], [width])
   *
   * Returns the range `rows`/`cols` away from the anchor reference, sized height×width (both
   * default to 1). Returned as an ArrayResult, so it spills standalone, collapses to a scalar when
   * 1×1, and composes with aggregates (e.g. SUM(OFFSET(...))). Out-of-bounds or non-positive size
   * yields #REF!. Note: a cross-sheet anchor's sheet is not tracked (same-sheet result).
   */
  val offset: FunctionSpec[ArrayResult] { type Args = OffsetArgs } =
    FunctionSpec.simple[ArrayResult, OffsetArgs]("OFFSET", Arity.Range(3, 5)) { (args, ctx) =>
      val (refExpr, rowsExpr, colsExpr, hOpt, wOpt) = args
      extractARef(refExpr) match
        case None =>
          Left(EvalError.EvalFailed("OFFSET requires a cell reference", Some("OFFSET(...)")))
        case Some(anchor) =>
          for
            dRows <- ctx.evalExpr(rowsExpr)
            dCols <- ctx.evalExpr(colsExpr)
            height <- hOpt.map(e => ctx.evalExpr(e)).getOrElse(Right(1))
            width <- wOpt.map(e => ctx.evalExpr(e)).getOrElse(Right(1))
            result <-
              val r0 = anchor.row.index0 + dRows
              val c0 = anchor.col.index0 + dCols
              if height < 1 || width < 1 ||
                r0 < 0 || c0 < 0 ||
                r0 + height - 1 > Row.MaxIndex0 || c0 + width - 1 > Column.MaxIndex0
              then Right(ArrayResult.single(CellValue.Error(CellError.Ref)))
              else
                val range =
                  CellRange(ARef.from0(c0, r0), ARef.from0(c0 + width - 1, r0 + height - 1))
                Right(ArrayResult(extractRangeAsMatrix(range, ctx.sheet)))
          yield result
    }

  /**
   * SEQUENCE(rows, [cols], [start], [step])
   *
   * Generates a row-major grid of sequential numbers. Defaults: cols=1, start=1, step=1.
   */
  val sequence: FunctionSpec[ArrayResult] { type Args = SequenceArgs } =
    FunctionSpec.simple[ArrayResult, SequenceArgs]("SEQUENCE", Arity.Range(1, 4)) { (args, ctx) =>
      val (rowsExpr, colsOpt, startOpt, stepOpt) = args
      for
        nRows <- ctx.evalExpr(rowsExpr)
        nCols <- colsOpt.map(e => ctx.evalExpr(e)).getOrElse(Right(1))
        start <- startOpt.map(e => ctx.evalExpr(e)).getOrElse(Right(BigDecimal(1)))
        step <- stepOpt.map(e => ctx.evalExpr(e)).getOrElse(Right(BigDecimal(1)))
        result <-
          if nRows < 1 || nCols < 1 then
            Left(EvalError.EvalFailed("SEQUENCE: rows and columns must be >= 1", None))
          else if nRows.toLong * nCols.toLong > 1048576L then
            Left(EvalError.EvalFailed("SEQUENCE: result exceeds 1,048,576 cells", None))
          else
            val values = (0 until nRows).toVector.map { r =>
              (0 until nCols).toVector.map { c =>
                (CellValue.Number(start + step * BigDecimal(r * nCols + c)): CellValue)
              }
            }
            Right(ArrayResult(values))
      yield result
    }

  /**
   * SORT(array, [sort_index], [sort_order])
   *
   * Sorts the rows of a range by a 1-based column index (default 1). sort_order 1 = ascending
   * (default), -1 = descending. Sort key: numbers < text (case-insensitive) < booleans.
   */
  val sortFn: FunctionSpec[ArrayResult] { type Args = SortArgs } =
    FunctionSpec.simple[ArrayResult, SortArgs]("SORT", Arity.Range(1, 3)) { (args, ctx) =>
      val (location, idxOpt, orderOpt) = args
      for
        sortIndex <- idxOpt.map(e => ctx.evalExpr(e)).getOrElse(Right(1))
        sortOrder <- orderOpt.map(e => ctx.evalExpr(e)).getOrElse(Right(1))
        targetSheet <- Evaluator.resolveRangeLocation(location, ctx.sheet, ctx.workbook)
        result <-
          val matrix = extractRangeAsMatrix(location.range, targetSheet)
          val width = matrix.headOption.map(_.size).getOrElse(0)
          val colIdx = sortIndex - 1
          if matrix.isEmpty then Right(ArrayResult(matrix))
          else if colIdx < 0 || colIdx >= width then
            Left(EvalError.EvalFailed(s"SORT: sort_index $sortIndex is outside 1..$width", None))
          else
            val asc = matrix.sortBy(row => cellSortKey(row(colIdx)))
            Right(ArrayResult(if sortOrder < 0 then asc.reverse else asc))
      yield result
    }

  /**
   * UNIQUE(array, [by_col], [exactly_once])
   *
   * Returns distinct rows (or columns when by_col=TRUE), preserving first-seen order. When
   * exactly_once=TRUE, returns only the rows that occur exactly once. Text comparison is
   * case-insensitive (Excel parity).
   */
  val unique: FunctionSpec[ArrayResult] { type Args = UniqueArgs } =
    FunctionSpec.simple[ArrayResult, UniqueArgs]("UNIQUE", Arity.Range(1, 3)) { (args, ctx) =>
      val (location, byColOpt, onceOpt) = args
      for
        byCol <- byColOpt.map(e => ctx.evalExpr(e)).getOrElse(Right(false))
        exactlyOnce <- onceOpt.map(e => ctx.evalExpr(e)).getOrElse(Right(false))
        targetSheet <- Evaluator.resolveRangeLocation(location, ctx.sheet, ctx.workbook)
      yield
        val matrix0 = extractRangeAsMatrix(location.range, targetSheet)
        val matrix = if byCol then matrix0.transpose else matrix0
        val resultRows =
          if exactlyOnce then
            val keys = matrix.map(_.map(cellKey))
            val counts = keys.groupBy(identity).view.mapValues(_.size).toMap
            matrix.zip(keys).collect { case (row, k) if counts.getOrElse(k, 0) == 1 => row }
          else
            matrix
              .foldLeft((Vector.empty[Vector[CellValue]], Set.empty[Vector[String]])) {
                case ((acc, seen), row) =>
                  val k = row.map(cellKey)
                  if seen.contains(k) then (acc, seen) else (acc :+ row, seen + k)
              }
              ._1
        ArrayResult(if byCol then resultRows.transpose else resultRows)
    }

  /**
   * FILTER(array, include, [if_empty])
   *
   * Keeps the rows of `array` whose corresponding entry in the single-column `include` range is
   * truthy (TRUE or a non-zero number). Returns `if_empty` (or #N/A) when nothing matches.
   */
  val filterFn: FunctionSpec[ArrayResult] { type Args = FilterArgs } =
    FunctionSpec.simple[ArrayResult, FilterArgs]("FILTER", Arity.Range(2, 3)) { (args, ctx) =>
      val (arrayLoc, includeLoc, ifEmptyOpt) = args
      for
        arraySheet <- Evaluator.resolveRangeLocation(arrayLoc, ctx.sheet, ctx.workbook)
        includeSheet <- Evaluator.resolveRangeLocation(includeLoc, ctx.sheet, ctx.workbook)
        result <-
          val matrix = extractRangeAsMatrix(arrayLoc.range, arraySheet)
          val flags = extractRangeAsMatrix(includeLoc.range, includeSheet)
            .map(row => row.headOption.exists(isTruthy))
          val kept = matrix.zip(flags).collect { case (row, true) => row }
          if kept.nonEmpty then Right(ArrayResult(kept))
          else
            ifEmptyOpt match
              case Some(expr) => evalValue(ctx, expr).map(v => ArrayResult.single(toCellValue(v)))
              case None => Right(ArrayResult.single(CellValue.Error(CellError.NA)))
      yield result
    }

  /** Sort key: numbers (by value) < text (case-insensitive) < booleans < everything else. */
  private def cellSortKey(cv: CellValue): (Int, BigDecimal, String) =
    cv match
      case CellValue.Number(n) => (0, n, "")
      case CellValue.Text(s) => (1, BigDecimal(0), s.toLowerCase)
      case CellValue.Bool(b) => (2, if b then BigDecimal(1) else BigDecimal(0), "")
      case CellValue.Formula(_, Some(c)) => cellSortKey(c)
      case _ => (3, BigDecimal(0), "")

  /** Canonical equality key for UNIQUE (case-insensitive text, normalized numbers). */
  private def cellKey(cv: CellValue): String =
    cv match
      case CellValue.Number(n) => "n:" + n.bigDecimal.stripTrailingZeros.toPlainString
      case CellValue.Text(s) => "t:" + s.toLowerCase
      case CellValue.Bool(b) => "b:" + b
      case CellValue.Empty => "e:"
      case CellValue.Formula(_, Some(c)) => cellKey(c)
      case other => "o:" + other.toString

  /** Truthiness for FILTER include flags. */
  private def isTruthy(cv: CellValue): Boolean =
    cv match
      case CellValue.Bool(b) => b
      case CellValue.Number(n) => n.signum != 0
      case CellValue.Formula(_, Some(c)) => isTruthy(c)
      case _ => false

  /**
   * Extract a range from sheet as a 2D matrix of CellValues.
   *
   * @param range
   *   The cell range to extract
   * @param sheet
   *   The sheet to read from
   * @return
   *   Row-major Vector[Vector[CellValue]]
   */
  private def extractRangeAsMatrix(range: CellRange, sheet: Sheet): Vector[Vector[CellValue]] =
    (range.rowStart.index0 to range.rowEnd.index0).map { rowIdx =>
      (range.colStart.index0 to range.colEnd.index0).map { colIdx =>
        val ref = ARef.from0(colIdx, rowIdx)
        val cell = sheet(ref)
        // For formulas with cached values, use the cached value
        cell.value match
          case CellValue.Formula(_, Some(cachedValue)) => cachedValue
          case other => other
      }.toVector
    }.toVector
