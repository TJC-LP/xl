package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.Arity

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
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
