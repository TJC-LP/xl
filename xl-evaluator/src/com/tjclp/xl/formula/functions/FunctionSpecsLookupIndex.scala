package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{ArrayArithmetic, ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue

trait FunctionSpecsLookupIndex extends FunctionSpecsBase:
  private def coerceToBigDecimal(value: ExprValue): BigDecimal =
    value match
      case ExprValue.Number(n) => n
      case ExprValue.Text(s) => scala.util.Try(BigDecimal(s.trim)).getOrElse(BigDecimal(0))
      case ExprValue.Bool(b) => ArrayArithmetic.boolToNumeric(b)
      case ExprValue.Cell(CellValue.Number(n)) => n
      case ExprValue.Cell(CellValue.Text(s)) =>
        scala.util.Try(BigDecimal(s.trim)).getOrElse(BigDecimal(0))
      case ExprValue.Cell(CellValue.Bool(b)) => ArrayArithmetic.boolToNumeric(b)
      case ExprValue.Cell(CellValue.Formula(_, Some(cached), _)) =>
        coerceToBigDecimal(ExprValue.Cell(cached))
      case _ => BigDecimal(0)

  /**
   * INDEX(array, row_num, [column_num])
   *
   * GH-654: a reference-returning function, typed `ArrayResult` like OFFSET and INDIRECT. The
   * selected cell is a 1×1 result that collapses in scalar positions; Excel's 0 — or the omitted
   * slot, `INDEX(rng,,2)` (GH-603) — selects the whole row or column (`SUM(INDEX(A1:B2,,2))` is
   * B1+B2, `INDEX(rng,0,0)` the whole array), spilling standalone and folding under aggregates. The
   * two-argument form on a one-row array reads its argument as column_num, on a one-column array as
   * row_num, and on a 2-D array as row_num with the whole row selected (collapsing to that row's
   * first cell in a scalar position — the value it always returned there). A whole-axis selection
   * is bounded to the sheet's used range (as INDIRECT bounds "A:A"), so its cost follows the data
   * rather than the reference; a position outside the array is a descriptive `#REF!`.
   */
  val index: FunctionSpec[ArrayResult] { type Args = IndexArgs } =
    FunctionSpec.simple[ArrayResult, IndexArgs]("INDEX", Arity.Range(2, 3)) { (args, ctx) =>
      val (array, rowNumExpr, colNumOpt) = args
      for
        rowNum <- ctx.evalExpr(rowNumExpr)
        colNum <- colNumOpt match
          case Some(expr) => ctx.evalExpr(expr).map(Some(_))
          case None => Right(None)
        resolved <- Evaluator.resolveRangeLocation(array, ctx.sheet, ctx.workbook)
        (targetSheet, arrayRange) = resolved
        result <- {
          val startCol = arrayRange.colStart.index0
          val startRow = arrayRange.rowStart.index0
          val numCols = arrayRange.width
          val numRows = arrayRange.height
          val call = s"INDEX(${array.toA1}, $rowNum${colNum.map(c => s", $c").getOrElse("")})"

          // Excel's two-argument form: the single position reads along a vector's long axis; on
          // a 2-D array it is row_num and the whole row is selected (column_num 0)
          val (rowPos, colPos): (Int, Int) = colNum match
            case Some(c) => (rowNum.toInt, c.toInt)
            case None if numRows == 1 => (1, rowNum.toInt)
            case None => (rowNum.toInt, 0)

          // None = the whole axis (Excel's 0), Some(i) = one 0-based line of it
          def axis(
            pos: Int,
            size: Int,
            label: String,
            noun: String
          ): Either[EvalError, Option[Int]] =
            if pos == 0 then Right(None)
            else if pos < 0 || pos > size then
              Left(
                EvalError.EvalFailed(
                  s"INDEX: $label $pos is out of bounds (array has $size $noun, valid range: 1-$size) (#REF!)",
                  Some(call)
                )
              )
            else Right(Some(pos - 1))

          // the 0-based inclusive span an axis selection covers; a whole axis is bounded to the
          // used range (None when the two do not meet), so `INDEX($A$1:$A$100000,0)` costs the
          // data, not the reference — cells past the used range are blank either way
          def span(
            sel: Option[Int],
            start: Int,
            size: Int,
            used: Option[(Int, Int)]
          ): Option[(Int, Int)] =
            sel match
              case Some(i) => Some((start + i, start + i))
              case None =>
                used
                  .map { case (lo, hi) => (math.max(lo, start), math.min(hi, start + size - 1)) }
                  .filter { case (lo, hi) => lo <= hi }

          for
            rowSel <- axis(rowPos, numRows, "row_num", "rows")
            colSel <- axis(colPos, numCols, "col_num", "columns")
            values <-
              val used = targetSheet.usedRange
              val rowSpan =
                span(rowSel, startRow, numRows, used.map(u => (u.rowStart.index0, u.rowEnd.index0)))
              val colSpan =
                span(colSel, startCol, numCols, used.map(u => (u.colStart.index0, u.colEnd.index0)))
              (rowSpan, colSpan) match
                case (Some((r0, r1)), Some((c0, c1))) =>
                  val selected = CellRange(ARef.from0(c0, r0), ARef.from0(c1, r1))
                  extractRangeAsMatrixEval(selected, targetSheet, ctx).map(ArrayResult(_))
                case _ => Right(ArrayResult.empty)
          yield values
        }
      yield result
    }

  val matchFn: FunctionSpec[BigDecimal] { type Args = MatchArgs } =
    FunctionSpec.simple[BigDecimal, MatchArgs](
      "MATCH",
      Arity.Range(2, 3),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (args, ctx) =>
      val (lookupValue, lookupArray, matchTypeOpt) = args
      val matchTypeExpr = matchTypeOpt.getOrElse(TExpr.Lit(BigDecimal(1)))
      for
        lookupValueEval <- evalValue(ctx, lookupValue)
        matchType <- ctx.evalExpr(matchTypeExpr)
        resolved <- Evaluator.resolveRangeLocation(lookupArray, ctx.sheet, ctx.workbook)
        (targetSheet, lookupRange) = resolved
        result <- {
          val matchTypeInt = matchType.toInt
          // GH-467: dereference cell-ref lookup values and coerce dates to serials up front —
          // positions below are RANGE positions (idx + 1), blanks included, never compacted.
          val lookup = normalizeLookupValue(lookupValueEval)
          val cells: List[(Int, CellValue)] =
            lookupRange.cells.toList.zipWithIndex.map { case (ref, idx) =>
              (idx + 1, targetSheet(ref).value)
            }

          val positionOpt: Option[Int] = matchTypeInt match
            case 0 =>
              cells
                .find { case (_, cv) =>
                  matchesLookupExact(cv, lookup)
                }
                .map(_._1)
            case 1 =>
              val numericLookup = coerceToBigDecimal(lookup)
              // GH-467: extractNumericValue skips blanks/text/bools (they are NOT zero) and
              // yields serials for DateTime cells, so date columns work in approximate mode.
              val candidates = cells.flatMap { case (pos, cv) =>
                extractNumericValue(cv).filter(_ <= numericLookup).map(n => (pos, n))
              }
              candidates.maxByOption(_._2).map(_._1)
            case -1 =>
              val numericLookup = coerceToBigDecimal(lookup)
              val candidates = cells.flatMap { case (pos, cv) =>
                extractNumericValue(cv).filter(_ >= numericLookup).map(n => (pos, n))
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
