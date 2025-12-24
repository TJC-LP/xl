package com.tjclp.xl.formula

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue

trait FunctionSpecsLookupIndex extends FunctionSpecsBase:
  private def compareCellValues(cv: CellValue, value: Any): Int =
    (cv, value) match
      case (CellValue.Number(n1), n2: BigDecimal) => n1.compare(n2)
      case (CellValue.Number(n1), n2: Int) => n1.compare(BigDecimal(n2))
      case (CellValue.Number(n1), n2: Long) => n1.compare(BigDecimal(n2))
      case (CellValue.Number(n1), n2: Double) => n1.compare(BigDecimal(n2))
      case (CellValue.Text(s1), s2: String) => s1.compareToIgnoreCase(s2)
      case (CellValue.Bool(b1), b2: Boolean) => b1.compare(b2)
      case (CellValue.Error(e1), CellValue.Error(e2)) => e1.ordinal.compareTo(e2.ordinal)
      case (CellValue.Formula(_, Some(cached)), other) => compareCellValues(cached, other)
      case _ => 0

  private def coerceToBigDecimal(value: Any): BigDecimal =
    value match
      case n: BigDecimal => n
      case i: Int => BigDecimal(i)
      case l: Long => BigDecimal(l)
      case d: Double => BigDecimal(d)
      case s: String => scala.util.Try(BigDecimal(s.trim)).getOrElse(BigDecimal(0))
      case b: Boolean => if b then BigDecimal(1) else BigDecimal(0)
      case CellValue.Number(n) => n
      case CellValue.Text(s) => scala.util.Try(BigDecimal(s.trim)).getOrElse(BigDecimal(0))
      case CellValue.Bool(true) => BigDecimal(1)
      case CellValue.Bool(false) => BigDecimal(0)
      case CellValue.Formula(_, Some(cached)) => coerceToBigDecimal(cached)
      case _ => BigDecimal(0)

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
