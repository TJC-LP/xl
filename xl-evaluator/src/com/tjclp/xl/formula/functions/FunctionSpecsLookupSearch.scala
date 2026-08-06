package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator, CriteriaMatcher}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet

trait FunctionSpecsLookupSearch extends FunctionSpecsBase:
  private def performXLookup(
    lookupValue: ExprValue,
    lookupSheet: Sheet,
    lookupArray: CellRange,
    returnSheet: Sheet,
    returnArray: CellRange,
    ifNotFoundOpt: Option[AnyExpr],
    matchMode: Int,
    searchMode: Int,
    ctx: EvalContext
  ): Either[EvalError, CellValue] =
    val lookupCells = lookupArray.cells.toVector
    val returnCells = returnArray.cells.toVector

    // GH-467: dereference cell-ref lookup values and coerce dates to serials before comparing —
    // a ref to a cell holding "k2" must match exactly like the literal "k2".
    val lookup = normalizeLookupValue(lookupValue)

    // GH-55: accept binary-search modes 2 (ascending) and -2 (descending). Linear iteration in the
    // correct direction yields correct results; -2 iterates reversed like -1 (descending order).
    val indices =
      if searchMode == -1 || searchMode == -2 then lookupCells.indices.reverse
      else lookupCells.indices

    val wildcardCriterionOpt = lookup match
      case ExprValue.Text(text) =>
        CriteriaMatcher.parse(ExprValue.Text(text)) match
          case c: CriteriaMatcher.Wildcard => Some(c)
          case _ => None
      case _ => None

    val matchedIndexOpt = matchMode match
      case 0 =>
        indices.find { idx =>
          val cellValue = lookupSheet(lookupCells(idx)).value
          matchesLookupExact(cellValue, lookup)
        }
      case -1 =>
        findNextSmaller(lookup, lookupCells, lookupSheet, indices)
      case 1 =>
        findNextLarger(lookup, lookupCells, lookupSheet, indices)
      case 2 =>
        indices.find { idx =>
          val cellValue = lookupSheet(lookupCells(idx)).value
          matchesLookupExact(cellValue, lookup) ||
          wildcardCriterionOpt.exists(CriteriaMatcher.matches(cellValue, _))
        }
      case _ => None

    matchedIndexOpt match
      case Some(idx) => Right(returnSheet(returnCells(idx)).value)
      case None =>
        ifNotFoundOpt match
          case Some(expr) => evalValue(ctx, expr).map(toCellValue)
          case None => Right(CellValue.Error(CellError.NA))

  private def findNextSmaller(
    lookupValue: ExprValue,
    lookupCells: Vector[ARef],
    sheet: Sheet,
    indices: IndexedSeq[Int]
  ): Option[Int] =
    lookupValue match
      case ExprValue.Number(targetNum) =>
        val candidates = indices
          .flatMap { idx =>
            extractNumericValue(sheet(lookupCells(idx)).value).map(n => (idx, n))
          }
          .filter(_._2 <= targetNum)
        candidates.sortBy(_._2).lastOption.map(_._1)
      case _ => None

  private def findNextLarger(
    lookupValue: ExprValue,
    lookupCells: Vector[ARef],
    sheet: Sheet,
    indices: IndexedSeq[Int]
  ): Option[Int] =
    lookupValue match
      case ExprValue.Number(targetNum) =>
        val candidates = indices
          .flatMap { idx =>
            extractNumericValue(sheet(lookupCells(idx)).value).map(n => (idx, n))
          }
          .filter(_._2 >= targetNum)
        candidates.sortBy(_._2).headOption.map(_._1)
      case _ => None

  val vlookup: FunctionSpec[CellValue] { type Args = VlookupArgs } =
    FunctionSpec.simple[CellValue, VlookupArgs]("VLOOKUP", Arity.Range(3, 4)) { (args, ctx) =>
      val (lookupExpr, table, colIndexExpr, rangeLookupOpt) = args
      val rangeLookupExpr = rangeLookupOpt.getOrElse(TExpr.Lit(true))
      for
        lookupValue <- evalValue(ctx, lookupExpr)
        colIndex <- ctx.evalExpr(colIndexExpr)
        rangeMatch <- ctx.evalExpr(rangeLookupExpr)
        resolved <- Evaluator.resolveRangeLocation(table, ctx.sheet, ctx.workbook)
        (targetSheet, tableRange) = resolved
        result <-
          if colIndex < 1 || colIndex > tableRange.width then
            Left(
              EvalError.EvalFailed(
                s"VLOOKUP: col_index_num $colIndex is outside 1..${tableRange.width}",
                Some(s"VLOOKUP(…, ${table.toA1})")
              )
            )
          else
            val rowIndices = 0 until tableRange.height
            val keyCol0 = tableRange.colStart.index0
            val rowStart0 = tableRange.rowStart.index0
            val resultCol0 = keyCol0 + (colIndex - 1)

            def renderValue(value: ExprValue): String = value match
              case ExprValue.Text(s) => s
              case ExprValue.Number(n) => n.toString
              case ExprValue.Bool(b) => b.toString
              case ExprValue.Date(d) => d.toString
              case ExprValue.DateTime(dt) => dt.toString
              case ExprValue.Cell(cv) => cv.toString
              case ExprValue.Opaque(other) => other.toString

            // GH-488: one lookup plane for MATCH/XLOOKUP/VLOOKUP/HLOOKUP — the inline copy this
            // replaces lacked the DateTime→serial case, so a date key over a date column missed.
            val normalizedLookup: ExprValue = normalizeLookupValue(lookupValue)

            val isTextLookup = normalizedLookup match
              case ExprValue.Text(_) => true
              case ExprValue.Number(_) => false
              case ExprValue.Bool(_) => false
              case _ => true

            val chosenRowOpt: Option[Int] =
              if rangeMatch then
                val numericLookup: Option[BigDecimal] = normalizedLookup match
                  case ExprValue.Number(n) => Some(n)
                  case ExprValue.Text(s) => scala.util.Try(BigDecimal(s.trim)).toOption
                  case _ => None

                numericLookup.flatMap { lookup =>
                  val keyedRows: List[(Int, BigDecimal)] =
                    rowIndices.toList.flatMap { i =>
                      val keyRef = ARef.from0(keyCol0, rowStart0 + i)
                      extractNumericForMatch(targetSheet(keyRef).value).map(k => (i, k))
                    }
                  keyedRows
                    .filter(_._2 <= lookup)
                    .sortBy(_._2)
                    .lastOption
                    .map(_._1)
                }
              else if isTextLookup then
                val lookupText = renderValue(normalizedLookup).toLowerCase
                rowIndices.find { i =>
                  val keyRef = ARef.from0(keyCol0, rowStart0 + i)
                  extractTextForMatch(targetSheet(keyRef).value)
                    .exists(_.toLowerCase == lookupText)
                }
              else
                val numericLookup: Option[BigDecimal] = normalizedLookup match
                  case ExprValue.Number(n) => Some(n)
                  case _ => None
                numericLookup.flatMap { lookup =>
                  rowIndices.find { i =>
                    val keyRef = ARef.from0(keyCol0, rowStart0 + i)
                    extractNumericForMatch(targetSheet(keyRef).value).contains(lookup)
                  }
                }

            chosenRowOpt match
              case Some(rowIndex) =>
                val resultRef = ARef.from0(resultCol0, rowStart0 + rowIndex)
                Right(targetSheet(resultRef).value)
              case None =>
                Left(
                  EvalError.EvalFailed(
                    if rangeMatch then "VLOOKUP approximate match not found"
                    else "VLOOKUP exact match not found",
                    Some(
                      s"VLOOKUP(${renderValue(normalizedLookup)}, ${table.toA1}, $colIndex, $rangeMatch)"
                    )
                  )
                )
      yield result
    }

  /** Render an ExprValue to text for HLOOKUP text matching / diagnostics. */
  private def exprValueToText(value: ExprValue): String =
    value match
      case ExprValue.Text(s) => s
      case ExprValue.Number(n) => n.toString
      case ExprValue.Bool(b) => b.toString
      case ExprValue.Date(d) => d.toString
      case ExprValue.DateTime(dt) => dt.toString
      case ExprValue.Cell(cv) => cv.toString
      case ExprValue.Opaque(other) => other.toString

  /**
   * HLOOKUP(lookup, table, row_index_num, [range_lookup])
   *
   * Horizontal transpose of VLOOKUP: searches the first ROW of the table, returns the cell from the
   * 1-based row_index_num in the matched COLUMN. range_lookup TRUE (default) = approximate.
   */
  val hlookup: FunctionSpec[CellValue] { type Args = VlookupArgs } =
    FunctionSpec.simple[CellValue, VlookupArgs]("HLOOKUP", Arity.Range(3, 4)) { (args, ctx) =>
      val (lookupExpr, table, rowIndexExpr, rangeLookupOpt) = args
      val rangeLookupExpr = rangeLookupOpt.getOrElse(TExpr.Lit(true))
      for
        lookupValue <- evalValue(ctx, lookupExpr)
        rowIndex <- ctx.evalExpr(rowIndexExpr)
        rangeMatch <- ctx.evalExpr(rangeLookupExpr)
        resolved <- Evaluator.resolveRangeLocation(table, ctx.sheet, ctx.workbook)
        (targetSheet, tableRange) = resolved
        result <-
          if rowIndex < 1 || rowIndex > tableRange.height then
            Left(
              EvalError.EvalFailed(
                s"HLOOKUP: row_index_num $rowIndex is outside 1..${tableRange.height}",
                Some(s"HLOOKUP(…, ${table.toA1})")
              )
            )
          else
            val colIndices = 0 until tableRange.width
            val keyRow0 = tableRange.rowStart.index0
            val colStart0 = tableRange.colStart.index0
            val resultRow0 = keyRow0 + (rowIndex - 1)

            // GH-488: shared lookup plane (see VLOOKUP above)
            val normalizedLookup: ExprValue = normalizeLookupValue(lookupValue)

            val isTextLookup = normalizedLookup match
              case ExprValue.Text(_) => true
              case ExprValue.Number(_) => false
              case ExprValue.Bool(_) => false
              case _ => true

            val chosenColOpt: Option[Int] =
              if rangeMatch then
                val numericLookup: Option[BigDecimal] = normalizedLookup match
                  case ExprValue.Number(n) => Some(n)
                  case ExprValue.Text(s) => scala.util.Try(BigDecimal(s.trim)).toOption
                  case _ => None
                numericLookup.flatMap { lookup =>
                  val keyedCols: List[(Int, BigDecimal)] =
                    colIndices.toList.flatMap { i =>
                      val keyRef = ARef.from0(colStart0 + i, keyRow0)
                      extractNumericForMatch(targetSheet(keyRef).value).map(k => (i, k))
                    }
                  keyedCols.filter(_._2 <= lookup).sortBy(_._2).lastOption.map(_._1)
                }
              else if isTextLookup then
                val lookupText = exprValueToText(normalizedLookup).toLowerCase
                colIndices.find { i =>
                  val keyRef = ARef.from0(colStart0 + i, keyRow0)
                  extractTextForMatch(targetSheet(keyRef).value).exists(_.toLowerCase == lookupText)
                }
              else
                val numericLookup: Option[BigDecimal] = normalizedLookup match
                  case ExprValue.Number(n) => Some(n)
                  case _ => None
                numericLookup.flatMap { lookup =>
                  colIndices.find { i =>
                    val keyRef = ARef.from0(colStart0 + i, keyRow0)
                    extractNumericForMatch(targetSheet(keyRef).value).contains(lookup)
                  }
                }

            chosenColOpt match
              case Some(colIdx) =>
                Right(targetSheet(ARef.from0(colStart0 + colIdx, resultRow0)).value)
              case None =>
                Left(
                  EvalError.EvalFailed(
                    if rangeMatch then "HLOOKUP approximate match not found"
                    else "HLOOKUP exact match not found",
                    Some(
                      s"HLOOKUP(${exprValueToText(normalizedLookup)}, ${table.toA1}, $rowIndex, $rangeMatch)"
                    )
                  )
                )
      yield result
    }

  val xlookup: FunctionSpec[CellValue] { type Args = XLookupArgs } =
    FunctionSpec.simple[CellValue, XLookupArgs]("XLOOKUP", Arity.Range(3, 6)) { (args, ctx) =>
      val (lookupValue, lookupLoc, returnLoc, ifNotFoundOpt, matchModeOpt, searchModeOpt) =
        args
      val matchModeExpr = matchModeOpt.getOrElse(TExpr.Lit(0))
      val searchModeExpr = searchModeOpt.getOrElse(TExpr.Lit(1))
      // GH-394: resolve locations first (Name locations have no static range), then validate
      // dimensions on the resolved shapes
      for
        resolvedLookup <- Evaluator.resolveRangeLocation(lookupLoc, ctx.sheet, ctx.workbook)
        resolvedReturn <- Evaluator.resolveRangeLocation(returnLoc, ctx.sheet, ctx.workbook)
        (lookupSheet, lookupArray) = resolvedLookup
        (returnSheet, returnArray) = resolvedReturn
        _ <-
          if lookupArray.width != returnArray.width || lookupArray.height != returnArray.height
          then
            Left(
              EvalError.EvalFailed(
                s"XLOOKUP: lookup_array and return_array must have same dimensions (${lookupArray.height}×${lookupArray.width} vs ${returnArray.height}×${returnArray.width})",
                Some(s"XLOOKUP(..., ${lookupLoc.toA1}, ${returnLoc.toA1}, ...)")
              )
            )
          else Right(())
        lookupValueEval <- evalValue(ctx, lookupValue)
        matchModeRaw <- evalValue(ctx, matchModeExpr)
        searchModeRaw <- evalValue(ctx, searchModeExpr)
        matchMode <- toIntArg("XLOOKUP", matchModeRaw)
        searchMode <- toIntArg("XLOOKUP", searchModeRaw)
        result <- performXLookup(
          lookupValueEval,
          lookupSheet,
          lookupArray,
          returnSheet,
          returnArray,
          ifNotFoundOpt,
          matchMode,
          searchMode,
          ctx
        )
      yield result
    }
