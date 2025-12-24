package com.tjclp.xl.formula

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet

trait FunctionSpecsLookupSearch extends FunctionSpecsBase:
  private def performXLookup(
    lookupValue: Any,
    lookupArray: CellRange,
    returnArray: CellRange,
    ifNotFoundOpt: Option[AnyExpr],
    matchMode: Int,
    searchMode: Int,
    ctx: EvalContext
  ): Either[EvalError, CellValue] =
    val lookupCells = lookupArray.cells.toVector
    val returnCells = returnArray.cells.toVector

    val indices =
      if searchMode == -1 then lookupCells.indices.reverse
      else lookupCells.indices

    val matchedIndexOpt = matchMode match
      case 0 =>
        indices.find { idx =>
          val cellValue = ctx.sheet(lookupCells(idx)).value
          matchesExactForXLookup(cellValue, lookupValue)
        }
      case -1 =>
        findNextSmaller(lookupValue, lookupCells, ctx.sheet, indices)
      case 1 =>
        findNextLarger(lookupValue, lookupCells, ctx.sheet, indices)
      case 2 =>
        indices.find { idx =>
          val cellValue = ctx.sheet(lookupCells(idx)).value
          matchesExactForXLookup(cellValue, lookupValue) ||
          ((cellValue, lookupValue) match
            case (CellValue.Text(s), v: String) => s.toLowerCase.contains(v.toLowerCase)
            case _ => false)
        }
      case _ => None

    matchedIndexOpt match
      case Some(idx) => Right(ctx.sheet(returnCells(idx)).value)
      case None =>
        ifNotFoundOpt match
          case Some(expr) => evalAny(ctx, expr).map(toCellValue)
          case None => Left(EvalError.EvalFailed("XLOOKUP: no match found", None))

  private def matchesExactForXLookup(cellValue: CellValue, lookupValue: Any): Boolean =
    (cellValue, lookupValue) match
      case (CellValue.Number(n), v: BigDecimal) => n == v
      case (CellValue.Number(n), v: Int) => n == BigDecimal(v)
      case (CellValue.Number(n), v: Long) => n == BigDecimal(v)
      case (CellValue.Number(n), v: Double) => n == BigDecimal(v)
      case (CellValue.Text(s), v: String) => s.equalsIgnoreCase(v)
      case (CellValue.Bool(b), v: Boolean) => b == v
      case (CellValue.Formula(_, Some(cached)), v) => matchesExactForXLookup(cached, v)
      case _ => false

  private def findNextSmaller(
    lookupValue: Any,
    lookupCells: Vector[ARef],
    sheet: Sheet,
    indices: IndexedSeq[Int]
  ): Option[Int] =
    lookupValue match
      case targetNum: BigDecimal =>
        val candidates = indices
          .flatMap { idx =>
            extractNumericValue(sheet(lookupCells(idx)).value).map(n => (idx, n))
          }
          .filter(_._2 <= targetNum)
        candidates.sortBy(_._2).lastOption.map(_._1)
      case _ => None

  private def findNextLarger(
    lookupValue: Any,
    lookupCells: Vector[ARef],
    sheet: Sheet,
    indices: IndexedSeq[Int]
  ): Option[Int] =
    lookupValue match
      case targetNum: BigDecimal =>
        val candidates = indices
          .flatMap { idx =>
            extractNumericValue(sheet(lookupCells(idx)).value).map(n => (idx, n))
          }
          .filter(_._2 >= targetNum)
        candidates.sortBy(_._2).headOption.map(_._1)
      case _ => None

  private def extractNumericValue(value: CellValue): Option[BigDecimal] =
    value match
      case CellValue.Number(n) => Some(n)
      case CellValue.Formula(_, Some(CellValue.Number(n))) => Some(n)
      case _ => None

  val vlookup: FunctionSpec[CellValue] { type Args = VlookupArgs } =
    FunctionSpec.simple[CellValue, VlookupArgs]("VLOOKUP", Arity.Range(3, 4)) { (args, ctx) =>
      val (lookupExpr, table, colIndexExpr, rangeLookupOpt) = args
      val rangeLookupExpr = rangeLookupOpt.getOrElse(TExpr.Lit(true))
      for
        lookupValue <- evalAny(ctx, lookupExpr)
        colIndex <- ctx.evalExpr(colIndexExpr)
        rangeMatch <- ctx.evalExpr(rangeLookupExpr)
        targetSheet <- Evaluator.resolveRangeLocation(table, ctx.sheet, ctx.workbook)
        result <-
          if colIndex < 1 || colIndex > table.range.width then
            Left(
              EvalError.EvalFailed(
                s"VLOOKUP: col_index_num $colIndex is outside 1..${table.range.width}",
                Some(s"VLOOKUP(…, ${table.range.toA1})")
              )
            )
          else
            val rowIndices = 0 until table.range.height
            val keyCol0 = table.range.colStart.index0
            val rowStart0 = table.range.rowStart.index0
            val resultCol0 = keyCol0 + (colIndex - 1)

            def extractTextForMatch(cv: CellValue): Option[String] = cv match
              case CellValue.Text(s) => Some(s)
              case CellValue.Number(n) => Some(n.bigDecimal.stripTrailingZeros().toPlainString)
              case CellValue.Bool(b) => Some(if b then "TRUE" else "FALSE")
              case CellValue.Formula(_, Some(cached)) => extractTextForMatch(cached)
              case _ => None

            def extractNumericForMatch(cv: CellValue): Option[BigDecimal] = cv match
              case CellValue.Number(n) => Some(n)
              case CellValue.Text(s) => scala.util.Try(BigDecimal(s.trim)).toOption
              case CellValue.Bool(b) => Some(if b then BigDecimal(1) else BigDecimal(0))
              case CellValue.Formula(_, Some(cached)) => extractNumericForMatch(cached)
              case _ => None

            val normalizedLookup: Any = lookupValue match
              case cv: CellValue =>
                cv match
                  case CellValue.Number(n) => n
                  case CellValue.Text(s) => s
                  case CellValue.Bool(b) => b
                  case CellValue.Formula(_, Some(cached)) =>
                    cached match
                      case CellValue.Number(n) => n
                      case CellValue.Text(s) => s
                      case CellValue.Bool(b) => b
                      case other => other
                  case other => other
              case other => other

            val isTextLookup = normalizedLookup match
              case _: String => true
              case _: BigDecimal => false
              case _: Int => false
              case _: Boolean => false
              case _ => true

            val chosenRowOpt: Option[Int] =
              if rangeMatch then
                val numericLookup: Option[BigDecimal] = normalizedLookup match
                  case n: BigDecimal => Some(n)
                  case i: Int => Some(BigDecimal(i))
                  case s: String => scala.util.Try(BigDecimal(s.trim)).toOption
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
                val lookupText = normalizedLookup.toString.toLowerCase
                rowIndices.find { i =>
                  val keyRef = ARef.from0(keyCol0, rowStart0 + i)
                  extractTextForMatch(targetSheet(keyRef).value)
                    .exists(_.toLowerCase == lookupText)
                }
              else
                val numericLookup: Option[BigDecimal] = normalizedLookup match
                  case n: BigDecimal => Some(n)
                  case i: Int => Some(BigDecimal(i))
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
                      s"VLOOKUP($normalizedLookup, ${table.range.toA1}, $colIndex, $rangeMatch)"
                    )
                  )
                )
      yield result
    }

  val xlookup: FunctionSpec[CellValue] { type Args = XLookupArgs } =
    FunctionSpec.simple[CellValue, XLookupArgs]("XLOOKUP", Arity.Range(3, 6)) { (args, ctx) =>
      val (lookupValue, lookupArray, returnArray, ifNotFoundOpt, matchModeOpt, searchModeOpt) =
        args
      val matchModeExpr = matchModeOpt.getOrElse(TExpr.Lit(0))
      val searchModeExpr = searchModeOpt.getOrElse(TExpr.Lit(1))
      if lookupArray.width != returnArray.width || lookupArray.height != returnArray.height then
        Left(
          EvalError.EvalFailed(
            s"XLOOKUP: lookup_array and return_array must have same dimensions (${lookupArray.height}×${lookupArray.width} vs ${returnArray.height}×${returnArray.width})",
            Some(s"XLOOKUP(..., ${lookupArray.toA1}, ${returnArray.toA1}, ...)")
          )
        )
      else
        for
          lookupValueEval <- evalAny(ctx, lookupValue)
          matchModeRaw <- evalAny(ctx, matchModeExpr)
          searchModeRaw <- evalAny(ctx, searchModeExpr)
          matchMode = toInt(matchModeRaw)
          searchMode = toInt(searchModeRaw)
          result <- performXLookup(
            lookupValueEval,
            lookupArray,
            returnArray,
            ifNotFoundOpt,
            matchMode,
            searchMode,
            ctx
          )
        yield result
    }
