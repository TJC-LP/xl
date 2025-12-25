package com.tjclp.xl.formula

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet

trait FunctionSpecsLookupSearch extends FunctionSpecsBase:
  private def performXLookup(
    lookupValue: ExprValue,
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

    val wildcardCriterionOpt = lookupValue match
      case ExprValue.Text(text) =>
        CriteriaMatcher.parse(ExprValue.Text(text)) match
          case c: CriteriaMatcher.Wildcard => Some(c)
          case _ => None
      case ExprValue.Cell(CellValue.Text(text)) =>
        CriteriaMatcher.parse(ExprValue.Text(text)) match
          case c: CriteriaMatcher.Wildcard => Some(c)
          case _ => None
      case _ => None

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
          wildcardCriterionOpt.exists(CriteriaMatcher.matches(cellValue, _))
        }
      case _ => None

    matchedIndexOpt match
      case Some(idx) => Right(ctx.sheet(returnCells(idx)).value)
      case None =>
        ifNotFoundOpt match
          case Some(expr) => evalValue(ctx, expr).map(toCellValue)
          case None => Right(CellValue.Error(CellError.NA))

  private def matchesExactForXLookup(cellValue: CellValue, lookupValue: ExprValue): Boolean =
    (cellValue, lookupValue) match
      case (CellValue.Number(n), ExprValue.Number(v)) => n == v
      case (CellValue.Text(s), ExprValue.Text(v)) => s.equalsIgnoreCase(v)
      case (CellValue.Bool(b), ExprValue.Bool(v)) => b == v
      case (CellValue.Formula(_, Some(cached)), v) => matchesExactForXLookup(cached, v)
      case _ => false

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
        lookupValue <- evalValue(ctx, lookupExpr)
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

            def renderValue(value: ExprValue): String = value match
              case ExprValue.Text(s) => s
              case ExprValue.Number(n) => n.toString
              case ExprValue.Bool(b) => b.toString
              case ExprValue.Date(d) => d.toString
              case ExprValue.DateTime(dt) => dt.toString
              case ExprValue.Cell(cv) => cv.toString
              case ExprValue.Opaque(other) => other.toString

            val normalizedLookup: ExprValue = lookupValue match
              case ExprValue.Cell(cv) =>
                cv match
                  case CellValue.Number(n) => ExprValue.Number(n)
                  case CellValue.Text(s) => ExprValue.Text(s)
                  case CellValue.Bool(b) => ExprValue.Bool(b)
                  case CellValue.Formula(_, Some(cached)) =>
                    cached match
                      case CellValue.Number(n) => ExprValue.Number(n)
                      case CellValue.Text(s) => ExprValue.Text(s)
                      case CellValue.Bool(b) => ExprValue.Bool(b)
                      case other => ExprValue.Cell(other)
                  case other => ExprValue.Cell(other)
              case other => other

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
                      s"VLOOKUP(${renderValue(normalizedLookup)}, ${table.range.toA1}, $colIndex, $rangeMatch)"
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
          lookupValueEval <- evalValue(ctx, lookupValue)
          matchModeRaw <- evalValue(ctx, matchModeExpr)
          searchModeRaw <- evalValue(ctx, searchModeExpr)
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
