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

    // #670(g): the next-smaller/larger modes share the lookups' approximate plane (text keys
    // included); a linear search, so a tie keeps the first key in search order
    def approximate(direction: Int): Option[Int] =
      approximateMatch(
        indices.map(idx => (idx, lookupSheet(lookupCells(idx)).value)),
        lookup,
        direction,
        lastOnTie = false
      )

    val matchedIndexOpt = matchMode match
      case 0 =>
        indices.find { idx =>
          val cellValue = lookupSheet(lookupCells(idx)).value
          matchesLookupExact(cellValue, lookup)
        }
      case -1 => approximate(-1)
      case 1 => approximate(1)
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

  /**
   * VLOOKUP/HLOOKUP's search over the key line: approximate (range_lookup TRUE) through the shared
   * plane, the last of a run of equal keys winning as Excel's binary search lands; exact by the
   * key's own type, text case-insensitively. A blank lookup value matches nothing (LibreOffice:
   * `VLOOKUP(blank,{"Empty";"x";0},1,FALSE)` is `#N/A`).
   */
  private def legacyLookupIndex(
    keys: IndexedSeq[CellValue],
    lookup: ExprValue,
    rangeMatch: Boolean
  ): Option[Int] =
    if rangeMatch then approximateMatch(keys.zipWithIndex.map(_.swap), lookup, -1, lastOnTie = true)
    else
      lookup match
        case ExprValue.Text(text) =>
          keys.indexWhere(extractTextForMatch(_).exists(_.equalsIgnoreCase(text))) match
            case -1 => None
            case i => Some(i)
        case ExprValue.Number(n) =>
          keys.indexWhere(extractNumericForMatch(_).contains(n)) match
            case -1 => None
            case i => Some(i)
        case ExprValue.Bool(_) =>
          keys.indexWhere(matchesLookupExact(_, lookup)) match
            case -1 => None
            case i => Some(i)
        case _ => None

  val vlookup: FunctionSpec[CellValue] { type Args = VlookupArgs } =
    FunctionSpec.simple[CellValue, VlookupArgs](
      "VLOOKUP",
      Arity.Range(3, 4),
      flags = FunctionFlags(lift = ArrayLift.all)
    ) { (args, ctx) =>
      val (lookupExpr, table, colIndexExpr, rangeLookupOpt) = args
      val rangeLookupExpr = rangeLookupOpt.getOrElse(TExpr.Lit(true))
      for
        lookupValue <- evalValue(ctx, lookupExpr)
        // GH-488: one lookup plane for MATCH/XLOOKUP/VLOOKUP/HLOOKUP (dates as serials)
        normalizedLookup = normalizeLookupValue(lookupValue)
        _ <- lookupValueError("VLOOKUP", normalizedLookup)
        colIndex <- ctx.evalExpr(colIndexExpr)
        rangeMatch <- ctx.evalExpr(rangeLookupExpr)
        resolved <- Evaluator.resolveRangeLocation(table, ctx.sheet, ctx.workbook)
        (targetSheet, tableRange) = resolved
        result <-
          // GH-662: Excel's codes — an index below 1 is #VALUE!, one beyond the table is #REF!
          if colIndex < 1 then
            Left(
              EvalError.ErrorValue(
                CellError.Value,
                Some(s"VLOOKUP: col_index_num $colIndex is below 1: VLOOKUP(…, ${table.toA1})")
              )
            )
          else if colIndex > tableRange.width then
            Left(
              EvalError.ErrorValue(
                CellError.Ref,
                Some(
                  s"VLOOKUP: col_index_num $colIndex exceeds the table width ${tableRange.width}: VLOOKUP(…, ${table.toA1})"
                )
              )
            )
          else
            val keyCol0 = tableRange.colStart.index0
            val rowStart0 = tableRange.rowStart.index0
            val resultCol0 = keyCol0 + (colIndex - 1)
            val keys = (0 until tableRange.height).map { i =>
              targetSheet(ARef.from0(keyCol0, rowStart0 + i)).value
            }
            legacyLookupIndex(keys, normalizedLookup, rangeMatch) match
              case Some(rowIndex) =>
                Right(targetSheet(ARef.from0(resultCol0, rowStart0 + rowIndex)).value)
              case None =>
                // GH-662: a miss is Excel's #N/A — IFNA/ISNA-visible, cached when unguarded —
                // with the diagnostic kept as the error's context for putf/eval error text
                val mode = if rangeMatch then "approximate" else "exact"
                Left(
                  lookupNotFound(
                    s"VLOOKUP $mode match not found: VLOOKUP(${renderLookupValue(normalizedLookup)}, ${table.toA1}, $colIndex, ${renderBoolean(rangeMatch)})"
                  )
                )
      yield result
    }

  /**
   * HLOOKUP(lookup, table, row_index_num, [range_lookup])
   *
   * Horizontal transpose of VLOOKUP: searches the first ROW of the table, returns the cell from the
   * 1-based row_index_num in the matched COLUMN. range_lookup TRUE (default) = approximate.
   */
  val hlookup: FunctionSpec[CellValue] { type Args = VlookupArgs } =
    FunctionSpec.simple[CellValue, VlookupArgs](
      "HLOOKUP",
      Arity.Range(3, 4),
      flags = FunctionFlags(lift = ArrayLift.all)
    ) { (args, ctx) =>
      val (lookupExpr, table, rowIndexExpr, rangeLookupOpt) = args
      val rangeLookupExpr = rangeLookupOpt.getOrElse(TExpr.Lit(true))
      for
        lookupValue <- evalValue(ctx, lookupExpr)
        // GH-488: shared lookup plane (see VLOOKUP above)
        normalizedLookup = normalizeLookupValue(lookupValue)
        _ <- lookupValueError("HLOOKUP", normalizedLookup)
        rowIndex <- ctx.evalExpr(rowIndexExpr)
        rangeMatch <- ctx.evalExpr(rangeLookupExpr)
        resolved <- Evaluator.resolveRangeLocation(table, ctx.sheet, ctx.workbook)
        (targetSheet, tableRange) = resolved
        result <-
          // GH-662: Excel's codes, as VLOOKUP above
          if rowIndex < 1 then
            Left(
              EvalError.ErrorValue(
                CellError.Value,
                Some(s"HLOOKUP: row_index_num $rowIndex is below 1: HLOOKUP(…, ${table.toA1})")
              )
            )
          else if rowIndex > tableRange.height then
            Left(
              EvalError.ErrorValue(
                CellError.Ref,
                Some(
                  s"HLOOKUP: row_index_num $rowIndex exceeds the table height ${tableRange.height}: HLOOKUP(…, ${table.toA1})"
                )
              )
            )
          else
            val keyRow0 = tableRange.rowStart.index0
            val colStart0 = tableRange.colStart.index0
            val resultRow0 = keyRow0 + (rowIndex - 1)
            val keys = (0 until tableRange.width).map { i =>
              targetSheet(ARef.from0(colStart0 + i, keyRow0)).value
            }
            legacyLookupIndex(keys, normalizedLookup, rangeMatch) match
              case Some(colIdx) =>
                Right(targetSheet(ARef.from0(colStart0 + colIdx, resultRow0)).value)
              case None =>
                // GH-662: the typed #N/A, as VLOOKUP above
                val mode = if rangeMatch then "approximate" else "exact"
                Left(
                  lookupNotFound(
                    s"HLOOKUP $mode match not found: HLOOKUP(${renderLookupValue(normalizedLookup)}, ${table.toA1}, $rowIndex, ${renderBoolean(rangeMatch)})"
                  )
                )
      yield result
    }

  val xlookup: FunctionSpec[CellValue] { type Args = XLookupArgs } =
    FunctionSpec.simple[CellValue, XLookupArgs](
      "XLOOKUP",
      Arity.Range(3, 6),
      flags = FunctionFlags(lift = ArrayLift.slots(0))
    ) { (args, ctx) =>
      val (lookupValue, lookupLoc, returnLoc, ifNotFoundSlot, matchModeSlot, searchModeSlot) =
        args
      // GH-654: XLOOKUP reads an empty optional slot as omitted — `XLOOKUP(x,a,b,,0)` is #N/A
      // when nothing matches, not a blank result (Excel and LibreOffice agree)
      val ifNotFoundOpt = unlessOmitted(ifNotFoundSlot)
      val matchModeOpt = unlessOmitted(matchModeSlot)
      val searchModeOpt = unlessOmitted(searchModeSlot)
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
            // #670: Excel's #VALUE!, a cached error value rather than a host failure
            Left(
              EvalError.ErrorValue(
                CellError.Value,
                Some(
                  s"XLOOKUP: lookup_array and return_array must have same dimensions (${lookupArray.height}×${lookupArray.width} vs ${returnArray.height}×${returnArray.width}): XLOOKUP(…, ${lookupLoc.toA1}, ${returnLoc.toA1}, …)"
                )
              )
            )
          else Right(())
        lookupValueEval <- evalValue(ctx, lookupValue)
        _ <- lookupValueError("XLOOKUP", normalizeLookupValue(lookupValueEval))
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

  /**
   * LOOKUP(lookup_value, lookup_vector, [result_vector]) and LOOKUP(lookup_value, array) — #670(d).
   *
   * Always approximate over sorted keys: the largest key ≤ lookup_value (the last of a run of equal
   * keys, as Excel's binary search lands), numbers against numbers and text against text
   * case-insensitively. The array form searches the first row of an array wider than it is tall,
   * otherwise its first column, and returns from the last row or column. result_vector may lie on
   * the other axis; a position past its end is `#N/A` (LibreOffice). A miss raises through
   * [[lookupNotFound]].
   */
  val lookup: FunctionSpec[CellValue] { type Args = LookupArgs } =
    FunctionSpec.simple[CellValue, LookupArgs](
      "LOOKUP",
      Arity.Range(2, 3),
      flags = FunctionFlags(lift = ArrayLift.slots(0))
    ) { (args, ctx) =>
      val (lookupExpr, lookupLoc, resultLocOpt) = args
      // position i of a range's line: the first row of a range wider than it is tall, otherwise
      // the first column (a vector is its own line)
      def lineCell(range: CellRange, i: Int): ARef =
        if range.width > range.height then
          ARef.from0(range.colStart.index0 + i, range.rowStart.index0)
        else ARef.from0(range.colStart.index0, range.rowStart.index0 + i)
      val call = s"${lookupLoc.toA1}${resultLocOpt.fold("")(loc => s", ${loc.toA1}")}"
      for
        lookupValue <- evalValue(ctx, lookupExpr)
        normalized = normalizeLookupValue(lookupValue)
        _ <- lookupValueError("LOOKUP", normalized)
        resolved <- Evaluator.resolveRangeLocation(lookupLoc, ctx.sheet, ctx.workbook)
        (lookupSheet, lookupRange) = resolved
        resultTarget <- resultLocOpt match
          case Some(loc) =>
            Evaluator.resolveRangeLocation(loc, ctx.sheet, ctx.workbook).map(Some(_))
          case None => Right(None)
        values <- lookupRangeValues(lookupRange, lookupSheet, ctx)
        wide = lookupRange.width > lookupRange.height
        keys =
          if wide then values.headOption.getOrElse(Vector.empty) else values.flatMap(_.headOption)
        // the result cell of a key position: result_vector's own line (None past its end), or
        // the array's last row (wide) / last column (tall)
        target = (i: Int) =>
          resultTarget match
            case Some((resultSheet, resultRange)) =>
              Option.when(i < math.max(resultRange.width, resultRange.height))(
                (resultSheet, lineCell(resultRange, i))
              )
            case None if wide =>
              Some(
                (
                  lookupSheet,
                  ARef.from0(lookupRange.colStart.index0 + i, lookupRange.rowEnd.index0)
                )
              )
            case None =>
              Some(
                (
                  lookupSheet,
                  ARef.from0(lookupRange.colEnd.index0, lookupRange.rowStart.index0 + i)
                )
              )
        result <- approximateMatch(keys.zipWithIndex.map(_.swap), normalized, -1, lastOnTie = true)
          .flatMap(target) match
          case Some((resultSheet, at)) => rangeCellReader(resultSheet, ctx)(at)
          case None =>
            Left(
              lookupNotFound(
                s"LOOKUP: no match found for lookup value: LOOKUP(${renderLookupValue(normalized)}, $call)"
              )
            )
      yield result
    }
