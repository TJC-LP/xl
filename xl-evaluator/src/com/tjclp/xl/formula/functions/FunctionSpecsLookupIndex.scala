package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{ArrayArithmetic, ArrayResult, EvalError, RangeOperand}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.printer.FormulaPrinter

trait FunctionSpecsLookupIndex extends FunctionSpecsBase:
  /**
   * INDEX(array, row_num, [column_num], [area_num])
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
   *
   * GH-669: area_num picks one area of a union (`INDEX((A1:B2,D1:E2),1,1,2)` is D1) or of a
   * multi-area intersection; omitted or empty it is 1, below 1 `#VALUE!`, past the last area
   * `#REF!`. An array constant (`INDEX({1,2;3,4},2,1)` is 3) is indexed by value — it names no
   * cells — with the same row/column rules and one area.
   */
  val index: FunctionSpec[ArrayResult] { type Args = IndexArgs } =
    given ArgSpec[ReferenceOperators.Operand] = ReferenceOperators.indexOperand
    FunctionSpec.referencing[ArrayResult, IndexArgs](
      "INDEX",
      Arity.Range(2, 4),
      flags = FunctionFlags(lift = ArrayLift.all)
    )((args, ctx) => indexTarget(args, ctx).map(_.fold(identity, identity))) { (args, ctx) =>
      indexTarget(args, ctx).flatMap {
        case Left(values) => Right(values)
        case Right(ref) =>
          referenceResult(ref.range, ref.sheet, ctx)(boundedReferenceValues(ref, ctx))
      }
    }

  /**
   * The cells of `values` at the selected `rows` and `cols` ([[selection]]: each a whole axis or
   * one line, so a window): sliced, keeping a sparse array sparse (#750). The window lies inside
   * the array, so the slice is never refused.
   */
  private def sliceOf(
    values: ArrayResult,
    rows: Vector[Int],
    cols: Vector[Int]
  ): Either[EvalError, ArrayResult] =
    values.slice(rows.headOption.getOrElse(0), rows.size, cols.headOption.getOrElse(0), cols.size)

  /**
   * What INDEX selects: the values of an array constant (`Left`), or the reference into the chosen
   * area (`Right`) — one cell, or a whole row or column of it (unbounded — a reference position
   * reads it whole, ROWS counts it). A position outside the array is a descriptive `#REF!`.
   */
  private def indexTarget(
    args: IndexArgs,
    ctx: EvalContext
  ): Either[EvalError, Either[ArrayResult, RangeOperand]] =
    val (array, rowNumExpr, colNumOpt, areaNumOpt) = args
    for
      rowNum <- ctx.evalExpr(rowNumExpr)
      colNum <- colNumOpt match
        case Some(expr) => ctx.evalExpr(expr).map(Some(_))
        case None => Right(None)
      areaNum <- areaNumOpt match
        case Some(TExpr.Missing) | Some(TExpr.Coerced(TExpr.Missing, _)) | None => Right(1)
        case Some(expr) => ctx.evalExpr(expr).map(_.toInt)
      operandText = array match
        case Left(location) => location.toA1
        case Right(expr) => FormulaPrinter.printFileForm(expr)
      call = s"INDEX($operandText, $rowNum${colNum.map(c => s", $c").getOrElse("")}" +
        s"${areaNumOpt.fold("")(_ => s", $areaNum")})"
      // a value INDEX indexes: an array constant or any other expression that is no reference
      // (#725: a scalar is a 1x1 array, `INDEX(5,1)` is 5), or a LET name bound to a value
      // (LET(x, SEQUENCE(3), INDEX(x, 2)), LET(x, 5, INDEX(x, 1)))
      constant <- array match
        case Right(expr) if !ReferenceOperators.isReferenceShape(expr) =>
          // outside SUMPRODUCT's shared extent: the extent aligns arrays, and would cut the
          // positions INDEX counts (`INDEX(ROW(A:A),1000)`)
          ctx.evalBoundedArrayExpr(expr, identity).map(valuesOf)
        case Left(TExpr.RangeLocation.Binding(name)) =>
          Right(ctx.bindings.get(name).flatMap(valuesOf))
        case _ => Right(None)
      target <- constant match
        case Some(values) => indexValues(values, rowNum, colNum, areaNum, call)
        case None =>
          ReferenceOperators.areas(array, ctx) match
            case Right(areas) =>
              for
                _ <- pickArea(areas.size, areaNum, call)
                RangeOperand(targetSheet, arrayRange) = areas(areaNum - 1)
                (rows, cols) <- selection(rowNum, colNum, arrayRange.height, arrayRange.width, call)
              yield
                val startCol = arrayRange.colStart.index0
                val startRow = arrayRange.rowStart.index0
                Right(
                  RangeOperand(
                    targetSheet,
                    CellRange(
                      ARef.from0(
                        startCol + cols.headOption.getOrElse(0),
                        startRow + rows.headOption.getOrElse(0)
                      ),
                      ARef.from0(
                        startCol + cols.lastOption.getOrElse(0),
                        startRow + rows.lastOption.getOrElse(0)
                      )
                    )
                  )
                )
            // #725: a defined name bound to a value (`konst = 5`, `=SUM(A1:A3)`) is INDEX's
            // array: its range slot is `#VALUE!` ("does not refer to a range"), and the name is
            // evaluated for its value. Any other error (an OFFSET off the grid) stays the result.
            case Left(notAReference @ EvalError.ErrorValue(CellError.Value, _)) =>
              array match
                case Left(TExpr.RangeLocation.Name(name, scope)) =>
                  val named = scope.fold[TExpr[?]](TExpr.NameRef(name))(TExpr.SheetNameRef(_, name))
                  ctx
                    .evalReference(named.asInstanceOf[TExpr[Any]])
                    .flatMap(valuesOf(_).toRight(notAReference))
                    .flatMap(indexValues(_, rowNum, colNum, areaNum, call))
                case _ => Left(notAReference)
            case Left(err) => Left(err)
    yield target

  /** The values of INDEX's first slot when it holds a value, never a reference. */
  private def valuesOf(value: Any): Option[ArrayResult] = value match
    case _: RangeOperand => None
    case values: ArrayResult => Some(values)
    case scalar => Some(ArrayResult.single(ArrayArithmetic.anyToCellValue(scalar)))

  /** INDEX over values, which form one area: the selected window of them. */
  private def indexValues(
    values: ArrayResult,
    rowNum: BigDecimal,
    colNum: Option[BigDecimal],
    areaNum: Int,
    call: String
  ): Either[EvalError, Either[ArrayResult, RangeOperand]] =
    pickArea(1, areaNum, call).flatMap(_ =>
      selection(rowNum, colNum, values.rows, values.cols, call).flatMap { (rows, cols) =>
        sliceOf(values, rows, cols).map(Left(_))
      }
    )

  /** GH-669: area_num against the number of areas — `#VALUE!` below 1, `#REF!` past the last. */
  private def pickArea(count: Int, areaNum: Int, call: String): Either[EvalError, Unit] =
    if areaNum < 1 then
      Left(
        EvalError.ErrorValue(CellError.Value, Some(s"INDEX: area_num $areaNum is below 1 ($call)"))
      )
    else if areaNum > count then
      Left(
        EvalError.ErrorValue(
          CellError.Ref,
          Some(s"INDEX: area_num $areaNum is past the reference's $count area(s) ($call)")
        )
      )
    else Right(())

  /**
   * The 0-based rows and columns INDEX selects in an array of `numRows` × `numCols` — never empty.
   * Excel's two-argument form: the single position reads along a vector's long axis; on a 2-D array
   * it is row_num and the whole row is selected (column_num 0).
   */
  private def selection(
    rowNum: BigDecimal,
    colNum: Option[BigDecimal],
    numRows: Int,
    numCols: Int,
    call: String
  ): Either[EvalError, (Vector[Int], Vector[Int])] =
    val (rowPos, colPos) = colNum match
      case Some(c) => (rowNum.toInt, c.toInt)
      case None if numRows == 1 => (1, rowNum.toInt)
      case None => (rowNum.toInt, 0)
    for
      rowSel <- indexAxis(rowPos, numRows, "row_num", "rows", call)
      colSel <- indexAxis(colPos, numCols, "col_num", "columns", call)
    yield
      def line(sel: Option[Int], size: Int): Vector[Int] =
        sel.fold(Vector.range(0, size))(Vector(_))
      (line(rowSel, numRows), line(colSel, numCols))

  /** None = the whole axis (Excel's 0), Some(i) = one 0-based line of it. */
  private def indexAxis(
    pos: Int,
    size: Int,
    label: String,
    noun: String,
    call: String
  ): Either[EvalError, Option[Int]] =
    // #670: Excel's codes as cached error values — a negative position is #VALUE!, one past the
    // array #REF! (LibreOffice: Err:502, its #VALUE!, and #REF!)
    if pos == 0 then Right(None)
    else if pos < 0 then
      Left(EvalError.ErrorValue(CellError.Value, Some(s"INDEX: $label $pos is negative: $call")))
    else if pos > size then
      Left(
        EvalError.ErrorValue(
          CellError.Ref,
          Some(
            s"INDEX: $label $pos is out of bounds (array has $size $noun, valid range: 1-$size): $call"
          )
        )
      )
    else Right(Some(pos - 1))

  val matchFn: FunctionSpec[BigDecimal] { type Args = MatchArgs } =
    FunctionSpec.simple[BigDecimal, MatchArgs](
      "MATCH",
      Arity.Range(2, 3),
      flags = FunctionFlags(returnsNumeric = true, lift = ArrayLift.all)
    ) { (args, ctx) =>
      val (lookupValue, lookupArray, matchTypeOpt) = args
      val matchTypeExpr = matchTypeOpt.getOrElse(TExpr.Lit(BigDecimal(1)))
      for
        lookupValueEval <- evalValue(ctx, lookupValue)
        // GH-467: dereference cell-ref lookup values and coerce dates to serials up front —
        // positions below are RANGE positions (idx + 1), blanks included, never compacted.
        lookup = normalizeLookupValue(lookupValueEval)
        _ <- lookupValueError("MATCH", lookup)
        matchType <- ctx.evalExpr(matchTypeExpr)
        resolved <- ctx.resolveRange(lookupArray)
        (targetSheet, lookupRange) = resolved
        // #670(c): Excel reads match_type by its sign once truncated — 5 is 1, -3 is -1, and a
        // fraction below 1 in magnitude is 0 (LibreOffice agrees on 1.5, -1.5 and 0.5)
        matchTypeInt = matchType.setScale(0, BigDecimal.RoundingMode.DOWN).signum
        // the keys as references read them, lazily (GH-715: see LookupLine)
        line = lookupLine(lookupRange, targetSheet, ctx)
        found <- matchTypeInt match
          case 0 => line.find(forward = true)(matchesLookupExact(_, lookup))
          // #670(g): 1 is the largest key ≤ the value over ascending keys, -1 the smallest key ≥
          // it over descending keys — numbers or text (case-insensitive), the last of a run of
          // equal keys as Excel's binary search lands
          case sign =>
            line.storedKeys.map(keys =>
              approximateMatch(
                keys.zipWithIndex.map(_.swap),
                lookup,
                if sign > 0 then -1 else 1,
                lastOnTie = true
              )
            )
        result <- {
          // positions are 1-based
          val positionOpt = found.map(_ + 1)
          positionOpt match
            case Some(pos) => Right(BigDecimal(pos))
            case None =>
              // GH-662: the typed #N/A (Left channel: MATCH is a FunctionSpec[BigDecimal]), with
              // the real call rendered as its diagnostic context
              Left(
                lookupNotFound(
                  s"MATCH: no match found for lookup value: MATCH(${renderLookupValue(lookup)}, ${lookupArray.toA1}, $matchTypeInt)"
                )
              )
        }
      yield result
    }

  /**
   * XMATCH(lookup_value, lookup_array, [match_mode], [search_mode]) — #670(d), the position half of
   * XLOOKUP over a single row or column.
   *
   * match_mode 0 (default) exact, -1 exact or next smaller, 1 exact or next larger, 2 wildcard
   * (`*`, `?`, `~` escapes); search_mode 1 (default) first to last, -1 last to first, 2 / -2 the
   * binary searches (answered by the same linear search in their direction, which agrees on the
   * sorted data they require). Keys compare in the lookup plane: numbers with numbers, text with
   * text case-insensitively, logicals with logicals; a tie keeps the first key in search order. A
   * mode outside those tables or a 2-D lookup_array is `#VALUE!`; a miss raises through
   * [[lookupNotFound]]. An empty optional slot is omitted, as for XLOOKUP (GH-654).
   */
  val xmatch: FunctionSpec[BigDecimal] { type Args = XMatchArgs } =
    FunctionSpec.simple[BigDecimal, XMatchArgs](
      "XMATCH",
      Arity.Range(2, 4),
      flags = FunctionFlags(returnsNumeric = true, lift = ArrayLift.all)
    ) { (args, ctx) =>
      val (lookupExpr, lookupLoc, matchModeSlot, searchModeSlot) = args
      val matchModeExpr = unlessOmitted(matchModeSlot).getOrElse(TExpr.Lit(0))
      val searchModeExpr = unlessOmitted(searchModeSlot).getOrElse(TExpr.Lit(1))
      def invalid(detail: String): Either[EvalError, Unit] =
        Left(EvalError.ErrorValue(CellError.Value, Some(s"XMATCH: $detail")))
      for
        lookupValue <- evalValue(ctx, lookupExpr)
        lookup = normalizeLookupValue(lookupValue)
        _ <- lookupValueError("XMATCH", lookup)
        matchModeRaw <- evalValue(ctx, matchModeExpr)
        matchMode <- toIntArg("XMATCH", matchModeRaw)
        searchModeRaw <- evalValue(ctx, searchModeExpr)
        searchMode <- toIntArg("XMATCH", searchModeRaw)
        call =
          s"XMATCH(${renderLookupValue(lookup)}, ${lookupLoc.toA1}, $matchMode, $searchMode)"
        _ <- matchModeError("XMATCH", matchMode, call)
        _ <- searchModeError("XMATCH", searchMode, call)
        resolved <- ctx.resolveRange(lookupLoc)
        (targetSheet, lookupRange) = resolved
        _ <-
          if lookupRange.width == 1 || lookupRange.height == 1 then Right(())
          else invalid(s"lookup_array ${lookupRange.toA1} is not a single row or column: $call")
        // #692: the exact and wildcard modes walk the range lazily in search_mode's direction and
        // stop at the first hit (positions 1-based along the single row or column); -1 and 1 read
        // the whole (used) range for the approximate plane
        positionOpt <- matchMode match
          case 0 =>
            exactLookupPosition(lookupRange, targetSheet, ctx, lookup, None, searchMode)
              .map(_.map(_ + 1))
          case 2 =>
            exactLookupPosition(
              lookupRange,
              targetSheet,
              ctx,
              lookup,
              wildcardCriterion(lookup),
              searchMode
            ).map(_.map(_ + 1))
          case -1 | 1 =>
            lookupRangeValues(lookupRange, targetSheet, ctx).map { values =>
              // flattening is position-exact only because the single row/column check above ran
              val keys = values.flatten.zipWithIndex.map { case (cv, idx) => (idx + 1, cv) }
              val ordered = if searchMode < 0 then keys.reverse else keys
              approximateMatch(ordered, lookup, matchMode, lastOnTie = false)
            }
          // unreachable: matchModeError above validated the mode; kept total rather than sending
          // an unknown mode into the approximate search
          case other => matchModeError("XMATCH", other, call).map(_ => None)
        result <- positionOpt
          .map(pos => BigDecimal(pos))
          .toRight(lookupNotFound(s"XMATCH: no match found for lookup value: $call"))
      yield result
    }
