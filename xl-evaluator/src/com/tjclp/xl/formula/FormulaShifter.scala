package com.tjclp.xl.formula

import scala.annotation.nowarn

import com.tjclp.xl.{Anchor, CellRange}
import com.tjclp.xl.addressing.{ARef, Column, Row}

/**
 * AST transformer that shifts cell references for formula dragging.
 *
 * When a formula is applied across a range (e.g., B2:B100), references adjust based on their anchor
 * mode:
 *   - Relative (A1): Both column and row adjust
 *   - AbsCol ($A1): Column fixed, row adjusts
 *   - AbsRow (A$1): Column adjusts, row fixed
 *   - Absolute ($A$1): Both fixed
 *
 * Example: =A1*$B$1 applied from B2 to B5 becomes:
 *   - B2: =A1*$B$1 (no shift, delta = (0,0))
 *   - B3: =A2*$B$1 (A1→A2 shifted, $B$1 fixed)
 *   - B4: =A3*$B$1
 *   - B5: =A4*$B$1
 *
 * Laws:
 *   - Identity: shift(expr, 0, 0) == expr
 *   - Commutativity: shift(shift(expr, c1, r1), c2, r2) == shift(expr, c1+c2, r1+r2)
 *   - Anchor preservation: Anchor of shifted ref equals original anchor
 */
object FormulaShifter:

  /**
   * Shift all cell references in the expression by the given deltas.
   *
   * @param expr
   *   The formula AST to transform
   * @param colDelta
   *   Number of columns to shift (positive = right)
   * @param rowDelta
   *   Number of rows to shift (positive = down)
   * @return
   *   Transformed AST with shifted references
   */
  def shift[A](expr: TExpr[A], colDelta: Int, rowDelta: Int): TExpr[A] =
    if colDelta == 0 && rowDelta == 0 then expr
    else shiftInternal(expr, colDelta, rowDelta)

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  @nowarn(
    "msg=Unreachable case"
  ) // PolyRef extends TExpr[Nothing], reachable via asInstanceOf casts
  private def shiftInternal[A](expr: TExpr[A], colDelta: Int, rowDelta: Int): TExpr[A] =
    import TExpr.*

    expr match
      // Cell references - apply anchor-aware shifting
      case Ref(at, anchor, decode) =>
        val shiftedRef = shiftARef(at, anchor, colDelta, rowDelta)
        Ref(shiftedRef, anchor, decode)

      case PolyRef(at, anchor) =>
        val shiftedRef = shiftARef(at, anchor, colDelta, rowDelta)
        PolyRef(shiftedRef, anchor).asInstanceOf[TExpr[A]]

      // Sheet-qualified references - shift the cell ref but keep the sheet
      case SheetRef(sheet, at, anchor, decode) =>
        val shiftedRef = shiftARef(at, anchor, colDelta, rowDelta)
        SheetRef(sheet, shiftedRef, anchor, decode)

      case SheetPolyRef(sheet, at, anchor) =>
        val shiftedRef = shiftARef(at, anchor, colDelta, rowDelta)
        SheetPolyRef(sheet, shiftedRef, anchor).asInstanceOf[TExpr[A]]

      case SheetRange(sheet, range) =>
        val shiftedRange = shiftRange(range, colDelta, rowDelta)
        SheetRange(sheet, shiftedRange).asInstanceOf[TExpr[A]]

      // Literals - unchanged
      case lit: Lit[?] => lit.asInstanceOf[TExpr[A]]

      // Arithmetic operators
      case Add(x, y) =>
        Add(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Sub(x, y) =>
        Sub(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Mul(x, y) =>
        Mul(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Div(x, y) =>
        Div(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]

      // Boolean operators
      case And(x, y) =>
        And(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Or(x, y) =>
        Or(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Not(x) =>
        Not(shiftInternal(x, colDelta, rowDelta)).asInstanceOf[TExpr[A]]

      // Conditional
      case If(cond, ifTrue, ifFalse) =>
        If(
          shiftInternal(cond, colDelta, rowDelta),
          shiftInternal(ifTrue, colDelta, rowDelta),
          shiftInternal(ifFalse, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]

      // Comparison operators
      case Eq(x, y) =>
        Eq(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Neq(x, y) =>
        Neq(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Lt(x, y) =>
        Lt(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Lte(x, y) =>
        Lte(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Gt(x, y) =>
        Gt(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Gte(x, y) =>
        Gte(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]

      // Type conversion
      case ToInt(e) =>
        ToInt(shiftInternal(e, colDelta, rowDelta)).asInstanceOf[TExpr[A]]

      // Range aggregation - shift the range
      case FoldRange(range, z, step, decode) =>
        FoldRange(shiftRange(range, colDelta, rowDelta), z, step, decode).asInstanceOf[TExpr[A]]

      // Cross-sheet range aggregation - shift the range but preserve sheet name
      case SheetFoldRange(sheet, range, z, step, decode) =>
        SheetFoldRange(sheet, shiftRange(range, colDelta, rowDelta), z, step, decode)
          .asInstanceOf[TExpr[A]]

      // Arithmetic range functions
      case Min(range) =>
        Min(shiftRange(range, colDelta, rowDelta)).asInstanceOf[TExpr[A]]
      case Max(range) =>
        Max(shiftRange(range, colDelta, rowDelta)).asInstanceOf[TExpr[A]]
      case Average(range) =>
        Average(shiftRange(range, colDelta, rowDelta)).asInstanceOf[TExpr[A]]

      // Text functions
      case Concatenate(xs) =>
        Concatenate(xs.map(shiftInternal(_, colDelta, rowDelta))).asInstanceOf[TExpr[A]]
      case Left(text, n) =>
        Left(shiftInternal(text, colDelta, rowDelta), shiftInternal(n, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Right(text, n) =>
        Right(shiftInternal(text, colDelta, rowDelta), shiftInternal(n, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Len(text) =>
        Len(shiftInternal(text, colDelta, rowDelta)).asInstanceOf[TExpr[A]]
      case Upper(text) =>
        Upper(shiftInternal(text, colDelta, rowDelta)).asInstanceOf[TExpr[A]]
      case Lower(text) =>
        Lower(shiftInternal(text, colDelta, rowDelta)).asInstanceOf[TExpr[A]]

      // Date/Time functions - Today/Now have no refs to shift
      case t: Today => t.asInstanceOf[TExpr[A]]
      case n: Now => n.asInstanceOf[TExpr[A]]
      case Date(year, month, day) =>
        Date(
          shiftInternal(year, colDelta, rowDelta),
          shiftInternal(month, colDelta, rowDelta),
          shiftInternal(day, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]
      case Year(date) =>
        Year(shiftInternal(date, colDelta, rowDelta)).asInstanceOf[TExpr[A]]
      case Month(date) =>
        Month(shiftInternal(date, colDelta, rowDelta)).asInstanceOf[TExpr[A]]
      case Day(date) =>
        Day(shiftInternal(date, colDelta, rowDelta)).asInstanceOf[TExpr[A]]

      // Date calculation functions
      case Eomonth(startDate, months) =>
        Eomonth(
          shiftInternal(startDate, colDelta, rowDelta),
          shiftInternal(months, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]
      case Edate(startDate, months) =>
        Edate(
          shiftInternal(startDate, colDelta, rowDelta),
          shiftInternal(months, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]
      case Datedif(startDate, endDate, unit) =>
        Datedif(
          shiftInternal(startDate, colDelta, rowDelta),
          shiftInternal(endDate, colDelta, rowDelta),
          shiftInternal(unit, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]
      case Networkdays(startDate, endDate, holidays) =>
        Networkdays(
          shiftInternal(startDate, colDelta, rowDelta),
          shiftInternal(endDate, colDelta, rowDelta),
          holidays.map(shiftRange(_, colDelta, rowDelta))
        ).asInstanceOf[TExpr[A]]
      case Workday(startDate, days, holidays) =>
        Workday(
          shiftInternal(startDate, colDelta, rowDelta),
          shiftInternal(days, colDelta, rowDelta),
          holidays.map(shiftRange(_, colDelta, rowDelta))
        ).asInstanceOf[TExpr[A]]
      case Yearfrac(startDate, endDate, basis) =>
        Yearfrac(
          shiftInternal(startDate, colDelta, rowDelta),
          shiftInternal(endDate, colDelta, rowDelta),
          shiftInternal(basis, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]

      // Financial functions
      case Npv(rate, values) =>
        Npv(shiftInternal(rate, colDelta, rowDelta), shiftRange(values, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Irr(values, guess) =>
        Irr(
          shiftRange(values, colDelta, rowDelta),
          guess.map(shiftInternal(_, colDelta, rowDelta))
        ).asInstanceOf[TExpr[A]]
      case Xnpv(rate, values, dates) =>
        Xnpv(
          shiftInternal(rate, colDelta, rowDelta),
          shiftRange(values, colDelta, rowDelta),
          shiftRange(dates, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]
      case Xirr(values, dates, guess) =>
        Xirr(
          shiftRange(values, colDelta, rowDelta),
          shiftRange(dates, colDelta, rowDelta),
          guess.map(shiftInternal(_, colDelta, rowDelta))
        ).asInstanceOf[TExpr[A]]
      case VLookup(lookup, table, colIndex, rangeLookup) =>
        VLookup(
          shiftInternal(lookup, colDelta, rowDelta),
          shiftRange(table, colDelta, rowDelta),
          shiftInternal(colIndex, colDelta, rowDelta),
          shiftInternal(rangeLookup, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]

      // Conditional aggregation
      case SumIf(range, criteria, sumRange) =>
        SumIf(
          shiftRange(range, colDelta, rowDelta),
          shiftWildcard(criteria, colDelta, rowDelta),
          sumRange.map(shiftRange(_, colDelta, rowDelta))
        ).asInstanceOf[TExpr[A]]
      case CountIf(range, criteria) =>
        CountIf(
          shiftRange(range, colDelta, rowDelta),
          shiftWildcard(criteria, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]
      case SumIfs(sumRange, conditions) =>
        SumIfs(
          shiftRange(sumRange, colDelta, rowDelta),
          conditions.map { case (r, c) =>
            (shiftRange(r, colDelta, rowDelta), shiftWildcard(c, colDelta, rowDelta))
          }
        ).asInstanceOf[TExpr[A]]
      case CountIfs(conditions) =>
        CountIfs(
          conditions.map { case (r, c) =>
            (shiftRange(r, colDelta, rowDelta), shiftWildcard(c, colDelta, rowDelta))
          }
        ).asInstanceOf[TExpr[A]]

      // Array functions
      case SumProduct(arrays) =>
        SumProduct(arrays.map(shiftRange(_, colDelta, rowDelta))).asInstanceOf[TExpr[A]]
      case XLookup(lookupValue, lookupArray, returnArray, ifNotFound, matchMode, searchMode) =>
        XLookup(
          shiftWildcard(lookupValue, colDelta, rowDelta),
          shiftRange(lookupArray, colDelta, rowDelta),
          shiftRange(returnArray, colDelta, rowDelta),
          ifNotFound.map(shiftWildcard(_, colDelta, rowDelta)),
          shiftInternal(matchMode, colDelta, rowDelta),
          shiftInternal(searchMode, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]

      // Error handling functions
      case Iferror(value, valueIfError) =>
        Iferror(
          shiftInternal(value, colDelta, rowDelta),
          shiftInternal(valueIfError, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]
      case Iserror(value) =>
        Iserror(shiftInternal(value, colDelta, rowDelta)).asInstanceOf[TExpr[A]]

      // Rounding and math functions
      case Round(value, numDigits) =>
        Round(
          shiftInternal(value, colDelta, rowDelta),
          shiftInternal(numDigits, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]
      case RoundUp(value, numDigits) =>
        RoundUp(
          shiftInternal(value, colDelta, rowDelta),
          shiftInternal(numDigits, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]
      case RoundDown(value, numDigits) =>
        RoundDown(
          shiftInternal(value, colDelta, rowDelta),
          shiftInternal(numDigits, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]
      case Abs(value) =>
        Abs(shiftInternal(value, colDelta, rowDelta)).asInstanceOf[TExpr[A]]

      // Lookup functions
      case Index(array, rowNum, colNum) =>
        Index(
          shiftRange(array, colDelta, rowDelta),
          shiftInternal(rowNum, colDelta, rowDelta),
          colNum.map(shiftInternal(_, colDelta, rowDelta))
        ).asInstanceOf[TExpr[A]]
      case Match(lookupValue, lookupArray, matchType) =>
        Match(
          shiftWildcard(lookupValue, colDelta, rowDelta),
          shiftRange(lookupArray, colDelta, rowDelta),
          shiftInternal(matchType, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]

  /**
   * Shift a cell reference based on its anchor mode.
   *
   * @param ref
   *   The cell reference to shift
   * @param anchor
   *   The anchor mode determining which dimensions are fixed
   * @param colDelta
   *   Column shift amount
   * @param rowDelta
   *   Row shift amount
   * @return
   *   Shifted cell reference (clamped to valid bounds)
   */
  private def shiftARef(cellRef: ARef, anchor: Anchor, colDelta: Int, rowDelta: Int): ARef =
    // Extract column and row indices using companion object methods
    val colIdx = Column.index0(cellRef.col)
    val rowIdx = Row.index0(cellRef.row)

    val newCol =
      if anchor.isColAbsolute then colIdx
      else math.max(0, colIdx + colDelta)

    val newRow =
      if anchor.isRowAbsolute then rowIdx
      else math.max(0, rowIdx + rowDelta)

    ARef.from0(newCol, newRow)

  /**
   * Shift a cell range by the given deltas, respecting per-endpoint anchors.
   *
   * Each endpoint shifts according to its own anchor mode:
   *   - `$A$1:B10` → start ($A$1) is Absolute (fixed), end (B10) is Relative (shifts)
   *   - `$A1:B$10` → start has AbsCol (col fixed), end has AbsRow (row fixed)
   */
  private def shiftRange(range: CellRange, colDelta: Int, rowDelta: Int): CellRange =
    val newStart = shiftARef(range.start, range.startAnchor, colDelta, rowDelta)
    val newEnd = shiftARef(range.end, range.endAnchor, colDelta, rowDelta)
    new CellRange(newStart, newEnd, range.startAnchor, range.endAnchor)

  /**
   * Helper to shift TExpr[?] (wildcard type).
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def shiftWildcard(expr: TExpr[?], colDelta: Int, rowDelta: Int): TExpr[?] =
    shiftInternal(expr.asInstanceOf[TExpr[Any]], colDelta, rowDelta)
