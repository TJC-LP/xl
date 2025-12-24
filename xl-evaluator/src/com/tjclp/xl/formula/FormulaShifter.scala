package com.tjclp.xl.formula

import scala.annotation.nowarn

import com.tjclp.xl.{Anchor, CellRange}
import com.tjclp.xl.addressing.{ARef, Column, Row}
import TExpr.RangeLocation

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

      case RangeRef(range) =>
        val shiftedRange = shiftRange(range, colDelta, rowDelta)
        RangeRef(shiftedRange).asInstanceOf[TExpr[A]]

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

      // Arithmetic range functions (now using RangeLocation)
      case Aggregate(aggregatorId, location) =>
        Aggregate(aggregatorId, shiftLocation(location, colDelta, rowDelta)).asInstanceOf[TExpr[A]]

      case call: Call[?] =>
        val shifted =
          call.spec.argSpec.map(call.args)(
            expr => shiftInternal(expr, colDelta, rowDelta),
            loc => shiftLocation(loc, colDelta, rowDelta),
            range => shiftRange(range, colDelta, rowDelta)
          )
        Call(call.spec, shifted).asInstanceOf[TExpr[A]]

      // Date-to-serial converters - shift inner expression
      case DateToSerial(dateExpr) =>
        DateToSerial(shiftInternal(dateExpr, colDelta, rowDelta)).asInstanceOf[TExpr[A]]
      case DateTimeToSerial(dtExpr) =>
        DateTimeToSerial(shiftInternal(dtExpr, colDelta, rowDelta)).asInstanceOf[TExpr[A]]

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
   * Shift a RangeLocation by the given deltas.
   *
   * Handles both Local and CrossSheet locations, shifting the underlying CellRange.
   */
  private def shiftLocation(location: RangeLocation, colDelta: Int, rowDelta: Int): RangeLocation =
    location match
      case RangeLocation.Local(range) =>
        RangeLocation.Local(shiftRange(range, colDelta, rowDelta))
      case RangeLocation.CrossSheet(sheet, range) =>
        RangeLocation.CrossSheet(sheet, shiftRange(range, colDelta, rowDelta))

  /**
   * Helper to shift TExpr[?] (wildcard type).
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def shiftWildcard(expr: TExpr[?], colDelta: Int, rowDelta: Int): TExpr[?] =
    shiftInternal(expr.asInstanceOf[TExpr[Any]], colDelta, rowDelta)
