package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.{ARef, CellRange, Column, Row, SheetName}

trait TExprRangeLocation:
  /**
   * Where a range is located - same sheet or cross-sheet.
   *
   * This enum unifies local and cross-sheet range references, eliminating the need for paired TExpr
   * cases (e.g., Min + SheetMin). Used by TExpr.Aggregate for unified aggregation.
   */
  enum RangeLocation derives CanEqual:
    case Local(range: CellRange)
    case CrossSheet(sheet: SheetName, range: CellRange)

  object RangeLocation:
    extension (loc: RangeLocation)
      /** Extract the CellRange regardless of location */
      def range: CellRange = loc match
        case Local(r) => r
        case CrossSheet(_, r) => r

      /** Get sheet name for cross-sheet, None for local */
      def sheetName: Option[SheetName] = loc match
        case CrossSheet(s, _) => Some(s)
        case _ => None

      /** Get cells for local ranges only (for intra-sheet dependency graphs) */
      def localCells: Set[ARef] = loc match
        case Local(r) => r.cells.toSet
        case CrossSheet(_, _) => Set.empty

      /**
       * Get cells for local ranges, bounded by the sheet's used range.
       *
       * Preferred for dependency graph construction to avoid materializing 1M+ cells for full
       * column/row references like A:A or 1:1.
       *
       * @param bounds
       *   Optional bounding range (typically sheet.usedRange)
       * @return
       *   Set of cell references in the intersection of this range and bounds
       */
      def localCellsBounded(bounds: Option[CellRange]): Set[ARef] = loc match
        case Local(r) =>
          bounds match
            case Some(b) => r.intersect(b).map(_.cells.toSet).getOrElse(Set.empty)
            case None => r.cells.toSet
        case CrossSheet(_, _) => Set.empty

      /** Check if this is a cross-sheet reference */
      def isCrossSheet: Boolean = loc match
        case CrossSheet(_, _) => true
        case _ => false

      /** Get all cells in the range (delegates to underlying CellRange) */
      def cells: Iterator[ARef] = loc.range.cells

      /** Get width of the range */
      def width: Int = loc.range.width

      /** Get height of the range */
      def height: Int = loc.range.height

      /** Get A1 string representation */
      def toA1: String = loc match
        case Local(r) => r.toA1
        case CrossSheet(s, r) => s"${s.value}!${r.toA1}"

      /** Get starting column of the range */
      def colStart: Column = loc.range.colStart

      /** Get ending column of the range */
      def colEnd: Column = loc.range.colEnd

      /** Get starting row of the range */
      def rowStart: Row = loc.range.rowStart

      /** Get ending row of the range */
      def rowEnd: Row = loc.range.rowEnd

      /** Get start cell reference of the range */
      def start: ARef = loc.range.start

      /** Get end cell reference of the range */
      def end: ARef = loc.range.end
