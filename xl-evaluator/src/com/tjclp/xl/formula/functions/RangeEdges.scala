package com.tjclp.xl.formula.functions

import com.tjclp.xl.addressing.{CellRange, Column, Row}

/**
 * GH-746: the one rule for a range that runs to the sheet's edge. INDEX's reference values
 * (`FunctionSpecsBase.readsUsedRange`), SUMPRODUCT's shared bound, the criteria functions' trimmed
 * blank count and [[SumProductTail]] all read it here, so they cannot drift apart again.
 */
private[formula] object RangeEdges:

  /** Whether `range` runs to the sheet's last row (`A:A`, `A2:A1048576`): trimmable along rows. */
  def reachesLastRow(range: CellRange): Boolean = range.rowEnd.index0 == Row.MaxIndex0

  /**
   * Whether `range` runs to the sheet's last column (`2:2`, `C2:XFD2`): trimmable along columns.
   */
  def reachesLastColumn(range: CellRange): Boolean = range.colEnd.index0 == Column.MaxIndex0

  /** Whether `range` runs to the sheet's last row or last column. */
  def reachesEdge(range: CellRange): Boolean = reachesLastRow(range) || reachesLastColumn(range)
