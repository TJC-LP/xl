package com.tjclp.xl

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.error.{XLError, XLResult}

/**
 * Helper constructors and parsing extensions.
 *
 * Import with:
 * {{{
 * import com.tjclp.xl.syntax.*
 * import com.tjclp.xl.syntax.given
 * }}}
 */
object syntax:

  // ========== String Parsing Helpers ==========
  extension (s: String)
    /** Parse string as cell reference */
    def asCell: XLResult[ARef] =
      ARef.parse(s).left.map(err => XLError.InvalidCellRef(s, err))

    /** Parse string as range */
    def asRange: XLResult[CellRange] =
      CellRange.parse(s).left.map(err => XLError.InvalidRange(s, err))

    /** Parse string as sheet name */
    def asSheetName: XLResult[SheetName] =
      SheetName(s).left.map(err => XLError.InvalidSheetName(s, err))

  // ========== Convenience Constructors ==========
  /** Create column from 0-based index */
  def col(index: Int): Column = Column.from0(index)

  /** Create row from 0-based index */
  def row(index: Int): Row = Row.from0(index)

  /** Create cell reference from 0-based indices */
  def ref(col: Int, row: Int): ARef = ARef.from0(col, row)

export syntax.*
