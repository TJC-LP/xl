package com.tjclp

/** Pure Scala 3.7 Excel (OOXML) Library
  *
  * Core module providing:
  * - Pure domain model (Cell, Sheet, Workbook)
  * - Zero-overhead opaque types (Column, Row, ARef, SheetName)
  * - Type-safe operations with total APIs
  *
  * For compile-time validated literals (cell"", range""), import from xl-macros module.
  *
  * Usage:
  * {{{
  * import com.tjclp.xl.*
  * import com.tjclp.xl.macros.{cell, range}  // From xl-macros module
  *
  * val ref = cell"A1"
  * val rng = range"A1:B10"
  *
  * val sheet = Sheet("MySheet").map { s =>
  *   s.put(ref, CellValue.Text("Hello"))
  *    .put(cell"B1", CellValue.Number(42))
  * }
  * }}}
  */
package object xl:
  // Extension methods for ergonomic API
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

  // Convenience constructors
  object syntax:
    /** Create column from 0-based index */
    def col(index: Int): Column = Column.from0(index)

    /** Create row from 0-based index */
    def row(index: Int): Row = Row.from0(index)

    /** Create cell reference from 0-based indices */
    def ref(col: Int, row: Int): ARef = ARef.from0(col, row)
