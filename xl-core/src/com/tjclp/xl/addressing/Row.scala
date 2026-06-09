package com.tjclp.xl.addressing

/**
 * Row index with zero-based internal representation. Opaque type for zero-overhead wrapping.
 */
opaque type Row = Int

object Row:
  /** Maximum 0-based row index supported by Excel (1-1,048,576) */
  val MaxIndex0: Int = 1048575

  /** Create a Row from 0-based index (0 = row 1, 1 = row 2, ...) */
  inline def from0(index: Int): Row = index

  /** Create a Row from 1-based index (1 = row 1, 2 = row 2, ...) */
  inline def from1(index: Int): Row = index - 1

  extension (row: Row)
    /** Get 0-based index */
    def index0: Int = row

    /** Get 1-based index */
    def index1: Int = row + 1

    /** Shift row by offset */
    def +(offset: Int): Row = row + offset

    /** Shift row by negative offset */
    def -(offset: Int): Row = row - offset

end Row
