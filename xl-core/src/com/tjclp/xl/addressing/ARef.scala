package com.tjclp.xl.addressing

import java.util.Locale

/**
 * Absolute cell reference with 64-bit packed representation. Upper 32 bits: row index Lower 32
 * bits: column index
 *
 * This allows efficient storage and comparison.
 */
opaque type ARef = Long

object ARef:
  import Column.index0 as colIndex
  import Row.index0 as rowIndex

  /** Create cell reference from column and row */
  inline def apply(col: Column, row: Row): ARef =
    (rowIndex(row).toLong << 32) | (colIndex(col).toLong & 0xffffffffL)

  /** Create cell reference from 0-based indices */
  inline def from0(colIndex: Int, rowIndex: Int): ARef =
    apply(Column.from0(colIndex), Row.from0(rowIndex))

  /** Create cell reference from 1-based indices */
  inline def from1(colIndex: Int, rowIndex: Int): ARef =
    apply(Column.from1(colIndex), Row.from1(rowIndex))

  /** Parse cell reference from A1 notation */
  def parse(s: String): Either[String, ARef] =
    val normalized = s.toUpperCase(Locale.ROOT)
    val (letters, digits) = normalized.span(c => c >= 'A' && c <= 'Z')
    if letters.isEmpty then Left(s"No column letters in: $s")
    else if digits.isEmpty then Left(s"No row digits in: $s")
    else
      for
        col <- Column.fromLetter(letters)
        rowNum <- digits.toIntOption.toRight(s"Invalid row number: $digits")
        _ <- Either.cond(
          rowNum >= 1 && rowNum <= Row.MaxIndex0 + 1,
          (),
          s"Row out of range: $rowNum"
        )
        row = Row.from1(rowNum)
      yield apply(col, row)

  extension (ref: ARef)
    /** Extract column */
    def col: Column = Column.from0((ref & 0xffffffffL).toInt)

    /** Extract row */
    def row: Row = Row.from0((ref >> 32).toInt)

    /** Convert to A1 notation */
    def toA1: String =
      import Column.toLetter
      import Row.index1
      s"${toLetter(ref.col)}${index1(ref.row)}"

    /** Shift reference by column and row offsets */
    def shift(colOffset: Int, rowOffset: Int): ARef =
      ARef(ref.col + colOffset, ref.row + rowOffset)

end ARef
