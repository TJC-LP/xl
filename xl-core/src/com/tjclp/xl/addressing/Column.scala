package com.tjclp.xl.addressing

import java.util.Locale

/**
 * Column index with zero-based internal representation. Opaque type for zero-overhead wrapping.
 */
opaque type Column = Int

object Column:
  /** Maximum 0-based column index supported by Excel (A-XFD) */
  val MaxIndex0: Int = 16383

  /** Create a Column from 0-based index (0 = A, 1 = B, ...) */
  inline def from0(index: Int): Column = index

  /** Create a Column from 1-based index (1 = A, 2 = B, ...) */
  inline def from1(index: Int): Column = index - 1

  /** Create a Column from Excel letter notation (A, B, AA, etc.) */
  def fromLetter(input: String): Either[String, Column] =
    val normalized = input.toUpperCase(Locale.ROOT)
    if normalized.isEmpty then Left("Column letter cannot be empty")
    else if !normalized.forall(c => c >= 'A' && c <= 'Z') then
      Left(s"Invalid column letter: $input")
    else
      val index = normalized.foldLeft(0)((acc, c) => acc * 26 + (c - 'A' + 1)) - 1
      if index < 0 || index > MaxIndex0 then Left(s"Column index out of range: $index")
      else Right(index)

  extension (col: Column)
    /** Get 0-based index (0 = A, 1 = B, ...) */
    inline def index0: Int = col

    /** Get 1-based index (1 = A, 2 = B, ...) */
    inline def index1: Int = col + 1

    /** Convert to Excel letter notation (A, B, AA, etc.) */
    def toLetter: String =
      def loop(n: Int, acc: String): String =
        if n < 0 then acc
        else loop((n / 26) - 1, s"${((n % 26) + 'A').toChar}$acc")
      loop(col, "")

    /** Shift column by offset */
    inline def +(offset: Int): Column = col + offset

    /** Shift column by negative offset */
    inline def -(offset: Int): Column = col - offset

end Column
