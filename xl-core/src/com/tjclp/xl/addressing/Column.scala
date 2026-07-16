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
  def from0(index: Int): Column = index

  /** Create a Column from 1-based index (1 = A, 2 = B, ...) */
  def from1(index: Int): Column = index - 1

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

  /**
   * Parse a column from Excel letter notation at runtime (`"D"`, `"aa"` — case-insensitive).
   *
   * Runtime counterpart of the compile-time `ref` literal for code that computes column letters
   * dynamically (e.g. folding over `Vector("C", "D")` to set widths). Trailing row digits are
   * tolerated and discarded (`"D1"` parses as column D) so full A1 refs can be passed directly; any
   * other suffix (anchors, ranges, mixed letters) is rejected. Strict letters-only parsing is
   * [[fromLetter]].
   */
  def parse(input: String): Either[String, Column] =
    val normalized = input.toUpperCase(Locale.ROOT)
    val (letters, rest) = normalized.span(c => c >= 'A' && c <= 'Z')
    if letters.isEmpty then Left(s"No column letters in: $input")
    else if !rest.forall(c => c >= '0' && c <= '9') then Left(s"Invalid column: $input")
    else fromLetter(letters)

  extension (col: Column)
    /** Get 0-based index (0 = A, 1 = B, ...) */
    def index0: Int = col

    /** Get 1-based index (1 = A, 2 = B, ...) */
    def index1: Int = col + 1

    /** Convert to Excel letter notation (A, B, AA, etc.) */
    def toLetter: String =
      def loop(n: Int, acc: String): String =
        if n < 0 then acc
        else loop((n / 26) - 1, s"${((n % 26) + 'A').toChar}$acc")
      loop(col, "")

    /** Shift column by offset */
    def +(offset: Int): Column = col + offset

    /** Shift column by negative offset */
    def -(offset: Int): Column = col - offset

end Column
