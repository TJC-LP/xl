package com.tjclp.xl.addressing

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
  def apply(col: Column, row: Row): ARef =
    (rowIndex(row).toLong << 32) | (colIndex(col).toLong & 0xffffffffL)

  /** Create cell reference from 0-based indices */
  def from0(colIndex: Int, rowIndex: Int): ARef =
    apply(Column.from0(colIndex), Row.from0(rowIndex))

  /** Create cell reference from 1-based indices */
  def from1(colIndex: Int, rowIndex: Int): ARef =
    apply(Column.from1(colIndex), Row.from1(rowIndex))

  /** Parse cell reference from A1 notation */
  def parse(s: String): Either[String, ARef] =
    val normalized = AsciiCase.upper(s)
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
      // NOTE: every member here (and in Column/Row/SheetName/style units) that touches the
      // opaque representation must stay NON-inline: inline bodies fail to re-elaborate at call
      // sites outside this package, where the representation is hidden (issue #252). Do not
      // "optimize" these back to `inline` — the JIT/AOT inlines trivial static methods anyway.
      import Column.toLetter
      import Row.index1
      s"${toLetter(ref.col)}${index1(ref.row)}"

    /** Shift reference by column and row offsets */
    def shift(colOffset: Int, rowOffset: Int): ARef =
      ARef(ref.col + colOffset, ref.row + rowOffset)

    // ----- Bounded navigation (GH-465) -----
    // `shift` (and the DSL's down/up/left/right) are total but unchecked: they mint "A0" or
    // column -1 past the grid edge, which corrupts output if written. The forms below make the
    // edge explicit — `None` past it, or a clamp onto it. Offsets may be any Int; the arithmetic
    // is done in Long so extreme offsets cannot wrap around into the grid.

    /**
     * Bounded shift: `Some` of the shifted reference, or `None` when the target would leave the
     * grid (column 0..16383 = A..XFD, row 0..1048575 = 1..1048576). Agrees with [[shift]] whenever
     * it is `Some`; `tryShift(dc, dr).flatMap(_.tryShift(-dc, -dr))` is `Some(ref)` in bounds.
     */
    def tryShift(colOffset: Int, rowOffset: Int): Option[ARef] =
      val c = colIndex(ref.col).toLong + colOffset
      val r = rowIndex(ref.row).toLong + rowOffset
      if c < 0L || c > Column.MaxIndex0.toLong || r < 0L || r > Row.MaxIndex0.toLong then None
      else Some(from0(c.toInt, r.toInt))

    /** Bounded step along rows: `None` past row 1048576 (or above row 1 for a negative `n`). */
    def tryDown(n: Int): Option[ARef] = ref.tryShift(0, n)

    /** Bounded step along columns: `None` past column XFD (or left of A for a negative `n`). */
    def tryRight(n: Int): Option[ARef] = ref.tryShift(n, 0)

    /**
     * Shift, pinning each axis independently to the nearest grid edge instead of overrunning it:
     * `ref"C3".clampShift(-10, 5)` is `A8` (column pinned to A, row shifted). Agrees with [[shift]]
     * whenever [[tryShift]] is `Some`; once both axes are pinned, re-applying the same offsets is a
     * no-op.
     */
    def clampShift(colOffset: Int, rowOffset: Int): ARef =
      val c = math.max(0L, math.min(colIndex(ref.col).toLong + colOffset, Column.MaxIndex0.toLong))
      val r = math.max(0L, math.min(rowIndex(ref.row).toLong + rowOffset, Row.MaxIndex0.toLong))
      from0(c.toInt, r.toInt)

end ARef
