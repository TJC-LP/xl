package com.tjclp.xl.addressing

/**
 * Cell range from start to end (inclusive).
 *
 * @param start
 *   Top-left cell reference
 * @param end
 *   Bottom-right cell reference
 * @param startAnchor
 *   Anchor mode for start reference (for formula dragging)
 * @param endAnchor
 *   Anchor mode for end reference (for formula dragging)
 */
final case class CellRange(
  start: ARef,
  end: ARef,
  startAnchor: Anchor = Anchor.Relative,
  endAnchor: Anchor = Anchor.Relative
):
  /** Top-left column */
  inline def colStart: Column = start.col

  /** Top-left row */
  inline def rowStart: Row = start.row

  /** Bottom-right column */
  inline def colEnd: Column = end.col

  /** Bottom-right row */
  inline def rowEnd: Row = end.row

  /** Number of columns in range */
  inline def width: Int = colEnd.index0 - colStart.index0 + 1

  /** Number of rows in range */
  inline def height: Int = rowEnd.index0 - rowStart.index0 + 1

  /** Total number of cells */
  inline def size: Int = width * height

  /** Check if reference is within this range */
  inline def contains(ref: ARef): Boolean =
    ref.col.index0 >= colStart.index0 &&
      ref.col.index0 <= colEnd.index0 &&
      ref.row.index0 >= rowStart.index0 &&
      ref.row.index0 <= rowEnd.index0

  /** Check if this range intersects with another */
  def intersects(other: CellRange): Boolean =
    !(colEnd.index0 < other.colStart.index0 ||
      colStart.index0 > other.colEnd.index0 ||
      rowEnd.index0 < other.rowStart.index0 ||
      rowStart.index0 > other.rowEnd.index0)

  /** Expand range to include a reference */
  def expand(ref: ARef): CellRange =
    CellRange(
      ARef(
        Column.from0(math.min(colStart.index0, ref.col.index0)),
        Row.from0(math.min(rowStart.index0, ref.row.index0))
      ),
      ARef(
        Column.from0(math.max(colEnd.index0, ref.col.index0)),
        Row.from0(math.max(rowEnd.index0, ref.row.index0))
      ),
      startAnchor,
      endAnchor
    )

  /** Convert to A1:B2 notation (without anchors) */
  def toA1: String = s"${start.toA1}:${end.toA1}"

  /** Convert to A1:B2 notation with anchors (e.g., $A$1:B10) */
  def toA1Anchored: String =
    s"${formatRefWithAnchor(start, startAnchor)}:${formatRefWithAnchor(end, endAnchor)}"

  private def formatRefWithAnchor(ref: ARef, anchor: Anchor): String =
    val colStr = if anchor.isColAbsolute then s"$$${ref.col.toLetter}" else ref.col.toLetter
    val rowStr = if anchor.isRowAbsolute then s"$$${ref.row.index1}" else ref.row.index1.toString
    colStr + rowStr

  /** Iterate over all cell references in range (row-major order) */
  def cells: Iterator[ARef] =
    for
      row <- (rowStart.index0 to rowEnd.index0).iterator
      col <- (colStart.index0 to colEnd.index0).iterator
    yield ARef.from0(col, row)

object CellRange:
  /**
   * Create range from two references (order doesn't matter).
   *
   * Normalizes so start is top-left and end is bottom-right. Anchor modes default to Relative.
   */
  def apply(ref1: ARef, ref2: ARef): CellRange =
    val minCol = Column.from0(math.min(ref1.col.index0, ref2.col.index0))
    val maxCol = Column.from0(math.max(ref1.col.index0, ref2.col.index0))
    val minRow = Row.from0(math.min(ref1.row.index0, ref2.row.index0))
    val maxRow = Row.from0(math.max(ref1.row.index0, ref2.row.index0))
    new CellRange(ARef(minCol, minRow), ARef(maxCol, maxRow), Anchor.Relative, Anchor.Relative)

  /**
   * Create range from two references with explicit anchors.
   *
   * Note: Anchors are preserved even when refs are swapped for normalization.
   */
  def apply(
    ref1: ARef,
    ref2: ARef,
    startAnchor: Anchor,
    endAnchor: Anchor
  ): CellRange =
    val minCol = Column.from0(math.min(ref1.col.index0, ref2.col.index0))
    val maxCol = Column.from0(math.max(ref1.col.index0, ref2.col.index0))
    val minRow = Row.from0(math.min(ref1.row.index0, ref2.row.index0))
    val maxRow = Row.from0(math.max(ref1.row.index0, ref2.row.index0))
    // Determine which ref became start vs end after normalization
    val ref1IsStart = ref1.col.index0 == minCol.index0 && ref1.row.index0 == minRow.index0
    val (finalStartAnchor, finalEndAnchor) =
      if ref1IsStart then (startAnchor, endAnchor) else (endAnchor, startAnchor)
    new CellRange(ARef(minCol, minRow), ARef(maxCol, maxRow), finalStartAnchor, finalEndAnchor)

  /** Check if string is a column-only reference (letters only, no digits) */
  private def isColumnOnly(s: String): Boolean =
    s.nonEmpty && s.forall(c => c.isLetter)

  /** Check if string is a row-only reference (digits only, no letters) */
  private def isRowOnly(s: String): Boolean =
    s.nonEmpty && s.forall(c => c.isDigit)

  /**
   * Parse full column range like A:C or $A:$C. Returns range spanning all rows (0 to MaxIndex0) for
   * the specified columns.
   */
  private def parseFullColumnRange(
    startStr: String,
    endStr: String,
    startAnchor: Anchor,
    endAnchor: Anchor
  ): Either[String, CellRange] =
    for
      startCol <- Column.fromLetter(startStr)
      endCol <- Column.fromLetter(endStr)
    yield
      val minCol = Column.from0(math.min(startCol.index0, endCol.index0))
      val maxCol = Column.from0(math.max(startCol.index0, endCol.index0))
      new CellRange(
        ARef(minCol, Row.from0(0)),
        ARef(maxCol, Row.from0(Row.MaxIndex0)),
        startAnchor,
        endAnchor
      )

  /**
   * Parse full row range like 1:5 or $1:$5. Returns range spanning all columns (0 to MaxIndex0) for
   * the specified rows.
   */
  private def parseFullRowRange(
    startStr: String,
    endStr: String,
    startAnchor: Anchor,
    endAnchor: Anchor
  ): Either[String, CellRange] =
    for
      startRow <- startStr.toIntOption
        .filter(r => r >= 1 && r <= Row.MaxIndex0 + 1)
        .toRight(s"Invalid row number: $startStr")
      endRow <- endStr.toIntOption
        .filter(r => r >= 1 && r <= Row.MaxIndex0 + 1)
        .toRight(s"Invalid row number: $endStr")
    yield
      val minRow = Row.from1(math.min(startRow, endRow))
      val maxRow = Row.from1(math.max(startRow, endRow))
      new CellRange(
        ARef(Column.from0(0), minRow),
        ARef(Column.from0(Column.MaxIndex0), maxRow),
        startAnchor,
        endAnchor
      )

  /**
   * Parse range from A1:B2 notation, preserving anchors.
   *
   * Supports:
   *   - Standard ranges: A1:B10, $A$1:B10
   *   - Full column references: A:A, A:C, $A:$C (all rows of specified columns)
   *   - Full row references: 1:1, 1:5, $1:$5 (all columns of specified rows)
   */
  def parse(s: String): Either[String, CellRange] =
    if s.contains(':') then
      s.split(':') match
        case Array(startStr, endStr) if startStr.nonEmpty && endStr.nonEmpty =>
          val (cleanStart, startAnchor) = Anchor.parse(startStr)
          val (cleanEnd, endAnchor) = Anchor.parse(endStr)

          // Check for full column reference (A:C) - letters only
          if isColumnOnly(cleanStart) && isColumnOnly(cleanEnd) then
            parseFullColumnRange(cleanStart, cleanEnd, startAnchor, endAnchor)
          // Check for full row reference (1:5) - digits only
          else if isRowOnly(cleanStart) && isRowOnly(cleanEnd) then
            parseFullRowRange(cleanStart, cleanEnd, startAnchor, endAnchor)
          // Standard cell range (A1:B10)
          else
            for
              startRef <- ARef.parse(cleanStart)
              endRef <- ARef.parse(cleanEnd)
            yield CellRange(startRef, endRef, startAnchor, endAnchor)
        case _ =>
          Left(s"Invalid range format: $s")
    else if s.nonEmpty then
      val (clean, anchor) = Anchor.parse(s)
      ARef.parse(clean).map(ref => new CellRange(ref, ref, anchor, anchor))
    else Left(s"Invalid range format: $s")
