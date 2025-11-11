package com.tjclp.xl.addressing

/** Cell range from start to end (inclusive) */
case class CellRange(start: ARef, end: ARef):
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
      )
    )

  /** Convert to A1:B2 notation */
  def toA1: String = s"${start.toA1}:${end.toA1}"

  /** Iterate over all cell references in range (row-major order) */
  def cells: Iterator[ARef] =
    for
      row <- (rowStart.index0 to rowEnd.index0).iterator
      col <- (colStart.index0 to colEnd.index0).iterator
    yield ARef.from0(col, row)

object CellRange:
  /** Create range from two references (order doesn't matter) */
  def apply(ref1: ARef, ref2: ARef): CellRange =
    val minCol = Column.from0(math.min(ref1.col.index0, ref2.col.index0))
    val maxCol = Column.from0(math.max(ref1.col.index0, ref2.col.index0))
    val minRow = Row.from0(math.min(ref1.row.index0, ref2.row.index0))
    val maxRow = Row.from0(math.max(ref1.row.index0, ref2.row.index0))
    new CellRange(ARef(minCol, minRow), ARef(maxCol, maxRow))

  /** Parse range from A1:B2 notation */
  def parse(s: String): Either[String, CellRange] =
    if s.contains(':') then
      s.split(':') match
        case Array(start, end) if start.nonEmpty && end.nonEmpty =>
          for
            startRef <- ARef.parse(start)
            endRef <- ARef.parse(end)
          yield CellRange(startRef, endRef)
        case _ =>
          Left(s"Invalid range format: $s")
    else if s.nonEmpty then ARef.parse(s).map(ref => CellRange(ref, ref))
    else Left(s"Invalid range format: $s")
