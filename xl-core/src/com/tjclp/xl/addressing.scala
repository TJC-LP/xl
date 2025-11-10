package com.tjclp.xl

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
  def fromLetter(s: String): Either[String, Column] =
    if s.isEmpty then Left("Column letter cannot be empty")
    else if !s.forall(c => c >= 'A' && c <= 'Z') then Left(s"Invalid column letter: $s")
    else
      val index = s.foldLeft(0)((acc, c) => acc * 26 + (c - 'A' + 1)) - 1
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

/**
 * Row index with zero-based internal representation. Opaque type for zero-overhead wrapping.
 */
opaque type Row = Int

object Row:
  /** Maximum 0-based row index supported by Excel (1-1,048,576) */
  val MaxIndex0: Int = 1048575

  /** Create a Row from 0-based index (0 = row 1, 1 = row 2, ...) */
  def from0(index: Int): Row = index

  /** Create a Row from 1-based index (1 = row 1, 2 = row 2, ...) */
  def from1(index: Int): Row = index - 1

  extension (row: Row)
    /** Get 0-based index */
    inline def index0: Int = row

    /** Get 1-based index */
    inline def index1: Int = row + 1

    /** Shift row by offset */
    inline def +(offset: Int): Row = row + offset

    /** Shift row by negative offset */
    inline def -(offset: Int): Row = row - offset

end Row

/** Opaque sheet name with validation */
opaque type SheetName = String

object SheetName:
  private val InvalidChars = Set(':', '\\', '/', '?', '*', '[', ']')
  private val MaxLength = 31

  /** Create a validated sheet name */
  def apply(name: String): Either[String, SheetName] =
    if name.isEmpty then Left("Sheet name cannot be empty")
    else if name.length > MaxLength then Left(s"Sheet name too long (max $MaxLength): $name")
    else if name.exists(InvalidChars.contains) then
      Left(s"Sheet name contains invalid characters: $name")
    else Right(name)

  /** Create an unsafe sheet name (use only when validation is guaranteed) */
  def unsafe(name: String): SheetName = name

  extension (name: SheetName) inline def value: String = name

end SheetName

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
    val (letters, digits) = s.span(c => c >= 'A' && c <= 'Z')
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
    inline def col: Column = Column.from0((ref & 0xffffffffL).toInt)

    /** Extract row */
    inline def row: Row = Row.from0((ref >> 32).toInt)

    /** Convert to A1 notation */
    def toA1: String =
      import Column.toLetter
      import Row.index1
      s"${toLetter(ref.col)}${index1(ref.row)}"

    /** Shift reference by column and row offsets */
    inline def shift(colOffset: Int, rowOffset: Int): ARef =
      ARef(ref.col + colOffset, ref.row + rowOffset)

end ARef

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
