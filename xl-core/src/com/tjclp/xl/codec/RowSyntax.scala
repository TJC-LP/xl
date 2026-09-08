package com.tjclp.xl.codec

import java.util.Locale

import scala.annotation.tailrec

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.{CellStyle, StyleRegistry}
import com.tjclp.xl.tables.TableSpec

/**
 * Where a block of records landed (GH-590): the result of `putRows`, `putRowsWithHeader` and
 * `putTable`.
 *
 * @param sheet
 *   the updated sheet
 * @param headerRange
 *   the header row, when one was written
 * @param dataRange
 *   the records — one row each, [[RowCodec.width]] columns — or `None` when no record was given
 */
final case class RowsPlaced(
  sheet: Sheet,
  headerRange: Option[CellRange],
  dataRange: Option[CellRange]
) derives CanEqual:

  /** Header and records together; `None` only when nothing at all was written. */
  def range: Option[CellRange] = (headerRange, dataRange) match
    case (Some(header), Some(data)) => Some(CellRange(header.start, data.end))
    case (header, data) => header.orElse(data)

  /** Records written. */
  def count: Int = dataRange.fold(0)(_.height)

/**
 * Records as rows (GH-590): read and write case classes through a [[RowCodec]].
 *
 * Positional entry points (`readRows`, `putRows`) align `fields(i)` with the i-th column of the
 * range; header-driven ones (`readRowsByHeader`, `putRowsWithHeader`, `putTable`) go through a
 * header row of field names, matched by [[rowSyntax.column]]. Reads see a formula's cached value
 * (GH-477).
 */
object rowSyntax:
  extension (sheet: Sheet)

    /**
     * Decode one record per row of `range`, positionally: column i of the range holds `fields(i)`.
     * The range must be exactly [[RowCodec.width]] columns wide ([[RowCodecError.Width]]
     * otherwise); every row in it is a record, so a blank row is a [[RowCodecError.Missing]] unless
     * every field is an `Option`. Stops at the first failing cell, row-major.
     */
    def readRows[A](range: CellRange)(using codec: RowCodec[A]): Either[RowCodecError, Vector[A]] =
      if range.width != codec.width then Left(RowCodecError.Width(codec.width, range.width))
      else
        val columns = Vector.tabulate(codec.width)(i => range.start.col + i)
        decodeRows(sheet, columns, range.start.row.index0, range.end.row.index0, codec)

    /**
     * Decode the records under `headerRow`: each field is read from the column whose header matches
     * its name ([[column]]), so column order is free and extra columns are ignored. Reads the
     * contiguous block below the header — Excel's current region — and stops at the first row whose
     * record cells are all empty; `Right(Vector.empty)` when nothing follows the header. A field
     * with no header is a [[RowCodecError.HeaderNotFound]] listing the headers present.
     */
    def readRowsByHeader[A](headerRow: Row)(using
      codec: RowCodec[A]
    ): Either[RowCodecError, Vector[A]] =
      val present = headers(headerRow)
      resolveColumns(codec.fields, present, headerRow).flatMap { columns =>
        val firstRow = headerRow.index0 + 1
        val colSet = columns.toSet
        val lastRow = sheet.cells.valuesIterator
          .filter(c => c.row.index0 >= firstRow && colSet.contains(c.col) && c.nonEmpty)
          .map(_.row.index0)
          .maxOption
        lastRow match
          case None => Right(Vector.empty)
          case Some(last) =>
            val endRow = (firstRow to last)
              .find { r =>
                columns.forall(c => sheet(ARef(c, Row.from0(r))).effectiveValue == CellValue.Empty)
              }
              .map(_ - 1)
              .getOrElse(last)
            decodeRows(sheet, columns, firstRow, endRow, codec)
      }

    /**
     * Header texts in `row`, left to right, each with its column. A header is the cell's
     * [[Cell.effectiveValue]] as text (numbers and rich text included), verbatim; blank cells,
     * errors and uncached formulas are skipped.
     */
    def headers(row: Row): Vector[(Column, String)] =
      sheet.cells.valuesIterator
        .filter(_.row.index0 == row.index0)
        .flatMap(c => headerText(c.effectiveValue).map(text => (c.col, text)))
        .toVector
        .sortBy(_._1.index0)

    /**
     * The column whose header in `headerRow` is `header`: an exact match wins, otherwise the match
     * ignoring case, whitespace, `_` and `-` (`"Order ID"`, `order_id` and `orderId` all agree);
     * the leftmost of several. `None` when no header agrees.
     */
    def column(header: String, headerRow: Row): Option[Column] =
      columnIn(headers(headerRow), header)

    /**
     * Write one row per record starting at `at` (no header): `fields(i)` goes to column `at.col +
     * i`. Codec format hints register as styles and merge into any existing cell style the way
     * `put` does; a `None` field leaves its cell empty (an existing cell's value is cleared, none
     * is created). Only the records' cells are touched: rows below a previous, longer block
     * survive, so clear the old block first when regenerating a table in place. `OutOfBounds` when
     * the block would run past column XFD or row 1048576.
     */
    def putRows[A](at: ARef, rows: Iterable[A])(using codec: RowCodec[A]): XLResult[RowsPlaced] =
      place(sheet, at, rows, header = false, codec)

    /** [[putRows]] with the field names written as a header row at `at`; records start below. */
    def putRowsWithHeader[A](at: ARef, rows: Iterable[A])(using
      codec: RowCodec[A]
    ): XLResult[RowsPlaced] =
      place(sheet, at, rows, header = true, codec)

    /**
     * [[putRowsWithHeader]] plus an Excel table named `name` over header and records, its columns
     * named after the fields. With no records the table keeps the one blank data row Excel itself
     * insists on. `name` follows Excel's rules (letters, digits, `_`; unique per workbook — the
     * sheet-level check rejects a name this sheet already uses) and doubles as the display name.
     */
    def putTable[A](at: ARef, rows: Iterable[A], name: String)(using
      codec: RowCodec[A]
    ): XLResult[RowsPlaced] =
      if sheet.hasTable(name) then
        Left(
          XLError.InvalidTableName(
            name,
            s"sheet '${sheet.name.value}' already has a table named '$name'"
          )
        )
      else
        val records = rows.toVector
        val tableRows = 1 + math.max(1, records.size)
        for
          _ <- checkBounds(at, codec.width, tableRows)
          placed <- place(sheet, at, records, header = true, codec)
          tableRange = CellRange(at, ARef(at.col + codec.width - 1, at.row + tableRows - 1))
          spec <- TableSpec.fromColumnNames(name, name, tableRange, codec.fields)
        yield placed.copy(sheet = placed.sheet.withTable(spec))

  // ========== Internals ==========

  private def decodeRows[A](
    sheet: Sheet,
    columns: Vector[Column],
    firstRow: Int,
    lastRow: Int,
    codec: RowCodec[A]
  ): Either[RowCodecError, Vector[A]] =
    @tailrec def loop(r: Int, acc: Vector[A]): Either[RowCodecError, Vector[A]] =
      if r > lastRow then Right(acc)
      else
        val row = Row.from0(r)
        codec.read(columns.map(c => sheet(ARef(c, row)))) match
          case Right(record) => loop(r + 1, acc :+ record)
          case Left(err) => Left(err)
    loop(firstRow, Vector.empty)

  private def resolveColumns(
    fields: Vector[String],
    present: Vector[(Column, String)],
    headerRow: Row
  ): Either[RowCodecError, Vector[Column]] =
    @tailrec def loop(i: Int, acc: Vector[Column]): Either[RowCodecError, Vector[Column]] =
      if i == fields.size then Right(acc)
      else
        columnIn(present, fields(i)) match
          case Some(col) => loop(i + 1, acc :+ col)
          case None => Left(RowCodecError.HeaderNotFound(fields(i), headerRow, present.map(_._2)))
    loop(0, Vector.empty)

  private def columnIn(present: Vector[(Column, String)], header: String): Option[Column] =
    present.collectFirst { case (col, text) if text == header => col }.orElse {
      val key = normalize(header)
      present.collectFirst { case (col, text) if normalize(text) == key => col }
    }

  private def normalize(header: String): String =
    header.filterNot(c => c.isWhitespace || c == '_' || c == '-').toLowerCase(Locale.ROOT)

  private def headerText(value: CellValue): Option[String] = value match
    case CellValue.Text(s) if s.trim.nonEmpty => Some(s)
    case CellValue.RichText(rt) =>
      val plain = rt.toPlainText
      Option.when(plain.trim.nonEmpty)(plain)
    case CellValue.Number(n) => Some(n.bigDecimal.stripTrailingZeros.toPlainString)
    case CellValue.Bool(b) => Some(b.toString)
    case CellValue.DateTime(dt) => Some(dt.toString)
    case _ => None

  private def checkBounds(at: ARef, width: Int, rows: Int): XLResult[Unit] =
    if width == 0 then Left(XLError.ValueCountMismatch(1, 0, "RowCodec.fields"))
    else if at.col.index0 + width - 1 > Column.MaxIndex0 then
      Left(
        XLError.OutOfBounds(
          at.toA1,
          s"a $width-column record starting at ${at.col.toLetter} runs past column ${Column.from0(Column.MaxIndex0).toLetter}"
        )
      )
    else if at.row.index0 + rows - 1 > Row.MaxIndex0 then
      Left(
        XLError.OutOfBounds(
          at.toA1,
          s"$rows rows starting at row ${at.row.index1} run past row ${Row.MaxIndex0 + 1}"
        )
      )
    else Right(())

  // One pass over the records (GH-297's discipline for `put`): each record is encoded, its width
  // checked and its cells written before the next is touched, so no block-sized intermediate is
  // materialised. A width drift is still all-or-nothing: the caller's sheet is immutable and the
  // partially built map is simply dropped with the `Left`.
  private def place[A](
    sheet: Sheet,
    at: ARef,
    rows: Iterable[A],
    header: Boolean,
    codec: RowCodec[A]
  ): XLResult[RowsPlaced] =
    val width = codec.width
    val count = rows.size
    val headerRows = if header then 1 else 0
    checkBounds(at, width, headerRows + count).flatMap { _ =>
      val start: WriteState = (sheet.cells, sheet.styleRegistry)
      val afterHeader =
        if header then
          codec.fields.zipWithIndex.foldLeft(start) { case (state, (field, j)) =>
            writeCell(state, ARef(at.col + j, at.row), CellValue.Text(field), None)
          }
        else start
      val records = rows.iterator
      @tailrec def loop(i: Int, state: WriteState): XLResult[WriteState] =
        if !records.hasNext then Right(state)
        else
          val payloads = codec.write(records.next())
          if payloads.size != width then
            Left(XLError.ValueCountMismatch(width, payloads.size, s"RowCodec.write, record $i"))
          else loop(i + 1, writeRecord(state, at.col, at.row + headerRows + i, payloads))
      loop(0, afterHeader).map { (cells, registry) =>
        val lastCol = at.col + width - 1
        RowsPlaced(
          sheet = sheet.copy(cells = cells, styleRegistry = registry),
          headerRange = Option.when(header)(CellRange(at, ARef(lastCol, at.row))),
          dataRange = Option.when(count > 0)(
            CellRange(
              ARef(at.col, at.row + headerRows),
              ARef(lastCol, at.row + headerRows + count - 1)
            )
          )
        )
      }
    }

  private type WriteState = (Map[ARef, Cell], StyleRegistry)

  private def writeRecord(
    state: WriteState,
    firstCol: Column,
    row: Row,
    payloads: Vector[(CellValue, Option[CellStyle])]
  ): WriteState =
    @tailrec def loop(j: Int, acc: WriteState): WriteState =
      if j == payloads.size then acc
      else
        val (value, hint) = payloads(j)
        loop(j + 1, writeCell(acc, ARef(firstCol + j, row), value, hint))
    loop(0, state)

  // The single-cell `put` semantics, fused over a block (GH-297): the value replaces the
  // existing one, the codec's format hint merges into the existing style through the one policy
  // `put` uses (`Sheet.mergeStyles`), styles register once per distinct style, and an empty value
  // creates no cell.
  private def writeCell(
    state: WriteState,
    ref: ARef,
    value: CellValue,
    hint: Option[CellStyle]
  ): WriteState =
    val (cells, registry) = state
    val existing = cells.get(ref)
    (existing, value, hint) match
      case (None, CellValue.Empty, None) => state
      case _ =>
        val cell = existing.fold(Cell(ref, value))(_.withValue(value))
        hint match
          case None => (cells.updated(ref, cell), registry)
          case Some(codecStyle) =>
            val merged = existing.flatMap(_.styleId).flatMap(registry.get) match
              case Some(current) => Sheet.mergeStyles(current, codecStyle)
              case None => codecStyle
            val (nextRegistry, styleId) = registry.register(merged)
            (cells.updated(ref, cell.withStyle(styleId)), nextRegistry)
