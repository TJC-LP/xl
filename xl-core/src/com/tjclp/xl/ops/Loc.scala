package com.tjclp.xl.ops

import com.tjclp.xl.addressing.{ARef, CellRange, Column, RefType, Row, SheetName}
import com.tjclp.xl.error.{XLError, XLResult}

/**
 * A cell an [[Edit]] targets: the sheet it names (a qualifier or a batch `sheet` key), or `None`
 * for "the scope's default" (ADR-017 §2.5, THE sheet rule).
 */
final case class Loc(sheet: Option[SheetName], ref: ARef) derives CanEqual:
  /** `B7` or `'Q1 Data'!B7` — the form [[Loc.parse]] reads back. */
  def toA1: String = sheet.fold(ref.toA1)(s => RefType.QualifiedCell(s, ref).toA1)

object Loc:
  /** A bare or sheet-qualified single cell through `RefType.parseToXLError`; a range is refused. */
  def parse(s: String): XLResult[Loc] =
    RefType.parseToXLError(s).flatMap {
      case RefType.Cell(ref) => Right(Loc(None, ref))
      case RefType.QualifiedCell(sheet, ref) => Right(Loc(Some(sheet), ref))
      case RefType.Range(_) | RefType.QualifiedRange(_, _) =>
        Left(XLError.InvalidReference(s"'$s' is a range; expected a single cell"))
    }

/** A range an [[Edit]] targets, with the same sheet convention as [[Loc]]. */
final case class Area(sheet: Option[SheetName], range: CellRange) derives CanEqual:
  /** `A1:B2` or `'Q1 Data'!A1:B2` — the form [[Area.parse]] reads back. */
  def toA1: String = sheet.fold(range.toA1)(s => RefType.QualifiedRange(s, range).toA1)

object Area:
  /** A bare or qualified range; a single cell is a 1x1 area (the batch convention). */
  def parse(s: String): XLResult[Area] =
    RefType.parseToXLError(s).map {
      case RefType.Cell(ref) => Area(None, CellRange(ref, ref))
      case RefType.QualifiedCell(sheet, ref) => Area(Some(sheet), CellRange(ref, ref))
      case RefType.Range(range) => Area(None, range)
      case RefType.QualifiedRange(sheet, range) => Area(Some(sheet), range)
    }

  /** The whole-sheet-less area of one cell. */
  def cell(loc: Loc): Area = Area(loc.sheet, CellRange(loc.ref, loc.ref))

/** An inclusive, normalized column span (`E:H`); the sheet comes from the edit that carries it. */
final case class ColSpan(start: Column, end: Column) derives CanEqual:
  def columns: Vector[Column] = (start.index0 to end.index0).map(Column.from0).toVector
  def size: Int = end.index0 - start.index0 + 1

  /** `E` for a single column, `E:H` otherwise. */
  def render: String =
    if start == end then start.toLetter else s"${start.toLetter}:${end.toLetter}"

object ColSpan:
  /** Normalizing constructor: `ColSpan(H, E)` is `E:H`. */
  def apply(a: Column, b: Column): ColSpan =
    if a.index0 <= b.index0 then new ColSpan(a, b) else new ColSpan(b, a)

  def single(col: Column): ColSpan = new ColSpan(col, col)

  /**
   * `E` or `E:H` (case-insensitive). A row span, a cell range or a sheet-qualified span is refused:
   * the sheet belongs to the edit, not the span.
   */
  def parse(s: String): XLResult[ColSpan] =
    val trimmed = s.trim
    if trimmed.contains(':') then
      // `E:H` is a full-column range, which RefType (cells and cell ranges) does not spell
      CellRange.parse(trimmed).toOption.filter(_.isFullColumn) match
        case Some(range) => Right(ColSpan(range.colStart, range.colEnd))
        case None =>
          Left(XLError.InvalidReference(s"'$s' is not a column span (expected E or E:H)"))
    else
      Column
        .fromLetter(trimmed)
        .map(single)
        .left
        .map(reason => XLError.InvalidReference(s"'$s' is not a column span ($reason)"))

/** An inclusive, normalized row span (`10:20`); the sheet comes from the edit that carries it. */
final case class RowSpan(start: Row, end: Row) derives CanEqual:
  def rows: Vector[Row] = (start.index0 to end.index0).map(Row.from0).toVector
  def size: Int = end.index0 - start.index0 + 1

  /** `10` for a single row, `10:20` otherwise. */
  def render: String =
    if start == end then start.index1.toString else s"${start.index1}:${end.index1}"

object RowSpan:
  /** Normalizing constructor: `RowSpan(20, 10)` is `10:20`. */
  def apply(a: Row, b: Row): RowSpan =
    if a.index0 <= b.index0 then new RowSpan(a, b) else new RowSpan(b, a)

  def single(row: Row): RowSpan = new RowSpan(row, row)

  /** `10` or `10:20` (1-based). A column span, a cell range or a qualified span is refused. */
  def parse(s: String): XLResult[RowSpan] =
    val trimmed = s.trim
    if trimmed.contains(':') then
      CellRange.parse(trimmed).toOption.filter(_.isFullRow) match
        case Some(range) => Right(RowSpan(range.rowStart, range.rowEnd))
        case None =>
          Left(XLError.InvalidReference(s"'$s' is not a row span (expected 10 or 10:20)"))
    else
      trimmed.toIntOption
        .filter(r => r >= 1 && r <= Row.MaxIndex0 + 1)
        .map(r => single(Row.from1(r)))
        .toRight(
          XLError.InvalidReference(
            s"'$s' is not a row span (expected a 1-based row 1-${Row.MaxIndex0 + 1} or 10:20)"
          )
        )
