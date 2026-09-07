package com.tjclp.xl.codec

import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.styles.CellStyle

/**
 * How one field of a record maps to one cell (GH-590): the per-field arm of a derived [[RowCodec]].
 *
 * Instances come from the field type's [[CellCodec]] — `required` for a plain field (an empty cell
 * is a [[RowCodecError.Missing]]), `optional` for an `Option[T]` field (an empty cell is `None`,
 * `None` writes an empty cell). To support a new field type, define a `given CellCodec[T]`; write a
 * `FieldCodec[T]` directly only when a field needs different emptiness semantics than its codec.
 */
trait FieldCodec[A]:
  /** Decode the cell aligned with `field`; errors name the cell through its ref. */
  def read(cell: Cell, field: String): Either[RowCodecError, A]

  /** Encode the field as a cell value plus the codec's number-format hint. */
  def write(a: A): (CellValue, Option[CellStyle])

object FieldCodec:
  def apply[A](using fc: FieldCodec[A]): FieldCodec[A] = fc

  given required[A](using codec: CellCodec[A]): FieldCodec[A] = new FieldCodec[A]:
    def read(cell: Cell, field: String): Either[RowCodecError, A] = codec.read(cell) match
      case Right(Some(a)) => Right(a)
      case Right(None) => Left(RowCodecError.Missing(cell.row, cell.col, field))
      case Left(cause) => Left(RowCodecError.Field(cell.row, cell.col, field, cause))
    def write(a: A): (CellValue, Option[CellStyle]) = codec.write(a)

  given optional[A](using codec: CellCodec[A]): FieldCodec[Option[A]] = new FieldCodec[Option[A]]:
    def read(cell: Cell, field: String): Either[RowCodecError, Option[A]] =
      codec.read(cell).left.map(cause => RowCodecError.Field(cell.row, cell.col, field, cause))
    def write(a: Option[A]): (CellValue, Option[CellStyle]) =
      a.fold[(CellValue, Option[CellStyle])]((CellValue.Empty, None))(codec.write)
