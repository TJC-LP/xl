package com.tjclp.xl.codec

import scala.annotation.tailrec
import scala.compiletime.{constValueTuple, erasedValue, error, summonInline}
import scala.deriving.Mirror

import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.styles.CellStyle

/**
 * One record ↔ one row of cells (GH-590).
 *
 * `fields(i)` is the name of the field stored in the i-th cell of a row: the default header text
 * and the key `readRowsByHeader` matches against. Derive an instance for any case class whose
 * fields have a [[CellCodec]] (`Option[T]` fields are empty cells):
 *
 * {{{
 * final case class Order(id: Int, customer: String, qty: Int, price: BigDecimal, note: Option[String])
 *   derives RowCodec
 *
 * sheet.putRowsWithHeader(ref"A1", orders)   // header row + one row per record
 * sheet.readRowsByHeader[Order](Row.from1(1)) // Either[RowCodecError, Vector[Order]]
 * }}}
 *
 * Reads go through each cell's [[Cell.effectiveValue]] (GH-477): a recalculated formula decodes as
 * its cached value, an uncached one is a [[RowCodecError.Field]] naming the formula. Hand-written
 * instances are ordinary: implement the three members, keeping `write(a).size == fields.size`.
 */
trait RowCodec[A]:
  /** Field names in declaration order — the header row and the width of a record. */
  def fields: Vector[String]

  /** Cells a record spans. */
  def width: Int = fields.size

  /**
   * Decode one record. `cells(i)` is the cell aligned with `fields(i)` and carries its own ref, so
   * errors name the row and column. `cells.size != width` is a [[RowCodecError.Width]].
   */
  def read(cells: Vector[Cell]): Either[RowCodecError, A]

  /** Encode one record: `width` cell payloads aligned with `fields`, each with its format hint. */
  def write(a: A): Vector[(CellValue, Option[CellStyle])]

object RowCodec:
  def apply[A](using rc: RowCodec[A]): RowCodec[A] = rc

  /**
   * Derive a codec for a case class from its `Mirror`: one [[FieldCodec]] per field, resolved at
   * the `derives` site. Only the field-codec summons and the label read are inline — the codec
   * itself is a plain instance built once, so a derived `given` costs one allocation per class.
   */
  inline def derived[A <: Product](using m: Mirror.ProductOf[A]): RowCodec[A] =
    inline erasedValue[m.MirroredElemTypes] match
      case _: EmptyTuple => error("RowCodec.derived: a record needs at least one field")
      case _ =>
        val labels = constValueTuple[m.MirroredElemLabels].toList.map(_.toString).toVector
        build[A](labels, fieldCodecs[m.MirroredElemTypes].toVector)

  private inline def fieldCodecs[Elems <: Tuple]: List[FieldCodec[?]] =
    inline erasedValue[Elems] match
      case _: EmptyTuple => Nil
      case _: (h *: t) => summonInline[FieldCodec[h]] :: fieldCodecs[t]

  /** A field codec with its type erased, so a row can be decoded by index into a `Product`. */
  private final class Arm(
    val read: (Cell, String) => Either[RowCodecError, Any],
    val write: Any => (CellValue, Option[CellStyle])
  )

  // Mirror derivation hands fields back through Product's untyped iterator; the cast restores
  // the type the FieldCodec was summoned for.
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def erase[T](fc: FieldCodec[T]): Arm =
    Arm(fc.read, value => fc.write(value.asInstanceOf[T]))

  private def build[A <: Product](labels: Vector[String], codecs: Vector[FieldCodec[?]])(using
    m: Mirror.ProductOf[A]
  ): RowCodec[A] =
    val arms: Vector[Arm] = codecs.map(fc => erase(fc))
    new RowCodec[A]:
      val fields: Vector[String] = labels

      def read(cells: Vector[Cell]): Either[RowCodecError, A] =
        if cells.size != arms.size then Left(RowCodecError.Width(arms.size, cells.size))
        else
          @tailrec def loop(i: Int, acc: List[Any]): Either[RowCodecError, A] =
            if i == arms.size then Right(m.fromProduct(Tuple.fromArray(acc.reverse.toArray)))
            else
              arms(i).read(cells(i), labels(i)) match
                case Right(value) => loop(i + 1, value :: acc)
                case Left(err) => Left(err)
          loop(0, Nil)

      def write(a: A): Vector[(CellValue, Option[CellStyle])] =
        arms.zip(a.productIterator.toVector).map((arm, value) => arm.write(value))
