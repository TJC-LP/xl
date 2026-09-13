package com.tjclp.xl.codec

import scala.annotation.tailrec
import scala.compiletime.{constValueTuple, erasedValue, error, summonInline}
import scala.deriving.Mirror

import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.styles.CellStyle

/**
 * One record ↔ one row of cells (GH-590).
 *
 * `fields(i)` is the name of the field stored in the i-th cell of a row — the name every
 * [[RowCodecError]] carries — and `headers(i)` is the text above that column: what
 * `putRowsWithHeader`/`putTable` write and `readRowsByHeader` matches. The two agree unless a field
 * is annotated `@header` or the codec went through [[RowCodec.withHeaders]] (GH-614). Derive an
 * instance for any case class whose fields have a [[CellCodec]] (`Option[T]` fields are empty
 * cells):
 *
 * {{{
 * final case class Order(id: Int, customer: String, qty: Int, price: BigDecimal, note: Option[String])
 *   derives RowCodec
 *
 * sheet.putRowsWithHeader(ref"A1", orders)   // header row + one row per record
 * sheet.readRowsByHeader[Order](Row.from1(1)) // Either[RowCodecError, Vector[Order]]
 *
 * final case class Deal(@header("Portfolio Co.") portfolioCo: String, @header("Rev ($M)") rev: BigDecimal)
 *   derives RowCodec                          // headers no identifier can spell
 * }}}
 *
 * Reads go through each cell's [[Cell.effectiveValue]] (GH-477): a recalculated formula decodes as
 * its cached value, an uncached one is a [[RowCodecError.Field]] naming the formula. Hand-written
 * instances are ordinary: implement the three abstract members, keeping `write(a).size ==
 * fields.size`; `headers` defaults to `fields`.
 */
trait RowCodec[A]:
  /** Field names in declaration order — the width of a record and the names errors carry. */
  def fields: Vector[String]

  /**
   * Header text per field, aligned with `fields`: written by `putRowsWithHeader` and `putTable`,
   * matched by `readRowsByHeader`. The field names, unless `@header` or [[withHeaders]] said
   * otherwise (GH-614).
   */
  def headers: Vector[String] = fields

  /** Cells a record spans. */
  def width: Int = fields.size

  /**
   * Decode one record. `cells(i)` is the cell aligned with `fields(i)` and carries its own ref, so
   * errors name the row and column. `cells.size != width` is a [[RowCodecError.Width]].
   */
  def read(cells: Vector[Cell]): Either[RowCodecError, A]

  /** Encode one record: `width` cell payloads aligned with `fields`, each with its format hint. */
  def write(a: A): Vector[(CellValue, Option[CellStyle])]

  /**
   * This codec with some headers renamed — the runtime twin of `@header`, for a header known only
   * when the script runs. `overrides` maps a field name (as spelled in `fields`, exactly) to its
   * header and layers on the current [[headers]], so it composes with the annotation and with an
   * earlier `withHeaders`. Refused as [[XLError.InvalidArgument]] (`RowCodec.withHeaders`): a key
   * that is not a field, a blank header, or two fields left with the same header — that would bind
   * both to one column on read and fail `putTable` late.
   *
   * Spell the given with `derived`, never with the summon: `given RowCodec[Order] =
   * RowCodec[Order].withHeaders(…)` asks for the very instance it is defining.
   *
   * {{{
   * given RowCodec[Order] = orExit(RowCodec.derived[Order].withHeaders(Map("price" -> "Unit Price ($)")))
   * }}}
   */
  def withHeaders(overrides: Map[String, String]): XLResult[RowCodec[A]] =
    RowCodec.rename(this, overrides)

object RowCodec:
  def apply[A](using rc: RowCodec[A]): RowCodec[A] = rc

  /**
   * Derive a codec for a case class from its `Mirror`: one [[FieldCodec]] per field, resolved at
   * the `derives` site, and each field's `@header` text read off the constructor (GH-614). Only the
   * field-codec summons, the label read and the annotation read are inline — the codec itself is a
   * plain instance built once, so a derived `given` costs one allocation per class.
   */
  inline def derived[A <: Product](using m: Mirror.ProductOf[A]): RowCodec[A] =
    inline erasedValue[m.MirroredElemTypes] match
      case _: EmptyTuple => error("RowCodec.derived: a record needs at least one field")
      case _ =>
        val labels = constValueTuple[m.MirroredElemLabels].toList.map(_.toString).toVector
        val overrides = RowCodecMacros.headerOverrides[A].toVector
        val headers =
          labels.zipWithIndex.map((label, i) => overrides.lift(i).flatten.getOrElse(label))
        build[A](labels, headers, fieldCodecs[m.MirroredElemTypes].toVector)

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

  private def build[A <: Product](
    labels: Vector[String],
    headerTexts: Vector[String],
    codecs: Vector[FieldCodec[?]]
  )(using m: Mirror.ProductOf[A]): RowCodec[A] =
    val arms: Vector[Arm] = codecs.map(fc => erase(fc))
    new RowCodec[A]:
      val fields: Vector[String] = labels
      override val headers: Vector[String] = headerTexts

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

  /** A codec with its headers replaced; everything else delegates (GH-614). */
  private final class Renamed[A](underlying: RowCodec[A], override val headers: Vector[String])
      extends RowCodec[A]:
    def fields: Vector[String] = underlying.fields
    def read(cells: Vector[Cell]): Either[RowCodecError, A] = underlying.read(cells)
    def write(a: A): Vector[(CellValue, Option[CellStyle])] = underlying.write(a)

  private val WithHeaders = "RowCodec.withHeaders"

  private def rename[A](codec: RowCodec[A], overrides: Map[String, String]): XLResult[RowCodec[A]] =
    val fields = codec.fields
    val unknown = overrides.keys.filterNot(fields.contains).toVector.sorted
    val blank =
      overrides.collect { case (field, text) if text.trim.isEmpty => field }.toVector.sorted
    if unknown.nonEmpty then
      Left(
        XLError.InvalidArgument(
          WithHeaders,
          s"no field ${listed(unknown)}; fields: ${fields.mkString(", ")}"
        )
      )
    else if blank.nonEmpty then
      Left(XLError.InvalidArgument(WithHeaders, s"blank header for field ${listed(blank)}"))
    else
      val headers = fields.zip(codec.headers).map((field, text) => overrides.getOrElse(field, text))
      val shared =
        fields.zip(headers).groupMap(_._2)(_._1).filter(_._2.size > 1).toVector.sortBy(_._1)
      if shared.nonEmpty then
        Left(
          XLError.InvalidArgument(
            WithHeaders,
            shared
              .map((text, names) => s"fields ${listed(names)} share header '$text'")
              .mkString("; ")
          )
        )
      else Right(Renamed(codec, headers))

  /** `'a'`, `'a' and 'b'`, `'a', 'b' and 'c'`. */
  private def listed(names: Vector[String]): String =
    val quoted = names.map(n => s"'$n'")
    if quoted.size <= 1 then quoted.mkString
    else s"${quoted.dropRight(1).mkString(", ")} and ${quoted.takeRight(1).mkString}"
