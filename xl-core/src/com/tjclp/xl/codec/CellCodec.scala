package com.tjclp.xl.codec

import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import java.time.{LocalDate, LocalDateTime}

/**
 * Bidirectional codec for cell values.
 *
 * Note: CellCodec does NOT extend CellWriter to avoid ambiguous given instances. For write
 * operations, use CellWriter[CellWritable] which handles all built-in types via contravariance.
 * This separation enables:
 *   - CellCodec for read + write of specific types
 *   - CellWriter for write-only with user extensibility
 */
trait CellCodec[A] extends CellReader[A]:
  def write(a: A): (CellValue, Option[CellStyle])

object CellCodec:
  def apply[A](using cc: CellCodec[A]): CellCodec[A] = cc

  // Helper to create codecs
  private def codec[A](
    r: Cell => Either[CodecError, Option[A]],
    w: A => (CellValue, Option[CellStyle])
  ): CellCodec[A] = new CellCodec[A]:
    def read(cell: Cell) = r(cell)
    def write(a: A) = w(a)

  // ========== 8 Primitive Codec Instances ==========

  /** String codec - reads text and number cells as strings */
  inline given CellCodec[String] = codec(
    cell =>
      cell.value match
        case CellValue.Empty => Right(None)
        case CellValue.Text(s) => Right(Some(s))
        case CellValue.Number(n) => Right(Some(n.toString))
        case CellValue.Bool(b) => Right(Some(b.toString))
        case other => Left(CodecError.TypeMismatch("String", other)),
    s => (CellValue.Text(s), None)
  )

  /** Int codec - reads numeric cells as integers */
  inline given CellCodec[Int] = codec(
    cell =>
      cell.value match
        case CellValue.Empty => Right(None)
        case CellValue.Number(n) =>
          try Right(Some(n.toIntExact))
          catch
            case _: ArithmeticException =>
              Left(
                CodecError
                  .ParseError(n.toString, "Int", "Value out of Int range or has decimal places")
              )
        case other => Left(CodecError.TypeMismatch("Int", other)),
    i => (CellValue.Number(BigDecimal(i)), Some(CellStyle.default.withNumFmt(NumFmt.General)))
  )

  /** Long codec - reads numeric cells as longs */
  inline given CellCodec[Long] = codec(
    cell =>
      cell.value match
        case CellValue.Empty => Right(None)
        case CellValue.Number(n) =>
          try Right(Some(n.toLongExact))
          catch
            case _: ArithmeticException =>
              Left(
                CodecError
                  .ParseError(n.toString, "Long", "Value out of Long range or has decimal places")
              )
        case other => Left(CodecError.TypeMismatch("Long", other)),
    l => (CellValue.Number(BigDecimal(l)), Some(CellStyle.default.withNumFmt(NumFmt.General)))
  )

  /** Double codec - reads numeric cells as doubles */
  inline given CellCodec[Double] = codec(
    cell =>
      cell.value match
        case CellValue.Empty => Right(None)
        case CellValue.Number(n) => Right(Some(n.toDouble))
        case other => Left(CodecError.TypeMismatch("Double", other)),
    d => (CellValue.Number(BigDecimal(d)), Some(CellStyle.default.withNumFmt(NumFmt.General)))
  )

  /** BigDecimal codec - reads numeric cells with auto-inferred decimal format */
  inline given CellCodec[BigDecimal] = codec(
    cell =>
      cell.value match
        case CellValue.Empty => Right(None)
        case CellValue.Number(n) => Right(Some(n))
        case other => Left(CodecError.TypeMismatch("BigDecimal", other)),
    bd => (CellValue.Number(bd), Some(CellStyle.default.withNumFmt(NumFmt.Decimal)))
  )

  /** Boolean codec - reads boolean cells */
  inline given CellCodec[Boolean] = codec(
    cell =>
      cell.value match
        case CellValue.Empty => Right(None)
        case CellValue.Bool(b) => Right(Some(b))
        case CellValue.Number(n) if n == BigDecimal(0) || n == BigDecimal(1) =>
          Right(Some(n == BigDecimal(1)))
        case other => Left(CodecError.TypeMismatch("Boolean", other)),
    b => (CellValue.Bool(b), None)
  )

  /**
   * LocalDate codec - reads date/datetime cells and Excel serial numbers with auto-inferred date
   * format
   */
  inline given CellCodec[LocalDate] = codec(
    cell =>
      cell.value match
        case CellValue.Empty => Right(None)
        case CellValue.DateTime(dt) => Right(Some(dt.toLocalDate))
        case CellValue.Number(serial) =>
          try Right(Some(CellValue.excelSerialToDateTime(serial.toDouble).toLocalDate))
          catch
            case e: Exception =>
              Left(
                CodecError.ParseError(
                  serial.toString,
                  "LocalDate",
                  s"Invalid Excel serial number: ${e.getMessage}"
                )
              )
        case other => Left(CodecError.TypeMismatch("LocalDate", other)),
    date => (CellValue.DateTime(date.atStartOfDay), Some(CellStyle.default.withNumFmt(NumFmt.Date)))
  )

  /**
   * LocalDateTime codec - reads date/datetime cells and Excel serial numbers with auto-inferred
   * datetime format
   */
  inline given CellCodec[LocalDateTime] = codec(
    cell =>
      cell.value match
        case CellValue.Empty => Right(None)
        case CellValue.DateTime(dt) => Right(Some(dt))
        case CellValue.Number(serial) =>
          try Right(Some(CellValue.excelSerialToDateTime(serial.toDouble)))
          catch
            case e: Exception =>
              Left(
                CodecError.ParseError(
                  serial.toString,
                  "LocalDateTime",
                  s"Invalid Excel serial number: ${e.getMessage}"
                )
              )
        case other => Left(CodecError.TypeMismatch("LocalDateTime", other)),
    dt => (CellValue.DateTime(dt), Some(CellStyle.default.withNumFmt(NumFmt.DateTime)))
  )

  /**
   * RichText codec - reads rich text cells or converts plain text to RichText.
   *
   * Note: Rich text has formatting within the text runs (intra-cell formatting), not at the cell
   * level. Therefore, this codec does not return a CellStyle hint (formatting is in the TextRun
   * font properties).
   */
  inline given CellCodec[RichText] = codec(
    cell =>
      cell.value match
        case CellValue.Empty => Right(None)
        case CellValue.RichText(rt) => Right(Some(rt))
        case CellValue.Text(s) =>
          // Plain text can be read as RichText (single unformatted run)
          Right(Some(RichText.plain(s)))
        case other => Left(CodecError.TypeMismatch("RichText", other)),
    rt => (CellValue.RichText(rt), None) // No cell-level style, formatting is in runs
  )

  // ========== Union Type Master Writer ==========

  /**
   * Master writer for the CellWritable union type.
   *
   * This single instance handles all supported types via pattern matching. Due to contravariance of
   * `CellWriter[-A]`, this instance satisfies `CellWriter[T]` for any `T <: CellWritable`.
   *
   * This enables type-safe heterogeneous batch operations like:
   * {{{
   * sheet.put(ref"A1" -> "hello", ref"B1" -> 42, ref"C1" -> LocalDate.now)
   * // Inferred type: A = String | Int | LocalDate <: CellWritable ✓
   * }}}
   *
   * Performance (GH-297): this is the hot path for every codec-based `Sheet.put`. Branches
   * construct results directly instead of summoning the per-type codecs — the codecs are `inline
   * given`, so a summon inside `write` would allocate a fresh codec instance (and a fresh
   * `CellStyle` hint) on every call. Each branch MUST stay in lockstep with the corresponding
   * `CellCodec` write implementation above; `CellCodecSpec` ("GH-297: master writer matches codec
   * writes...") pins the equivalence. Common primitive branches are ordered first.
   */
  given CellWriter[CellWritable] with
    // Shared style hints: immutable, and CellStyle.canonicalKey is a lazy val — reusing one
    // instance per hint computes the style-dedup key once per JVM instead of once per write.
    private val generalNumberHint: Option[CellStyle] =
      Some(CellStyle.default.withNumFmt(NumFmt.General))
    private val decimalHint: Option[CellStyle] = Some(CellStyle.default.withNumFmt(NumFmt.Decimal))
    private val dateHint: Option[CellStyle] = Some(CellStyle.default.withNumFmt(NumFmt.Date))
    private val dateTimeHint: Option[CellStyle] = Some(
      CellStyle.default.withNumFmt(NumFmt.DateTime)
    )

    def write(value: CellWritable): (CellValue, Option[CellStyle]) = value match
      case s: String => (CellValue.Text(s), None)
      case i: Int => (CellValue.Number(BigDecimal(i)), generalNumberHint)
      case d: Double => (CellValue.Number(BigDecimal(d)), generalNumberHint)
      case b: Boolean => (CellValue.Bool(b), None)
      case ld: LocalDate => (CellValue.DateTime(ld.atStartOfDay), dateHint)
      case l: Long => (CellValue.Number(BigDecimal(l)), generalNumberHint)
      case bd: BigDecimal => (CellValue.Number(bd), decimalHint)
      case ldt: LocalDateTime => (CellValue.DateTime(ldt), dateTimeHint)
      case rt: RichText => (CellValue.RichText(rt), None)
      case tr: com.tjclp.xl.richtext.TextRun =>
        // TextRun from RichText DSL - convert to single-run RichText
        (CellValue.RichText(RichText(Vector(tr))), None)
      case cv: CellValue => (cv, None)
      case f: com.tjclp.xl.formatted.Formatted =>
        (f.value, Some(CellStyle.default.withNumFmt(f.numFmt)))
