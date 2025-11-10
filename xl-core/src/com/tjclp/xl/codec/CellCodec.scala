package com.tjclp.xl.codec

import com.tjclp.xl.*
import com.tjclp.xl.style.{CellStyle, NumFmt}
import java.time.{LocalDate, LocalDateTime}

/**
 * Read a typed value from a Cell.
 *
 * Returns Right(None) if the cell is empty, Right(Some(value)) if successfully decoded, or
 * Left(error) if there's a type mismatch or parse error.
 */
trait CellReader[A]:
  def read(cell: Cell): Either[CodecError, Option[A]]

/**
 * Write a typed value to produce cell data with optional style hint.
 *
 * Returns (CellValue, Optional[CellStyle]) where the style is auto-inferred based on the value type
 * (e.g., DateTime gets date format, BigDecimal gets decimal format).
 */
trait CellWriter[A]:
  def write(a: A): (CellValue, Option[CellStyle])

/** Bidirectional codec for cell values */
trait CellCodec[A] extends CellReader[A], CellWriter[A]

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

/**
 * Cell codec error types
 *
 * These errors are specific to cell-level encoding/decoding and can be converted to XLError when
 * needed.
 */
enum CodecError:
  /** Type mismatch when reading a cell */
  case TypeMismatch(expected: String, actual: CellValue)

  /** Parse error when converting cell value to target type */
  case ParseError(value: String, targetType: String, detail: String)

object CodecError:
  extension (error: CodecError)
    /** Convert CodecError to XLError for compatibility with general error handling */
    def toXLError(ref: ARef): XLError = error match
      case TypeMismatch(expected, actual) =>
        XLError.TypeMismatch(expected, actual.toString, ref.toA1)
      case ParseError(value, targetType, detail) =>
        XLError.ParseError(ref.toA1, s"Cannot parse '$value' as $targetType: $detail")
