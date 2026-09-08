package com.tjclp.xl.codec

import com.tjclp.xl.addressing.{ARef, Column, Row}
import com.tjclp.xl.error.XLError

/**
 * Why a row failed to decode as a record (GH-590).
 *
 * Every case names what an agent needs to fix the sheet or the record type: the cell (row and
 * column), the field, and the cell-level cause. [[RowCodecError.toXLError]] bridges into the
 * library-wide vocabulary for callers composing with `XLResult`.
 */
enum RowCodecError derives CanEqual:
  /** The cell at (`row`, `column`) held a value that `field`'s codec rejected. */
  case Field(row: Row, column: Column, field: String, cause: CodecError)

  /** `field` is not an `Option`, but its cell at (`row`, `column`) is empty. */
  case Missing(row: Row, column: Column, field: String)

  /** No cell in `headerRow` carries `header`; `available` lists the headers that are there. */
  case HeaderNotFound(header: String, headerRow: Row, available: Vector[String])

  /** A record spans `expected` cells; `actual` were supplied (a range of the wrong width). */
  case Width(expected: Int, actual: Int)

object RowCodecError:
  extension (error: RowCodecError)
    /** Human-readable, one line, cell first. */
    def message: String = error match
      case Field(row, column, field, cause) =>
        s"${ARef(column, row).toA1} ($field): ${describe(cause)}"
      case Missing(row, column, field) =>
        s"${ARef(column, row).toA1} ($field): required field is empty"
      case HeaderNotFound(header, headerRow, available) =>
        s"no header '$header' in row ${headerRow.index1}; found: ${available.mkString(", ")}"
      case Width(expected, actual) => s"a record spans $expected columns, got $actual"

    /** Bridge into the library-wide error vocabulary (the `CodecError.toXLError` precedent). */
    def toXLError: XLError = error match
      case Field(row, column, field, cause) =>
        // Matched on the cause alone so a new CodecError case fails to compile here, not widen.
        cause match
          case CodecError.TypeMismatch(expected, actual) =>
            XLError.TypeMismatch(
              s"$expected for field '$field'",
              actual.toString,
              ARef(column, row).toA1
            )
          case CodecError.ParseError(value, targetType, detail) =>
            XLError.ParseError(
              ARef(column, row).toA1,
              s"field '$field': cannot parse '$value' as $targetType: $detail"
            )
      case Missing(row, column, field) =>
        XLError.TypeMismatch(s"a value for field '$field'", "Empty", ARef(column, row).toA1)
      case HeaderNotFound(header, headerRow, available) =>
        XLError.InvalidReference(
          s"no header '$header' in row ${headerRow.index1}; found: ${available.mkString(", ")}"
        )
      case Width(expected, actual) => XLError.ValueCountMismatch(expected, actual, "record columns")

  private def describe(cause: CodecError): String = cause match
    case CodecError.TypeMismatch(expected, actual) => s"expected $expected, got $actual"
    case CodecError.ParseError(value, targetType, detail) =>
      s"cannot parse '$value' as $targetType: $detail"
