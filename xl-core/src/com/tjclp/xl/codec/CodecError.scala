package com.tjclp.xl.codec

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError

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
