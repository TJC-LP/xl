package com.tjclp.xl.error

/**
 * Error types for the XL library. All errors are pure values that can be composed and transformed.
 */
enum XLError:
  /** Invalid cell reference format */
  case InvalidCellRef(ref: String, reason: String)

  /** Invalid range format */
  case InvalidRange(range: String, reason: String)

  /** Invalid sheet name */
  case InvalidSheetName(name: String, reason: String)

  /** Cell reference out of bounds */
  case OutOfBounds(ref: String, reason: String)

  /** Sheet not found */
  case SheetNotFound(name: String)

  /** Duplicate sheet name */
  case DuplicateSheet(name: String)

  /** Invalid column index */
  case InvalidColumn(index: Int, reason: String)

  /** Invalid row index */
  case InvalidRow(index: Int, reason: String)

  /** Type mismatch when reading cell value */
  case TypeMismatch(expected: String, actual: String, ref: String)

  /** Formula parse error */
  case FormulaError(expression: String, reason: String)

  /** Style error */
  case StyleError(reason: String)

  /** Number format error */
  case NumberFormatError(format: String, reason: String)

  /** Color format error */
  case ColorError(color: String, reason: String)

  /** Invalid workbook structure */
  case InvalidWorkbook(reason: String)

  /** IO error (from xl-cats-effect layer) */
  case IOError(reason: String)

  /** Parse error (from xl-ooxml layer) */
  case ParseError(location: String, reason: String)

  /** Number of supplied values mismatched expectation */
  case ValueCountMismatch(expected: Int, actual: Int, context: String)

  /** Generic error for extensibility */
  case Other(message: String)

object XLError:
  extension (error: XLError)
    /** Get human-readable error message */
    def message: String = error match
      case InvalidCellRef(ref, reason) => s"Invalid cell reference '$ref': $reason"
      case InvalidRange(range, reason) => s"Invalid range '$range': $reason"
      case InvalidSheetName(name, reason) => s"Invalid sheet name '$name': $reason"
      case OutOfBounds(ref, reason) => s"Reference out of bounds '$ref': $reason"
      case SheetNotFound(name) => s"Sheet not found: '$name'"
      case DuplicateSheet(name) => s"Duplicate sheet name: '$name'"
      case InvalidColumn(index, reason) => s"Invalid column index $index: $reason"
      case InvalidRow(index, reason) => s"Invalid row index $index: $reason"
      case TypeMismatch(expected, actual, ref) =>
        s"Type mismatch at $ref: expected $expected, got $actual"
      case FormulaError(expr, reason) => s"Formula error in '$expr': $reason"
      case StyleError(reason) => s"Style error: $reason"
      case NumberFormatError(format, reason) => s"Number format error '$format': $reason"
      case ColorError(color, reason) => s"Color error '$color': $reason"
      case InvalidWorkbook(reason) => s"Invalid workbook: $reason"
      case IOError(reason) => s"IO error: $reason"
      case ParseError(location, reason) => s"Parse error at $location: $reason"
      case ValueCountMismatch(expected, actual, context) =>
        s"Expected $expected values for $context but received $actual"
      case Other(message) => message
