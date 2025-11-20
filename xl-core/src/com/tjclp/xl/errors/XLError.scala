package com.tjclp.xl.errors

/**
 * Error types for the XL library. All errors are pure values that can be composed and transformed.
 */
enum XLError:
  /** Invalid cell reference format */
  case InvalidCellRef(ref: String, reason: String)

  /** Invalid range format */
  case InvalidRange(range: String, reason: String)

  /** Invalid reference (e.g., unqualified ref used with workbooks) */
  case InvalidReference(reason: String)

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

  /** Formula parse errors */
  case FormulaError(expression: String, reason: String)

  /** Style errors */
  case StyleError(reason: String)

  /** Number format errors */
  case NumberFormatError(format: String, reason: String)

  /** Money format parse errors */
  case MoneyFormatError(value: String, reason: String)

  /** Percent format parse errors */
  case PercentFormatError(value: String, reason: String)

  /** Date format parse errors */
  case DateFormatError(value: String, reason: String)

  /** Accounting format parse errors */
  case AccountingFormatError(value: String, reason: String)

  /** Color format errors */
  case ColorError(color: String, reason: String)

  /** Invalid workbooks structure */
  case InvalidWorkbook(reason: String)

  /** IO errors (from xl-cats-effect layer) */
  case IOError(reason: String)

  /** Parse errors (from xl-ooxml layer) */
  case ParseError(location: String, reason: String)

  /** Number of supplied values mismatched expectation */
  case ValueCountMismatch(expected: Int, actual: Int, context: String)

  /** Unsupported type in batch put operation */
  case UnsupportedType(ref: String, typeName: String)

  /** Generic errors for extensibility */
  case Other(message: String)

object XLError:
  extension (error: XLError)
    /** Get human-readable errors message */
    def message: String = error match
      case InvalidCellRef(ref, reason) => s"Invalid cell reference '$ref': $reason"
      case InvalidRange(range, reason) => s"Invalid range '$range': $reason"
      case InvalidReference(reason) => s"Invalid reference: $reason"
      case InvalidSheetName(name, reason) => s"Invalid sheet name '$name': $reason"
      case OutOfBounds(ref, reason) => s"Reference out of bounds '$ref': $reason"
      case SheetNotFound(name) => s"Sheet not found: '$name'"
      case DuplicateSheet(name) => s"Duplicate sheet name: '$name'"
      case InvalidColumn(index, reason) => s"Invalid column index $index: $reason"
      case InvalidRow(index, reason) => s"Invalid row index $index: $reason"
      case TypeMismatch(expected, actual, ref) =>
        s"Type mismatch at $ref: expected $expected, got $actual"
      case FormulaError(expr, reason) => s"Formula errors in '$expr': $reason"
      case StyleError(reason) => s"Style errors: $reason"
      case NumberFormatError(format, reason) => s"Number format errors '$format': $reason"
      case MoneyFormatError(value, reason) => s"Invalid money format '$value': $reason"
      case PercentFormatError(value, reason) => s"Invalid percent format '$value': $reason"
      case DateFormatError(value, reason) => s"Invalid date format '$value': $reason"
      case AccountingFormatError(value, reason) => s"Invalid accounting format '$value': $reason"
      case ColorError(color, reason) => s"Color errors '$color': $reason"
      case InvalidWorkbook(reason) => s"Invalid workbooks: $reason"
      case IOError(reason) => s"IO errors: $reason"
      case ParseError(location, reason) => s"Parse errors at $location: $reason"
      case ValueCountMismatch(expected, actual, context) =>
        s"Expected $expected values for $context but received $actual"
      case UnsupportedType(ref, typeName) =>
        s"Unsupported type at $ref: $typeName. Supported types: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText, Formatted"
      case Other(message) => message

/** Type alias for common result type */
type XLResult[A] = Either[XLError, A]
