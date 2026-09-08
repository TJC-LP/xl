package com.tjclp.xl.error

import java.util.Locale

import scala.compiletime.constValueTuple
import scala.deriving.Mirror

/**
 * Error types for the XL library. All errors are pure values that can be composed and transformed.
 *
 * Every case carries a stable machine-readable [[XLError.code]] — the SCREAMING_SNAKE of its name —
 * plus an optional one-line [[XLError.hint]] and [[XLError.candidates]] for "did you mean" (ADR-017
 * §2.7). The CLI's `code:` line and the scripting prelude read these; nothing maps them by hand.
 */
enum XLError derives CanEqual:
  /** Invalid cell reference format */
  case InvalidCellRef(ref: String, reason: String)

  /** Invalid range format */
  case InvalidRange(range: String, reason: String)

  /** Invalid reference (e.g., unqualified ref used with workbook) */
  case InvalidReference(reason: String)

  /** Invalid sheet name */
  case InvalidSheetName(name: String, reason: String)

  /** Cell reference out of bounds */
  case OutOfBounds(ref: String, reason: String)

  /** Sheet not found */
  case SheetNotFound(name: String)

  /** Duplicate sheet name */
  case DuplicateSheet(name: String)

  /** Duplicate cell reference in batch operation */
  case DuplicateCellRef(refs: String)

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

  /** Money format parse error */
  case MoneyFormatError(value: String, reason: String)

  /** Percent format parse error */
  case PercentFormatError(value: String, reason: String)

  /** Date format parse error */
  case DateFormatError(value: String, reason: String)

  /** Accounting format parse error */
  case AccountingFormatError(value: String, reason: String)

  /** Color format error */
  case ColorError(color: String, reason: String)

  /** Invalid workbook structure */
  case InvalidWorkbook(reason: String)

  /** IO error (from xl-cats-effect layer) */
  case IOError(reason: String)

  /** Parse error (from xl-ooxml layer) */
  case ParseError(location: String, reason: String)

  /** Security error (ZIP bomb, formula injection, size limits) */
  case SecurityError(reason: String)

  /**
   * Number of supplied values mismatched expectation. `expected` is a `Long`: a range's cell count
   * (`CellRange.cellCount`) exceeds `Int.MaxValue` from `A1:XFD131072` on, and the message must
   * report it rather than its truncation.
   */
  case ValueCountMismatch(expected: Long, actual: Int, context: String)

  /** Unsupported type in batch put operation */
  case UnsupportedType(ref: String, typeName: String)

  /** Invalid table name (empty, spaces, invalid characters) */
  case InvalidTableName(name: String, reason: String)

  /** Invalid table display name (empty, spaces, invalid characters) */
  case InvalidTableDisplayName(displayName: String, reason: String)

  /** Invalid table range (too small, invalid bounds) */
  case InvalidTableRange(range: String, reason: String)

  /** Invalid table column configuration (empty, duplicates, count mismatch) */
  case InvalidTableColumns(reason: String)

  /** Generic error for extensibility */
  case Other(message: String)

  /**
   * The edit at position `index` (1-based, the batch's `Object N`) of a sequence failed; `cause` is
   * the domain error. Its [[XLError.code]], [[XLError.hint]] and [[XLError.candidates]] are the
   * cause's (ADR-017 §2.7).
   */
  case EditFailed(index: Int, op: String, cause: XLError)

  /** `op` needs `capability` (a mode, a backend, a loaded workbook) that the current run lacks. */
  case UnsupportedCapability(op: String, capability: String, hint: String)

  /** A sheet-scoped operation ran without a sheet; `available` lists the workbook's sheet names. */
  case SheetRequired(context: String, available: Vector[String])

object XLError:
  extension (error: XLError)
    /** Get human-readable error message */
    def message: String = error match
      case InvalidCellRef(ref, reason) => s"Invalid cell reference '$ref': $reason"
      case InvalidRange(range, reason) => s"Invalid range '$range': $reason"
      case InvalidReference(reason) => s"Invalid reference: $reason"
      case InvalidSheetName(name, reason) => s"Invalid sheet name '$name': $reason"
      case OutOfBounds(ref, reason) => s"Reference out of bounds '$ref': $reason"
      case SheetNotFound(name) => s"Sheet not found: '$name'"
      case DuplicateSheet(name) => s"Duplicate sheet name: '$name'"
      case DuplicateCellRef(refs) => s"Duplicate cell references not allowed: $refs"
      case InvalidColumn(index, reason) => s"Invalid column index $index: $reason"
      case InvalidRow(index, reason) => s"Invalid row index $index: $reason"
      case TypeMismatch(expected, actual, ref) =>
        s"Type mismatch at $ref: expected $expected, got $actual"
      case FormulaError(expr, reason) => s"Formula error in '$expr': $reason"
      case StyleError(reason) => s"Style error: $reason"
      case NumberFormatError(format, reason) => s"Number format error '$format': $reason"
      case MoneyFormatError(value, reason) => s"Invalid money format '$value': $reason"
      case PercentFormatError(value, reason) => s"Invalid percent format '$value': $reason"
      case DateFormatError(value, reason) => s"Invalid date format '$value': $reason"
      case AccountingFormatError(value, reason) => s"Invalid accounting format '$value': $reason"
      case ColorError(color, reason) => s"Color error '$color': $reason"
      case InvalidWorkbook(reason) => s"Invalid workbook: $reason"
      case IOError(reason) => s"IO error: $reason"
      case ParseError(location, reason) => s"Parse error at $location: $reason"
      case SecurityError(reason) => s"Security error: $reason"
      case ValueCountMismatch(expected, actual, context) =>
        s"Expected $expected values for $context but received $actual"
      case UnsupportedType(ref, typeName) =>
        s"Unsupported type at $ref: $typeName. Supported types: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText, Formatted"
      case InvalidTableName(name, reason) => s"Invalid table name '$name': $reason"
      case InvalidTableDisplayName(displayName, reason) =>
        s"Invalid table displayName '$displayName': $reason"
      case InvalidTableRange(range, reason) => s"Invalid table range '$range': $reason"
      case InvalidTableColumns(reason) => s"Invalid table columns: $reason"
      case Other(message) => message
      case EditFailed(index, op, cause) => s"op $index ($op): ${cause.message}"
      case UnsupportedCapability(op, capability, hint) => s"$op requires $capability: $hint"
      case SheetRequired(context, available) =>
        s"$context requires a sheet: pass -s <name> or qualify the ref (Sheet!A1). Available: ${available
            .mkString(", ")}"

    /**
     * Stable machine-readable code: the SCREAMING_SNAKE of the case name (`SheetNotFound` →
     * `SHEET_NOT_FOUND`, `IOError` → `IO_ERROR`, `Other` → `OTHER`). [[EditFailed]] reports its
     * cause's code so a batch failure is classified by what went wrong, not by where.
     */
    def code: String = error match
      case EditFailed(_, _, cause) => cause.code
      case other => caseCode(other.productPrefix)

    /** A one-line next step for the cases that have an obvious one; `None` otherwise. */
    def hint: Option[String] = error match
      case SheetNotFound(_) => Some("list sheets with `xl -f <file> sheets`")
      case SheetRequired(_, _) => Some("use -s <name> or a qualified ref like 'Name'!A1")
      case ValueCountMismatch(expected, _, _) =>
        Some(s"provide exactly $expected values, or 1 to fill the range")
      case FormulaError(_, _) => Some("check the formula with `xl eval`")
      case SecurityError(_) => Some("re-run with --max-size 0 or --stream")
      case OutOfBounds(_, _) => Some("valid cells are A1..XFD1048576")
      case UnsupportedCapability(_, _, hint) => Some(hint)
      case EditFailed(_, _, cause) => cause.hint
      case _ => None

    /** Names to offer as "did you mean": the available sheets of a [[SheetRequired]]. */
    def candidates: Vector[String] = error match
      case SheetRequired(_, available) => available
      case EditFailed(_, _, cause) => cause.candidates
      case _ => Vector.empty

    /** The innermost error: [[EditFailed]] unwrapped, anything else itself. */
    def root: XLError = error match
      case EditFailed(_, _, cause) => cause.root
      case other => other

    /** The batch position of an [[EditFailed]]. */
    def opIndex: Option[Int] = error match
      case EditFailed(index, _, _) => Some(index)
      case _ => None

  /** SCREAMING_SNAKE of a CamelCase case name; acronyms stay together (`IOError` → `IO_ERROR`). */
  private def caseCode(name: String): String =
    name
      .replaceAll("([a-z0-9])([A-Z])", "$1_$2")
      .replaceAll("([A-Z]+)([A-Z][a-z])", "$1_$2")
      .toUpperCase(Locale.ROOT)

  /** The case names in declaration order, read from the enum's mirror so they cannot drift. */
  private inline def labelsOf[T](using m: Mirror.SumOf[T]): List[String] =
    constValueTuple[m.MirroredElemLabels].toList.map(_.toString)

  private val caseNames: Vector[String] = labelsOf[XLError].toVector

  /**
   * One code per case, unique, in declaration order — the complete vocabulary `code` can return
   * (plus `EDIT_FAILED`, which [[EditFailed]] itself never reports). Feeds the CLI's
   * `ErrorCode.all` and `schema --json`.
   */
  val codes: Vector[String] = caseNames.map(caseCode)

/** Type alias for common result type */
type XLResult[A] = Either[XLError, A]
