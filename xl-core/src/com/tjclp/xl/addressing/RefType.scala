package com.tjclp.xl.addressing

import com.tjclp.xl.addressing.RefParser.ParsedRef

/**
 * Unified reference type for all Excel addressing formats.
 *
 * Supports:
 *   - Single cells: `A1`
 *   - Cell ranges: `A1:B10`
 *   - Sheet-qualified cells: `Sales!A1` or `'Q1 Sales'!A1`
 *   - Sheet-qualified ranges: `Sales!A1:B10` or `'Q1 Sales'!A1:B10`
 *   - Escaped quotes in sheet names: `'It''s Q1'!A1` ('' → ')
 *
 * **Note on macro vs runtime parsing:**
 *   - **Compile-time (`ref` macro)**: Simple refs return **unwrapped** types (ARef/CellRange) for
 *     backwards compatibility. Only sheet-qualified refs return RefType.
 *   - **Runtime (`RefType.parse`)**: Always returns RefType enum with Cell/Range cases for
 *     unqualified refs.
 *
 * Usage:
 * {{{
 * import com.tjclp.xl.macros.ref
 *
 * // Compile-time macro (unwrapped for simple refs)
 * val r1: ARef = ref"A1"              // Unwrapped ARef
 * val r2: CellRange = ref"A1:B10"     // Unwrapped CellRange
 * val r3 = ref"Sales!A1"              // RefType.QualifiedCell
 * val r4 = ref"'It''s Q1'!A1:B10"     // RefType.QualifiedRange
 *
 * // Runtime parsing (always wrapped)
 * RefType.parse("A1")                 // Right(RefType.Cell(ARef(...)))
 * RefType.parse("Sales!A1")           // Right(RefType.QualifiedCell(...))
 * }}}
 */
enum RefType derives CanEqual:
  /**
   * Single cell reference (e.g., `A1`).
   *
   * **Usage**: Returned by `RefType.parse` for unqualified cells. The `ref` macro returns unwrapped
   * `ARef` instead for backwards compatibility.
   */
  case Cell(ref: ARef)

  /**
   * Cell range reference (e.g., `A1:B10`).
   *
   * **Usage**: Returned by `RefType.parse` for unqualified ranges. The `ref` macro returns
   * unwrapped `CellRange` instead for backwards compatibility.
   */
  case Range(range: CellRange)

  /** Sheet-qualified cell reference (e.g., `Sales!A1` or `'Q1 Sales'!A1`) */
  case QualifiedCell(sheet: SheetName, ref: ARef)

  /** Sheet-qualified range reference (e.g., `Sales!A1:B10` or `'Q1 Sales'!A1:B10`) */
  case QualifiedRange(sheet: SheetName, range: CellRange)

  /**
   * Convert to A1 notation.
   *
   * For qualified refs, includes sheet name (with quotes if needed). Escapes single quotes as ''
   * (Excel convention).
   */
  def toA1: String = this match
    case Cell(ref) => ref.toA1
    case Range(range) => range.toA1
    case QualifiedCell(sheet, ref) =>
      val sheetStr = formatSheetName(sheet.value)
      s"$sheetStr!${ref.toA1}"
    case QualifiedRange(sheet, range) =>
      val sheetStr = formatSheetName(sheet.value)
      s"$sheetStr!${range.toA1}"

  /** Format sheet name with proper quoting and escaping */
  private def formatSheetName(name: String): String =
    if needsQuoting(name) then
      // Escape ' → '' and wrap in quotes
      val escaped = name.replace("'", "''")
      s"'$escaped'"
    else name

  /** Check if sheet name needs quoting (has spaces or special chars) */
  private def needsQuoting(name: String): Boolean =
    name.exists(c => c == ' ' || c == '\'' || !c.isLetterOrDigit && c != '_')

end RefType

object RefType:
  /**
   * Parse any reference format from string.
   *
   * Supports:
   *   - `A1` → Cell
   *   - `A1:B10` → Range
   *   - `Sales!A1` → QualifiedCell
   *   - `'Q1 Sales'!A1` → QualifiedCell (quoted sheet)
   *   - `Sales!A1:B10` → QualifiedRange
   *
   * @param s
   *   The reference string to parse
   * @return
   *   Either error message or parsed RefType
   */
  def parse(s: String): Either[String, RefType] =
    RefParser.parse(s).map {
      case ParsedRef.Cell(None, col0, row0) => Cell(ARef.from0(col0, row0))
      case ParsedRef.Range(None, cs, rs, ce, re) =>
        val range = CellRange(ARef.from0(cs, rs), ARef.from0(ce, re))
        Range(range)
      case ParsedRef.Cell(Some(sheetName), col0, row0) =>
        QualifiedCell(SheetName.unsafe(sheetName), ARef.from0(col0, row0))
      case ParsedRef.Range(Some(sheetName), cs, rs, ce, re) =>
        val range = CellRange(ARef.from0(cs, rs), ARef.from0(ce, re))
        QualifiedRange(SheetName.unsafe(sheetName), range)
    }

  /**
   * Parse ref string with XLError wrapping.
   *
   * Used by runtime string interpolation macro.
   */
  def parseToXLError(s: String): Either[com.tjclp.xl.error.XLError, RefType] =
    parse(s).left.map { err =>
      com.tjclp.xl.error.XLError.InvalidReference(s"Failed to parse '$s': $err")
    }

end RefType
