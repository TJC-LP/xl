package com.tjclp.xl.addressing

/**
 * Unified reference type for all Excel addressing formats.
 *
 * Supports:
 *   - Single cells: `A1`
 *   - Cell ranges: `A1:B10`
 *   - Sheet-qualified cells: `Sales!A1` or `'Q1 Sales'!A1`
 *   - Sheet-qualified ranges: `Sales!A1:B10` or `'Q1 Sales'!A1:B10`
 *
 * Usage:
 * {{{
 * import com.tjclp.xl.macros.ref
 *
 * val r1 = ref"A1"              // RefType.Cell(ARef(...))
 * val r2 = ref"A1:B10"          // RefType.Range(CellRange(...))
 * val r3 = ref"Sales!A1"        // RefType.QualifiedCell(SheetName("Sales"), ARef(...))
 * val r4 = ref"Sales!A1:B10"    // RefType.QualifiedRange(SheetName("Sales"), CellRange(...))
 * }}}
 */
enum RefType derives CanEqual:
  /** Single cell reference (e.g., `A1`) */
  case Cell(ref: ARef)

  /** Cell range reference (e.g., `A1:B10`) */
  case Range(range: CellRange)

  /** Sheet-qualified cell reference (e.g., `Sales!A1` or `'Q1 Sales'!A1`) */
  case QualifiedCell(sheet: SheetName, ref: ARef)

  /** Sheet-qualified range reference (e.g., `Sales!A1:B10` or `'Q1 Sales'!A1:B10`) */
  case QualifiedRange(sheet: SheetName, range: CellRange)

  /**
   * Convert to A1 notation.
   *
   * For qualified refs, includes sheet name (with quotes if needed).
   */
  def toA1: String = this match
    case Cell(ref) => ref.toA1
    case Range(range) => range.toA1
    case QualifiedCell(sheet, ref) =>
      val sheetStr = if needsQuoting(sheet.value) then s"'${sheet.value}'" else sheet.value
      s"$sheetStr!${ref.toA1}"
    case QualifiedRange(sheet, range) =>
      val sheetStr = if needsQuoting(sheet.value) then s"'${sheet.value}'" else sheet.value
      s"$sheetStr!${range.toA1}"

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
    if s.isEmpty then return Left("Empty reference")

    // Check for sheet qualifier (!)
    val bangIdx = findUnquotedBang(s)
    if bangIdx < 0 then
      // No sheet qualifier, parse as cell or range
      if s.contains(':') then CellRange.parse(s).map(Range.apply)
      else ARef.parse(s).map(Cell.apply)
    else
      // Has sheet qualifier
      val sheetPart = s.substring(0, bangIdx)
      val refPart = s.substring(bangIdx + 1)

      if refPart.isEmpty then return Left(s"Missing reference after '!' in: $s")

      // Parse sheet name (handle quotes)
      val sheetName = if sheetPart.startsWith("'") && sheetPart.endsWith("'") then
        val unquoted = sheetPart.substring(1, sheetPart.length - 1)
        if unquoted.isEmpty then return Left("Empty sheet name in quotes")
        SheetName(unquoted)
      else SheetName(sheetPart)

      sheetName match
        case Left(err) => Left(s"Invalid sheet name: $err")
        case Right(sheet) =>
          // Parse ref part as cell or range
          if refPart.contains(':') then CellRange.parse(refPart).map(QualifiedRange(sheet, _))
          else ARef.parse(refPart).map(QualifiedCell(sheet, _))

  /**
   * Find index of unquoted '!' (not inside 'quotes').
   *
   * Returns -1 if no unquoted bang found.
   */
  private def findUnquotedBang(s: String): Int =
    var i = 0
    var inQuote = false
    while i < s.length do
      val c = s.charAt(i)
      if c == '\'' then inQuote = !inQuote
      else if c == '!' && !inQuote then return i
      i += 1
    -1

end RefType
