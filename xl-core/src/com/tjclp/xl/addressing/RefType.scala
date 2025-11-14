package com.tjclp.xl.addressing

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
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.Return"))
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

      // Parse sheet name (handle quotes and escaping)
      val sheetName = if sheetPart.startsWith("'") then
        if !sheetPart.endsWith("'") then
          return Left(s"Unbalanced quotes in sheet name: $sheetPart (missing closing quote)")
        val quoted = sheetPart.substring(1, sheetPart.length - 1)
        if quoted.isEmpty then return Left("Empty sheet name in quotes")
        // Unescape '' → ' (Excel convention)
        val unescaped = quoted.replace("''", "'")
        SheetName(unescaped)
      else
        if sheetPart.contains("'") then
          return Left(s"Misplaced quote in sheet name: $sheetPart (quotes must wrap entire name)")
        SheetName(sheetPart)

      sheetName match
        case Left(err) => Left(s"Invalid sheet name: $err")
        case Right(sheet) =>
          // Parse ref part as cell or range
          if refPart.contains(':') then CellRange.parse(refPart).map(QualifiedRange(sheet, _))
          else ARef.parse(refPart).map(QualifiedCell(sheet, _))

  /**
   * Find index of unquoted '!' (not inside 'quotes').
   *
   * Uses a toggle approach: each ' flips the inQuote state. This handles escaped quotes ('')
   * correctly because Excel's escaping convention uses two consecutive quotes to represent a single
   * literal quote.
   *
   * Examples:
   *   - "Sales!A1" → returns 5 (unquoted)
   *   - "'Sales!A1'!B1" → returns 11 (second ! is unquoted)
   *   - "'It''s Q1'!A1" → returns 10 ('' toggles twice, staying inside quotes)
   *
   * Returns -1 if no unquoted bang found.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.Return"))
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
