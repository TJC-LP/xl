package com.tjclp.xl.ooxml

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}

/**
 * OOXML shared-formula support shared by the DOM and streaming readers.
 *
 * A shared-formula dependent stores only a group index; its formula is the master's expression
 * copied from the master cell to the dependent cell. Copying is anchor-aware, so this translator
 * rewrites A1 references without requiring the evaluator's deliberately smaller formula grammar. In
 * particular, unsupported Excel functions and syntax remain opaque while references still move.
 */
private[xl] object SharedFormula:

  type Index = Long

  private val MaxUnsignedInt = 4_294_967_295L

  /** Normalize the xsd:unsignedInt lexical space used by `<f si="...">`. */
  def parseIndex(raw: String): Option[Index] =
    raw.trim.toLongOption.filter(index => index >= 0L && index <= MaxUnsignedInt)

  final case class Master(
    origin: ARef,
    expression: String,
    range: Option[CellRange] = None
  )

  /** Expand `master` for `target`, following Excel fill/copy reference semantics. */
  def translate(master: Master, target: ARef): String =
    val colDelta = target.col.index0 - master.origin.col.index0
    val rowDelta = target.row.index0 - master.origin.row.index0
    if colDelta == 0 && rowDelta == 0 then master.expression
    else translateExpression(master.expression, colDelta, rowDelta)

  private final case class ColPart(index0: Int, absolute: Boolean, end: Int)
  private final case class RowPart(index0: Int, absolute: Boolean, end: Int)
  private final case class CellPart(
    colIndex0: Int,
    rowIndex0: Int,
    colAbsolute: Boolean,
    rowAbsolute: Boolean,
    end: Int
  )

  /**
   * Lexically translate references while leaving strings, quoted sheet/workbook names, and
   * structured-reference brackets untouched. This intentionally does not parse the whole formula:
   * shared formulas may contain any function Excel supports, not just xl-evaluator's typed subset.
   * Structured-reference payloads stay deliberately opaque: horizontal table-column translation
   * requires table schema/context that this low-level OOXML reader does not have.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def translateExpression(expression: String, colDelta: Int, rowDelta: Int): String =
    val out = new StringBuilder(expression.length + 16)
    var index = 0
    var bracketDepth = 0

    while index < expression.length do
      val ch = expression.charAt(index)

      if bracketDepth > 0 then
        out.append(ch)
        if ch == '[' then bracketDepth += 1
        else if ch == ']' then bracketDepth -= 1
        index += 1
      else
        ch match
          case '"' =>
            val end = quotedEnd(expression, index, '"')
            out.append(expression.substring(index, end))
            index = end

          case '\'' =>
            // Single quotes delimit sheet/workbook qualifiers. A doubled quote escapes a quote in
            // the name, just as doubled double-quotes do in string literals.
            val end = quotedEnd(expression, index, '\'')
            out.append(expression.substring(index, end))
            index = end

          case '[' =>
            // Covers both external-workbook qualifiers ([Book.xlsx]) and structured references.
            // Neither may have its contents translated as ordinary A1 references.
            out.append(ch)
            bracketDepth = 1
            index += 1

          case _ =>
            val replacement =
              parseCellRange(expression, index)
                .map { case (start, end) =>
                  val rendered = renderCellRange(start, end, colDelta, rowDelta)
                  (end.end, rendered)
                }
                .orElse {
                  parseColumnRange(expression, index).map { case (start, end) =>
                    val rendered = renderColumnRange(start, end, colDelta)
                    (end.end, rendered)
                  }
                }
                .orElse {
                  parseRowRange(expression, index).map { case (start, end) =>
                    val rendered = renderRowRange(start, end, rowDelta)
                    (end.end, rendered)
                  }
                }
                .orElse {
                  // Avoid rescanning every suffix of a long defined name. Once the first
                  // character has failed to form a reference, every later character in that
                  // name has a non-boundary predecessor and can be rejected in O(1).
                  if !hasBoundaryBefore(expression, index) then None
                  else
                    parseCellPart(expression, index)
                      .filter(cell => hasBoundaryAfter(expression, cell.end))
                      .map { cell =>
                        val original = expression.substring(index, cell.end)
                        val rendered =
                          if isNonReferenceUse(expression, cell.end) then original
                          else renderCell(cell, colDelta, rowDelta)
                        (cell.end, rendered)
                      }
                }

            replacement match
              case Some((end, text)) =>
                out.append(text)
                index = end
              case None =>
                out.append(ch)
                index += 1

    out.result()

  /** End (exclusive) of a quoted formula segment, tolerating malformed/unclosed input. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def quotedEnd(input: String, start: Int, quote: Char): Int =
    var index = start + 1
    var closed = false
    while index < input.length && !closed do
      if input.charAt(index) == quote then
        if index + 1 < input.length && input.charAt(index + 1) == quote then index += 2
        else
          index += 1
          closed = true
      else index += 1
    index

  private def parseCellRange(input: String, start: Int): Option[(CellPart, CellPart)] =
    if !hasBoundaryBefore(input, start) then None
    else
      for
        first <- parseCellPart(input, start)
        if first.end < input.length && input.charAt(first.end) == ':'
        second <- parseCellPart(input, first.end + 1)
        if hasBoundaryAfter(input, second.end)
        if !isNonReferenceUse(input, second.end)
      yield (first, second)

  private def parseColumnRange(input: String, start: Int): Option[(ColPart, ColPart)] =
    if !hasBoundaryBefore(input, start) then None
    else
      for
        first <- parseColPart(input, start)
        if first.end < input.length && input.charAt(first.end) == ':'
        second <- parseColPart(input, first.end + 1)
        if hasBoundaryAfter(input, second.end)
        if !isNonReferenceUse(input, second.end)
      yield (first, second)

  private def parseRowRange(input: String, start: Int): Option[(RowPart, RowPart)] =
    if !hasBoundaryBefore(input, start) then None
    else
      for
        first <- parseRowPart(input, start)
        if first.end < input.length && input.charAt(first.end) == ':'
        second <- parseRowPart(input, first.end + 1)
        if hasBoundaryAfter(input, second.end)
        if !isNonReferenceUse(input, second.end)
      yield (first, second)

  private def parseCellPart(input: String, start: Int): Option[CellPart] =
    val colAbsolute = start < input.length && input.charAt(start) == '$'
    val colStart = if colAbsolute then start + 1 else start

    val lettersEnd = spanLetters(input, colStart)
    val letterCount = lettersEnd - colStart
    if letterCount < 1 || letterCount > 3 then None
    else
      val rowAbsolute = lettersEnd < input.length && input.charAt(lettersEnd) == '$'
      val rowStart = if rowAbsolute then lettersEnd + 1 else lettersEnd
      val digitsEnd = spanDigits(input, rowStart)
      if digitsEnd == rowStart then None
      else
        for
          col <- Column.fromLetter(input.substring(colStart, lettersEnd)).toOption
          row1 <- input.substring(rowStart, digitsEnd).toIntOption
          if row1 >= 1 && row1 <= Row.MaxIndex0 + 1
        yield CellPart(col.index0, row1 - 1, colAbsolute, rowAbsolute, digitsEnd)

  private def parseColPart(input: String, start: Int): Option[ColPart] =
    val absolute = start < input.length && input.charAt(start) == '$'
    val lettersStart = if absolute then start + 1 else start
    val end = spanLetters(input, lettersStart)
    val count = end - lettersStart
    if count < 1 || count > 3 then None
    else
      Column
        .fromLetter(input.substring(lettersStart, end))
        .toOption
        .map(col => ColPart(col.index0, absolute, end))

  private def parseRowPart(input: String, start: Int): Option[RowPart] =
    val absolute = start < input.length && input.charAt(start) == '$'
    val digitsStart = if absolute then start + 1 else start
    val end = spanDigits(input, digitsStart)
    if end == digitsStart then None
    else
      input
        .substring(digitsStart, end)
        .toIntOption
        .filter(row1 => row1 >= 1 && row1 <= Row.MaxIndex0 + 1)
        .map(row1 => RowPart(row1 - 1, absolute, end))

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def spanLetters(input: String, start: Int): Int =
    var end = start
    while end < input.length && isAsciiLetter(input.charAt(end)) do end += 1
    end

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def spanDigits(input: String, start: Int): Int =
    var end = start
    while end < input.length && isAsciiDigit(input.charAt(end)) do end += 1
    end

  private def isAsciiLetter(ch: Char): Boolean =
    (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z')

  private def isAsciiDigit(ch: Char): Boolean = ch >= '0' && ch <= '9'

  private def renderCellRange(
    start: CellPart,
    end: CellPart,
    colDelta: Int,
    rowDelta: Int
  ): String =
    (shiftedCell(start, colDelta, rowDelta), shiftedCell(end, colDelta, rowDelta)) match
      case (Some(first), Some(second)) => s"$first:$second"
      case _ => "#REF!"

  private def renderColumnRange(start: ColPart, end: ColPart, colDelta: Int): String =
    (shiftedColumn(start, colDelta), shiftedColumn(end, colDelta)) match
      case (Some(first), Some(second)) => s"$first:$second"
      case _ => "#REF!"

  private def renderRowRange(start: RowPart, end: RowPart, rowDelta: Int): String =
    (shiftedRow(start, rowDelta), shiftedRow(end, rowDelta)) match
      case (Some(first), Some(second)) => s"$first:$second"
      case _ => "#REF!"

  private def renderCell(cell: CellPart, colDelta: Int, rowDelta: Int): String =
    shiftedCell(cell, colDelta, rowDelta).getOrElse("#REF!")

  private def shiftedCell(cell: CellPart, colDelta: Int, rowDelta: Int): Option[String] =
    val col = shiftedIndex(cell.colIndex0, cell.colAbsolute, colDelta, Column.MaxIndex0)
    val row = shiftedIndex(cell.rowIndex0, cell.rowAbsolute, rowDelta, Row.MaxIndex0)
    for
      colIndex <- col
      rowIndex <- row
    yield
      val colAnchor = if cell.colAbsolute then "$" else ""
      val rowAnchor = if cell.rowAbsolute then "$" else ""
      s"$colAnchor${Column.from0(colIndex).toLetter}$rowAnchor${rowIndex + 1}"

  private def shiftedColumn(column: ColPart, delta: Int): Option[String] =
    shiftedIndex(column.index0, column.absolute, delta, Column.MaxIndex0).map { index =>
      val anchor = if column.absolute then "$" else ""
      s"$anchor${Column.from0(index).toLetter}"
    }

  private def shiftedRow(row: RowPart, delta: Int): Option[String] =
    shiftedIndex(row.index0, row.absolute, delta, Row.MaxIndex0).map { index =>
      val anchor = if row.absolute then "$" else ""
      s"$anchor${index + 1}"
    }

  private def shiftedIndex(index: Int, absolute: Boolean, delta: Int, max: Int): Option[Int] =
    val shifted = if absolute then index.toLong else index.toLong + delta.toLong
    Option.when(shifted >= 0L && shifted <= max.toLong)(shifted.toInt)

  private def hasBoundaryBefore(input: String, start: Int): Boolean =
    start == 0 || !isNameChar(input.charAt(start - 1))

  private def hasBoundaryAfter(input: String, end: Int): Boolean =
    end >= input.length || !isNameChar(input.charAt(end))

  private def isNameChar(ch: Char): Boolean =
    ch.isLetterOrDigit || ch == '_' || ch == '.' || ch == '\\' || ch == '?' || ch.toInt > 127

  /**
   * A cell-shaped token is not a reference when it names a function (`LOG10(`), is a short table
   * name before a structured reference (`T1[...]`), or belongs to a sheet/3-D qualifier (`S1!` or
   * `S1:S2!`). Defined names cannot otherwise have the shape of a valid Excel cell reference.
   */
  private def isNonReferenceUse(input: String, end: Int): Boolean =
    nextNonWhitespace(input, end) match
      case Some(('(', _)) | Some(('[', _)) | Some(('!', _)) => true
      case Some((':', colon)) => qualifierBangFollows(input, colon + 1)
      case _ => false

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def nextNonWhitespace(input: String, start: Int): Option[(Char, Int)] =
    var index = start
    while index < input.length && input.charAt(index).isWhitespace do index += 1
    Option.when(index < input.length)((input.charAt(index), index))

  /** True when the remainder of `token:token!` is a 3-D sheet qualifier, not a cell range. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def qualifierBangFollows(input: String, start: Int): Boolean =
    var index = start
    var found = false
    var stopped = false
    while index < input.length && !found && !stopped do
      input.charAt(index) match
        case '!' => found = true
        case '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<' | '>' | ',' | ';' | '(' | ')' | '{' |
            '}' | ' ' | '\t' | '\r' | '\n' =>
          stopped = true
        case _ => index += 1
    found
