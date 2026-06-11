package com.tjclp.xl.cli.helpers

/**
 * Parser for GFM (GitHub Flavored Markdown) pipe tables (GH-159).
 *
 * Recognizes the first table in the input: a header row containing at least one pipe, followed by a
 * delimiter row (cells of `-` with optional `:` alignment markers), followed by body rows. The
 * table ends at the first blank line or the first line without a pipe.
 *
 * Normalization rules:
 *   - Outer (leading/trailing) pipes are optional
 *   - `\|` inside a cell is a literal pipe (other backslash sequences are kept verbatim)
 *   - Cell content is trimmed of surrounding whitespace
 *   - The delimiter row fixes the column count; header/body rows are padded or truncated to it
 *
 * Total: never throws; malformed input yields `Left` with a description.
 */
object MarkdownTableParser:

  /** Column alignment from GFM delimiter markers: `:---` left, `:---:` center, `---:` right. */
  enum ColumnAlignment derives CanEqual:
    case None, Left, Center, Right

  /**
   * Parsed table: `rows` contains the header row first, then body rows. All rows have exactly
   * `alignments.length` cells.
   */
  final case class MdTable(
    rows: Vector[Vector[String]],
    alignments: Vector[ColumnAlignment]
  ):
    def columnCount: Int = alignments.length

  /** Parse the first GFM pipe table found in `input`. */
  def parse(input: String): Either[String, MdTable] =
    val lines = input.linesIterator.toVector
    val headerIdx = lines.indices.find { i =>
      lines(i).contains('|') && lines.lift(i + 1).exists(isDelimiterRow)
    }
    headerIdx match
      case scala.None =>
        Left(
          "no markdown table found (expected a header row followed by a delimiter row like |---|---|)"
        )
      case Some(idx) =>
        val delimiterCells = rowCells(lines(idx + 1))
        val alignments = delimiterCells.map(alignmentOf)
        val width = alignments.length
        val header = normalizeRow(rowCells(lines(idx)), width)
        val body = lines
          .drop(idx + 2)
          .takeWhile(line => line.trim.nonEmpty && line.contains('|'))
          .map(line => normalizeRow(rowCells(line), width))
        Right(MdTable(header +: body, alignments))

  /** True if the line is a GFM delimiter row: cells of `:?-+:?` with at least one pipe. */
  private def isDelimiterRow(line: String): Boolean =
    line.contains('|') && {
      val cells = rowCells(line)
      cells.nonEmpty && cells.forall(_.matches(":?-+:?"))
    }

  private def alignmentOf(cell: String): ColumnAlignment =
    val left = cell.startsWith(":")
    val right = cell.endsWith(":")
    if left && right then ColumnAlignment.Center
    else if left then ColumnAlignment.Left
    else if right then ColumnAlignment.Right
    else ColumnAlignment.None

  /** Pad with empty cells or truncate to exactly `width` cells. */
  private def normalizeRow(cells: Vector[String], width: Int): Vector[String] =
    cells.padTo(width, "").take(width)

  /**
   * Split a table line into trimmed cells, honoring `\|` escapes and dropping the empty segments
   * produced by optional outer pipes.
   */
  private def rowCells(line: String): Vector[String] =
    val segments = splitOnUnescapedPipes(line)
    val trimmedLine = line.trim
    val withoutLeading =
      if trimmedLine.startsWith("|") && segments.headOption.exists(_.trim.isEmpty) then
        segments.drop(1)
      else segments
    val withoutOuter =
      if trimmedLine.endsWith("|") && withoutLeading.lastOption.exists(_.trim.isEmpty) then
        withoutLeading.dropRight(1)
      else withoutLeading
    withoutOuter.map(_.trim)

  /** Split on `|`, treating `\|` as a literal pipe. Other backslash pairs pass through verbatim. */
  private def splitOnUnescapedPipes(line: String): Vector[String] =
    @annotation.tailrec
    def loop(i: Int, current: String, acc: Vector[String], escaped: Boolean): Vector[String] =
      if i >= line.length then
        val last = if escaped then current + "\\" else current
        acc :+ last
      else
        val c = line.charAt(i)
        if escaped then
          val piece = if c == '|' then "|" else s"\\$c"
          loop(i + 1, current + piece, acc, escaped = false)
        else if c == '\\' then loop(i + 1, current, acc, escaped = true)
        else if c == '|' then loop(i + 1, "", acc :+ current, escaped = false)
        else loop(i + 1, current + c, acc, escaped = false)
    loop(0, "", Vector.empty, escaped = false)
