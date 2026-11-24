package com.tjclp.xl.formula.printer

import com.tjclp.xl.addressing.{ARef, Column}

/**
 * The byte spans of formula text that a text-level rewrite may touch (GH-936, GH-935): sheet
 * qualifiers and cell references, found by a small scanner that skips string literals, bracketed
 * parts (external-workbook indices, structured references) and error literals. A rewrite replaces
 * only these spans, so every other byte of what the author typed (spacing, parentheses, number
 * spelling, function-name case) is kept, as Excel keeps it.
 *
 * The scanner is deliberately conservative: anything it cannot classify with certainty (an
 * unterminated literal or quoted name, a 3-D reference `Sheet1:Sheet3!A1`, a quoted name not
 * followed by `!`, a qualified range end) makes [[scan]] return `None`, and callers then leave the
 * text alone or fall back to the parser. Callers also verify a rewrite against the parser.
 */
private[printer] object FormulaSpans:

  /** A sheet qualifier: `[start, end)` covers the name as typed (quotes included), not the `!`. */
  final case class Qualifier(start: Int, end: Int, name: String) derives CanEqual

  /** A cell reference as typed: `[start, end)`, its cell, and its `$` anchors. */
  final case class CellToken(start: Int, end: Int, ref: ARef, colAbs: Boolean, rowAbs: Boolean)
      derives CanEqual:
    /** The same anchors at another cell. */
    def spelledAt(at: ARef): String =
      (if colAbs then "$" else "") + Column.from0(at.col.index0).toLetter +
        (if rowAbs then "$" else "") + (at.row.index0 + 1).toString

  /** A cell or cell-to-cell range reference with its qualifier, when it has one. */
  final case class Reference(
    qualifier: Option[Qualifier],
    first: CellToken,
    second: Option[CellToken]
  ) derives CanEqual

  final case class Scan(qualifiers: Vector[Qualifier], references: Vector[Reference])
      derives CanEqual

  private val CellShape = "^(\\$?)([A-Za-z]{1,3})(\\$?)([0-9]+)$".r

  private def isWordChar(c: Char): Boolean = c.isLetterOrDigit || c == '_' || c == '.' || c == '$'

  private def cellToken(word: String, start: Int): Option[CellToken] =
    word match
      case CellShape(colAbs, letters, rowAbs, digits) =>
        ARef
          .parse(letters + digits)
          .toOption
          .map(ref => CellToken(start, start + word.length, ref, colAbs.nonEmpty, rowAbs.nonEmpty))
      case _ => None

  /** The spans of `text`, or `None` when any part of it is ambiguous to a text-level rewrite. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def scan(text: String): Option[Scan] =
    val n = text.length
    val quals = Vector.newBuilder[Qualifier]
    val refs = Vector.newBuilder[Reference]
    var i = 0
    var ok = true
    // the next qualifier belongs to another workbook (`[1]Sheet1!A1`, `'[1]Sheet1'!A1`)
    var external = false

    def charAt(j: Int): Char = if j >= 0 && j < n then text.charAt(j) else '\u0000'
    def wordEnd(from: Int): Int =
      var j = from
      while j < n && isWordChar(text.charAt(j)) do j += 1
      j

    /**
     * After a `!` at `bang`: the reference it qualifies, recorded under `q`; returns the next
     * index.
     */
    def qualified(q: Option[Qualifier], bang: Int): Int =
      val start = bang + 1
      val end = wordEnd(start)
      cellToken(text.substring(start, end), start) match
        case Some(first) if charAt(end) != '(' =>
          if charAt(end) == ':' then
            val s2 = end + 1
            val e2 = wordEnd(s2)
            if charAt(e2) == '!' || charAt(e2) == '\'' then
              ok = false
              e2
            else
              cellToken(text.substring(s2, e2), s2) match
                case Some(second) =>
                  if !external then refs += Reference(q, first, Some(second))
                  e2
                case None => end
          else
            if !external then refs += Reference(q, first, None)
            end
        case _ => start

    while ok && i < n do
      val c = text.charAt(i)
      if c == '"' then
        var j = i + 1
        var closed = false
        while !closed && j < n do
          if text.charAt(j) == '"' then
            if charAt(j + 1) == '"' then j += 2
            else
              closed = true
              j += 1
          else j += 1
        if !closed then ok = false
        i = j
      else if c == '[' then
        var depth = 0
        var j = i
        var closed = false
        while !closed && j < n do
          val d = text.charAt(j)
          if d == '[' then depth += 1
          else if d == ']' then
            depth -= 1
            if depth == 0 then closed = true
          j += 1
        if !closed then ok = false
        external = true
        i = j
      else if c == '\'' then
        var j = i + 1
        var closed = false
        val name = new StringBuilder
        while !closed && j < n do
          if text.charAt(j) == '\'' then
            if charAt(j + 1) == '\'' then
              name.append('\'')
              j += 2
            else
              closed = true
              j += 1
          else
            name.append(text.charAt(j))
            j += 1
        if !closed || charAt(j) != '!' || charAt(i - 1) == ':' then ok = false
        else
          val ext = external || name.headOption.contains('[')
          val q = Qualifier(i, j, name.toString)
          if !ext then quals += q
          external = ext
          i = qualified(Some(q), j)
          external = false
      else if c == '#' then
        var j = i + 1
        while j < n && (text.charAt(j).isLetterOrDigit || text.charAt(j) == '/' ||
            text.charAt(j) == '_')
        do j += 1
        if charAt(j) == '!' || charAt(j) == '?' then j += 1
        i = j
      else if isWordChar(c) then
        val end = wordEnd(i)
        val word = text.substring(i, end)
        if charAt(end) == '!' then
          if charAt(i - 1) == ':' then ok = false
          else
            val q = Qualifier(i, end, word)
            if !external then quals += q
            i = qualified(Some(q), end)
            external = false
        else if charAt(end) == ':' && {
            val e2 = wordEnd(end + 1)
            charAt(e2) == '!'
          }
        then ok = false // a 3-D reference: Sheet1:Sheet3!A1
        else
          external = false
          cellToken(word, i) match
            case Some(first) if charAt(end) != '(' =>
              if charAt(end) == ':' then
                val s2 = end + 1
                val e2 = wordEnd(s2)
                cellToken(text.substring(s2, e2), s2) match
                  case Some(second) =>
                    refs += Reference(None, first, Some(second))
                    i = e2
                  case None => i = end
              else
                refs += Reference(None, first, None)
                i = end
            case _ => i = end
      else
        if !c.isWhitespace then external = false
        i += 1
    Option.when(ok)(Scan(quals.result(), refs.result()))

  /** `text` with each `(start, end, replacement)` applied; the spans must not overlap. */
  def replace(text: String, edits: Vector[(Int, Int, String)]): String =
    edits.sortBy(_._1).reverse.foldLeft(text) { case (acc, (start, end, by)) =>
      acc.substring(0, start) + by + acc.substring(end)
    }
