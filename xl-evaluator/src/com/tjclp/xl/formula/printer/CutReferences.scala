package com.tjclp.xl.formula.printer

import com.tjclp.xl.addressing.{ARef, SheetName}

/**
 * GH-935: Excel's rule for the formulas and names that refer to a cell CUT from `from` and PASTED
 * onto `to` on sheet `sheet` — what a merge over an empty upper-left cell does ("cut the winner
 * onto the anchor, then clear the rest"). Pinned against Excel for Mac 16.115 by the merge oracle
 * (`excel-oracle/merge-references.json`: 93 scenarios, 69 of them with a cut-and-paste control that
 * produced the same text character for character).
 *
 * Only references to `sheet` count: qualified with it, or unqualified in a formula on it (a defined
 * name's unqualified references count too). Each such reference becomes:
 *   - exactly `from` (a cell, or the one-cell range `from:from`) → `to`, `$` flags kept per end;
 *   - exactly `to` → `#REF!` (the pasted cell overwrote it; a qualified one keeps its qualifier);
 *   - a one-row (horizontal move) or one-column (vertical move) range on `from`'s line with `from`
 *     at an end: `from` first → the start moves to `to` (`B1:C1` → `A1:C1`); `from` last with `to`
 *     inside the range → the end moves back one (`A1:A5` → `A1:A4`); `from` last with `to` outside
 *     → unchanged;
 *   - anything else unchanged: 2-D ranges, `from` in the middle, diagonal moves, whole columns or
 *     rows, union and intersection areas, INDIRECT text.
 *
 * A formula is rewritten at all only when one of its references DEPENDS on `from` or `to` (Excel
 * leaves every other formula byte-identical): a cell reference depends on its cell, a range that is
 * a whole function argument (`SUM(B1:C1)`) or sits in an array formula on every cell in it, and a
 * bare range Excel implicitly intersects (`=B1:C1+0`, `=(B1:C1)`) only on its intersection with the
 * formula's own row or column. A defined name has no cell and is always rewritten.
 *
 * Text-level: only the changed reference spans are replaced ([[FormulaSpans]]); text the scanner
 * finds ambiguous (a 3-D reference) is returned unchanged.
 */
object CutReferences:

  /** Where the text being rewritten lives. */
  enum Host derives CanEqual:
    /**
     * A cell formula on `sheet` at `at`; `array` for an array formula (no implicit intersection).
     */
    case Cell(sheet: SheetName, at: ARef, array: Boolean)

    /**
     * A defined name's `refersTo`: always rewritten, its unqualified references on the cut sheet.
     */
    case Name

    /**
     * Conditional-format or data-validation text on `sheet` (not covered by the oracle): only exact
     * references to `from` follow it, as before; nothing becomes `#REF!`.
     */
    case Rule(sheet: SheetName)

  private final case class Box(r1: Int, c1: Int, r2: Int, c2: Int):
    def contains(r: Int, c: Int): Boolean = r1 <= r && r <= r2 && c1 <= c && c <= c2
    def isCell: Boolean = r1 == r2 && c1 == c2

  private enum Outcome:
    case Keep
    case Refs
    case To(box: Box)

  def rewrite(text: String, host: Host, sheet: SheetName, from: ARef, to: ARef): String =
    if from == to then text
    else
      FormulaSpans.scan(text).fold(text) { scan =>
        val (fr, fc) = (from.row.index0, from.col.index0)
        val (tr, tc) = (to.row.index0, to.col.index0)
        val masked = maskStrings(text)
        val ruleHost = host match
          case Host.Rule(_) => true
          case _ => false

        def onSheet(r: FormulaSpans.Reference): Boolean =
          r.qualifier.fold(host match
            case Host.Cell(s, _, _) => s.value.equalsIgnoreCase(sheet.value)
            case Host.Rule(s) => s.value.equalsIgnoreCase(sheet.value)
            case Host.Name => true)(_.name.equalsIgnoreCase(sheet.value))

        def box(r: FormulaSpans.Reference): Box =
          val a = r.first.ref
          val b = r.second.fold(a)(_.ref)
          Box(
            a.row.index0.min(b.row.index0),
            a.col.index0.min(b.col.index0),
            a.row.index0.max(b.row.index0),
            a.col.index0.max(b.col.index0)
          )

        def start(r: FormulaSpans.Reference): Int = r.qualifier.fold(r.first.start)(_.start)
        def end(r: FormulaSpans.Reference): Int = r.second.fold(r.first.end)(_.end)

        // the cells a bare (implicitly intersected) area depends on: its intersection with the
        // formula's own row or column; a single cell is itself
        def implicitly(b: Box, at: ARef): Option[Box] =
          val (row, col) = (at.row.index0, at.col.index0)
          if b.isCell then Some(b)
          else if b.c1 == b.c2 && b.r1 <= row && row <= b.r2 then Some(Box(row, b.c1, row, b.c1))
          else if b.r1 == b.r2 && b.c1 <= col && col <= b.c2 then Some(Box(b.r1, col, b.r1, col))
          else if b.contains(row, col) then Some(Box(row, col, row, col))
          else None

        def outcome(b: Box): Outcome =
          if b == Box(fr, fc, fr, fc) then Outcome.To(Box(tr, tc, tr, tc))
          else if ruleHost then Outcome.Keep
          else if b == Box(tr, tc, tr, tc) then Outcome.Refs
          else
            val horizontal = fr == tr && fc != tc
            val vertical = fc == tc && fr != tr
            if horizontal && b.r1 == b.r2 && b.r1 == fr then
              if b.c1 == fc then Outcome.To(b.copy(c1 = tc))
              else if b.c2 == fc && b.contains(tr, tc) then Outcome.To(b.copy(c2 = b.c2 - 1))
              else Outcome.Keep
            else if vertical && b.c1 == b.c2 && b.c1 == fc then
              if b.r1 == fr then Outcome.To(b.copy(r1 = tr))
              else if b.r2 == fr && b.contains(tr, tc) then Outcome.To(b.copy(r2 = b.r2 - 1))
              else Outcome.Keep
            else Outcome.Keep

        val refs = scan.references.filter(onSheet)
        // rule 5: a formula is touched only when one of its reference expressions depends on the
        // cut or the paste cell. An expression is a reference, an intersection of references
        // (`B1 B1:C1`: the cells they share) or a parenthesised union of those (`(B1:C1,D5)`: every
        // cell of every area); as a whole function argument (or in an array formula) it depends on
        // all the cells it denotes, as a bare value only on their intersection with the formula's
        // own row or column
        val affected = host match
          case Host.Name | Host.Rule(_) => true
          case Host.Cell(_, at, array) =>
            expressions(masked, scan.references).exists { (span, areasOf) =>
              val areas = areasOf.flatMap { group =>
                group
                  .filter(onSheet)
                  .map(box)
                  .reduceOption((a, b) => intersect(a, b).getOrElse(Box(-1, -1, -2, -2)))
                  .filter(b => b.r1 <= b.r2 && b.c1 <= b.c2)
                  // an intersection with an area on another sheet (or none on this one) is skipped
                  .filter(_ => group.forall(onSheet))
              }
              val whole = array || wholeArgument(masked, span._1, span._2)
              val deps = if whole then areas else areas.flatMap(implicitly(_, at))
              deps.exists(d => d.contains(fr, fc) || d.contains(tr, tc))
            }
        if !affected then text
        else
          val edits = refs.flatMap { r =>
            outcome(box(r)) match
              case Outcome.Keep => Vector.empty
              case Outcome.Refs =>
                val qualifier =
                  r.qualifier.fold("")(q => text.substring(q.start, q.end) + "!")
                Vector((start(r), end(r), qualifier + "#REF!"))
              case Outcome.To(nb) =>
                val first = r.first.spelledAt(ARef.from0(nb.c1, nb.r1))
                r.second match
                  case None => Vector((r.first.start, r.first.end, first))
                  case Some(second) =>
                    Vector(
                      (r.first.start, r.first.end, first),
                      (second.start, second.end, second.spelledAt(ARef.from0(nb.c2, nb.r2)))
                    )
          }
          FormulaSpans.replace(text, edits)
      }

  private def intersect(a: Box, b: Box): Option[Box] =
    val r = Box(a.r1.max(b.r1), a.c1.max(b.c1), a.r2.min(b.r2), a.c2.min(b.c2))
    Option.when(r.r1 <= r.r2 && r.c1 <= r.c2)(r)

  /**
   * The reference expressions of a formula, each with its text span and its areas: a vector of
   * intersection groups (references joined only by spaces), more than one when a non-call
   * parenthesis holds nothing but comma-separated groups (a union). Every reference belongs to
   * exactly one expression.
   */
  private def expressions(
    masked: String,
    refs: Vector[FormulaSpans.Reference]
  ): Vector[((Int, Int), Vector[Vector[FormulaSpans.Reference]])] =
    def start(r: FormulaSpans.Reference): Int = r.qualifier.fold(r.first.start)(_.start)
    def end(r: FormulaSpans.Reference): Int = r.second.fold(r.first.end)(_.end)
    // intersection groups: consecutive references separated by whitespace only
    val groups = refs.foldLeft(Vector.empty[Vector[FormulaSpans.Reference]]) { (acc, r) =>
      acc.lastOption match
        case Some(g) if {
              val gap = masked.substring(g.lastOption.fold(0)(end), start(r))
              gap.nonEmpty && gap.forall(_.isWhitespace)
            } =>
          acc.dropRight(1) :+ (g :+ r)
        case _ => acc :+ Vector(r)
    }
    def span(g: Vector[FormulaSpans.Reference]): (Int, Int) =
      (g.headOption.fold(0)(start), g.lastOption.fold(0)(end))
    // non-call parenthesis pairs whose content is only comma-separated groups: unions
    val pairs = masked.zipWithIndex
      .foldLeft((List.empty[(Int, Boolean)], Vector.empty[(Int, Int)])) {
        case ((stack, found), (c, i)) =>
          if c == '(' then
            val before = if i > 0 then masked.charAt(i - 1) else ' '
            ((i, before.isLetterOrDigit || before == '.' || before == '_') :: stack, found)
          else if c == ')' then
            stack match
              case (open, call) :: rest => (rest, if call then found else found :+ (open, i))
              case Nil => (stack, found)
          else (stack, found)
      }
      ._2
    val unions = pairs.flatMap { (open, close) =>
      val inner = groups.filter(g => span(g)._1 > open && span(g)._2 <= close)
      val pieces = masked.substring(open + 1, close).split(",", -1).map(_.trim).toVector
      val matches = inner.map(g => masked.substring(span(g)._1, span(g)._2).trim)
      Option.when(inner.size >= 2 && pieces == matches)(((open, close + 1), inner))
    }
    val inUnion = unions.flatMap(_._2).toSet
    unions ++ groups.filterNot(inUnion.contains).map(g => (span(g), Vector(g)))

  private def maskStrings(text: String): String =
    text
      .foldLeft((new StringBuilder(text.length), false)) { case ((sb, inString), c) =>
        if c == '"' then (sb.append(' '), !inString)
        else (sb.append(if inString then ' ' else c), inString)
      }
      ._1
      .toString

  /**
   * Whether the reference at `[start, end)` is a whole argument of a function call: the innermost
   * open parenthesis before it belongs to a call (a name right before it), and the reference is the
   * whole argument (`(` or `,` before it, `)` or `,` after it, spaces aside).
   */
  private def wholeArgument(masked: String, start: Int, end: Int): Boolean =
    val calls = masked
      .substring(0, start)
      .zipWithIndex
      .foldLeft(List.empty[Boolean]) { case (stack, (c, i)) =>
        if c == '(' then
          val before = if i > 0 then masked.charAt(i - 1) else ' '
          (before.isLetterOrDigit || before == '.' || before == '_') :: stack
        else if c == ')' then stack.drop(1)
        else stack
      }
    val inCall = calls.headOption.contains(true)
    val prev = masked.substring(0, start).reverse.find(!_.isWhitespace)
    val next = masked.substring(end).find(!_.isWhitespace)
    inCall && prev.exists(c => c == '(' || c == ',') && next.exists(c => c == ')' || c == ',')
