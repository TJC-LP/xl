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

        // the cells a reference depends on (rule 5): a bare range Excel implicitly intersects only
        // depends on its intersection with the formula's own row or column
        def dependsOn(r: FormulaSpans.Reference): Option[Box] =
          val b = box(r)
          host match
            case Host.Cell(_, at, array)
                if r.second.isDefined && !b.isCell && !array &&
                  !wholeArgument(masked, start(r), end(r)) =>
              val (row, col) = (at.row.index0, at.col.index0)
              if b.c1 == b.c2 && b.r1 <= row && row <= b.r2 then Some(Box(row, b.c1, row, b.c1))
              else if b.r1 == b.r2 && b.c1 <= col && col <= b.c2 then
                Some(Box(b.r1, col, b.r1, col))
              else if b.contains(row, col) then Some(Box(row, col, row, col))
              else None
            case _ => Some(b)

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
        val affected = host match
          case Host.Name | Host.Rule(_) => true
          case Host.Cell(_, _, _) =>
            refs.exists(r => dependsOn(r).exists(d => d.contains(fr, fc) || d.contains(tr, tc)))
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

  /** `text` with string-literal contents blanked (same length), so parentheses inside are inert. */
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
