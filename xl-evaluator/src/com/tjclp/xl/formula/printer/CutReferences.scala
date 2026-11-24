package com.tjclp.xl.formula.printer

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.formula.ast.{BinarySpine, TExpr}
import com.tjclp.xl.formula.functions.ArgValue
import com.tjclp.xl.formula.parser.FormulaParser

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
 * Text-level: only the changed reference spans are replaced ([[FormulaSpans]]); a 3-D reference is
 * one opaque span (Excel leaves it alone), and text the scanner finds ambiguous is returned
 * unchanged. Data-validation and conditional-format formulas are hosted at their rule's first cell.
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

        def onSheet(r: FormulaSpans.Reference): Boolean =
          r.qualifier.fold(host match
            case Host.Cell(s, _, _) => s.value.equalsIgnoreCase(sheet.value)
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
        // (`B1 B1:C1`) or a parenthesised union of those (`(B1:C1,D5)`), and Excel's dependency
        // for the last two is conservative (checked in Excel for Mac): every cell of every operand
        // area, not the cells the areas share, wherever the formula sits. Only a lone range that
        // is not a whole function argument (in a plain cell) narrows to its implicit intersection
        // with the formula's own row or column (the oracle's rule)
        val affected = host match
          case Host.Name => true
          case Host.Cell(hostSheet, at, array) =>
            expressions(masked, scan.references).exists { (span, areasOf) =>
              val operands = areasOf.flatten
              val areas = operands.filter(onSheet).map(box)
              val whole = array || wholeArgument(masked, span._1, span._2)
              // only a lone bare range is implicitly intersected; an intersection or union depends
              // on every cell of its operand areas wherever it sits (checked in Excel for Mac)
              val deps =
                if whole || operands.size > 1 then areas
                else areas.flatMap(implicitly(_, at))
              deps.exists(d => d.contains(fr, fc) || d.contains(tr, tc))
            } || (!array && arrayArgumentsTouch(text, hostSheet, sheet, fr, fc, tr, tc)) ||
            (!array && refs.exists { r =>
              val b = box(r)
              (b.contains(fr, fc) || b.contains(tr, tc)) && inArrayArgument(masked, start(r))
            })
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

  /**
   * GH-935 review 4 (checked in Excel for Mac): an argument whose parameter Excel types as an array
   * (SUMPRODUCT's, INDEX's first, ROWS', COLUMNS': `FunctionFlags.arrayArguments`, the registry's
   * own declaration) is evaluated whole in a plain cell too, so every range anywhere inside it —
   * `SUMPRODUCT((B1:C1>0)*B1:C1)`, `INDEX(B1:C1*1,1,1)` — depends on all its cells. Whether one of
   * those ranges on the cut sheet holds the cut or the paste cell. Text that does not parse has no
   * such argument here (the text gate still applies).
   */
  private def arrayArgumentsTouch(
    text: String,
    host: SheetName,
    sheet: SheetName,
    fr: Int,
    fc: Int,
    tr: Int,
    tc: Int
  ): Boolean =
    FormulaParser.parse(text).toOption.exists { expr =>
      arrayArgumentRanges(expr).exists { (qualifier, range) =>
        qualifier.getOrElse(host).value.equalsIgnoreCase(sheet.value) &&
        (range.contains(ARef.from0(fc, fr)) || range.contains(ARef.from0(tc, tr)))
      }
    }

  /** The ranges inside the array-typed arguments of every call in `expr`. */
  private def arrayArgumentRanges(expr: TExpr[?]): List[(Option[SheetName], CellRange)] =
    import TExpr.*
    def whole(value: ArgValue): List[(Option[SheetName], CellRange)] = value match
      case ArgValue.Expr(e) => TExpr.collectRanges(e)
      case ArgValue.Range(RangeLocation.Local(range, _)) => List((Option.empty[SheetName], range))
      case ArgValue.Range(RangeLocation.CrossSheet(s, range, _)) => List((Some(s), range))
      case ArgValue.Range(RangeLocation.Computed(e)) => TExpr.collectRanges(e)
      case ArgValue.Range(_) => Nil
    def inner(value: ArgValue): List[(Option[SheetName], CellRange)] = value match
      case ArgValue.Expr(e) => arrayArgumentRanges(e)
      case ArgValue.Range(RangeLocation.Computed(e)) => arrayArgumentRanges(e)
      case ArgValue.Range(_) => Nil
    expr match
      case call: Call[?] =>
        call.spec.argSpec.toValues(call.args).zipWithIndex.flatMap { (value, position) =>
          (if call.spec.flags.arrayArguments.contains(position) then whole(value) else Nil) ++
            inner(value)
        }
      case chain @ (_: Add | _: Sub | _: Mul | _: Div | _: Pow | _: Concat | _: Eq[?] | _: Neq[?] |
          _: Lt[?] | _: Lte[?] | _: Gt[?] | _: Gte[?]) =>
        BinarySpine.operandList(chain).flatMap(arrayArgumentRanges)
      case ToInt(e) => arrayArgumentRanges(e)
      case DateToSerial(e) => arrayArgumentRanges(e)
      case DateTimeToSerial(e) => arrayArgumentRanges(e)
      case UnaryPlus(e) => arrayArgumentRanges(e)
      case Percent(e) => arrayArgumentRanges(e)
      case Let(bindings, body) =>
        bindings.flatMap((_, value) => arrayArgumentRanges(value)) ++ arrayArgumentRanges(body)
      case Coerced(e, _) => arrayArgumentRanges(e)
      case _ => Nil

  /**
   * GH-935 review 5/6: Excel's array-typed parameters by function, as 0-based argument positions
   * (`None`: every argument), for formulas read as text. Checked in Excel for Mac (merge A1:B1 with
   * the value in B1 rewrites `B1:C1*1` in each): SUMPRODUCT, INDEX (first), ROWS, COLUMNS, LOOKUP
   * (lookup and result vectors), IRR, MIRR (values), XLOOKUP (lookup and return arrays), XMATCH
   * (lookup array), SORT, UNIQUE, MMULT, SUMX2MY2, CORREL, FILTER (array, include), SORTBY, TOCOL,
   * VSTACK, FORECAST.LINEAR (known ys and xs), COVARIANCE.P. From Excel's function reference, not
   * checked by a merge: SUMX2PY2, SUMXMY2, PEARSON, RSQ, SLOPE, INTERCEPT, COVAR, FORECAST (known
   * ys and xs), FREQUENCY, TREND, GROWTH, LINEST, LOGEST, MDETERM, MINVERSE, HSTACK, TOROW, TAKE,
   * DROP, CHOOSECOLS, CHOOSEROWS, WRAPROWS, WRAPCOLS, EXPAND (array), COVARIANCE.S, the
   * FORECAST.ETS family (values, timeline). Excel leaves the value-typed ones alone, so they are
   * not here (checked): LARGE, SMALL, PERCENTILE, QUARTILE, TRIMMEAN, MATCH, TRANSPOSE, SUM, LET, a
   * LAMBDA call. Every function the registry flags (`FunctionFlags.arrayArguments`) is here too (a
   * law pins it).
   */
  val ArrayParameters: Map[String, Option[Set[Int]]] = Map(
    "SUMPRODUCT" -> None,
    "INDEX" -> Some(Set(0)),
    "ROWS" -> None,
    "COLUMNS" -> None,
    "LOOKUP" -> Some(Set(1, 2)),
    "IRR" -> Some(Set(0)),
    "MIRR" -> Some(Set(0)),
    "XLOOKUP" -> Some(Set(1, 2)),
    "XMATCH" -> Some(Set(1)),
    "SORT" -> Some(Set(0)),
    "UNIQUE" -> Some(Set(0)),
    "MMULT" -> None,
    "SUMX2MY2" -> None,
    "CORREL" -> None,
    "SUMX2PY2" -> None,
    "SUMXMY2" -> None,
    "PEARSON" -> None,
    "RSQ" -> None,
    "SLOPE" -> None,
    "INTERCEPT" -> None,
    "COVAR" -> None,
    "FORECAST" -> Some(Set(1, 2)),
    "FREQUENCY" -> None,
    "TREND" -> Some(Set(0, 1, 2)),
    "GROWTH" -> Some(Set(0, 1, 2)),
    "LINEST" -> Some(Set(0, 1)),
    "LOGEST" -> Some(Set(0, 1)),
    "MDETERM" -> None,
    "MINVERSE" -> None,
    "FILTER" -> Some(Set(0, 1)),
    "SORTBY" -> None,
    "TOCOL" -> Some(Set(0)),
    "VSTACK" -> None,
    "FORECAST.LINEAR" -> Some(Set(1, 2)),
    "COVARIANCE.P" -> None,
    "HSTACK" -> None,
    "TOROW" -> Some(Set(0)),
    "TAKE" -> Some(Set(0)),
    "DROP" -> Some(Set(0)),
    "CHOOSECOLS" -> Some(Set(0)),
    "CHOOSEROWS" -> Some(Set(0)),
    "WRAPROWS" -> Some(Set(0)),
    "WRAPCOLS" -> Some(Set(0)),
    "EXPAND" -> Some(Set(0)),
    "COVARIANCE.S" -> None,
    "FORECAST.ETS" -> Some(Set(1, 2)),
    "FORECAST.ETS.CONFINT" -> Some(Set(1, 2)),
    "FORECAST.ETS.SEASONALITY" -> Some(Set(0, 1)),
    "FORECAST.ETS.STAT" -> Some(Set(0, 1))
  )

  /**
   * Whether the text at `pos` lies, at any depth, inside an argument whose parameter Excel types as
   * an array ([[ArrayParameters]]): walks the parentheses before it, tracking each call's name and
   * argument index. Works on text the parser rejects (XLOOKUP, IRR, LOOKUP with expression
   * arguments), which is where it matters; strings, quoted sheet names, array constants and
   * brackets are masked ([[maskStrings]]), so only the formula's own commas count.
   */
  private def inArrayArgument(masked: String, pos: Int): Boolean =
    // a frame: the function's array positions (if it is a call) and the current argument index
    final case class Frame(arrays: Option[Option[Set[Int]]], arg: Int)
    val frames =
      masked.substring(0, pos).zipWithIndex.foldLeft(List.empty[Frame]) { case (stack, (c, i)) =>
        if c == '(' then
          val name = masked
            .substring(0, i)
            .reverse
            .takeWhile(ch => ch.isLetterOrDigit || ch == '.' || ch == '_')
            .reverse
            .toUpperCase(java.util.Locale.ROOT)
            .stripPrefix("_XLFN.")
            .stripPrefix("_XLWS.")
          Frame(ArrayParameters.get(name), 0) :: stack
        else if c == ')' then stack.drop(1)
        else if c == ',' then
          stack match
            case top :: rest => top.copy(arg = top.arg + 1) :: rest
            case Nil => Nil
        else stack
      }
    frames.exists(f => f.arrays.exists(_.forall(_.contains(f.arg))))

  /** Where [[maskStrings]] is in the text: plain formula, a string literal, a quoted sheet name. */
  private enum Mask derives CanEqual:
    case Plain, Str, Quoted

  /**
   * `text` with the punctuation that is not formula structure blanked (same length, so every span
   * keeps its position): string literals and quoted sheet names (`'a,b'!`, `'Q)1'!`, `''` escapes
   * included) whole, and the contents of array constants (`{1,2}`) and brackets (structured
   * references `T[[#This Row],[a]]`, external-book prefixes; both nest, and a `'` inside brackets
   * escapes the next character), whose braces and brackets stay. So the commas and parentheses left
   * are the formula's own: argument separators, unions and calls (GH-935 review 6).
   */
  private def maskStrings(text: String): String =
    final case class State(mode: Mask, braces: Int, brackets: Int, escape: Boolean)
    val sb = new StringBuilder(text.length)
    text.foldLeft(State(Mask.Plain, 0, 0, escape = false)) { (st, c) =>
      st.mode match
        case Mask.Str =>
          sb.append(' ')
          if c == '"' then st.copy(mode = Mask.Plain) else st
        case Mask.Quoted =>
          sb.append(' ')
          if c == '\'' then st.copy(mode = Mask.Plain) else st
        case Mask.Plain =>
          val nested = st.braces > 0 || st.brackets > 0
          val structural = c == '{' || c == '}' || c == '[' || c == ']'
          if st.escape then
            sb.append(' ')
            st.copy(escape = false)
          else if st.brackets > 0 then
            sb.append(if structural then c else ' ')
            c match
              case '\'' => st.copy(escape = true)
              case '[' => st.copy(brackets = st.brackets + 1)
              case ']' => st.copy(brackets = st.brackets - 1)
              case _ => st
          else
            c match
              case '"' =>
                sb.append(' ')
                st.copy(mode = Mask.Str)
              case '\'' if !nested =>
                sb.append(' ')
                st.copy(mode = Mask.Quoted)
              case '{' =>
                sb.append(c)
                st.copy(braces = st.braces + 1)
              case '}' =>
                sb.append(c)
                st.copy(braces = (st.braces - 1).max(0))
              case '[' =>
                sb.append(c)
                st.copy(brackets = 1)
              case _ =>
                sb.append(if nested then ' ' else c)
                st
    }
    sb.toString

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
