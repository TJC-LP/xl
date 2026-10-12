package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{RangeForm, TExpr}

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}

/**
 * GH-631: Excel's rule for `SUMIF`'s `sum_range` and `AVERAGEIF`'s `average_range` — the third
 * argument "does not have to be the same size and shape as `range`": the cells actually summed
 * start at its upper-left cell and take `range`'s size and shape. `SUMIF(A1:A10, ">0", C1)` reads
 * C1:C10; `SUMIF(A1:A3, ">1", C1:D5)` reads C1:C3. The rule is applied in two places that must
 * agree — evaluation ([[FunctionSpecsAggregate]]) and dependency extraction
 * ([[com.tjclp.xl.formula.graph.DependencyGraph]]), so an edit to C5 dirties the formula that reads
 * it — and this object is the one statement of it. A pairing that would run past the grid edge
 * (`SUMIF(A1:A3, ">1", Z1048575)`) contributes nothing there, as LibreOffice computes it.
 *
 * The `*IFS` family is different: Excel requires every `criteria_range` to match `sum_range`'s
 * shape and returns `#VALUE!` otherwise; nothing is resized.
 */
object CriteriaRangeResize:

  /** The functions whose third argument is sized to their first. */
  val resizingFunctions: Set[String] = Set("SUMIF", "AVERAGEIF")

  /**
   * Whether resizing can read cells the static dependency graph cannot know. Resizing is static
   * when every range it uses is: a written range, or a defined name bound to one (the graph
   * resolves it through the workbook and records the resized target, #733). It is dynamic when the
   * shape or the target's origin is known only at evaluation: a LET name, or a defined name whose
   * definition computes the reference (OFFSET, INDIRECT, IF over references). Such readers take the
   * dynamic-recalculation path, which invalidates them on edits and evaluates them after static
   * precedents with dynamic caches cleared. A two-argument call does not resize.
   *
   * @param name
   *   the call's canonical (upper-case) function name, `FunctionSpec.name`
   * @param nameIsComputed
   *   whether a defined name in a range slot computes its reference; a name that does not resolve
   *   at all reads nothing (its evaluation fails), so it is not dynamic either
   */
  private[formula] def hasDynamicResize(
    name: String,
    values: List[ArgValue],
    nameIsComputed: TExpr.RangeLocation.Name => Boolean
  ): Boolean =
    def unresolved(location: TExpr.RangeLocation): Boolean = location match
      case named: TExpr.RangeLocation.Name => nameIsComputed(named)
      case TExpr.RangeLocation.Binding(_) => true
      case _ => false

    resizingFunctions.contains(name) && (values match
      case ArgValue.Range(criteria) :: _ :: ArgValue.Range(target) :: _ =>
        unresolved(criteria) || unresolved(target)
      case _ => false)

  /**
   * A call to SUMIF or AVERAGEIF — not SUMIFS or AVERAGEIFS, which never resize. Match the parser's
   * Unicode case folding and Char.isWhitespace so valid spellings cannot evade discovery.
   */
  private val ResizingCall = java.util.regex.Pattern.compile(
    "\\b(?:SUMIF|AVERAGEIF)\\p{javaWhitespace}*\\(",
    java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.UNICODE_CASE |
      java.util.regex.Pattern.UNICODE_CHARACTER_CLASS
  )

  /**
   * The text pre-filter for [[hasDynamicResize]]: whether a formula's text could hold a SUMIF or
   * AVERAGEIF call whose criteria or target range is a defined name or a LET name. The dependency
   * graph parses only such formulas when it looks for dynamic cells, so books full of SUMIFS, or of
   * SUMIF over written ranges, pay no parse for it.
   *
   * Sound: a false negative would leave a dynamic reader on the static schedule, so the test only
   * rules a formula out when no token in it can parse as a name. A defined name and a LET name are
   * both name-shaped identifiers, as [[com.tjclp.xl.formula.parser.FormulaParser]] reads them: a
   * letter, an underscore or a backslash first, then letters, digits, underscores, dots and
   * backslashes, and not a cell reference. An identifier is not a name when it is a function (`(`
   * follows), a sheet (`!` follows at once), a range corner (`:` follows or precedes it), holds `$`
   * (an anchored reference), is a cell reference, or is `TRUE` or `FALSE`; string literals, quoted
   * sheet names, `[…]` and error literals hold no names. Anything else counts as a possible name: a
   * false positive costs one parse.
   */
  private[formula] def mayResizeThroughName(formula: String): Boolean =
    mayResizeThroughName(formula, formula.toUpperCase(java.util.Locale.ROOT))

  /**
   * [[mayResizeThroughName]] for a caller already holding `upperCased`, the formula's
   * `toUpperCase(Locale.ROOT)`. A substring test on it rules out almost every formula before the
   * Unicode-aware regex runs, and it cannot rule out a call the parser reads: the parser
   * upper-cases a function name with `String.toUpperCase` before the registry lookup, so every
   * spelling it accepts as SUMIF or AVERAGEIF upper-cases to that text. That includes the long s
   * (`ſ` is `S`) and the dotless i (`ı` is `I`); a whitespace character between the name and `(`
   * lies outside the substring.
   */
  private[formula] def mayResizeThroughName(formula: String, upperCased: String): Boolean =
    (upperCased.contains("SUMIF") || upperCased.contains("AVERAGEIF")) &&
      ResizingCall.matcher(formula).find() && hasNameShapedToken(formula)

  private def isIdentStart(c: Char): Boolean = c.isLetter || c == '_' || c == '\\'

  private def isIdentPart(c: Char): Boolean =
    c.isLetterOrDigit || c == '_' || c == '.' || c == '\\' || c == '$'

  /** The index after the run of characters from `from` that satisfy `p`. */
  @scala.annotation.tailrec
  private def skipWhile(text: String, from: Int, p: Char => Boolean): Int =
    if from < text.length && p(text.charAt(from)) then skipWhile(text, from + 1, p) else from

  /** The index after a `quote`-delimited literal opening at `from` (a doubled quote escapes). */
  @scala.annotation.tailrec
  private def skipQuoted(text: String, from: Int, quote: Char): Int =
    val close = text.indexOf(quote.toInt, from + 1)
    if close < 0 then text.length
    else if close + 1 < text.length && text.charAt(close + 1) == quote then
      skipQuoted(text, close + 1, quote)
    else close + 1

  private def nextNonSpace(text: String, from: Int): Option[Char] =
    val at = skipWhile(text, from, _.isWhitespace)
    Option.when(at < text.length)(text.charAt(at))

  private def isName(token: String, text: String, end: Int): Boolean =
    !token.contains('$') &&
      ARef.parse(token).isLeft &&
      !token.equalsIgnoreCase("TRUE") && !token.equalsIgnoreCase("FALSE") &&
      !(end < text.length && text.charAt(end) == '!') &&
      !nextNonSpace(text, end).exists(c => c == '(' || c == ':')

  @scala.annotation.tailrec
  private def hasNameShapedToken(text: String, from: Int = 0): Boolean =
    if from >= text.length then false
    else
      text.charAt(from) match
        case '"' => hasNameShapedToken(text, skipQuoted(text, from, '"'))
        case '\'' => hasNameShapedToken(text, skipQuoted(text, from, '\''))
        case '[' =>
          val close = text.indexOf(']'.toInt, from)
          hasNameShapedToken(text, if close < 0 then text.length else close + 1)
        // an error literal (#REF!, #N/A, #DIV/0!) or the spill operator
        case '#' =>
          hasNameShapedToken(
            text,
            skipWhile(text, from + 1, c => c.isLetterOrDigit || "/_!?".contains(c))
          )
        // a range's second corner (A1:B2, A:C, 1:3) is never a name, as the parser reads it
        // today. GH-713's range operator (`A1:myName`, `myName:B3`) will make both corners
        // possible names: this arm, and the `:` look-ahead in isName, must accept them then
        case ':' =>
          hasNameShapedToken(text, skipWhile(text, from + 1, c => c.isLetterOrDigit || c == '$'))
        // a number (1, 2.5, 1E5): never a name
        case c if c.isDigit || c == '.' =>
          hasNameShapedToken(text, skipWhile(text, from, isIdentPart))
        case c if isIdentStart(c) || c == '$' =>
          val end = skipWhile(text, from, isIdentPart)
          isName(text.substring(from, end), text, end) || hasNameShapedToken(text, end)
        case _ => hasNameShapedToken(text, from + 1)

  /**
   * `target` resized to `shape`'s height and width from its upper-left cell, clipped to the grid.
   */
  def resize(target: CellRange, shape: CellRange): CellRange =
    val col0 = target.colStart.index0
    val row0 = target.rowStart.index0
    val colEnd = math.min(col0.toLong + shape.width - 1, Column.MaxIndex0.toLong).toInt
    val rowEnd = math.min(row0.toLong + shape.height - 1, Row.MaxIndex0.toLong).toInt
    CellRange(ARef.from0(col0, row0), ARef.from0(colEnd, rowEnd))

  /**
   * The cell of `target` paired with `testRef` of `shape`: the same offset from `target`'s
   * upper-left as `testRef` has from `shape`'s. None past the grid edge (that pairing reads
   * nothing).
   */
  def pairedCell(testRef: ARef, shape: CellRange, target: CellRange): Option[ARef] =
    val col = target.colStart.index0.toLong + (testRef.col.index0 - shape.colStart.index0)
    val row = target.rowStart.index0.toLong + (testRef.row.index0 - shape.rowStart.index0)
    Option.when(col <= Column.MaxIndex0 && row <= Row.MaxIndex0)(
      ARef.from0(col.toInt, row.toInt)
    )

  /**
   * A range-slot location resized to `shape` when it carries a range (local, cross-sheet or
   * external); a defined name, a LET name or an error is returned as it is.
   */
  def resizeLocation(location: TExpr.RangeLocation, shape: CellRange): TExpr.RangeLocation =
    location match
      case TExpr.RangeLocation.Local(range, _) =>
        TExpr.RangeLocation.Local(resize(range, shape), RangeForm.Cells)
      case TExpr.RangeLocation.CrossSheet(sheet, range, _) =>
        TExpr.RangeLocation.CrossSheet(sheet, resize(range, shape), RangeForm.Cells)
      case TExpr.RangeLocation.External(index, name, range, _) =>
        TExpr.RangeLocation.External(index, name, resize(range, shape), RangeForm.Cells)
      case other @ (TExpr.RangeLocation.Name(_, _) | TExpr.RangeLocation.Binding(_) |
          TExpr.RangeLocation.Error(_, _)) =>
        other

  /**
   * The argument values of a call to `name` with the resize applied: for a resizing function whose
   * first argument's shape `shapeOf` can tell, the third argument's range is resized to it. Every
   * other call — and a call whose first argument's shape is unknown (a defined name without a
   * workbook to resolve it) — is returned as given. The dependency extractors read a call's
   * arguments through this so the graph records the cells evaluation actually reads.
   *
   * @param name
   *   the function's name, in any case (the published signature since 0.22.0; the evaluator passes
   *   the canonical `FunctionSpec.name`)
   * @param targetOf
   *   the third argument as a location that carries its range, when one can be known (#733: a
   *   defined name bound to a range, resolved through the workbook); a location it cannot resolve
   *   is returned as it is and keeps its own cells
   */
  def resizedArgs(
    name: String,
    values: List[ArgValue],
    shapeOf: TExpr.RangeLocation => Option[CellRange] = _.staticRange,
    targetOf: TExpr.RangeLocation => TExpr.RangeLocation = identity
  ): List[ArgValue] =
    if !resizingFunctions.contains(name.toUpperCase(java.util.Locale.ROOT)) then values
    else
      values match
        case (first @ ArgValue.Range(criteria)) :: criterion :: ArgValue.Range(target) :: rest =>
          shapeOf(criteria) match
            case Some(shape) =>
              first :: criterion :: ArgValue.Range(resizeLocation(targetOf(target), shape)) :: rest
            case None => values
        case _ => values
