package com.tjclp.xl.formula.graph

import java.util.Locale

import scala.annotation.tailrec

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.Evaluator
import com.tjclp.xl.formula.functions.FunctionRegistry
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.ooxml.FormulaStorage
import com.tjclp.xl.workbooks.{DefinedName, Workbook}

/**
 * GH-606: a sound textual over-approximation of what a formula the parser rejects may read.
 *
 * The static graph has no edges for a formula it cannot parse (omitted arguments, `SINGLE`, a union
 * in a defined name — GH-507's "unresolved readers"), so an edit's dirty cone used to treat every
 * such reader as dirty on every edit: evaluated, failed, cache withdrawn. On a real model that
 * withdrew 273 Excel-written caches for a `put` on a formula-free cover sheet. The graph still
 * cannot prove those readers independent, but their TEXT can: a formula that spells only
 * `Summary!A1:A5` cannot read `Cover!B16`.
 *
 * [[reach]] scans the text character by character. Every character falls into a token class that
 * either contributes references, is provably not a reference, or is unknown — and one unknown token
 * makes the whole reach [[Reach.Unbounded]] (today's always-dirty behaviour). The classes:
 *
 *   - `"…"` string literals (`""` escape): no references — a literal `"B16"` is not a dependency.
 *   - `'…'` quoted and bare sheet qualifiers followed by `!`, applying to the immediately following
 *     reference only. `'S1:S3'!` and `S1:S3!` are 3-D spans over every sheet between the two
 *     (either end unknown → unbounded). A single qualifier naming no sheet reads `#REF!` and
 *     contributes nothing. Sheet lookup is case-insensitive, as Excel resolves qualifiers.
 *   - `[` anywhere outside a string: unbounded — structured (`Table1[Col]`) and external
 *     (`[1]Sheet!A1`) references stay coarse; readers with a pinned closed-workbook cache never
 *     reach this scanner (`DependencyGraph.unresolvedReaders` excludes them).
 *   - `#` error literals (`#REF!`, `#N/A`, `#DIV/0!`, …): dead operands, no references. Any other
 *     `#…` token is unknown.
 *   - numbers: none — except `n:m` between two integers, a whole-row range.
 *   - identifiers: a call (`NAME(`) contributes nothing unless the bare name (`_xlfn.` / `_xlws.`
 *     stripped) is a dynamic-reference function, which is unbounded; `TRUE`/`FALSE` are literals;
 *     an A1 shape within the grid is a cell (`$` anchors stripped), `A:A` a whole-column range,
 *     `A1:B2` the bounding rectangle; anything else is a defined name, resolved with the
 *     evaluator's case-insensitive sheet-scoped shadowing from the qualifier sheet or the reader's
 *     sheet. A missing name is unbounded. A found name reads what its definition reads: through the
 *     parsed expression when the definition parses and names no other name or dynamic call,
 *     otherwise textually with the definition's scope sheet as home — a workbook-scoped definition
 *     has none, so an unqualified reference inside it is unbounded. Name chains carry the same
 *     cycle/depth guard as `unresolvedReaders`.
 *   - `:` anywhere else — after a name, a call or a literal — is unbounded (`INDEX(…):INDEX(…)`
 *     spans cells only evaluation can locate).
 *   - `@`, the operators, parentheses, separators, array braces and whitespace (the intersection
 *     operator is harmless to an over-approximation): no references.
 *   - any other character: unbounded.
 *
 * Coarse on purpose: `LET`/`LAMBDA` parameter names look like defined names (unbounded when they do
 * not resolve), and no attempt is made to shrink a range by what a function actually consumes. Pure
 * and total; the scanner never throws.
 */
private[xl] object ReferenceScan:

  /** One rectangle on one sheet a blind reader may read. */
  type Area = (SheetName, CellRange)

  /** What a blind reader can read: anything, or cells inside a set of areas. */
  enum Reach derives CanEqual:
    case Unbounded
    case Areas(areas: Set[Area])

    /** Whether any of `cells` (grouped by sheet) lies inside the reach. */
    def touches(cellsBySheet: Map[SheetName, Set[ARef]]): Boolean = this match
      case Unbounded => true
      case Areas(areas) =>
        areas.exists((sheet, range) => cellsBySheet.get(sheet).exists(_.exists(range.contains)))

  /** The reach of formula `text` (leading `=` optional) as read from `sheet`. */
  def reach(workbook: Workbook, sheet: SheetName, text: String): Reach =
    Scanner(workbook).reach(text, Context(Some(sheet), sheet, Set.empty))

  /** The reach of every reader, each scanned once against its own sheet with one shared memo. */
  def reaches(workbook: Workbook, readers: Set[QualifiedRef]): Map[QualifiedRef, Reach] =
    val scanner = Scanner(workbook)
    readers.iterator.map { reader =>
      val formula = workbook(reader.sheet).toOption.flatMap(_.cells.get(reader.ref)).map(_.value)
      val reach = formula match
        case Some(CellValue.Formula(text, _, _)) =>
          scanner.reach(text, Context(Some(reader.sheet), reader.sheet, Set.empty))
        case _ => Reach.Unbounded
      reader -> reach
    }.toMap

  /** Name-chain depth guard, as in `DependencyGraph.unresolvedReaders`. */
  private val MaxNameDepth = 100

  /** Excel's error literals; each is a dead operand that reads nothing. */
  private val ErrorLiterals: Vector[String] = Vector(
    "#N/A",
    "#REF!",
    "#DIV/0!",
    "#NAME?",
    "#VALUE!",
    "#NUM!",
    "#NULL!",
    "#SPILL!",
    "#CALC!",
    "#GETTING_DATA",
    "#BLOCKED!",
    "#CONNECT!",
    "#FIELD!",
    "#UNKNOWN!",
    "#BUSY!",
    "#PYTHON!",
    "#EXTERNAL!"
  )

  /** Characters that separate references without contributing any. */
  private val Inert = "+-*/^&=<>%,;(){}@"

  /**
   * Where an unqualified reference lands and how an unqualified name is looked up.
   *
   * @param ambient
   *   the home sheet of an unqualified cell/row/column reference; None inside a workbook-scoped
   *   definition, where such a reference is unbounded
   * @param fallback
   *   the sheet an unqualified name is looked up from when there is no home sheet — the reader's
   *   sheet, or the scope sheet of the enclosing sheet-scoped definition (the evaluator's rule)
   * @param visiting
   *   name-chain cycle guard: (lookup sheet, fallback sheet, upper-cased name) on the current path
   */
  private final case class Context(
    ambient: Option[SheetName],
    fallback: SheetName,
    visiting: Set[(SheetName, SheetName, String)]
  )

  private final class Scanner(workbook: Workbook):
    private type NameKey = (SheetName, SheetName, String)

    private val order: Vector[SheetName] = workbook.sheets.map(_.name)
    // Reverse insertion so a duplicated sheet name keeps its FIRST position, like indexWhere
    private val positions: Map[SheetName, Int] =
      order.zipWithIndex.reverseIterator.map((name, i) => name -> i).toMap
    private val byFold: Map[String, SheetName] =
      order.reverseIterator.map(name => fold(name.value) -> name).toMap
    private val dynamicFunctions: Set[String] = FunctionRegistry.dynamicFunctionNames.toSet
    private val canonicalSheet: SheetName => SheetName =
      DependencyGraph.sheetCanonicaliser(workbook)
    // Build-local memo of resolved names; a guard-truncated result is only ever more conservative
    private val names = scala.collection.mutable.HashMap.empty[NameKey, Reach]

    def reach(text: String, context: Context): Reach =
      scan(text.stripPrefix("="), 0, context, Set.empty)

    @tailrec
    private def scan(s: String, i: Int, context: Context, acc: Set[Area]): Reach =
      if i >= s.length then Reach.Areas(acc)
      else
        val c = s.charAt(i)
        if c == '"' then
          val end = closingQuote(s, i + 1, '"')
          if end < 0 then Reach.Unbounded else scan(s, end, context, acc)
        else if c == '#' then
          errorLiteralEnd(s, i) match
            case Some(end) => scan(s, end, context, acc)
            case None => Reach.Unbounded
        else if c == '[' then Reach.Unbounded
        else if c.isWhitespace || Inert.indexOf(c.toInt) >= 0 then scan(s, i + 1, context, acc)
        else
          unit(s, i, context) match
            case Some((Reach.Areas(areas), end)) => scan(s, end, context, acc ++ areas)
            case _ => Reach.Unbounded

    /** One reference unit at `i`: an optional sheet qualifier, then the reference it applies to. */
    private def unit(s: String, i: Int, context: Context): Option[(Reach, Int)] =
      val c = s.charAt(i)
      if c == '\'' then
        val end = closingQuote(s, i + 1, '\'')
        if end < 0 || at(s, end) != '!' then None
        else
          val raw = s.substring(i + 1, end - 1).replace("''", "'")
          if raw.contains('[') then None
          else qualifier(raw).flatMap(q => reference(s, end + 1, context, Some(q)))
      else if isIdentStart(c) then
        val end = tokenEnd(s, i)
        if at(s, end) == '!' then
          qualifier(s.substring(i, end)).flatMap(q => reference(s, end + 1, context, Some(q)))
        else if at(s, end) == ':' && isIdentStart(at(s, end + 1)) &&
          at(s, tokenEnd(s, end + 1)) == '!'
        then
          val spanEnd = tokenEnd(s, end + 1)
          qualifier(s.substring(i, spanEnd))
            .flatMap(q => reference(s, spanEnd + 1, context, Some(q)))
        else reference(s, i, context, None)
      else if c == '$' || isDigit(c) || (c == '.' && isDigit(at(s, i + 1))) then
        reference(s, i, context, None)
      else None

    /**
     * The sheets a qualifier names: one, none (a dead qualifier reads `#REF!`), or the sheets of a
     * 3-D span; None when a span end is unknown.
     */
    private def qualifier(raw: String): Option[Vector[SheetName]] =
      raw.indexOf(':') match
        case -1 => Some(byFold.get(fold(raw)).toList.toVector)
        case k =>
          for
            first <- byFold.get(fold(raw.substring(0, k)))
            last <- byFold.get(fold(raw.substring(k + 1)))
            from <- positions.get(first)
            to <- positions.get(last)
          yield order.slice(math.min(from, to), math.max(from, to) + 1)

    /** The reference starting at `j` — a call, literal, cell, range, rows, columns or name. */
    private def reference(
      s: String,
      j: Int,
      context: Context,
      qualifier: Option[Vector[SheetName]]
    ): Option[(Reach, Int)] =
      val end = tokenEnd(s, j)
      if end == j then None
      else
        val token = s.substring(j, end)
        val bare = token.replace("$", "")
        val next = at(s, end)
        if bare.isEmpty || next == '!' then None
        else if next == '(' then
          // A call: the arguments are scanned by the caller; only the function itself matters
          if qualifier.isDefined || token.contains('$') then None
          else if dynamicFunctions.contains(
              FormulaStorage.bareFunctionName(bare).toUpperCase(Locale.ROOT)
            )
          then None
          else Some((Reach.Areas(Set.empty), end))
        else if isCell(bare) then
          if next == ':' then
            val end2 = tokenEnd(s, end + 1)
            val bare2 = s.substring(end + 1, end2).replace("$", "")
            if end2 > end + 1 && isCell(bare2) && at(s, end2) != '!' && at(s, end2) != '(' then
              for
                start <- ARef.parse(bare).toOption
                stop <- ARef.parse(bare2).toOption
                placed <- place(CellRange(start, stop), qualifier, context)
              yield (placed, end2)
            else None
          else
            ARef
              .parse(bare)
              .toOption
              .flatMap(ref => place(CellRange(ref, ref), qualifier, context))
              .map(_ -> end)
        else if isColumn(bare) && next == ':' then
          val end2 = tokenEnd(s, end + 1)
          val bare2 = s.substring(end + 1, end2).replace("$", "")
          if end2 > end + 1 && isColumn(bare2) && at(s, end2) != '!' && at(s, end2) != '(' then
            for
              first <- Column.fromLetter(bare).toOption
              last <- Column.fromLetter(bare2).toOption
              placed <- place(columns(first, last), qualifier, context)
            yield (placed, end2)
          else None
        else if isNumber(bare) then
          if next != ':' then
            if qualifier.isDefined || token.contains('$') then None
            else Some((Reach.Areas(Set.empty), end))
          else if !isInteger(bare) then None
          else
            val end2 = tokenEnd(s, end + 1)
            val bare2 = s.substring(end + 1, end2).replace("$", "")
            if end2 > end + 1 && isInteger(bare2) && at(s, end2) != '!' && at(s, end2) != '(' then
              for
                first <- bare.toIntOption
                last <- bare2.toIntOption
                range <- rows(first, last)
                placed <- place(range, qualifier, context)
              yield (placed, end2)
            else None
        else if bare.equalsIgnoreCase("TRUE") || bare.equalsIgnoreCase("FALSE") then
          if qualifier.isDefined || token.contains('$') then None
          else Some((Reach.Areas(Set.empty), end))
        else if token.contains('$') || next == ':' then None
        else
          // A defined name, looked up from its qualifier or from the reader's home
          val lookupFrom = qualifier match
            case None => Some(context.ambient.getOrElse(context.fallback))
            case Some(Vector(single)) => Some(single)
            case Some(_) => None
          lookupFrom.map(from => (resolveName(bare, from, context), end))

    /** Qualified areas, or the home sheet's; None when an unqualified reference has no home. */
    private def place(
      range: CellRange,
      qualifier: Option[Vector[SheetName]],
      context: Context
    ): Option[Reach] =
      qualifier match
        case Some(sheets) => Some(Reach.Areas(sheets.iterator.map(sheet => (sheet, range)).toSet))
        case None => context.ambient.map(sheet => Reach.Areas(Set((sheet, range))))

    private def resolveName(name: String, lookupFrom: SheetName, context: Context): Reach =
      val key: NameKey = (lookupFrom, context.fallback, name.toUpperCase(Locale.ROOT))
      names.get(key) match
        case Some(known) => known
        case None if context.visiting.contains(key) || context.visiting.size >= MaxNameDepth =>
          Reach.Unbounded
        case None =>
          val defined = positions
            .get(lookupFrom)
            .flatMap(position => Evaluator.lookupDefinedNameAt(workbook, Some(position), name))
          val resolved = defined match
            case None => Reach.Unbounded
            case Some(dn) => definitionReach(dn, context.copy(visiting = context.visiting + key))
          names(key) = resolved
          resolved

    /**
     * What a definition reads. A definition that parses and names neither another name nor a
     * dynamic call is read off its expression — the same extraction the dependency index uses, so
     * the areas are exactly its declared references. Otherwise the definition text is scanned like
     * a formula: the textual rules resolve nested names (an unparseable one reached through a
     * parseable alias included) and make a dynamic call unbounded.
     */
    private def definitionReach(dn: DefinedName, context: Context): Reach =
      val scope = Evaluator.definedNameScope(workbook, dn).map(_.name)
      val home = scope.getOrElse(context.fallback)
      val text = dn.formula.stripPrefix("=")
      FormulaParser.parse(text) match
        case Right(expr)
            if !DependencyGraph.referencesMatching(
              expr,
              (_, _) => true,
              includeDynamicCalls = true
            ) =>
          val ranges = scala.collection.mutable.HashSet.empty[Area]
          val points = DependencyGraph.extractQualifiedDependencies(
            expr,
            home,
            cellsFor = (sheet, range) =>
              // anchors are print detail: one rectangle, whichever way the definition spelt it
              ranges += ((sheet, CellRange(range.start, range.end)))
              Set.empty
            ,
            workbook = Some(workbook),
            canonicalSheet = canonicalSheet
          )
          Reach.Areas(ranges.toSet ++ points.map(q => (q.sheet, CellRange(q.ref, q.ref))))
        case _ => scan(text, 0, Context(scope, home, context.visiting), Set.empty)

    // ----- lexical helpers -----

    /**
     * Index just past the quote closing the literal whose body starts at `j`; -1 if unterminated.
     */
    @tailrec
    private def closingQuote(s: String, j: Int, quote: Char): Int =
      if j >= s.length then -1
      else if s.charAt(j) != quote then closingQuote(s, j + 1, quote)
      else if at(s, j + 1) == quote then closingQuote(s, j + 2, quote)
      else j + 1

    private def errorLiteralEnd(s: String, i: Int): Option[Int] =
      ErrorLiterals.collectFirst {
        case literal if s.regionMatches(true, i, literal, 0, literal.length) => i + literal.length
      }

    /** End of the maximal run of identifier/reference characters starting at `j`. */
    @tailrec
    private def tokenEnd(s: String, j: Int): Int =
      if j < s.length && isTokenChar(s.charAt(j)) then tokenEnd(s, j + 1) else j

    private def at(s: String, k: Int): Char = if k >= 0 && k < s.length then s.charAt(k) else ' '

    private def isIdentStart(c: Char): Boolean = Character.isLetter(c) || c == '_' || c == '\\'

    private def isTokenChar(c: Char): Boolean =
      Character.isLetterOrDigit(c) || c == '_' || c == '.' || c == '?' || c == '\\' || c == '$'

    private def isDigit(c: Char): Boolean = c >= '0' && c <= '9'

    private def isInteger(s: String): Boolean = s.nonEmpty && s.forall(isDigit)

    /**
     * Digits and at most one point, then an optional exponent marker (its sign follows as inert).
     */
    private def isNumber(token: String): Boolean =
      val mantissa = token.takeWhile(c => isDigit(c) || c == '.')
      val exponent = token.drop(mantissa.length)
      mantissa.nonEmpty && mantissa.count(_ == '.') <= 1 && mantissa.exists(isDigit) &&
      (exponent.isEmpty ||
        ((exponent.charAt(0) == 'e' || exponent.charAt(0) == 'E') && isInteger(exponent.drop(1)) ||
          exponent.equalsIgnoreCase("e")))

    /** An A1 cell inside the grid: one to three letters, then digits. */
    private def isCell(bare: String): Boolean =
      val letters = bare.takeWhile(Character.isLetter)
      letters.nonEmpty && letters.length <= 3 && bare.length > letters.length &&
      bare.drop(letters.length).forall(isDigit) && ARef.parse(bare).isRight

    private def isColumn(bare: String): Boolean =
      bare.nonEmpty && bare.length <= 3 && Column.fromLetter(bare).isRight

    private def columns(first: Column, last: Column): CellRange =
      CellRange(ARef(first, Row.from0(0)), ARef(last, Row.from0(Row.MaxIndex0)))

    private def rows(first: Int, last: Int): Option[CellRange] =
      val maxRow = Row.MaxIndex0 + 1
      Option.when(first >= 1 && last >= 1 && first <= maxRow && last <= maxRow)(
        CellRange(ARef.from1(1, first), ARef(Column.from0(Column.MaxIndex0), Row.from1(last)))
      )

    private def fold(s: String): String = s.toLowerCase(Locale.ROOT)
