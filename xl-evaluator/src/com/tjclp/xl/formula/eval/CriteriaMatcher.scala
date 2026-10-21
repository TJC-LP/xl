package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.Clock

import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.ArrayArithmetic
import scala.util.boundary
import boundary.break
import scala.util.matching.Regex

/**
 * Excel-compatible criteria matching for SUMIF/COUNTIF functions.
 *
 * Supports these criteria patterns (Excel-compatible):
 *   - **Exact**: `"Apple"` or `42` → Cells equal to value. `"=op"` reads `op` as a bare operand
 *     (#773), so `"=a*"` is the wildcard `a*`; a bare `"="` matches truly empty cells only.
 *   - **Wildcard**: `"A*"`, `"*pple"`, `"A?ple"` → Pattern matching (* = any chars, ? = single
 *     char) against text cells only (#780): numbers, booleans, errors and blanks never match. In an
 *     operand that contains `*` or `?`, `~x` stands for a literal `x` (`~*`, `~?`, `~~`). In an
 *     operand with neither, `~` is an ordinary character (GH-765): `"a~~b"` matches only the text
 *     `a~~b`.
 *   - **Greater**: `">100"`, `">=50"` → Numeric comparison; `">m"` → Text comparison
 *   - **Less**: `"<10"`, `"<=5"` → Numeric comparison; `"<m"` → Text comparison
 *   - **Not Equal**: `"<>0"` → every cell but a number equal to 0 (#731); `"<>x"` → Text
 *     inequality; bare `"<>"` → every cell but a truly empty one; `"<>A*"` → negated wildcard.
 *     Every `<>` criterion matches error cells.
 *   - **Error literal** (GH-565): `"#N/A"`, `"#DIV/0!"`, ... → Cells holding that error VALUE (not
 *     the text); `"<>#N/A"` → cells not holding that error
 *
 * Laws:
 *   1. Exact match: `matches(Text("Apple"), Exact(ExprValue.Text("Apple"))) == true`
 *   2. Wildcard: `matches(Text("Apple"), Wildcard("A*")) == true`
 *   3. Case-insensitive text: `matches(Text("APPLE"), Exact(ExprValue.Text("apple"))) == true`
 *   4. Numeric comparison: `matches(Number(150), Compare(Gt, 100)) == true`
 *   5. Text comparison: `matches(Text("x"), CompareText(Gt, "m")) == true`
 */
object CriteriaMatcher:

  /** Parsed criterion for cell matching */
  sealed trait Criterion derives CanEqual

  /** Exact match against text, number, or boolean */
  case class Exact(value: ExprValue) extends Criterion

  /** Numeric comparison */
  case class Compare(op: CompareOp, value: BigDecimal) extends Criterion

  /**
   * Case-insensitive text comparison (GH-466).
   *
   * Produced when the operand after `<>`/`>=`/`<=`/`>`/`<` is not numeric, e.g. `"<>x"` or `">m"`.
   * Bare `"<>"` parses to `CompareText(Neq, "")` — Excel's non-blank idiom.
   */
  case class CompareText(op: CompareOp, value: String) extends Criterion

  /** Wildcard pattern (* = any chars, ? = single char) */
  case class Wildcard(pattern: String) extends Criterion

  /**
   * #773: a bare `"="`, which matches truly empty cells only. A formula's empty text (`=""`) is not
   * empty here, unlike the criterion `""`, which matches both.
   */
  case object Blank extends Criterion

  /** Negated wildcard pattern, from `<>` followed by a pattern with unescaped wildcards */
  case class NotWildcard(pattern: String) extends Criterion

  /**
   * GH-565: `<>` followed by an error literal (`"<>#N/A"`) — every cell that does not hold that
   * error value, blanks and other errors included.
   */
  case class NotError(err: CellError) extends Criterion

  /** Comparison operators for numeric criteria */
  enum CompareOp derives CanEqual:
    case Gt // >
    case Gte // >=
    case Lt // <
    case Lte // <=
    case Neq // <>

  /**
   * Parse a criterion from an evaluated expression result.
   *
   * Handles:
   *   - Strings starting with operators: ">100", ">=50", "<10", "<=5", "<>0"
   *   - Strings with wildcards: "A*", "*pple", "A?ple"
   *   - Plain strings: exact match
   *   - Numbers/Booleans: exact match
   *
   * @param value
   *   The evaluated criterion value
   * @return
   *   Parsed Criterion
   */
  def parse(value: ExprValue): Criterion = value match
    case ExprValue.Text(text) => parseString(text)
    case ExprValue.Number(number) => Exact(ExprValue.Number(number))
    case ExprValue.Bool(bool) => Exact(ExprValue.Bool(bool))
    case ExprValue.Date(date) => Exact(ExprValue.Date(date))
    case ExprValue.DateTime(dateTime) => Exact(ExprValue.DateTime(dateTime))
    case ExprValue.Cell(cellValue) =>
      // Unwrap CellValue to appropriate ExprValue for proper matching.
      // This enables cell references used as criteria (e.g., =SUMIFS(..., A2))
      cellValue match
        case CellValue.Text(s) => parseString(s)
        case CellValue.Number(n) => Exact(ExprValue.Number(n))
        case CellValue.Bool(b) => Exact(ExprValue.Bool(b))
        case CellValue.DateTime(dt) => Exact(ExprValue.DateTime(dt))
        case CellValue.RichText(rt) => parseString(rt.toPlainText)
        case CellValue.Formula(_, Some(cached), _) => parse(ExprValue.Cell(cached))
        case CellValue.Formula(_, None, _) => Exact(ExprValue.Cell(cellValue))
        case CellValue.Empty => Exact(ExprValue.Text(""))
        case CellValue.Error(_) => Exact(ExprValue.Cell(cellValue))
    case ExprValue.Opaque(other) => Exact(ExprValue.Opaque(other))

  /**
   * Parse string criterion into operator, wildcard, or exact match.
   */
  private def parseString(s: String): Criterion =
    s match
      // GH-565: a criteria string that spells an Excel error literal IS that error value — Excel
      // parses criteria like formula text, so "#N/A" matches #N/A error cells, never the text
      // "#N/A". Checked before the wildcard rule: "#NAME?" contains a '?' but is not a pattern.
      case ErrorLiteral(err) => Exact(ExprValue.Cell(CellValue.Error(err)))
      // Not equal: <>
      // GH-466: non-numeric operands are text inequality (bare "<>" = non-blank);
      // wildcard operands negate the pattern (Excel "<>A*" = not-like "A*").
      case _ if s.startsWith("<>") =>
        val operand = s.drop(2)
        operand match
          case ErrorLiteral(err) => NotError(err) // GH-565
          case _ =>
            // a number past Excel's range ("1E+999999999") is text here, as in desktop Excel
            parseNumeric(operand).filter(_.abs <= MaxExcelNumber) match
              case Some(n) => Compare(CompareOp.Neq, n)
              case None if hasUnescapedWildcard(operand) => NotWildcard(operand)
              case None if hasWildcardChar(operand) =>
                CompareText(CompareOp.Neq, unescapePattern(operand))
              // GH-765: no `*` or `?`, so `~` is an ordinary character ("<>a~~b" is "a~~b")
              case None => CompareText(CompareOp.Neq, operand)
      // Greater than or equal: >=
      case _ if s.startsWith(">=") =>
        parseNumeric(s.drop(2)) match
          case Some(n) => Compare(CompareOp.Gte, n)
          case None => CompareText(CompareOp.Gte, s.drop(2)) // GH-466: text ordering
      // Less than or equal: <=
      case _ if s.startsWith("<=") =>
        parseNumeric(s.drop(2)) match
          case Some(n) => Compare(CompareOp.Lte, n)
          case None => CompareText(CompareOp.Lte, s.drop(2)) // GH-466: text ordering
      // Greater than: >
      case _ if s.startsWith(">") =>
        parseNumeric(s.drop(1)) match
          case Some(n) => Compare(CompareOp.Gt, n)
          case None => CompareText(CompareOp.Gt, s.drop(1)) // GH-466: text ordering
      // Less than: <
      case _ if s.startsWith("<") =>
        parseNumeric(s.drop(1)) match
          case Some(n) => Compare(CompareOp.Lt, n)
          case None => CompareText(CompareOp.Lt, s.drop(1)) // GH-466: text ordering
      // Equal: = (#773: the operand reads as a bare operand, wildcards included; a bare "="
      // matches truly empty cells only)
      case _ if s.startsWith("=") =>
        s.drop(1) match
          case "" => Blank
          case ErrorLiteral(err) => Exact(ExprValue.Cell(CellValue.Error(err))) // GH-565
          case operand =>
            parseNumeric(operand) match
              case Some(n) => Exact(ExprValue.Number(n))
              case None => parseOperand(operand)
      case _ => parseOperand(s)

  /** A criterion operand with no comparison operator, bare or after `=` (#773). */
  private def parseOperand(s: String): Criterion =
    // Wildcard pattern
    if hasUnescapedWildcard(s) then Wildcard(s)
    // Every `*` and `?` is escaped: an exact match on the unescaped text ("~*" is "*")
    else if hasWildcardChar(s) then Exact(ExprValue.Text(unescapePattern(s)))
    // GH-765: no `*` or `?` at all, so Excel compares the text as written and `~` is an
    // ordinary character: "a~~b" matches only "a~~b", and a cell used as its own criterion
    // matches itself.
    else Exact(ExprValue.Text(s))

  /** The largest magnitude an Excel number has (an IEEE double's). */
  private val MaxExcelNumber: BigDecimal = BigDecimal(Double.MaxValue)

  /**
   * GH-565: extractor for the seven Excel error literals (`#N/A`, `#DIV/0!`, ...). Letter case is
   * ignored, like the rest of criteria parsing (and like Excel's own cell input, where typing
   * `#n/a` yields the #N/A error): `"#n/a"` is the error value, not the text.
   */
  private object ErrorLiteral:
    def unapply(s: String): Option[CellError] =
      CellError.parse(s.toUpperCase(java.util.Locale.ROOT)).toOption

  /**
   * Check if string contains unescaped wildcards (* or ?).
   *
   * Escaped wildcards (~* or ~?) don't count as wildcards.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def hasUnescapedWildcard(s: String): Boolean =
    boundary {
      var i = 0
      while i < s.length do
        val c = s.charAt(i)
        if c == '~' && i + 1 < s.length then
          // Skip escaped character
          i += 2
        else if c == '*' || c == '?' then break(true)
        else i += 1
      false
    }

  /**
   * GH-765: whether the operand contains `*` or `?`, escaped or not. Excel reads `~` as an escape
   * only in such an operand; without one the operand is compared as written.
   */
  private def hasWildcardChar(s: String): Boolean =
    s.exists(c => c == '*' || c == '?')

  /**
   * Remove escape characters from a pattern: `~x` is `x` for every character `x` (`~*`, `~?`, `~~`
   * and `~b` alike); a `~` at the end stays a literal tilde. Desktop Excel reads `~` this way in a
   * criterion operand with a `*` or `?` and in an XMATCH/XLOOKUP match_mode 2 lookup value.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def unescapePattern(s: String): String =
    val sb = new StringBuilder
    var i = 0
    while i < s.length do
      val c = s.charAt(i)
      if c == '~' && i + 1 < s.length then
        sb.append(s.charAt(i + 1))
        i += 2
      else
        sb.append(c)
        i += 1
    sb.toString

  /**
   * #774: an XMATCH/XLOOKUP match_mode 2 lookup value, read with the escape grammar whether or not
   * it holds a wildcard. Unlike a criterion operand (GH-765), a lookup value always drops each `~`
   * before a character: `"a~~b"` finds `a~b` and `"~4"` finds the text `4`.
   */
  enum LookupPattern derives CanEqual:
    /** No unescaped `*` or `?`: the key must equal this text, every `~x` read as `x`. */
    case Literal(text: String)

    /** An unescaped `*` or `?` remains: the key must be text that matches the pattern. */
    case Pattern(wildcard: Wildcard)

  /** #774: how a match_mode 2 lookup value spelled `text` matches a key ([[LookupPattern]]). */
  def lookupPattern(text: String): LookupPattern =
    if hasUnescapedWildcard(text) then LookupPattern.Pattern(Wildcard(text))
    else LookupPattern.Literal(unescapePattern(text))

  /**
   * Parse string as BigDecimal, handling Excel's flexible numeric syntax.
   *
   * `BigDecimal(String)` accepts only text that starts with a sign, a point or a digit (any
   * `Character.isDigit` digit, not only ASCII), so text that starts with anything else is not a
   * number and is refused before the parse, which would throw and catch an exception for it.
   * `CriteriaIndexSpec` pins this to [[parseNumericByParse]] on random and adversarial text. The
   * criteria index (#767) keys numeric text with this same function.
   */
  private[formula] def parseNumeric(s: String): Option[BigDecimal] =
    val trimmed = s.trim
    if trimmed.isEmpty then None
    else
      val first = trimmed.charAt(0)
      if first == '+' || first == '-' || first == '.' || Character.isDigit(first) then
        scala.util.Try(BigDecimal(trimmed)).toOption
      else None

  /** [[parseNumeric]] without the leading-character check: the reference its test compares to. */
  private[eval] def parseNumericByParse(s: String): Option[BigDecimal] =
    val trimmed = s.trim
    if trimmed.isEmpty then None
    else scala.util.Try(BigDecimal(trimmed)).toOption

  /**
   * Test if a cell value matches a criterion.
   *
   * Excel matching rules:
   *   - Text comparisons are case-insensitive
   *   - Empty cells don't match any criteria (except exact match on empty)
   *   - Type coercion: Text "42" matches Number 42 for exact match
   *   - Wildcards only apply to text values
   *
   * @param cellValue
   *   The cell value to test
   * @param criterion
   *   The criterion to match against
   * @return
   *   true if cell value matches criterion
   */
  def matches(cellValue: CellValue, criterion: Criterion): Boolean =
    criterion match
      case Exact(expected) => matchesExact(cellValue, expected)
      case Compare(op, threshold) => matchesCompare(cellValue, op, threshold)
      case CompareText(op, operand) => matchesCompareText(cellValue, op, operand)
      case Wildcard(pattern) => matchesWildcard(cellValue, pattern)
      case NotWildcard(pattern) => matchesNotWildcard(cellValue, pattern)
      case NotError(err) => matchesNotError(cellValue, err)
      case Blank =>
        cellValue match
          case CellValue.Empty => true
          case _ => false

  /**
   * GH-565: `<>#N/A` — everything but that error value. Blanks, values and other errors match
   * (consistent with `<>text` matching blanks); uncached formulas never match.
   */
  private def matchesNotError(cellValue: CellValue, err: CellError): Boolean =
    cellValue match
      case CellValue.Error(e) => e != err
      case CellValue.Formula(_, Some(cached), _) => matchesNotError(cached, err)
      case CellValue.Formula(_, None, _) => false
      case _ => true

  /**
   * Exact matching with type coercion.
   *
   * Follows Excel's type coercion rules:
   *   - Text vs Text: case-insensitive
   *   - Text "42" vs Number 42: matches
   *   - Boolean TRUE vs Text "TRUE": matches
   */
  private def matchesExact(cellValue: CellValue, expected: ExprValue): Boolean =
    cellValue match
      case CellValue.Empty =>
        expected match
          case ExprValue.Text(text) => text.isEmpty
          case _ => false

      case CellValue.Text(text) =>
        expected match
          case ExprValue.Text(s) => text.equalsIgnoreCase(s)
          case ExprValue.Number(n) => parseNumeric(text).contains(n)
          case ExprValue.Bool(b) => text.equalsIgnoreCase(b.toString)
          case _ => false

      case CellValue.Number(value) =>
        expected match
          case ExprValue.Number(n) => value == n
          case ExprValue.Text(s) => parseNumeric(s).contains(value)
          case _ => false

      case CellValue.Bool(value) =>
        expected match
          case ExprValue.Bool(b) => value == b
          case ExprValue.Text(s) =>
            s.equalsIgnoreCase("TRUE") && value || s.equalsIgnoreCase("FALSE") && !value
          case ExprValue.Number(n) =>
            (n == BigDecimal(1) && value) || (n == BigDecimal(0) && !value)
          case _ => false

      case CellValue.DateTime(dt) =>
        expected match
          case ExprValue.Text(s) => dt.toString == s || dt.toLocalDate.toString == s
          case ExprValue.Date(ld) =>
            // DATE() returns LocalDate - compare date portion
            dt.toLocalDate == ld
          case ExprValue.Number(n) =>
            // Excel serial number comparison (e.g., from numeric criteria)
            val cellSerial = BigDecimal(CellValue.dateTimeToExcelSerial(dt))
            // Compare with tolerance for floating point (dates are integers, times add fractions)
            if n.scale <= 0 then
              // Integer serial = date only, compare date portion
              cellSerial.setScale(0, BigDecimal.RoundingMode.FLOOR) == n
            else
              // Fractional serial includes time
              (cellSerial - n).abs < BigDecimal("0.00001")
          case _ => false

      case CellValue.Formula(_, Some(cached), _) =>
        // Match against cached value
        matchesExact(cached, expected)

      case CellValue.Formula(_, None, _) =>
        false

      case CellValue.RichText(rt) =>
        expected match
          case ExprValue.Text(s) => rt.toPlainText.equalsIgnoreCase(s)
          case _ => false

      // GH-565: an error cell matches exactly the criterion carrying the same error value (an
      // error literal, or a reference to a cell holding that error); never text or numbers
      case CellValue.Error(err) =>
        expected match
          case ExprValue.Cell(CellValue.Error(e)) => e == err
          case _ => false

  /**
   * Numeric comparison matching. `"<>n"` (#731) matches every cell but a number (or date) equal to
   * `n`: blanks, text (the text `"4"` included, although `"=4"` matches it too), booleans and
   * errors all match, as in desktop Excel. Uncached formulas never match.
   */
  private def matchesCompare(
    cellValue: CellValue,
    op: CompareOp,
    threshold: BigDecimal
  ): Boolean =
    op match
      case CompareOp.Neq => !isNumberEqual(cellValue, threshold)
      case _ =>
        extractNumeric(cellValue) match
          case Some(value) =>
            op match
              case CompareOp.Gt => value > threshold
              case CompareOp.Gte => value >= threshold
              case CompareOp.Lt => value < threshold
              case CompareOp.Lte => value <= threshold
              case CompareOp.Neq => value != threshold // Unreachable: Neq handled above
          case None => false

  /**
   * #731: whether a cell is a number equal to `n`, the one kind of cell `"<>n"` leaves out. An
   * uncached formula counts as equal, so it never matches `"<>n"`.
   */
  private def isNumberEqual(cellValue: CellValue, n: BigDecimal): Boolean =
    cellValue match
      case CellValue.Number(value) => value == n
      case CellValue.DateTime(dt) => BigDecimal(CellValue.dateTimeToExcelSerial(dt)) == n
      case CellValue.Formula(_, Some(cached), _) => isNumberEqual(cached, n)
      case CellValue.Formula(_, None, _) => true
      case _ => false

  /**
   * Text comparison matching (GH-466), for criteria like "<>x" and ">m".
   *
   * Excel semantics:
   *   - `<>text` is a negated exact match with the same coercion rules as [[Exact]], so numbers,
   *     booleans, blank cells and error cells all match `"<>x"` (they are not equal to "x"). Bare
   *     `"<>"` (operand "") matches every cell but a truly empty one, a formula's empty text
   *     (`=""`) and errors included (#731).
   *   - Ordering operators (`>`, `>=`, `<`, `<=`) compare text case-insensitively and only apply to
   *     text values: numbers, booleans, blanks and errors never satisfy a text ordering criterion
   *     (Excel criteria comparisons are type-segregated).
   *   - Uncached formulas never match.
   */
  private def matchesCompareText(
    cellValue: CellValue,
    op: CompareOp,
    operand: String
  ): Boolean =
    cellValue match
      case CellValue.Error(_) => op == CompareOp.Neq
      case CellValue.Formula(_, Some(cached), _) => matchesCompareText(cached, op, operand)
      case CellValue.Formula(_, None, _) => false
      case other =>
        op match
          case CompareOp.Neq if operand.isEmpty =>
            other match
              case CellValue.Empty => false
              case _ => true
          case CompareOp.Neq => !matchesExact(other, ExprValue.Text(operand))
          case CompareOp.Gt | CompareOp.Gte | CompareOp.Lt | CompareOp.Lte =>
            extractOrderableText(other) match
              case Some(text) =>
                val cmp = text.compareToIgnoreCase(operand)
                op match
                  case CompareOp.Gt => cmp > 0
                  case CompareOp.Gte => cmp >= 0
                  case CompareOp.Lt => cmp < 0
                  case CompareOp.Lte => cmp <= 0
                  case CompareOp.Neq => cmp != 0 // Unreachable: Neq handled above
              case None => false

  /**
   * Extract text for ordering comparisons (GH-466).
   *
   * Deliberately narrower than [[extractText]]: only genuine text participates in text ordering,
   * matching Excel (a number never satisfies `">m"`).
   */
  private def extractOrderableText(cellValue: CellValue): Option[String] =
    cellValue match
      case CellValue.Text(s) => Some(s)
      case CellValue.RichText(rt) => Some(rt.toPlainText)
      case CellValue.Formula(_, Some(cached), _) => extractOrderableText(cached)
      case _ => None

  /**
   * Negated wildcard matching (GH-466), for criteria like "<>A*".
   *
   * Matches cells that do NOT match the pattern: every cell that is not text (numbers, booleans,
   * blanks and errors, #780) and text that does not match. Uncached formulas never match.
   */
  private def matchesNotWildcard(cellValue: CellValue, pattern: String): Boolean =
    cellValue match
      case CellValue.Formula(_, Some(cached), _) => matchesNotWildcard(cached, pattern)
      case CellValue.Formula(_, None, _) => false
      case other => !matchesWildcard(other, pattern)

  /**
   * Extract numeric value from cell, with type coercion.
   */
  private def extractNumeric(cellValue: CellValue): Option[BigDecimal] =
    cellValue match
      case CellValue.Number(n) => Some(n)
      case CellValue.Text(s) => parseNumeric(s)
      case CellValue.Bool(b) => Some(ArrayArithmetic.boolToNumeric(b))
      case CellValue.DateTime(dt) =>
        // Convert to Excel serial number for comparison
        Some(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case CellValue.Formula(_, Some(cached), _) => extractNumeric(cached)
      // GH-197: Uncached formula - can't extract numeric without evaluation context
      case CellValue.Formula(_, None, _) => None
      case _ => None

  /**
   * Wildcard pattern matching.
   *
   * Excel wildcards:
   *   - `*` matches any sequence of characters (including empty)
   *   - `?` matches exactly one character
   *   - Matching is case-insensitive
   *   - `~x` is a literal `x`, for every character `x`
   */
  private def matchesWildcard(cellValue: CellValue, pattern: String): Boolean =
    extractText(cellValue) match
      case Some(text) =>
        val regex = wildcardToRegex(pattern)
        regex.matches(text)
      case None => false

  /**
   * Extract text value from cell for wildcard matching (#780): only text matches a wildcard, as in
   * desktop Excel, where `"*"` counts no number, boolean, error or blank. A formula's empty text
   * (`=""`) is text.
   */
  private def extractText(cellValue: CellValue): Option[String] =
    cellValue match
      case CellValue.Text(s) => Some(s)
      case CellValue.RichText(rt) => Some(rt.toPlainText)
      case CellValue.Formula(_, Some(cached), _) => extractText(cached)
      // GH-197: Uncached formula - can't extract text without evaluation context
      case CellValue.Formula(_, None, _) => None
      case _ => None

  /**
   * Convert Excel wildcard pattern to regex.
   *
   * Handles:
   *   - `*` → `.*`
   *   - `?` → `.`
   *   - `~x` → literal `x`, for every character `x` ([[patternRegexBody]])
   *   - Other regex metacharacters are escaped
   *   - Matching is case-insensitive
   */
  private def wildcardToRegex(pattern: String): Regex =
    s"(?i)^${patternRegexBody(pattern)}$$".r // (?i) for case-insensitive

  /**
   * GH-476: the UNANCHORED regex body for an Excel wildcard pattern, as SEARCH reads it.
   *
   * SEARCH needs a substring search rather than the whole-value match criteria use. Here `~`
   * escapes only `*`, `?` and `~`, the three escapes Microsoft documents; before any other
   * character, and at the end, it is a literal tilde (not checked against Excel for SEARCH).
   * Criteria and lookups drop the tilde before any character ([[patternRegexBody]]).
   */
  def wildcardRegexBody(pattern: String): String =
    regexBody(pattern, tildeEscapesAny = false)

  /**
   * The regex body of a criterion or match_mode 2 pattern: `~x` is a literal `x` for every
   * character `x`, so desktop Excel's `"a~b*"` matches `abc` and `ab` and not `a~b`; a `~` at the
   * end is a literal tilde.
   */
  private def patternRegexBody(pattern: String): String =
    regexBody(pattern, tildeEscapesAny = true)

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def regexBody(pattern: String, tildeEscapesAny: Boolean): String =
    val sb = new StringBuilder
    var i = 0
    while i < pattern.length do
      val c = pattern.charAt(i)
      if c == '~' && i + 1 < pattern.length then
        // Escape sequence
        val next = pattern.charAt(i + 1)
        if tildeEscapesAny || next == '*' || next == '?' || next == '~' then
          sb.append(escapeRegexChar(next))
          i += 2
        else
          sb.append(escapeRegexChar(c))
          i += 1
      else if c == '*' then
        sb.append(".*")
        i += 1
      else if c == '?' then
        sb.append(".")
        i += 1
      else
        sb.append(escapeRegexChar(c))
        i += 1
    sb.toString

  /**
   * Escape a single character for use in a regex pattern.
   */
  private def escapeRegexChar(c: Char): String =
    if "\\^$.|?*+()[]{}".contains(c) then s"\\$c"
    else c.toString
