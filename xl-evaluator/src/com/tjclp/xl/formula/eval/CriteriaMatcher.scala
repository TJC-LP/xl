package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.Clock

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.ArrayArithmetic
import scala.util.boundary
import boundary.break
import scala.util.matching.Regex

/**
 * Excel-compatible criteria matching for SUMIF/COUNTIF functions.
 *
 * Supports 5 criteria patterns (Excel-compatible):
 *   - **Exact**: `"Apple"` or `42` → Cells equal to value
 *   - **Wildcard**: `"A*"`, `"*pple"`, `"A?ple"` → Pattern matching (* = any chars, ? = single
 *     char)
 *   - **Greater**: `">100"`, `">=50"` → Numeric comparison
 *   - **Less**: `"<10"`, `"<=5"` → Numeric comparison
 *   - **Not Equal**: `"<>0"` → Numeric inequality
 *
 * Laws:
 *   1. Exact match: `matches(Text("Apple"), Exact(ExprValue.Text("Apple"))) == true`
 *   2. Wildcard: `matches(Text("Apple"), Wildcard("A*")) == true`
 *   3. Case-insensitive text: `matches(Text("APPLE"), Exact(ExprValue.Text("apple"))) == true`
 *   4. Numeric comparison: `matches(Number(150), Compare(Gt, 100)) == true`
 */
object CriteriaMatcher:

  /** Parsed criterion for cell matching */
  sealed trait Criterion derives CanEqual

  /** Exact match against text, number, or boolean */
  case class Exact(value: ExprValue) extends Criterion

  /** Numeric comparison */
  case class Compare(op: CompareOp, value: BigDecimal) extends Criterion

  /** Wildcard pattern (* = any chars, ? = single char) */
  case class Wildcard(pattern: String) extends Criterion

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
        case CellValue.Formula(_, Some(cached)) => parse(ExprValue.Cell(cached))
        case CellValue.Formula(_, None) => Exact(ExprValue.Cell(cellValue))
        case CellValue.Empty => Exact(ExprValue.Text(""))
        case CellValue.Error(_) => Exact(ExprValue.Cell(cellValue))
    case ExprValue.Opaque(other) => Exact(ExprValue.Opaque(other))

  /**
   * Parse string criterion into operator, wildcard, or exact match.
   */
  private def parseString(s: String): Criterion =
    s match
      // Not equal: <>
      case _ if s.startsWith("<>") =>
        parseNumeric(s.drop(2)) match
          case Some(n) => Compare(CompareOp.Neq, n)
          case None => Exact(ExprValue.Text(s)) // Couldn't parse number, treat as literal
      // Greater than or equal: >=
      case _ if s.startsWith(">=") =>
        parseNumeric(s.drop(2)) match
          case Some(n) => Compare(CompareOp.Gte, n)
          case None => Exact(ExprValue.Text(s))
      // Less than or equal: <=
      case _ if s.startsWith("<=") =>
        parseNumeric(s.drop(2)) match
          case Some(n) => Compare(CompareOp.Lte, n)
          case None => Exact(ExprValue.Text(s))
      // Greater than: >
      case _ if s.startsWith(">") =>
        parseNumeric(s.drop(1)) match
          case Some(n) => Compare(CompareOp.Gt, n)
          case None => Exact(ExprValue.Text(s))
      // Less than: <
      case _ if s.startsWith("<") =>
        parseNumeric(s.drop(1)) match
          case Some(n) => Compare(CompareOp.Lt, n)
          case None => Exact(ExprValue.Text(s))
      // Equal: =
      case _ if s.startsWith("=") =>
        parseNumeric(s.drop(1)) match
          case Some(n) => Exact(ExprValue.Number(n))
          case None => Exact(ExprValue.Text(s.drop(1))) // Keep text after =
      // Wildcard pattern
      case _ if hasUnescapedWildcard(s) =>
        Wildcard(s)
      // Has escape sequences but no wildcards - unescape and exact match
      case _ if hasEscapeSequence(s) =>
        Exact(ExprValue.Text(unescapePattern(s)))
      // Plain string - exact match
      case _ =>
        Exact(ExprValue.Text(s))

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
   * Check if string contains escape sequences (~*, ~?, or ~~).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def hasEscapeSequence(s: String): Boolean =
    boundary {
      var i = 0
      while i < s.length - 1 do
        if s.charAt(i) == '~' then
          val next = s.charAt(i + 1)
          if next == '*' || next == '?' || next == '~' then break(true)
        i += 1
      false
    }

  /**
   * Remove escape characters from pattern (unescape ~* ~? ~~).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def unescapePattern(s: String): String =
    val sb = new StringBuilder
    var i = 0
    while i < s.length do
      val c = s.charAt(i)
      if c == '~' && i + 1 < s.length then
        val next = s.charAt(i + 1)
        if next == '*' || next == '?' || next == '~' then
          sb.append(next)
          i += 2
        else
          sb.append(c)
          i += 1
      else
        sb.append(c)
        i += 1
    sb.toString

  /**
   * Parse string as BigDecimal, handling Excel's flexible numeric syntax.
   */
  private def parseNumeric(s: String): Option[BigDecimal] =
    val trimmed = s.trim
    if trimmed.isEmpty then None
    else
      scala.util
        .Try(BigDecimal(trimmed))
        .toOption

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
      case Wildcard(pattern) => matchesWildcard(cellValue, pattern)

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

      case CellValue.Formula(_, Some(cached)) =>
        // Match against cached value
        matchesExact(cached, expected)

      case CellValue.Formula(_, None) =>
        false

      case CellValue.RichText(rt) =>
        expected match
          case ExprValue.Text(s) => rt.toPlainText.equalsIgnoreCase(s)
          case _ => false

      case CellValue.Error(_) =>
        false

  /**
   * Numeric comparison matching.
   */
  private def matchesCompare(
    cellValue: CellValue,
    op: CompareOp,
    threshold: BigDecimal
  ): Boolean =
    extractNumeric(cellValue) match
      case Some(value) =>
        op match
          case CompareOp.Gt => value > threshold
          case CompareOp.Gte => value >= threshold
          case CompareOp.Lt => value < threshold
          case CompareOp.Lte => value <= threshold
          case CompareOp.Neq => value != threshold
      case None => false

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
      case CellValue.Formula(_, Some(cached)) => extractNumeric(cached)
      // GH-197: Uncached formula - can't extract numeric without evaluation context
      case CellValue.Formula(_, None) => None
      case _ => None

  /**
   * Wildcard pattern matching.
   *
   * Excel wildcards:
   *   - `*` matches any sequence of characters (including empty)
   *   - `?` matches exactly one character
   *   - Matching is case-insensitive
   *   - `~*` escapes literal asterisk
   *   - `~?` escapes literal question mark
   *   - `~~` escapes literal tilde
   */
  private def matchesWildcard(cellValue: CellValue, pattern: String): Boolean =
    extractText(cellValue) match
      case Some(text) =>
        val regex = wildcardToRegex(pattern)
        regex.matches(text)
      case None => false

  /**
   * Extract text value from cell for wildcard matching.
   */
  private def extractText(cellValue: CellValue): Option[String] =
    cellValue match
      case CellValue.Text(s) => Some(s)
      case CellValue.Number(n) =>
        // Format without trailing zeros for matching
        val plain = n.bigDecimal.stripTrailingZeros().toPlainString
        Some(plain)
      case CellValue.Bool(b) => Some(if b then "TRUE" else "FALSE")
      case CellValue.RichText(rt) => Some(rt.toPlainText)
      case CellValue.Formula(_, Some(cached)) => extractText(cached)
      // GH-197: Uncached formula - can't extract text without evaluation context
      case CellValue.Formula(_, None) => None
      case _ => None

  /**
   * Convert Excel wildcard pattern to regex.
   *
   * Handles:
   *   - `*` → `.*`
   *   - `?` → `.`
   *   - `~*` → literal asterisk
   *   - `~?` → literal question mark
   *   - `~~` → literal tilde
   *   - Other regex metacharacters are escaped
   *   - Matching is case-insensitive
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def wildcardToRegex(pattern: String): Regex =
    val sb = new StringBuilder("(?i)^") // (?i) for case-insensitive
    var i = 0
    while i < pattern.length do
      val c = pattern.charAt(i)
      if c == '~' && i + 1 < pattern.length then
        // Escape sequence
        val next = pattern.charAt(i + 1)
        if next == '*' || next == '?' || next == '~' then
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
    sb.append("$")
    sb.toString.r

  /**
   * Escape a single character for use in a regex pattern.
   */
  private def escapeRegexChar(c: Char): String =
    if "\\^$.|?*+()[]{}".contains(c) then s"\\$c"
    else c.toString
