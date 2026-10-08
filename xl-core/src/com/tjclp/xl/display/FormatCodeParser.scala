package com.tjclp.xl.display

import java.math.BigInteger
import java.time.LocalDateTime

import scala.collection.mutable.ArrayBuffer
import scala.util.boundary, boundary.break

/**
 * Parser for Excel custom number format codes.
 *
 * Excel format codes have up to 4 sections separated by `;`:
 * {{{
 * positive ; negative ; zero ; text
 * }}}
 *
 * Key patterns:
 *   - `#,##0.00` - thousands + decimals
 *   - `"$"#,##0` - literal text (currency symbol)
 *   - `[Red]` - color modifier
 *   - `(#,##0)` - parentheses for negative
 *   - `_)` - space placeholder for alignment
 *   - `[$-409]` - locale code
 *
 * @since 0.2.0
 */
@SuppressWarnings(
  Array(
    "org.wartremover.warts.Var",
    "org.wartremover.warts.While"
  )
)
object FormatCodeParser:

  // ========== AST Types ==========

  /**
   * A complete format code with up to 4 sections.
   *
   * @param positive
   *   Format for positive numbers (required)
   * @param negative
   *   Format for negative numbers (optional)
   * @param zero
   *   Format for zero (optional)
   * @param text
   *   Format for text values (optional)
   */
  case class FormatCode(
    positive: FormatSection,
    negative: Option[FormatSection] = None,
    zero: Option[FormatSection] = None,
    text: Option[FormatSection] = None
  )

  /**
   * A single format section with stacked conditions and a pattern.
   *
   * @param conditions
   *   Leading bracket modifiers in order (e.g. `[Red][<=100]` stacks a color and a compare
   *   condition on the same section)
   * @param pattern
   *   The formatting pattern
   */
  case class FormatSection(
    conditions: Vector[Condition] = Vector.empty,
    pattern: FormatPattern
  ):
    /** The last parsed condition, if any (legacy accessor; conditions stack since GH-285). */
    def condition: Option[Condition] = conditions.lastOption

  /**
   * A condition modifier for a format section.
   */
  enum Condition derives CanEqual:
    /** Color condition: [Red], [Blue], [Green], [Magenta], [Cyan], [Yellow], [Black], [White] */
    case Color(name: String)

    /** Value comparison: [>100], [<0], [=5], [>=0], [<=100], [<>0] */
    case Compare(op: String, value: BigDecimal)

    /** Locale code: [$-409] (US English), [$€-407] (German Euro) */
    case Locale(code: String, symbol: Option[String])

  /**
   * A pattern consisting of format tokens.
   *
   * @param tokens
   *   Ordered sequence of tokens
   * @param hasThousands
   *   Whether the pattern uses thousands grouping
   * @param hasPercent
   *   Whether the pattern has percent (value * 100)
   */
  case class FormatPattern(
    tokens: Vector[FormatToken],
    hasThousands: Boolean = false,
    hasPercent: Boolean = false
  )

  /**
   * Individual tokens in a format pattern.
   */
  enum FormatToken derives CanEqual:
    /** Digit placeholder: 0 (show 0), # (hide 0), ? (space for 0) */
    case Digit(placeholder: Char)

    /** Decimal point */
    case Decimal

    /** Thousands separator (when in digit sequence) */
    case Thousands

    /**
     * Scaling comma: a comma after a digit placeholder with no integer-part placeholder after it
     * divides the value by 1000 (`#,##0,` and `#,##0,.0` show thousands, `0.0,,` millions; ECMA-376
     * §18.8.31, GH-666, #672).
     */
    case Scale

    /** Percent symbol - multiplies value by 100 */
    case Percent

    /**
     * The `General` keyword inside a section (ECMA-376 §18.8.30): renders the value in General
     * style in place, so `General"A"` on 2021 is `2021A` (GH-666).
     */
    case General

    /** Literal text (from "text" or escaped chars) */
    case Literal(text: String)

    /** Space placeholder: _x reserves space width of char x */
    case Spacer(char: Char)

    /** Fill: *x repeats char x to fill column width */
    case Fill(char: Char)

    /** Date/time part: m, d, y, h, s, etc. */
    case DatePart(part: String)

    /** AM/PM marker */
    case AmPm(format: String)

    /**
     * Fraction: digit-placeholder runs around `/` (e.g. `?/?`, `??/??`) or a fixed literal
     * denominator (e.g. `?/8`). `numerator`/`denominator` hold the raw placeholder runs;
     * `fixedDenominator` is defined when the denominator is a literal integer.
     */
    case Fraction(numerator: String, denominator: String, fixedDenominator: Option[Long])

    /** Elapsed time: [h], [m], [s] */
    case Elapsed(unit: Char)

    /** At sign: @ = text placeholder */
    case TextPlaceholder

    /**
     * Scientific exponent: `E`/`e` + sign convention + digit-placeholder run (GH-410). `E+` always
     * writes the exponent sign, `E-` only when the exponent is negative.
     */
    case Exponent(letter: Char, sign: Char, digits: String)

  // ========== Parser ==========

  /**
   * Parse an Excel format code string into structured AST.
   *
   * @param code
   *   The format code (e.g., "#,##0.00")
   * @return
   *   Either a parse error message or the parsed FormatCode
   */
  def parse(code: String): Either[String, FormatCode] =
    boundary:
      if code.isEmpty then break(Left("Empty format code"))

      val sections = splitSections(code)
      if sections.isEmpty then break(Left("No format sections found"))

      // Parse each section
      val parsedSections = sections.map(parseSection)
      val errors = parsedSections.collect { case Left(e) => e }
      if errors.nonEmpty then break(Left(errors.mkString("; ")))

      val validSections = parsedSections.collect { case Right(s) => s }

      // Build FormatCode from 1-4 sections
      validSections match
        case Vector(pos) => Right(FormatCode(pos))
        case Vector(pos, neg) => Right(FormatCode(pos, Some(neg)))
        case Vector(pos, neg, zero) => Right(FormatCode(pos, Some(neg), Some(zero)))
        case Vector(pos, neg, zero, txt) =>
          Right(FormatCode(pos, Some(neg), Some(zero), Some(txt)))
        case _ => Left(s"Invalid number of sections: ${validSections.size}")

  /**
   * Split format code into sections by semicolon. Respects quoted strings and brackets.
   *
   * Empty sections are preserved — including trailing ones — because an empty section means
   * "display nothing" for that value class (e.g. `0.0;;` hides negative and zero, GH-262).
   */
  private def splitSections(code: String): Vector[String] =
    val sections = ArrayBuffer[String]()
    val current = new StringBuilder
    var inQuotes = false
    var inBracket = false
    var i = 0

    while i < code.length do
      val c = code(i)
      c match
        case '"' if !inBracket =>
          inQuotes = !inQuotes
          current += c
        case '[' if !inQuotes =>
          inBracket = true
          current += c
        case ']' if !inQuotes =>
          inBracket = false
          current += c
        case '\\' | '_' | '*' if !inQuotes && i + 1 < code.length =>
          // An escape, spacer or fill takes the next character as its operand, so neither
          // `0\;;-0;0` nor `0_\;-0` splits at that `;` (#693)
          current += c += code(i + 1)
          i += 1
        case ';' if !inQuotes && !inBracket =>
          sections += current.toString
          current.clear()
        case _ =>
          current += c
      i += 1

    // Keep the final segment whenever a separator created it (sections.nonEmpty),
    // so trailing empty sections survive: "0.0;;" splits into three sections.
    if sections.nonEmpty || current.nonEmpty then sections += current.toString
    sections.toVector

  /**
   * Parse a single format section.
   */
  private def parseSection(section: String): Either[String, FormatSection] =
    boundary:
      var remaining = section
      var conditions = Vector.empty[Condition]

      // Extract leading conditions like [Red], [>100], [$-409] — they stack (GH-285)
      while remaining.startsWith("[") do
        val endBracket = remaining.indexOf(']')
        if endBracket < 0 then break(Left(s"Unclosed bracket in: $section"))

        val bracketContent = remaining.substring(1, endBracket)
        parseCondition(bracketContent) match
          case Some(cond) =>
            conditions = conditions :+ cond
          case None =>
            // Unknown bracket content - might be elapsed time [h], [m], [s]
            // Pass through as part of pattern
            val pattern = parsePattern(remaining)
            break(Right(FormatSection(conditions, pattern)))

        remaining = remaining.substring(endBracket + 1)

      val pattern = parsePattern(remaining)
      Right(FormatSection(conditions, pattern))

  /**
   * Parse a condition from bracket content.
   *
   * @param content
   *   The content inside brackets (without [ ])
   * @return
   *   Some(condition) if recognized, None if unknown
   */
  private def parseCondition(content: String): Option[Condition] =
    boundary:
      val lower = content.toLowerCase
      // Colors
      val colors = Set("red", "blue", "green", "yellow", "cyan", "magenta", "black", "white")
      if colors.contains(lower) then break(Some(Condition.Color(content.capitalize)))

      // Comparisons: >100, <0, =5, >=0, <=100, <>0
      val compPattern = "^([<>=]{1,2})(-?[0-9.]+)$".r
      content match
        case compPattern(op, num) =>
          scala.util.Try(BigDecimal(num)).toOption.map(n => Condition.Compare(op, n))
        case _ =>
          // Locale codes: $-409, $€-407
          if content.startsWith("$") then
            val rest = content.drop(1)
            val dashIdx = rest.indexOf('-')
            if dashIdx >= 0 then
              val symbol = if dashIdx > 0 then Some(rest.substring(0, dashIdx)) else None
              val code = rest.substring(dashIdx + 1)
              Some(Condition.Locale(code, symbol))
            else if rest.nonEmpty then Some(Condition.Locale(rest, None))
            else None
          else None

  /**
   * Parse the pattern portion of a format section.
   */
  private def parsePattern(pattern: String): FormatPattern =
    val tokens = ArrayBuffer[FormatToken]()
    var hasPercent = false
    var i = 0

    while i < pattern.length do
      val c = pattern(i)
      c match
        case '0' | '#' | '?' =>
          tokens += FormatToken.Digit(c)
          i += 1

        case '.' =>
          tokens += FormatToken.Decimal
          i += 1

        case ',' =>
          // Lexed as grouping; the post-pass below gives each comma its role: grouping, Scale
          // (÷1000 each, GH-666) or literal text (#672)
          tokens += FormatToken.Thousands
          i += 1

        case '%' =>
          hasPercent = true
          tokens += FormatToken.Percent
          i += 1

        case '"' =>
          // Quoted literal
          val endQuote = pattern.indexOf('"', i + 1)
          if endQuote > i then
            val text = pattern.substring(i + 1, endQuote)
            tokens += FormatToken.Literal(text)
            i = endQuote + 1
          else
            // Unclosed quote, take rest as literal
            tokens += FormatToken.Literal(pattern.substring(i + 1))
            i = pattern.length

        case '\\' =>
          // Escaped character
          if i + 1 < pattern.length then
            tokens += FormatToken.Literal(pattern(i + 1).toString)
            i += 2
          else i += 1

        case '_' =>
          // Spacer: _x
          if i + 1 < pattern.length then
            tokens += FormatToken.Spacer(pattern(i + 1))
            i += 2
          else i += 1

        case '*' =>
          // Fill: *x
          if i + 1 < pattern.length then
            tokens += FormatToken.Fill(pattern(i + 1))
            i += 2
          else i += 1

        case '@' =>
          tokens += FormatToken.TextPlaceholder
          i += 1

        case '[' =>
          // Elapsed time or already-parsed condition
          val endBracket = pattern.indexOf(']', i)
          if endBracket > i then
            val content = pattern.substring(i + 1, endBracket).toLowerCase
            if content == "h" || content == "m" || content == "s" then
              tokens += FormatToken.Elapsed(content.charAt(0))
            // Other bracket content (already processed as condition) - skip
            i = endBracket + 1
          else i += 1

        case 'G' | 'g' if pattern.regionMatches(true, i, "General", 0, 7) =>
          // ECMA-376 §18.8.30: the General keyword is a token wherever it appears in a section
          // (`General"A"`, `"FY"General`); quoted or escaped letters never reach here (GH-666)
          tokens += FormatToken.General
          i += 7

        case 'y' | 'Y' =>
          // Date year: y, yy, yyy, yyyy
          val start = i
          while i < pattern.length && (pattern(i) == 'y' || pattern(i) == 'Y') do i += 1
          tokens += FormatToken.DatePart(pattern.substring(start, i).toLowerCase)

        case 'm' | 'M' =>
          // Date month or time minute (context-dependent)
          val start = i
          while i < pattern.length && (pattern(i) == 'm' || pattern(i) == 'M') do i += 1
          tokens += FormatToken.DatePart(pattern.substring(start, i).toLowerCase)

        case 'd' | 'D' =>
          // Date day
          val start = i
          while i < pattern.length && (pattern(i) == 'd' || pattern(i) == 'D') do i += 1
          tokens += FormatToken.DatePart(pattern.substring(start, i).toLowerCase)

        case 'h' | 'H' =>
          // Time hour
          val start = i
          while i < pattern.length && (pattern(i) == 'h' || pattern(i) == 'H') do i += 1
          tokens += FormatToken.DatePart(pattern.substring(start, i).toLowerCase)

        case 's' | 'S' =>
          // Time second
          val start = i
          while i < pattern.length && (pattern(i) == 's' || pattern(i) == 'S') do i += 1
          tokens += FormatToken.DatePart(pattern.substring(start, i).toLowerCase)

        case 'A' | 'a' if pattern.regionMatches(true, i, "AM/PM", 0, 5) =>
          tokens += FormatToken.AmPm("AM/PM")
          i += 5

        case 'A' | 'a' if pattern.regionMatches(true, i, "A/P", 0, 3) =>
          tokens += FormatToken.AmPm("A/P")
          i += 3

        case 'E' | 'e'
            if i + 2 < pattern.length && (pattern(i + 1) == '+' || pattern(i + 1) == '-') &&
              isDigitPlaceholder(pattern(i + 2)) =>
          // Scientific exponent (GH-410): E/e, a sign convention, then digit placeholders.
          // A bare E without sign+placeholder stays a literal (date-era codes, plain text).
          var j = i + 2
          while j < pattern.length && isDigitPlaceholder(pattern(j)) do j += 1
          tokens += FormatToken.Exponent(c, pattern(i + 1), pattern.substring(i + 2, j))
          i = j

        case '/' =>
          // Fraction when digit placeholders directly flank the slash (GH-243): pop the
          // numerator run and consume the denominator (placeholder run or fixed integer).
          // Date separators (m/d/yy) reach here between DatePart tokens and stay literal.
          val denominator =
            val sb = new StringBuilder
            var j = i + 1
            while j < pattern.length && isFractionChar(pattern(j)) do
              sb += pattern(j)
              j += 1
            sb.toString
          var numeratorLen = 0
          while numeratorLen < tokens.length && (tokens(tokens.length - 1 - numeratorLen) match
              case FormatToken.Digit(_) => true
              case _ => false)
          do numeratorLen += 1
          if denominator.nonEmpty && numeratorLen > 0 then
            val numerator = tokens
              .takeRight(numeratorLen)
              .collect { case FormatToken.Digit(ch) => ch }
              .mkString
            tokens.dropRightInPlace(numeratorLen)
            val fixed =
              if denominator.forall(_.isDigit) && denominator.exists(_ != '0') then
                denominator.toLongOption
              else None
            tokens += FormatToken.Fraction(numerator, denominator, fixed)
            i += 1 + denominator.length
          else
            tokens += FormatToken.Literal("/")
            i += 1

        case ':' =>
          // Time separator
          tokens += FormatToken.Literal(":")
          i += 1

        case '-' | '+' | '(' | ')' | ' ' =>
          // Common literal characters
          tokens += FormatToken.Literal(c.toString)
          i += 1

        case _ =>
          // Other characters as literals
          tokens += FormatToken.Literal(c.toString)
          i += 1

    val classified = classifyScalingCommas(tokens.toVector)
    val hasThousands = classified.contains(FormatToken.Thousands)
    FormatPattern(classified, hasThousands, hasPercent)

  /**
   * Give every comma its role (ECMA-376 §18.8.31; GH-666, #672). A comma directly after a digit
   * placeholder — or after a scaling comma, an exponent or a fraction, which end in placeholders —
   * groups when a digit placeholder of the integer part follows it, and otherwise divides the value
   * by 1000 (Excel: "a comma that follows a digit placeholder scales the number by 1,000"). The
   * decimal point ends the integer part, so `#,##0,.0` and `0,,.0` scale like `#,##0,` and
   * `0.0,,"mm"`. Any other comma — after a literal, a space, `%`, the decimal point, `General`, a
   * date part, `@`, or nothing (a leading `,0`) — is literal text (LibreOffice: `0 ,` →
   * `12345678 ,`, `0"x",` → `12345678x,`, `,0` → `,1234`, `General,` → `12345,`, `mmm d, yyyy` →
   * `Mar 4, 2021`).
   */
  private def classifyScalingCommas(tokens: Vector[FormatToken]): Vector[FormatToken] =
    val decimalIdx = tokens.indexOf(FormatToken.Decimal)
    def isDigit(token: FormatToken): Boolean = token match
      case FormatToken.Digit(_) => true
      case _ => false
    tokens.zipWithIndex.foldLeft(Vector.empty[FormatToken]) { case (acc, (token, idx)) =>
      if token != FormatToken.Thousands then acc :+ token
      else
        val followsPlaceholder = acc.lastOption.exists {
          case FormatToken.Digit(_) | FormatToken.Scale | _: FormatToken.Exponent |
              _: FormatToken.Fraction =>
            true
          case _ => false
        }
        val integerEnd = if decimalIdx > idx then decimalIdx else tokens.length
        val groups = tokens.slice(idx + 1, integerEnd).exists(isDigit)
        acc :+ (
          if !followsPlaceholder then FormatToken.Literal(",")
          else if groups then FormatToken.Thousands
          else FormatToken.Scale
        )
    }

  /** Characters that may appear in a fraction numerator/denominator run. */
  private def isFractionChar(c: Char): Boolean =
    c == '#' || c == '?' || (c >= '0' && c <= '9')

  /** Digit placeholder characters: 0 (forced), # (optional), ? (space-padded). */
  private def isDigitPlaceholder(c: Char): Boolean =
    c == '0' || c == '#' || c == '?'

  // ========== Formatter ==========

  /**
   * Apply a parsed format code to a numeric value.
   *
   * Section selection follows Excel's rules (see [[selectSection]]):
   *   - 1 section: all numbers use it (negatives keep their default minus sign)
   *   - 2 sections: positive and zero use the 1st, negative uses the 2nd
   *   - 3+ sections: positive uses the 1st, negative the 2nd, zero the 3rd
   *   - compare conditions on the first two sections override positional routing (GH-285)
   *   - a trailing `@` section among fewer than 4 sections is the text section and never receives
   *     numbers
   *
   * Multi-section formats render the absolute value (any minus sign must be written in the
   * pattern); only single-section formats receive the default leading minus.
   *
   * @param value
   *   The number to format
   * @param format
   *   The parsed format code
   * @return
   *   Tuple of (formatted string, optional color)
   */
  def applyFormat(value: BigDecimal, format: FormatCode): (String, Option[String]) =
    applyFormat(value, format, NumFmtFormatter.GeneralRule.CellDisplay)

  /**
   * [[applyFormat]] with the `General` keyword rendered under `rule` (#672): cell display by
   * default, the text-conversion rule for TEXT(x, fmt).
   */
  def applyFormat(
    value: BigDecimal,
    format: FormatCode,
    rule: NumFmtFormatter.GeneralRule
  ): (String, Option[String]) =
    val section = selectSection(value, format).getOrElse(format.positive)
    val color = section.conditions.collectFirst { case Condition.Color(c) => c }
    val formatted = applyPattern(value, section.pattern, rule)
    val withDefaultSign =
      if value < 0 && numericSections(format).sizeIs <= 1 && formatted.nonEmpty &&
        !formatted.startsWith("-")
      then s"-$formatted"
      else formatted
    (withDefaultSign, color)

  /**
   * Route a numeric value to its format section per Excel semantics (GH-254/262/283/285).
   *
   * Mirrors SheetJS/SSF `choose_fmt` (reverse-engineered from Excel):
   *   - a trailing `@` section among fewer than 4 sections is the text section and drops out of
   *     numeric routing
   *   - without compare conditions, routing is positional with padding: 1 section serves all
   *     values, 2 sections serve [pos+zero, neg], 3+ serve [pos, neg, zero]
   *   - with compare conditions (honored on the first two sections only): the first matching
   *     condition wins; an unmatched value falls back to the third padded section when both leading
   *     sections carry conditions, otherwise to the second
   *
   * Returns None when the code has no numeric section (a lone `@` text format) — Excel renders such
   * numbers in General format.
   */
  def selectSection(value: BigDecimal, format: FormatCode): Option[FormatSection] =
    val sections = numericSections(format)
    if sections.isEmpty then None
    else
      val padded = sections.size match
        case 1 => Vector(sections(0), sections(0), sections(0))
        case 2 => Vector(sections(0), sections(1), sections(0))
        case _ => Vector(sections(0), sections(1), sections(2))
      val cmp1 = compareCondition(padded(0))
      val cmp2 = compareCondition(padded(1))
      val chosen =
        if cmp1.isEmpty && cmp2.isEmpty then
          if value > 0 then padded(0)
          else if value < 0 then padded(1)
          else padded(2) // Excel: zero uses the positive section when unpadded (GH-254)
        else if cmp1.exists(conditionMatches(value, _)) then padded(0)
        else if cmp2.exists(conditionMatches(value, _)) then padded(1)
        else if cmp1.isDefined && cmp2.isDefined then padded(2)
        else padded(1)
      Some(chosen)

  /**
   * The sections that participate in numeric routing: all sections except a trailing `@` text
   * section among fewer than 4 sections (SSF `choose_fmt` `lat` adjustment).
   */
  private def numericSections(format: FormatCode): Vector[FormatSection] =
    val all = Vector(format.positive) ++ format.negative ++ format.zero ++ format.text
    val trailingAt = all.sizeIs < 4 &&
      all.lastOption.exists(_.pattern.tokens.contains(FormatToken.TextPlaceholder))
    if trailingAt then all.dropRight(1) else all

  /**
   * True when the code has no numeric section — a lone `@` text format. Excel renders a number
   * under such a code as General text and lays the cell out as text: left-aligned, overflowing into
   * empty neighbours, never `####`. A 4-section code whose last arm is `@` keeps all four sections
   * and is NOT text-only (GH-501).
   */
  def isTextOnly(format: FormatCode): Boolean = numericSections(format).isEmpty

  private def compareCondition(section: FormatSection): Option[Condition.Compare] =
    section.conditions.collectFirst { case c: Condition.Compare => c }

  private def conditionMatches(value: BigDecimal, condition: Condition.Compare): Boolean =
    val cmp = value.compare(condition.value)
    condition.op match
      case "=" => cmp == 0
      case ">" => cmp > 0
      case "<" => cmp < 0
      case ">=" => cmp >= 0
      case "<=" => cmp <= 0
      case "<>" => cmp != 0
      case _ => false

  /**
   * Apply a format pattern to a number.
   *
   * Uses a simplified approach: collect pre-number literals, format the number, collect post-number
   * literals.
   */
  private def applyPattern(
    value: BigDecimal,
    pattern: FormatPattern,
    rule: NumFmtFormatter.GeneralRule
  ): String =
    val fracIdx = pattern.tokens.indexWhere {
      case _: FormatToken.Fraction => true
      case _ => false
    }
    pattern.tokens.lift(fracIdx) match
      case Some(f: FormatToken.Fraction) =>
        applyFractionPattern(value, pattern.tokens, fracIdx, f, rule)
      case _ =>
        val expIdx = pattern.tokens.indexWhere {
          case _: FormatToken.Exponent => true
          case _ => false
        }
        pattern.tokens.lift(expIdx) match
          case Some(e: FormatToken.Exponent) =>
            applyScientificPattern(value, pattern, expIdx, e)
          case _ => applyNumericPattern(value, pattern, rule)

  /**
   * Longest integer part, grouping commas included, that a digit pattern spells out in full (#689):
   * Excel's 32,767-character limit on a cell's text. Every number Excel can store (|x| < 1.8E308)
   * renders far inside it. A larger magnitude — only a corrupt or hostile file, or exact BigDecimal
   * arithmetic, carries one — renders its digit block in General form instead (`0` on
   * `1E+2147483647` is `1E+2147483647`), never materializing billions of digits.
   */
  val MaxDigitBlockLength: Int = 32767

  /**
   * Integer digits of `unscaled × 10^-scale` (unscaled ≥ 0) before rounding: at least 1. A Long,
   * because a percent or scaling comma moves the scale past the Int range at the extremes.
   */
  private def integerDigits(unscaled: BigInteger, scale: Long): Long =
    if unscaled.signum == 0 then 1L
    else math.max(1L, NumFmtFormatter.digitCount(unscaled).toLong - scale)

  /**
   * `unscaled × 10^-scale` (unscaled ≥ 0) rounded HALF_UP to an integer (#689). A division never
   * builds a power of ten longer than the value's own digits — anything below 0.1 is zero outright,
   * where BigDecimal.setScale asked for 10^2147483645 on `1E-2147483647`. A multiplication is as
   * long as the result, which callers bound first through [[integerDigits]].
   */
  private def roundToInteger(unscaled: BigInteger, scale: Long): BigInteger =
    if unscaled.signum == 0 then BigInteger.ZERO
    else if scale <= 0L then unscaled.multiply(BigInteger.TEN.pow((-scale).toInt))
    else if scale > NumFmtFormatter.digitCount(unscaled) then BigInteger.ZERO
    else
      val divisor = BigInteger.TEN.pow(scale.toInt)
      unscaled.add(divisor.shiftRight(1)).divide(divisor)

  private def applyNumericPattern(
    value: BigDecimal,
    pattern: FormatPattern,
    rule: NumFmtFormatter.GeneralRule
  ): String =
    val tokens = pattern.tokens
    // Percent multiplies by 100; each scaling comma divides by 1000 (exact: a decimal-point
    // move, so the rounding below sees the true scaled value, GH-666). The move is a Long power
    // of ten: a BigDecimal scale overflows Int at `1E±2147483647` (#689)
    val pow10 =
      (if pattern.hasPercent then 2L else 0L) - 3L * tokens.count(_ == FormatToken.Scale)
    val magnitude = value.bigDecimal.unscaledValue.abs
    // |adjusted value| = magnitude × 10^-scale
    val scale = value.scale.toLong - pow10

    // Count decimal places from pattern
    val decimalIdx = tokens.indexWhere(_ == FormatToken.Decimal)
    val hasDecimal = decimalIdx >= 0

    val decimalDigits =
      if hasDecimal then
        tokens.drop(decimalIdx + 1).count {
          case FormatToken.Digit(_) => true
          case _ => false
        }
      else 0

    // The integer part's placeholders: `0` pads with zeros, `?` with spaces, `#` with nothing
    val intPlaceholders = (if hasDecimal then tokens.take(decimalIdx) else tokens).collect {
      case FormatToken.Digit(c) => c
    }.mkString

    def groupedLength(digits: Long): Long =
      if pattern.hasThousands then digits + (digits - 1) / 3 else digits

    // The integer part (grouped, zero-padded) and the decimals, rounded HALF_UP; None when the
    // integer part would pass MaxDigitBlockLength. Bounded before anything is built (#689)
    lazy val digitBlock: Option[(String, String)] =
      if groupedLength(integerDigits(magnitude, scale)) > MaxDigitBlockLength then None
      else
        val rounded = roundToInteger(magnitude, scale - decimalDigits)
        val unit = BigInteger.TEN.pow(decimalDigits)
        val digits = rounded.divide(unit).toString
        // a rounding carry can add the one digit the bound above did not count
        Option.when(groupedLength(digits.length.toLong) <= MaxDigitBlockLength) {
          // A zero integer part shows only through a `0` placeholder (#693: `??` on 0 is blank)
          val shown = if digits == "0" && !intPlaceholders.contains('0') then "" else digits
          val intStr = if pattern.hasThousands then formatWithThousands(shown) else shown
          val paddedInt =
            padPlaceholders(shown, intPlaceholders, alignRight = true).dropRight(shown.length) +
              intStr
          val decimals = rounded.remainder(unit).toString
          val decStr =
            if decimalDigits > 0 then "0" * (decimalDigits - decimals.length) + decimals
            else ""
          (paddedInt, decStr)
        }

    // The General rendering of |adjusted value|: the General keyword, `@`, and the digit block
    // past MaxDigitBlockLength
    lazy val general = NumFmtFormatter.generalKeyword(magnitude, scale, rule)

    // Build result: prefix + number + suffix. A section with the General keyword renders the number
    // through it alone; placeholders beside it emit nothing (#681: `General0` printed the number
    // twice; Excel leaves such codes undefined, LibreOffice shows the General value alone). An `@`
    // renders the number in General form only where nothing else in the section does — no digit
    // placeholder, decimal point or General keyword (`@;@`, `@;0`): SSF emits the value for it
    // (#689: `TEXT(123456789012,"@;@")` was empty; Excel unverified). Beside them it renders
    // nothing, so `0.00 @` keeps its digits and `General @` shows the value once
    val atRendersNumber =
      tokens.contains(FormatToken.TextPlaceholder) && !tokens.exists {
        case FormatToken.Digit(_) | FormatToken.Decimal | FormatToken.General => true
        case _ => false
      }
    val result = new StringBuilder
    var numberEmitted = tokens.contains(FormatToken.General)

    for token <- tokens do
      token match
        case FormatToken.Digit(_) | FormatToken.Decimal | FormatToken.Thousands |
            FormatToken.Scale =>
          if !numberEmitted then
            // Emit the formatted number
            digitBlock match
              case Some((paddedInt, decStr)) =>
                result ++= paddedInt
                if hasDecimal then
                  result += '.'
                  result ++= decStr
              case None => result ++= general
            numberEmitted = true
          // Skip additional digit/decimal tokens

        case FormatToken.General =>
          // The keyword renders |x| in General style in place; applyFormat owns the sign, as
          // for digit patterns (GH-666)
          result ++= general

        case FormatToken.TextPlaceholder =>
          // `@` stands in for the number only in a section with no other way to render it (#689)
          if atRendersNumber then result ++= general

        case FormatToken.Percent =>
          result += '%'

        case FormatToken.Literal(text) =>
          result ++= text

        case FormatToken.Spacer(char) =>
          result += ' '

        case FormatToken.Fill(_) =>
          // Skip fill characters
          ()

        case _ =>
          // Date/time tokens not applicable
          ()

    result.toString

  /**
   * Render a scientific-notation pattern (GH-410).
   *
   * Excel semantics (ECMA-376 §18.8.30 ids 11 `0.00E+00` and 48 `##0.0E+0`, cross-checked against
   * the Excel-corpus-tested SheetJS/SSF):
   *   - the exponent snaps to the largest multiple of the mantissa's integer-placeholder count not
   *     exceeding floor(log10 |x|) — one placeholder is normalized scientific (mantissa in [1,
   *     10)), three is engineering notation
   *   - `E+` always writes the exponent sign; `E-` writes it only when the exponent is negative
   *   - exponent digits pad to the placeholder run and never truncate
   *   - a mantissa that rounds up to 10^placeholders renormalizes (9.999 → 1.00E+01)
   *
   * The sign of the value is handled by [[applyFormat]] (default leading minus).
   */
  private def applyScientificPattern(
    value: BigDecimal,
    pattern: FormatPattern,
    expIdx: Int,
    exp: FormatToken.Exponent
  ): String =
    val tokens = pattern.tokens
    // |x| (× 100 under a percent) = magnitude × 10^-scale; the scale is a Long, as in the plain
    // renderer (#689)
    val magnitude = value.bigDecimal.unscaledValue.abs
    val scale = value.scale.toLong - (if pattern.hasPercent then 2L else 0L)

    def digitCount(ts: Vector[FormatToken]): Int = ts.count {
      case FormatToken.Digit(_) => true
      case _ => false
    }
    val mantissaTokens = tokens.take(expIdx)
    val decimalIdx = mantissaTokens.indexWhere {
      case FormatToken.Decimal => true
      case _ => false
    }
    val intTokens = if decimalIdx >= 0 then mantissaTokens.take(decimalIdx) else mantissaTokens
    val intPlaceholders = math.max(digitCount(intTokens), 1)
    val decimalDigits =
      if decimalIdx >= 0 then digitCount(mantissaTokens.drop(decimalIdx + 1)) else 0
    val minIntDigits = intTokens.count {
      case FormatToken.Digit('0') => true
      case _ => false
    }

    // The mantissa in units of 10^-decimalDigits, and the exponent. Long exponent arithmetic: the
    // Int one wrapped at scale Int.MinValue (`1E+2147483648` printed E+2147483647, #689)
    val (mantissa, exponent) =
      if magnitude.signum == 0 then (BigInteger.ZERO, 0L)
      else
        // floor(log10 |x|), exact from the decimal representation (no Double detour)
        val e10 = NumFmtFormatter.digitCount(magnitude).toLong - scale - 1L
        val e = Math.floorDiv(e10, intPlaceholders.toLong) * intPlaceholders
        // |x| / 10^e to decimalDigits places: a shift within the mantissa's own few digits
        val rounded = roundToInteger(magnitude, scale + e - decimalDigits)
        // rounding can reach 10^placeholders exactly, never pass it
        if rounded.compareTo(BigInteger.TEN.pow(intPlaceholders + decimalDigits)) >= 0 then
          (rounded.divide(BigInteger.TEN.pow(intPlaceholders)), e + intPlaceholders)
        else (rounded, e)

    // "int" ++ "dec", zero-padded so the integer part has at least one digit
    val digits = mantissa.toString
    val plain =
      if digits.length <= decimalDigits then "0" * (decimalDigits + 1 - digits.length) + digits
      else digits
    val rawInt = plain.dropRight(decimalDigits)
    val decStr = plain.takeRight(decimalDigits)
    val paddedInt =
      if rawInt.length < minIntDigits then "0" * (minIntDigits - rawInt.length) + rawInt
      else rawInt
    val mantissaStr = if decimalIdx >= 0 then s"$paddedInt.$decStr" else paddedInt

    val absExp = math.abs(exponent).toString
    val paddedExp =
      if absExp.length < exp.digits.length then "0" * (exp.digits.length - absExp.length) + absExp
      else absExp
    val expSign =
      if exponent < 0 then "-"
      else if exp.sign == '+' then "+"
      else ""

    val result = new StringBuilder
    var numberEmitted = false
    tokens.foreach {
      case _: FormatToken.Exponent =>
        result ++= s"${exp.letter}$expSign$paddedExp"
      // Scale (GH-666) is a pass-through here: `0.0E+00,` does not divide and prints no comma
      // (LibreOffice: `0.0E+00,,` on 12345 is 1.2E+04; #672). Excel unverified; only the plain
      // numeric renderer scales
      case FormatToken.Digit(_) | FormatToken.Decimal | FormatToken.Thousands | FormatToken.Scale =>
        if !numberEmitted then
          result ++= mantissaStr
          numberEmitted = true
      case FormatToken.Percent =>
        result += '%'
      case FormatToken.Literal(text) =>
        result ++= text
      case FormatToken.Spacer(_) =>
        result += ' '
      case _ =>
        ()
    }
    result.toString

  /**
   * Format number string with thousands separators.
   */
  private def formatWithThousands(s: String): String =
    val result = new StringBuilder(s.length + s.length / 3)
    s.indices.foreach { i =>
      result += s(i)
      val posFromEnd = s.length - 1 - i
      if posFromEnd > 0 && posFromEnd % 3 == 0 then result += ','
    }
    result.toString

  /**
   * Render a fraction pattern (GH-243).
   *
   * Excel semantics (verified against the Excel-corpus-tested SheetJS/SSF algorithm):
   *   - variable denominators (`?/?`, `??/??`) use the last continued-fraction convergent whose
   *     denominator fits the placeholder budget (`10^digits - 1`, digits capped at 7)
   *   - fixed denominators (`?/8`) round to that denominator and never reduce (4/8 stays 4/8)
   *   - a whole value blanks the fraction area with spaces to preserve column alignment
   *   - unfilled `?` placeholders render as spaces, `0` as zeros, `#` as nothing; numerators
   *     right-align within their placeholders, denominators left-align
   *
   * The sign is handled by [[applyFormat]] (section literals or the default leading minus).
   *
   * Known divergences from Excel/SSF (corpus-checked, asymmetric-exotica only): codes with unequal
   * numerator/denominator placeholder widths (`# ??/?????????`) align per-placeholder here, while
   * Excel pads the numerator to the shared budget width `min(max(numLen, denLen), 7)` and
   * blank-fills `2*width + 1`; codes with spaces around the slash (`# ?? / ??`) are not tokenized
   * as fractions. Symmetric codes — everything Excel's Format Cells dialog offers — match the
   * Excel-generated corpus exactly.
   */
  private def applyFractionPattern(
    value: BigDecimal,
    tokens: Vector[FormatToken],
    fracIdx: Int,
    frac: FormatToken.Fraction,
    rule: NumFmtFormatter.GeneralRule
  ): String =
    val abs = value.abs
    // |value| = magnitude × 10^-scale, rounded through roundToInteger (#689)
    val magnitude = value.bigDecimal.unscaledValue.abs
    val scale = value.scale.toLong
    val wholePlaceholders = tokens
      .take(fracIdx)
      .collect { case FormatToken.Digit(ch) => ch }
      .mkString
    val mixed = wholePlaceholders.nonEmpty

    // The whole part's digits, numerator, denominator and the improper numerator's digits
    val (wholeText, num, den, improperText) =
      if integerDigits(magnitude, scale) > MaxDigitBlockLength then
        // A whole part past MaxDigitBlockLength renders in General form with no fractional part
        // (#689: `?/8` on `1E+2147483647` asked for two billion digits)
        val den = frac.fixedDenominator.getOrElse(1L)
        (
          NumFmtFormatter.generalKeyword(magnitude, scale, rule),
          BigInt(0),
          den,
          NumFmtFormatter.generalKeyword(magnitude.multiply(BigInteger.valueOf(den)), scale, rule)
        )
      else
        val (whole, num, den) = frac.fixedDenominator match
          case Some(d) =>
            val rr = BigInt(roundToInteger(magnitude.multiply(BigInteger.valueOf(d)), scale))
            (rr / d, rr % d, d)
          case None =>
            val digits = math.min(math.max(frac.numerator.length, frac.denominator.length), 7)
            val maxDen = math.pow(10, digits.toDouble) - 1
            // Excel stores values as IEEE-754 doubles and runs the search on the FULL value:
            // the whole part's binary noise is observable (12.3 → 12 1/3, but 0.3 → 2/7).
            val d = abs.toDouble
            if d.isInfinite then
              // |value| overflows Double (BigDecimal admits > ~1.8E308): the convergent search
              // cannot run, and no fractional part is representable at that magnitude anyway —
              // render the whole-number form (num == 0 blanks the fraction area), staying total.
              (BigInt(roundToInteger(magnitude, scale)), BigInt(0), 1L)
            else
              val (p, q) = nearestFraction(d, maxDen)
              val wholeD = math.floor(p / q)
              // p/q are exact-integer doubles for all values below 2^53 (Excel's own precision);
              // the max(0, _) keeps the numerator total for astronomically large inputs.
              (BigDecimal(wholeD).toBigInt, BigInt(math.max(0.0, p - wholeD * q).toLong), q.toLong)
        (whole.toString, num, den, (whole * den + num).toString)

    val denWidth = frac.fixedDenominator
      .fold(visibleWidth(frac.denominator))(_.toString.length)
    val fractionPart =
      if num == 0 && mixed then " " * (visibleWidth(frac.numerator) + 1 + denWidth)
      else
        val numerator = if mixed then num.toString else improperText
        val denStr = frac.fixedDenominator match
          case Some(d) => d.toString
          case None => padPlaceholders(den.toString, frac.denominator, alignRight = false)
        padPlaceholders(numerator, frac.numerator, alignRight = true) + "/" + denStr

    val wholeStr =
      if !mixed then ""
      else if wholeText != "0" then padPlaceholders(wholeText, wholePlaceholders, alignRight = true)
      else if num == 0 then "0"
      else padPlaceholders("", wholePlaceholders, alignRight = true)

    val result = new StringBuilder
    var wholeEmitted = false
    tokens.zipWithIndex.foreach { case (token, idx) =>
      token match
        case FormatToken.Digit(_) if idx < fracIdx =>
          if !wholeEmitted then
            result ++= wholeStr
            wholeEmitted = true
        case _: FormatToken.Fraction =>
          result ++= fractionPart
        case FormatToken.Literal(text) =>
          result ++= text
        case FormatToken.Spacer(_) =>
          result += ' '
        case _ =>
          () // fills, percent, date tokens: not meaningful inside fraction patterns
    }
    result.toString

  /**
   * Last continued-fraction convergent P/Q of `x` (non-negative) with Q <= maxDen.
   *
   * A verbatim port of the SheetJS/SSF `frac` algorithm (reverse-engineered from Excel and
   * validated against an Excel-generated corpus): convergents are generated in IEEE-754 double
   * arithmetic until the denominator budget is exceeded, then the previous convergent wins. Double
   * state is deliberate — Excel stores values as doubles and the binary noise is observable in the
   * chosen convergent (12.3 → 12 1/3 but 0.3 → 2/7).
   */
  private def nearestFraction(x: Double, maxDen: Double): (Double, Double) =
    var b = x
    var p2 = 0.0
    var p1 = 1.0
    var q2 = 1.0
    var q1 = 0.0
    var p = 0.0
    var q = 0.0
    var continue = true
    while continue && q1 < maxDen do
      val a = math.floor(b)
      p = a * p1 + p2
      q = a * q1 + q2
      if b - a < 0.00000005 then continue = false
      else
        b = 1.0 / (b - a)
        p2 = p1
        p1 = p
        q2 = q1
        q1 = q
    if q > maxDen then
      if q1 > maxDen then
        q = q2
        p = p2
      else
        q = q1
        p = p1
    (p, q)

  /**
   * Width a placeholder run occupies when blanked out: `?`, `0` and literal digits reserve one
   * space each; `#` reserves nothing.
   */
  private def visibleWidth(placeholders: String): Int =
    placeholders.count(c => c == '?' || (c >= '0' && c <= '9'))

  /**
   * Align digits within a placeholder run: unfilled `?` positions become spaces, `0` becomes zeros,
   * `#` adds nothing. Numerators/wholes right-align (pad left), denominators left-align (pad
   * right).
   */
  private def padPlaceholders(digits: String, placeholders: String, alignRight: Boolean): String =
    val diff = placeholders.length - digits.length
    if diff <= 0 then digits
    else
      val unfilled = if alignRight then placeholders.take(diff) else placeholders.takeRight(diff)
      val fill = unfilled.flatMap {
        case '?' => " "
        case '0' => "0"
        case _ => ""
      }
      if alignRight then fill + digits else digits + fill

  /**
   * Apply a format code to text.
   *
   * The text section is the 4th section when present; otherwise a trailing `@` section among fewer
   * than 4 sections (GH-285). Without either, text echoes unchanged.
   */
  def applyTextFormat(text: String, format: FormatCode): String =
    val all = Vector(format.positive) ++ format.negative ++ format.zero ++ format.text
    val textSection = format.text.orElse(
      if all.sizeIs < 4 then
        all.lastOption.filter(_.pattern.tokens.contains(FormatToken.TextPlaceholder))
      else None
    )
    textSection match
      case Some(section) =>
        section.pattern.tokens.map {
          case FormatToken.TextPlaceholder | FormatToken.General => text
          case FormatToken.Literal(s) => s
          case _ => ""
        }.mkString
      case None => text

  // ========== Date/Time Formatter ==========

  /**
   * Check if a format code's first section contains date/time tokens.
   */
  def hasDateTokens(format: FormatCode): Boolean =
    hasDateTokens(format.positive)

  /**
   * Check if a single format section contains date/time tokens (GH-283: sections of one code may
   * mix calendar and numeric patterns, so callers route per section).
   */
  def hasDateTokens(section: FormatSection): Boolean =
    section.pattern.tokens.exists {
      case FormatToken.DatePart(_) | FormatToken.AmPm(_) | FormatToken.Elapsed(_) => true
      case _ => false
    }

  /**
   * Apply a parsed format code's first section to a LocalDateTime value.
   *
   * @param dt
   *   The datetime to format
   * @param format
   *   The parsed format code
   * @return
   *   Formatted date/time string
   */
  def applyDateFormat(dt: LocalDateTime, format: FormatCode): String =
    applyDateFormat(dt, format.positive)

  /**
   * Apply a single format section to a LocalDateTime value (GH-283: the section is chosen by
   * [[selectSection]] on the date's serial number).
   *
   * The date renders in Excel's 1900 calendar (#688), as the cell would show it: 1899-12-31 is
   * `1900-01-00` (serial 0), and dates before 1900-03-01 take Excel's weekday, one before the real
   * one. Other dates render their own fields.
   */
  def applyDateFormat(dt: LocalDateTime, section: FormatSection): String =
    applyDateFormat(ExcelCalendar.of(dt), section)

  /**
   * The calendar fields an Excel date displays (#688). Excel's 1900 date system has two days no
   * LocalDateTime can hold: serial 0 shows as `1/0/1900` and the phantom serial 60 as `2/29/1900`.
   * Its weekdays also run off the serial (serial 1 is a Sunday), so before 1900-03-01 every day
   * shows the weekday before the real one.
   *
   * 1900 date system only: the formatter has no 1904 context, and none of these quirks exist there.
   *
   * @param weekday
   *   ISO day of week, 1 = Monday to 7 = Sunday
   */
  private[xl] final case class ExcelCalendar(
    year: Int,
    month: Int,
    day: Int,
    weekday: Int,
    hour: Int,
    minute: Int,
    second: Int
  ) derives CanEqual

  private[xl] object ExcelCalendar:
    private val serialZero = java.time.LocalDate.of(1899, 12, 31)
    private val phantomLeapDayEnd = java.time.LocalDate.of(1900, 3, 1)

    /**
     * The view of a LocalDateTime in Excel's 1900 date system: 1899-12-31 (serial 0) is
     * `1900-01-00`, and dates before 1900-03-01 take the weekday Excel's serial count gives them.
     * Other dates are their own calendar fields.
     */
    def of(dt: LocalDateTime): ExcelCalendar =
      val date = dt.toLocalDate
      val real = ExcelCalendar(
        dt.getYear,
        dt.getMonthValue,
        dt.getDayOfMonth,
        dt.getDayOfWeek.getValue,
        dt.getHour,
        dt.getMinute,
        dt.getSecond
      )
      if date.isBefore(serialZero) || !date.isBefore(phantomLeapDayEnd) then real
      else
        // Excel counts 1900-02-29, so its weekdays before the phantom day run one behind
        val shifted = real.copy(weekday = if real.weekday == 1 then 7 else real.weekday - 1)
        if date == serialZero then shifted.copy(year = 1900, month = 1, day = 0) else shifted

    /** Excel's phantom 1900-02-29 in the 1900 date system. */
    private val PhantomLeapDaySerial = 60L

    /**
     * The calendar Excel displays for a date serial (#688): the day comes from
     * [[com.tjclp.xl.cells.CellValue.excelSerialToDateTime]], which carries the leap-year offset
     * for serials below 60, so 1 is 1900-01-01 and 59 is 1900-02-28. Serial 0 shows as `1/0/1900`
     * and serial 60 as `2/29/1900` (a Wednesday), the day Excel counts and LibreOffice does not
     * (LibreOffice shows 2/28/1900). The time of day is the fraction truncated to the second.
     *
     * @param serial
     *   Excel date serial number, in Excel's displayable range [0, 2958466)
     */
    def fromSerial(serial: BigDecimal): ExcelCalendar =
      val days = serial.toLong
      val timeFraction = (serial % 1).toDouble
      val hours = (timeFraction * 24).toInt
      val minutes = ((timeFraction * 24 * 60) % 60).toInt
      val seconds = (((timeFraction * 24 * 60 * 60) % 60)).toInt
      if days == PhantomLeapDaySerial then ExcelCalendar(1900, 2, 29, 3, hours, minutes, seconds)
      else
        val date = com.tjclp.xl.cells.CellValue.excelSerialToDateTime(days.toDouble).toLocalDate
        of(date.atTime(hours, minutes, seconds))

  /** [[applyDateFormat]] on Excel calendar fields (#688: the days a LocalDateTime cannot hold). */
  private[xl] def applyDateFormat(dt: ExcelCalendar, section: FormatSection): String =
    val tokens = section.pattern.tokens
    val minutePositions = findMinutePositions(tokens)
    // ECMA-376 §18.8.31: the hour uses the 12-hour clock only when the code contains an
    // AM/PM marker; otherwise it is the 24-hour clock (GH-410).
    val twelveHour = tokens.exists {
      case FormatToken.AmPm(_) => true
      case _ => false
    }
    tokens.zipWithIndex.map { case (token, idx) =>
      renderDateToken(dt, token, minutePositions.contains(idx), twelveHour)
    }.mkString

  /**
   * Find positions where 'm'/'mm' should be interpreted as minute (not month).
   *
   * Excel rule: 'm' after 'h' or before 's' = minute, otherwise month.
   */
  private def findMinutePositions(tokens: Vector[FormatToken]): Set[Int] =
    val positions = scala.collection.mutable.Set[Int]()

    // Find all 'm'/'mm' token positions
    val mPositions = tokens.zipWithIndex.collect {
      case (FormatToken.DatePart(p), idx) if p == "m" || p == "mm" => idx
    }

    // For each 'm' token, check if it's in time context
    for mIdx <- mPositions do
      // Look backwards for 'h'/'hh' (skip literals)
      val hasHourBefore = tokens.take(mIdx).reverse.exists {
        case FormatToken.DatePart(p) if p.startsWith("h") => true
        case FormatToken.DatePart(_) => false // Found date token first
        case _ => false // Skip literals
      }

      // Look forwards for 's'/'ss' (skip literals)
      val hasSecondAfter = tokens.drop(mIdx + 1).exists {
        case FormatToken.DatePart(p) if p.startsWith("s") => true
        case FormatToken.DatePart(_) => false // Found date token first
        case _ => false // Skip literals
      }

      if hasHourBefore || hasSecondAfter then positions += mIdx

    positions.toSet

  // English month/weekday tables, indexed by getMonthValue - 1 / DayOfWeek.getValue - 1
  // (Monday-first). Hardcoded rather than TextStyle.getDisplayName(…, Locale.US): the locale was
  // already pinned to US English, and locale data diverges across platforms (Scala Native renders
  // these differently, ADR-016). FormatCodeParserSpec pins every entry at all reachable widths.
  private val MonthsShort: Vector[String] =
    Vector("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
  private val MonthsFull: Vector[String] = Vector(
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
  )
  private val MonthsNarrow: Vector[String] =
    Vector("J", "F", "M", "A", "M", "J", "J", "A", "S", "O", "N", "D")
  private val WeekdaysShort: Vector[String] =
    Vector("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
  private val WeekdaysFull: Vector[String] =
    Vector("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")

  /**
   * Render a single date/time token.
   *
   * @param dt
   *   The datetime value
   * @param token
   *   The format token
   * @param isMinute
   *   Whether 'm'/'mm' should render as minute (vs month)
   * @param twelveHour
   *   Whether 'h'/'hh' use the 12-hour clock (the section has an AM/PM marker, GH-410)
   */
  private def renderDateToken(
    dt: ExcelCalendar,
    token: FormatToken,
    isMinute: Boolean,
    twelveHour: Boolean
  ): String =
    token match
      // Year
      case FormatToken.DatePart("y") =>
        (dt.year % 100).toString
      case FormatToken.DatePart("yy") =>
        f"${dt.year % 100}%02d"
      case FormatToken.DatePart("yyy" | "yyyy") =>
        dt.year.toString

      // Month (or minute if isMinute)
      case FormatToken.DatePart("m") =>
        if isMinute then dt.minute.toString
        else dt.month.toString
      case FormatToken.DatePart("mm") =>
        if isMinute then f"${dt.minute}%02d"
        else f"${dt.month}%02d"
      case FormatToken.DatePart("mmm") =>
        MonthsShort(dt.month - 1)
      case FormatToken.DatePart("mmmm") =>
        MonthsFull(dt.month - 1)
      case FormatToken.DatePart("mmmmm") =>
        // First letter only (J, F, M, A, ...)
        MonthsNarrow(dt.month - 1)

      // Day
      case FormatToken.DatePart("d") =>
        dt.day.toString
      case FormatToken.DatePart("dd") =>
        f"${dt.day}%02d"
      case FormatToken.DatePart("ddd") =>
        WeekdaysShort(dt.weekday - 1)
      case FormatToken.DatePart("dddd") =>
        WeekdaysFull(dt.weekday - 1)

      // Hour: 12-hour clock only when the section carries AM/PM (ECMA-376 §18.8.31)
      case FormatToken.DatePart("h") =>
        if twelveHour then
          val hour = dt.hour % 12
          (if hour == 0 then 12 else hour).toString
        else dt.hour.toString
      case FormatToken.DatePart("hh") =>
        if twelveHour then
          val hour = dt.hour % 12
          f"${if hour == 0 then 12 else hour}%02d"
        else f"${dt.hour}%02d"

      // Second
      case FormatToken.DatePart("s") =>
        dt.second.toString
      case FormatToken.DatePart("ss") =>
        f"${dt.second}%02d"

      // AM/PM
      case FormatToken.AmPm("AM/PM") =>
        if dt.hour < 12 then "AM" else "PM"
      case FormatToken.AmPm("A/P") =>
        if dt.hour < 12 then "A" else "P"
      case FormatToken.AmPm(_) =>
        if dt.hour < 12 then "AM" else "PM"

      // Elapsed time (for duration formatting - just show value for now)
      case FormatToken.Elapsed('h') =>
        dt.hour.toString
      case FormatToken.Elapsed('m') =>
        dt.minute.toString
      case FormatToken.Elapsed('s') =>
        dt.second.toString

      // Literals and other tokens
      case FormatToken.Literal(text) =>
        text
      case FormatToken.Spacer(_) =>
        " "
      case FormatToken.Thousands | FormatToken.Scale =>
        // In date context, comma is a literal (not thousands separator)
        ","
      case _ =>
        ""
