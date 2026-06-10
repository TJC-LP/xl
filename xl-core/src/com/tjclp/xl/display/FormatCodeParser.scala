package com.tjclp.xl.display

import java.time.LocalDateTime
import java.time.format.TextStyle
import java.util.Locale

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
   * A single format section with optional condition and pattern.
   *
   * @param condition
   *   Optional color or value condition (e.g., [Red], [>100])
   * @param pattern
   *   The formatting pattern
   */
  case class FormatSection(
    condition: Option[Condition] = None,
    pattern: FormatPattern
  )

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

    /** Percent symbol - multiplies value by 100 */
    case Percent

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
      var condition: Option[Condition] = None

      // Extract leading conditions like [Red], [>100], [$-409]
      while remaining.startsWith("[") do
        val endBracket = remaining.indexOf(']')
        if endBracket < 0 then break(Left(s"Unclosed bracket in: $section"))

        val bracketContent = remaining.substring(1, endBracket)
        parseCondition(bracketContent) match
          case Some(cond) =>
            condition = Some(cond)
          case None =>
            // Unknown bracket content - might be elapsed time [h], [m], [s]
            // Pass through as part of pattern
            val pattern = parsePattern(remaining)
            break(Right(FormatSection(condition, pattern)))

        remaining = remaining.substring(endBracket + 1)

      val pattern = parsePattern(remaining)
      Right(FormatSection(condition, pattern))

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
    var hasThousands = false
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
          // Thousands separator (when between digits)
          // Also can mean scale by 1000 at end of number
          hasThousands = true
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

    FormatPattern(tokens.toVector, hasThousands, hasPercent)

  /** Characters that may appear in a fraction numerator/denominator run. */
  private def isFractionChar(c: Char): Boolean =
    c == '#' || c == '?' || (c >= '0' && c <= '9')

  // ========== Formatter ==========

  /**
   * Apply a parsed format code to a numeric value.
   *
   * Section selection follows Excel's rules:
   *   - 1 section: all numbers use it (negatives keep their default minus sign)
   *   - 2 sections: positive and zero use the 1st, negative uses the 2nd
   *   - 3+ sections: positive uses the 1st, negative the 2nd, zero the 3rd
   *
   * @param value
   *   The number to format
   * @param format
   *   The parsed format code
   * @return
   *   Tuple of (formatted string, optional color)
   */
  def applyFormat(value: BigDecimal, format: FormatCode): (String, Option[String]) =
    // Select appropriate section based on value sign
    val (section, effectiveValue) =
      if value > 0 then (format.positive, value)
      else if value < 0 then
        format.negative match
          case Some(neg) => (neg, value.abs) // Negative section uses absolute value
          case None => (format.positive, value) // Fall back to positive
      else
        format.zero match
          case Some(z) => (z, value)
          case None => (format.positive, value) // Excel: zero uses positive section (GH-254)

    val color = section.condition.collect { case Condition.Color(c) => c }
    val formatted = applyPattern(effectiveValue, section.pattern)
    val withDefaultSign =
      if value < 0 && format.negative.isEmpty && !formatted.startsWith("-") then s"-$formatted"
      else formatted
    (withDefaultSign, color)

  /**
   * Apply a format pattern to a number.
   *
   * Uses a simplified approach: collect pre-number literals, format the number, collect post-number
   * literals.
   */
  private def applyPattern(value: BigDecimal, pattern: FormatPattern): String =
    val fracIdx = pattern.tokens.indexWhere {
      case _: FormatToken.Fraction => true
      case _ => false
    }
    pattern.tokens.lift(fracIdx) match
      case Some(f: FormatToken.Fraction) => applyFractionPattern(value, pattern.tokens, fracIdx, f)
      case _ => applyNumericPattern(value, pattern)

  private def applyNumericPattern(value: BigDecimal, pattern: FormatPattern): String =
    // Handle percent: multiply by 100
    val adjustedValue = if pattern.hasPercent then value * 100 else value

    // Count decimal places from pattern
    val tokens = pattern.tokens
    val decimalIdx = tokens.indexWhere(_ == FormatToken.Decimal)
    val hasDecimal = decimalIdx >= 0

    val decimalDigits =
      if hasDecimal then
        tokens.drop(decimalIdx + 1).count {
          case FormatToken.Digit(_) => true
          case _ => false
        }
      else 0

    // Count minimum integer digits (0 placeholders)
    val minIntDigits = tokens.count {
      case FormatToken.Digit('0') => true
      case _ => false
    } - decimalDigits

    // Round to decimal places
    val rounded = adjustedValue.setScale(decimalDigits, BigDecimal.RoundingMode.HALF_UP)
    val absValue = rounded.abs
    val intPart = absValue.toBigInt
    val decPart =
      if decimalDigits > 0 then
        ((absValue - BigDecimal(intPart)) * BigDecimal(10).pow(decimalDigits))
          .setScale(0, BigDecimal.RoundingMode.HALF_UP)
          .toBigInt
      else BigInt(0)

    // Format integer part with thousands grouping
    val intStr =
      if pattern.hasThousands then formatWithThousands(intPart.toString)
      else intPart.toString

    // Pad integer with leading zeros if needed
    val paddedInt =
      if minIntDigits > 0 && intStr.length < minIntDigits then
        "0" * (minIntDigits - intStr.length) + intStr
      else intStr

    // Format decimal part with trailing zeros
    val decStr =
      if decimalDigits > 0 then decPart.toString.reverse.padTo(decimalDigits, '0').reverse
      else ""

    // Build result: prefix + number + suffix
    val result = new StringBuilder
    var inNumber = false
    var numberEmitted = false

    for token <- tokens do
      token match
        case FormatToken.Digit(_) | FormatToken.Decimal | FormatToken.Thousands =>
          if !numberEmitted then
            // Emit the formatted number
            result ++= paddedInt
            if hasDecimal then
              result += '.'
              result ++= decStr
            numberEmitted = true
          // Skip additional digit/decimal tokens

        case FormatToken.Percent =>
          result += '%'

        case FormatToken.Literal(text) =>
          result ++= text

        case FormatToken.Spacer(char) =>
          result += ' '

        case FormatToken.Fill(_) =>
          // Skip fill characters
          ()

        case FormatToken.TextPlaceholder =>
          // @ not applicable for numbers
          ()

        case _ =>
          // Date/time tokens not applicable
          ()

    result.toString

  /**
   * Format number string with thousands separators.
   */
  private def formatWithThousands(s: String): String =
    val chars = s.toVector
    val len = chars.length
    chars.zipWithIndex.map { case (c, i) =>
      val posFromEnd = len - 1 - i
      if posFromEnd > 0 && posFromEnd % 3 == 0 then s"$c,"
      else c.toString
    }.mkString

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
   */
  private def applyFractionPattern(
    value: BigDecimal,
    tokens: Vector[FormatToken],
    fracIdx: Int,
    frac: FormatToken.Fraction
  ): String =
    val abs = value.abs
    val wholePlaceholders = tokens
      .take(fracIdx)
      .collect { case FormatToken.Digit(ch) => ch }
      .mkString
    val mixed = wholePlaceholders.nonEmpty

    val (whole, num, den) = frac.fixedDenominator match
      case Some(d) =>
        val rr = (abs * d).setScale(0, BigDecimal.RoundingMode.HALF_UP).toBigInt
        (rr / d, rr % d, d)
      case None =>
        val digits = math.min(math.max(frac.numerator.length, frac.denominator.length), 7)
        val maxDen = math.pow(10, digits.toDouble) - 1
        // Excel stores values as IEEE-754 doubles and runs the search on the FULL value:
        // the whole part's binary noise is observable (12.3 → 12 1/3, but 0.3 → 2/7).
        val (p, q) = nearestFraction(abs.toDouble, maxDen)
        val wholeD = math.floor(p / q)
        // p/q are exact-integer doubles for all values below 2^53 (Excel's own precision);
        // the max(0, _) keeps the numerator total for astronomically large inputs.
        (BigDecimal(wholeD).toBigInt, BigInt(math.max(0.0, p - wholeD * q).toLong), q.toLong)

    val improperNumerator = whole * den + num

    val denWidth = frac.fixedDenominator
      .fold(visibleWidth(frac.denominator))(_.toString.length)
    val fractionPart =
      if num == 0 && mixed then " " * (visibleWidth(frac.numerator) + 1 + denWidth)
      else
        val numerator = if mixed then num else improperNumerator
        val denStr = frac.fixedDenominator match
          case Some(d) => d.toString
          case None => padPlaceholders(den.toString, frac.denominator, alignRight = false)
        padPlaceholders(numerator.toString, frac.numerator, alignRight = true) + "/" + denStr

    val wholeStr =
      if !mixed then ""
      else if whole != 0 then padPlaceholders(whole.toString, wholePlaceholders, alignRight = true)
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
   * arithmetic until the denominator budget is exceeded, then the previous convergent wins.
   * Double state is deliberate — Excel stores values as doubles and the binary noise is
   * observable in the chosen convergent (12.3 → 12 1/3 but 0.3 → 2/7).
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
   * Align digits within a placeholder run: unfilled `?` positions become spaces, `0` becomes
   * zeros, `#` adds nothing. Numerators/wholes right-align (pad left), denominators left-align
   * (pad right).
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
   */
  def applyTextFormat(text: String, format: FormatCode): String =
    format.text match
      case Some(section) =>
        section.pattern.tokens.map {
          case FormatToken.TextPlaceholder => text
          case FormatToken.Literal(s) => s
          case _ => ""
        }.mkString
      case None => text

  // ========== Date/Time Formatter ==========

  /**
   * Check if a format code contains date/time tokens.
   */
  def hasDateTokens(format: FormatCode): Boolean =
    format.positive.pattern.tokens.exists {
      case FormatToken.DatePart(_) | FormatToken.AmPm(_) | FormatToken.Elapsed(_) => true
      case _ => false
    }

  /**
   * Apply a parsed format code to a LocalDateTime value.
   *
   * @param dt
   *   The datetime to format
   * @param format
   *   The parsed format code
   * @return
   *   Formatted date/time string
   */
  def applyDateFormat(dt: LocalDateTime, format: FormatCode): String =
    val tokens = format.positive.pattern.tokens
    val minutePositions = findMinutePositions(tokens)
    tokens.zipWithIndex.map { case (token, idx) =>
      renderDateToken(dt, token, minutePositions.contains(idx))
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

  /**
   * Render a single date/time token.
   *
   * @param dt
   *   The datetime value
   * @param token
   *   The format token
   * @param isMinute
   *   Whether 'm'/'mm' should render as minute (vs month)
   */
  private def renderDateToken(dt: LocalDateTime, token: FormatToken, isMinute: Boolean): String =
    token match
      // Year
      case FormatToken.DatePart("y") =>
        (dt.getYear % 100).toString
      case FormatToken.DatePart("yy") =>
        f"${dt.getYear % 100}%02d"
      case FormatToken.DatePart("yyy" | "yyyy") =>
        dt.getYear.toString

      // Month (or minute if isMinute)
      case FormatToken.DatePart("m") =>
        if isMinute then dt.getMinute.toString
        else dt.getMonthValue.toString
      case FormatToken.DatePart("mm") =>
        if isMinute then f"${dt.getMinute}%02d"
        else f"${dt.getMonthValue}%02d"
      case FormatToken.DatePart("mmm") =>
        dt.getMonth.getDisplayName(TextStyle.SHORT, Locale.US)
      case FormatToken.DatePart("mmmm") =>
        dt.getMonth.getDisplayName(TextStyle.FULL, Locale.US)
      case FormatToken.DatePart("mmmmm") =>
        // First letter only (J, F, M, A, ...)
        dt.getMonth.getDisplayName(TextStyle.NARROW, Locale.US)

      // Day
      case FormatToken.DatePart("d") =>
        dt.getDayOfMonth.toString
      case FormatToken.DatePart("dd") =>
        f"${dt.getDayOfMonth}%02d"
      case FormatToken.DatePart("ddd") =>
        dt.getDayOfWeek.getDisplayName(TextStyle.SHORT, Locale.US)
      case FormatToken.DatePart("dddd") =>
        dt.getDayOfWeek.getDisplayName(TextStyle.FULL, Locale.US)

      // Hour (12-hour format assumed when AM/PM present)
      case FormatToken.DatePart("h") =>
        val hour = dt.getHour % 12
        (if hour == 0 then 12 else hour).toString
      case FormatToken.DatePart("hh") =>
        val hour = dt.getHour % 12
        f"${if hour == 0 then 12 else hour}%02d"

      // Second
      case FormatToken.DatePart("s") =>
        dt.getSecond.toString
      case FormatToken.DatePart("ss") =>
        f"${dt.getSecond}%02d"

      // AM/PM
      case FormatToken.AmPm("AM/PM") =>
        if dt.getHour < 12 then "AM" else "PM"
      case FormatToken.AmPm("A/P") =>
        if dt.getHour < 12 then "A" else "P"
      case FormatToken.AmPm(_) =>
        if dt.getHour < 12 then "AM" else "PM"

      // Elapsed time (for duration formatting - just show value for now)
      case FormatToken.Elapsed('h') =>
        dt.getHour.toString
      case FormatToken.Elapsed('m') =>
        dt.getMinute.toString
      case FormatToken.Elapsed('s') =>
        dt.getSecond.toString

      // Literals and other tokens
      case FormatToken.Literal(text) =>
        text
      case FormatToken.Spacer(_) =>
        " "
      case FormatToken.Thousands =>
        // In date context, comma is a literal (not thousands separator)
        ","
      case _ =>
        ""
