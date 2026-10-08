package com.tjclp.xl.display

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.styles.numfmt.NumFmt

import java.time.LocalDateTime

/**
 * Formats cell values according to Excel number format codes.
 *
 * Implements Excel-accurate display formatting for all NumFmt types (Currency, Percent, Date,
 * etc.).
 *
 * @since 0.2.0
 */
object NumFmtFormatter:

  /**
   * Parsed format codes for every built-in (non-Custom, non-General) NumFmt variant.
   *
   * The built-in enum arms render through FormatCodeParser on [[NumFmt.formatCode]], so a
   * programmatic NumFmt.PercentDecimal and a file-declared "0.00%" are provably the same display
   * (GH-410). Every code in the canonical table parses, making the map total over the variants
   * listed; the lookup fallbacks below are defensive only.
   */
  private val builtInFormats: Map[NumFmt, FormatCodeParser.FormatCode] =
    List(
      NumFmt.Integer,
      NumFmt.Decimal,
      NumFmt.ThousandsSeparator,
      NumFmt.ThousandsDecimal,
      NumFmt.Currency,
      NumFmt.Percent,
      NumFmt.PercentDecimal,
      NumFmt.Scientific,
      NumFmt.Fraction,
      NumFmt.Date,
      NumFmt.DateTime,
      NumFmt.Time,
      NumFmt.Text
    ).flatMap(fmt => FormatCodeParser.parse(NumFmt.formatCode(fmt)).toOption.map(fmt -> _)).toMap

  /**
   * Which General rule a number takes where a format says General — the whole code, the `General`
   * keyword inside a section, a text-only code or an unparseable one (#672).
   *   - CellDisplay: what a cell shows on screen, Excel's 11-character rule ([[formatGeneral]])
   *   - Text: the width-independent text conversion ([[generalText]]); TEXT(x, fmt) renders through
   *     it, so `TEXT(x,"General;-General")` agrees with `TEXT(x,"General")`
   */
  enum GeneralRule derives CanEqual:
    case CellDisplay, Text

  private def general(n: BigDecimal, rule: GeneralRule): String =
    rule match
      case GeneralRule.CellDisplay => formatGeneral(n)
      case GeneralRule.Text => generalText(n)

  /**
   * The `General` keyword token inside a custom section under `rule`, on the magnitude `unscaled ×
   * 10^-scale` (unscaled ≥ 0; the section owns the sign). The scale is a Long, so a section's
   * percent and scaling commas never have to fit a BigDecimal scale (#689: `General%` on
   * `1E+2147483647`).
   */
  private[display] def generalKeyword(
    unscaled: java.math.BigInteger,
    scale: Long,
    rule: GeneralRule
  ): String =
    if unscaled.signum == 0 then "0"
    else
      rule match
        case GeneralRule.CellDisplay => generalDisplayUnsigned(unscaled, scale)
        case GeneralRule.Text => generalTextUnsigned(unscaled, scale)

  /**
   * Format a cell value according to its number format.
   *
   * @param value
   *   The cell value to format
   * @param numFmt
   *   The number format to apply
   * @return
   *   Formatted string matching Excel display conventions
   */
  def formatValue(value: CellValue, numFmt: NumFmt): String =
    formatValue(value, numFmt, GeneralRule.CellDisplay)

  /** [[formatValue]] with General rendered under `rule` (#672: TEXT passes GeneralRule.Text). */
  def formatValue(value: CellValue, numFmt: NumFmt, rule: GeneralRule): String =
    value match
      case CellValue.Number(n) => formatNumber(n, numFmt, rule)
      // Only a custom code can carry a text section; built-in codes echo the text (#693)
      case CellValue.Text(s) =>
        numFmt match
          case NumFmt.Custom(code) =>
            FormatCodeParser.parse(code).fold(_ => s, FormatCodeParser.applyTextFormat(s, _))
          case _ => s
      case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
      case CellValue.DateTime(dt) => formatDateTime(dt, numFmt, rule)
      case CellValue.Empty => ""
      case CellValue.Error(err) => formatError(err)
      case CellValue.Formula(expr, _, _) =>
        s"=$expr" // Fallback - should be handled by FormulaDisplayStrategy
      case CellValue.RichText(rt) => rt.toPlainText

  /**
   * Format a numeric value according to Excel number format.
   *
   * @param n
   *   The number to format
   * @param numFmt
   *   The format to apply
   * @return
   *   Formatted number string
   */
  def formatNumber(n: BigDecimal, numFmt: NumFmt): String =
    formatNumber(n, numFmt, GeneralRule.CellDisplay)

  /** [[formatNumber]] with General rendered under `rule` (#672). */
  def formatNumber(n: BigDecimal, numFmt: NumFmt, rule: GeneralRule): String =
    numFmt match
      case NumFmt.General => general(n, rule)

      case NumFmt.Custom(code) if isGeneralCode(code) =>
        // The literal code "General" (ECMA-376 §18.8.30, case-insensitive) means General
        // rendering, not the literal characters. Reached since GH-404: file-declared
        // <numFmt formatCode="General"/> entries (LibreOffice writes them) stay Custom.
        general(n, rule)

      case NumFmt.Custom(code) =>
        FormatCodeParser.parse(code) match
          case Right(fmt) => formatCustom(n, serialToDateTime(n), fmt, rule)
          case Left(_) => general(n, rule) // Fallback for unparseable formats

      case builtin =>
        // Built-in arms render through FormatCodeParser on their canonical format code, so
        // the programmatic enum and a file-declared equal code display identically (GH-410).
        builtInFormats.get(builtin) match
          case Some(fmt) => formatCustom(n, serialToDateTime(n), fmt, rule)
          case None => general(n, rule) // unreachable: the map covers every such variant

  /** ECMA-376 §18.8.30: the whole-code "General" keyword, matched case-insensitively. */
  def isGeneralCode(code: String): Boolean = code.equalsIgnoreCase("General")

  /** Significant digits Excel keeps when a number becomes text (GH-665). */
  val GeneralTextDigits: Int = 15

  /**
   * Longest unsigned plain (non-E) rendering Excel's text conversion produces (GH-665): up to 20
   * integer digits, or `0.` plus leading zeros plus significant digits within 20 characters;
   * anything longer switches to E notation. The sign is never counted.
   */
  val GeneralTextMaxLength: Int = 20

  /**
   * Excel's width-independent number → text conversion (GH-665): the rule behind `&`, CONCATENATE,
   * text-typed function arguments, TEXT(x,"General") and the formula bar. NOT the column-width
   * cell-display General ([[formatGeneral]], what a cell shows on screen).
   *
   * The rule (Excel-verified via Apache POI's NumberToTextConverter):
   *   - zero of any scale or sign is "0"
   *   - round to [[GeneralTextDigits]] significant digits (HALF_UP), strip trailing zeros
   *   - |n| ≥ 1: plain while the integer part has at most [[GeneralTextMaxLength]] digits
   *     (`1E19` → `10000000000000000000`, `1E20` → `1E+20`)
   *   - |n| < 1: plain while `0.` + leading zeros + significant digits fits in
   *     [[GeneralTextMaxLength]] characters (`0.000123456789012346` stays plain at exactly 20;
   *     `0.000012345678901234568` → `1.23456789012346E-05`; `1E-16` → `0.0000000000000001`)
   *   - E form: mantissa with the decimal point only when more than one digit remains, exponent
   *     signed and zero-padded to at least two digits (`1E+20`, `9.5367431640625E-07`, `1E-100`)
   *
   * Output grammar: `-?[0-9]+(\.[0-9]+)?(E[+-][0-9]{2,})?`, bounded length (`1E+1000000` renders as
   * itself, never as a million zeros). Pure BigDecimal/BigInteger arithmetic with Long exponents:
   * no Double, no Locale, and no stored exponent throws. Note that LibreOffice's `&` conversion
   * diverges (three-digit exponents, E notation from about 1E16), so it is an oracle only for the
   * plain range.
   */
  def generalText(n: BigDecimal): String =
    if n.signum == 0 then "0"
    else
      val body = generalTextUnsigned(n.bigDecimal.unscaledValue.abs, n.scale.toLong)
      if n.signum < 0 then "-" + body else body

  /** [[generalText]] of the positive magnitude `unscaled × 10^-scale`. */
  private def generalTextUnsigned(unscaled: java.math.BigInteger, scale: Long): String =
    // Rounded as a BigInteger, not through BigDecimal.round: dropping digits lowers the scale,
    // which threw past Int.MinValue for a 16-digit value near `1E+2147483648` (#689)
    val (digits, roundedScale) = roundSignificant(unscaled, scale, GeneralTextDigits)
    val sig = digits.length
    // adjusted exponent: 1234.5 → 3, 0.0012 → -3. In Long: a scale near Int.MaxValue
    // (`1E-2147483647`) overflowed the Int length sum below to a negative, chose the plain form
    // and asked toPlainString for two billion zeros (PR #679 review) — every step is Long and
    // toPlainString is reached only once the length is proven to fit.
    val exp: Long = sig.toLong - roundedScale - 1L
    val plain =
      if exp >= 0 then exp <= GeneralTextMaxLength - 1L
      else 2L + (-exp - 1L) + sig.toLong <= GeneralTextMaxLength.toLong
    if plain then
      // the plain form bounds the scale to [-19, 18]
      new java.math.BigDecimal(new java.math.BigInteger(digits), roundedScale.toInt).toPlainString
    else
      val mantissa =
        if sig == 1 then digits else digits.substring(0, 1) + "." + digits.substring(1)
      val absExp = math.abs(exp)
      val expDigits = if absExp < 10L then s"0$absExp" else absExp.toString
      val expSign = if exp < 0 then "-" else "+"
      s"${mantissa}E$expSign$expDigits"

  /**
   * `unscaled × 10^-scale` (positive) rounded HALF_UP to `significant` digits, trailing zeros
   * stripped: the digit string and its scale. The divisor is a power of ten no longer than the
   * value's own digits, and the scale is a Long, so no stored exponent overflows it (#689).
   */
  private def roundSignificant(
    unscaled: java.math.BigInteger,
    scale: Long,
    significant: Int
  ): (String, Long) =
    val precision = digitCount(unscaled)
    val (rounded, roundedScale) =
      if precision <= significant then (unscaled, scale)
      else
        val drop = precision - significant
        val divisor = java.math.BigInteger.TEN.pow(drop)
        (unscaled.add(divisor.shiftRight(1)).divide(divisor), scale - drop)
    // a carry (999… → 1000…) only adds trailing zeros, which the strip removes
    val text = rounded.toString
    val stripped = text.reverse.dropWhile(_ == '0').reverse
    (stripped, roundedScale - (text.length - stripped.length))

  /** Decimal digits of a non-negative BigInteger (1 for zero). */
  private[display] def digitCount(n: java.math.BigInteger): Int =
    if n.signum == 0 then 1 else new java.math.BigDecimal(n).precision

  /**
   * Whole-value form of [[generalText]] (GH-665): Number → the rule; DateTime → its Excel serial
   * through the rule (GH-561: dates are numbers in text positions); Bool → TRUE/FALSE; Text →
   * itself; RichText → plain text; Empty → ""; Error → its Excel code; Formula → its cached value
   * through this same table, "" when uncached.
   */
  def generalText(value: CellValue): String =
    value match
      case CellValue.Number(n) => generalText(n)
      case CellValue.DateTime(dt) => generalText(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
      case CellValue.Text(s) => s
      case CellValue.RichText(rt) => rt.toPlainText
      case CellValue.Empty => ""
      case CellValue.Error(err) => formatError(err)
      case CellValue.Formula(_, Some(cached), _) => generalText(cached)
      case CellValue.Formula(_, None, _) => ""

  /** Characters a General cell shows, the sign uncounted (GH-672). */
  val GeneralDisplayWidth: Int = 11

  /**
   * Format in General style for CELL DISPLAY: what a General cell shows on screen, not what a
   * number becomes as text — text conversion (`&`, CONCATENATE, TEXT(x,"General")) is
   * [[generalText]].
   *
   * Excel's 11-character rule (GH-672), the sign uncounted:
   *   - plain while the adjusted exponent is in [-4, 10]: rounded HALF_UP to the decimals left
   *     after the integer digits and the point (`12345678.901` → `12345678.9`, `1/3` →
   *     `0.333333333`), trailing zeros stripped
   *   - E notation otherwise (`123456789012` → `1.23457E+11`, `0.00001` → `1E-05`), and when the
   *     plain rounding carries past 11 digits (`99999999999.5` → `1E+11`): as many significant
   *     digits as fit in 11 characters (6 with a two-digit exponent, 5 with three), trailing zeros
   *     and a bare point dropped, exponent signed and at least two digits
   *   - like %G, the switch reads the exponent after rounding: `0.0000999999999999` is `0.0001`
   *
   * Pure BigDecimal/BigInteger arithmetic with Long exponents: no Double, no Locale, and bounded
   * output for any stored exponent (`1E+2147483647` renders as itself). The column-width shrinking
   * Excel applies to narrow columns is not modelled.
   */
  private def formatGeneral(n: BigDecimal): String =
    if n.signum == 0 then "0"
    else
      val body = generalDisplayUnsigned(n.bigDecimal.unscaledValue.abs, n.scale.toLong)
      if n.signum < 0 then "-" + body else body

  /** [[formatGeneral]] of the positive magnitude `digits × 10^-scale`. */
  private def generalDisplayUnsigned(digits: java.math.BigInteger, scale: Long): String =
    val precision = digitCount(digits)
    // adjusted exponent in Long: a scale near either end of the Int range overflows Int
    val exp: Long = precision.toLong - scale - 1L
    val plain =
      // exp in [-4, 10] puts the scale within the value's own digit count (+3), an Int
      if exp >= -4L && exp <= 10L && scale.isValidInt then
        val decimals = if exp >= 0L then math.max(0, 9 - exp.toInt) else GeneralDisplayWidth - 2
        val text = new java.math.BigDecimal(digits, scale.toInt)
          .setScale(decimals, java.math.RoundingMode.HALF_UP)
          .stripTrailingZeros
          .toPlainString
        Option.when(text.length <= GeneralDisplayWidth)(text)
      else None
    plain.getOrElse(generalDisplayScientific(digits, precision, exp))

  /**
   * The E form of [[formatGeneral]]. Rounds the unscaled digits as a BigInteger rather than the
   * BigDecimal, whose scale would overflow Int at the extremes.
   */
  private def generalDisplayScientific(
    digits: java.math.BigInteger,
    precision: Int,
    exp: Long
  ): String =
    def exponentDigits(e: Long): Int = math.max(2, math.abs(e).toString.length)
    // mantissa `d.dddd` + `E±` + exponent within the width; one digit (no point) at the least
    def significantFor(e: Long): Int = math.max(1, GeneralDisplayWidth - 3 - exponentDigits(e))
    def roundTo(sig: Int): (java.math.BigInteger, Long) =
      if precision <= sig then (digits, exp)
      else
        val divisor = java.math.BigInteger.TEN.pow(precision - sig)
        val q = digits.add(divisor.shiftRight(1)).divide(divisor)
        if q.toString.length > sig then (q.divide(java.math.BigInteger.TEN), exp + 1L)
        else (q, exp)
    val first = roundTo(significantFor(exp))
    // a carry into a longer exponent (9.999999E+99 → 1E+100) leaves room for one digit fewer
    val (mantissa, e) =
      if significantFor(first._2) < significantFor(exp) then roundTo(significantFor(first._2))
      else first
    val sig = mantissa.toString.reverse.dropWhile(_ == '0').reverse
    if e >= -4L && e <= 10L then
      // a carry back into the plain range (0.0000999999999999 → 0.0001); e is small here
      new java.math.BigDecimal(
        new java.math.BigInteger(sig),
        sig.length - 1 - e.toInt
      ).toPlainString
    else
      val mantissaText =
        if sig.length == 1 then sig else s"${sig.substring(0, 1)}.${sig.substring(1)}"
      val absExp = math.abs(e)
      val expText = if absExp < 10L then s"0$absExp" else absExp.toString
      s"${mantissaText}E${if e < 0L then "-" else "+"}$expText"

  /**
   * Format a date/time value.
   *
   * @param dt
   *   The LocalDateTime to format
   * @param numFmt
   *   The format to apply
   * @return
   *   Formatted date/time string
   */
  def formatDateTime(dt: LocalDateTime, numFmt: NumFmt): String =
    formatDateTime(dt, numFmt, GeneralRule.CellDisplay)

  /** [[formatDateTime]] with General rendered under `rule` (#672). */
  def formatDateTime(dt: LocalDateTime, numFmt: NumFmt, rule: GeneralRule): String =
    numFmt match
      case NumFmt.Date | NumFmt.DateTime | NumFmt.Time =>
        // The calendar variants render straight off the LocalDateTime through their parsed
        // canonical codes (GH-410) — no serial round-trip, which truncates seconds. Bare 'h'
        // is the 24-hour clock (no AM/PM in these codes, ECMA-376 §18.8.31).
        // A date outside Excel's displayable range fills with `#`, as its serial does (#688).
        builtInFormats.get(numFmt) match
          case Some(_) if !isDisplayableSerial(dateTimeSerial(dt)) => "######"
          case Some(fmt) => FormatCodeParser.applyDateFormat(dt, fmt)
          case None => dt.toString // unreachable: the map covers the three variants

      case NumFmt.Custom(code) if isGeneralCode(code) =>
        // "General" keyword code: dates ARE numbers in Excel, so render the serial (GH-283)
        formatNumber(dateTimeSerial(dt), NumFmt.General, rule)

      case NumFmt.Custom(code) =>
        // Route through section selection on the serial (GH-283): ';;;' hides dates,
        // numeric sections render the serial, conditional codes pick sections by serial
        FormatCodeParser.parse(code) match
          case Right(fmt) =>
            val serial = dateTimeSerial(dt)
            val calendar =
              Option.when(isDisplayableSerial(serial))(FormatCodeParser.ExcelCalendar.of(dt))
            formatCustom(serial, calendar, fmt, rule)
          case Left(_) => dt.toString // Fallback for parse errors

      case other =>
        // Dates ARE numbers in Excel: any numeric format (General included) displays
        // the underlying serial number, never ISO text (GH-283)
        formatNumber(dateTimeSerial(dt), other, rule)

  /**
   * Render a numeric value through a parsed custom code with full section routing (GH-283/285): the
   * section chosen for the value decides between calendar rendering (date tokens), General
   * (text-only codes like a lone `@`) and numeric pattern rendering.
   *
   * @param n
   *   The numeric value (a date serial when the value is date-typed)
   * @param dt
   *   The calendar view of `n` for date-token sections; None marks a serial outside Excel's
   *   displayable date range, rendered as `######` like Excel's unrepresentable-date fill
   * @param fmt
   *   The parsed format code
   */
  private def formatCustom(
    n: BigDecimal,
    dt: => Option[FormatCodeParser.ExcelCalendar],
    fmt: FormatCodeParser.FormatCode,
    rule: GeneralRule
  ): String =
    FormatCodeParser.selectSection(n, fmt) match
      case None => general(n, rule)
      case Some(section) if FormatCodeParser.hasDateTokens(section) =>
        dt match
          case Some(d) => FormatCodeParser.applyDateFormat(d, section)
          case None => "######"
      case Some(_) => FormatCodeParser.applyFormat(n, fmt, rule)._1

  /** Exclusive upper bound of Excel's displayable date serials (9999-12-31 is 2958465). */
  private val maxDateSerialExclusive = BigDecimal(2958466)

  /**
   * Calendar view of a date serial, or None when the serial lies outside Excel's displayable range
   * (negative or on/after 10000-01-01) — Excel fills such cells with `#` (GH-283).
   */
  private def serialToDateTime(serial: BigDecimal): Option[FormatCodeParser.ExcelCalendar] =
    Option.when(isDisplayableSerial(serial))(FormatCodeParser.ExcelCalendar.fromSerial(serial))

  /** Whether a date serial lies in Excel's displayable range, [0, 10000-01-01). */
  private def isDisplayableSerial(serial: BigDecimal): Boolean =
    serial >= 0 && serial < maxDateSerialExclusive

  /** Excel serial number (days since 1899-12-30 + day fraction) of a LocalDateTime. */
  private def dateTimeSerial(dt: LocalDateTime): BigDecimal =
    BigDecimal(CellValue.dateTimeToExcelSerial(dt))

  /**
   * Format error values in Excel style.
   *
   * @param err
   *   The error value
   * @return
   *   Formatted error string (e.g., "#DIV/0!")
   */
  private def formatError(err: com.tjclp.xl.cells.CellError): String =
    import com.tjclp.xl.cells.CellError.*
    err.toExcel
