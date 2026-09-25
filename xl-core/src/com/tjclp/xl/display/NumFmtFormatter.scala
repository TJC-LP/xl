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
    value match
      case CellValue.Number(n) => formatNumber(n, numFmt)
      case CellValue.Text(s) => s
      case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
      case CellValue.DateTime(dt) => formatDateTime(dt, numFmt)
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
    numFmt match
      case NumFmt.General => formatGeneral(n)

      case NumFmt.Custom(code) if isGeneralCode(code) =>
        // The literal code "General" (ECMA-376 §18.8.30, case-insensitive) means General
        // rendering, not the literal characters. Reached since GH-404: file-declared
        // <numFmt formatCode="General"/> entries (LibreOffice writes them) stay Custom.
        formatGeneral(n)

      case NumFmt.Custom(code) =>
        FormatCodeParser.parse(code) match
          case Right(fmt) => formatCustom(n, serialToDateTime(n), fmt)
          case Left(_) => formatGeneral(n) // Fallback for unparseable formats

      case builtin =>
        // Built-in arms render through FormatCodeParser on their canonical format code, so
        // the programmatic enum and a file-declared equal code display identically (GH-410).
        builtInFormats.get(builtin) match
          case Some(fmt) => formatCustom(n, serialToDateTime(n), fmt)
          case None => formatGeneral(n) // unreachable: the map covers every such variant

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
   * itself, never as a million zeros). Pure BigDecimal arithmetic: no Double, no Locale. Note that
   * LibreOffice's `&` conversion diverges (three-digit exponents, E notation from about 1E16), so
   * it is an oracle only for the plain range.
   */
  def generalText(n: BigDecimal): String =
    if n.signum == 0 then "0"
    else
      val rounded = n.bigDecimal
        .round(new java.math.MathContext(GeneralTextDigits, java.math.RoundingMode.HALF_UP))
        .stripTrailingZeros
      val a = rounded.abs
      val digits = a.unscaledValue.toString
      val sig = digits.length
      // adjusted exponent: 1234.5 → 3, 0.0012 → -3. In Long: a scale near Int.MaxValue
      // (`1E-2147483647`) overflowed the Int length sum below to a negative, chose the plain form
      // and asked toPlainString for two billion zeros (PR #679 review) — every step is Long and
      // toPlainString is reached only once the length is proven to fit.
      val exp: Long = a.precision.toLong - a.scale.toLong - 1L
      val plain =
        if exp >= 0 then exp <= GeneralTextMaxLength - 1L
        else 2L + (-exp - 1L) + sig.toLong <= GeneralTextMaxLength.toLong
      val body =
        if plain then a.toPlainString
        else
          val mantissa =
            if sig == 1 then digits else digits.substring(0, 1) + "." + digits.substring(1)
          val absExp = math.abs(exp)
          val expDigits = if absExp < 10L then s"0$absExp" else absExp.toString
          val expSign = if exp < 0 then "-" else "+"
          s"${mantissa}E$expSign$expDigits"
      if rounded.signum < 0 then "-" + body else body

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

  /**
   * The `General` keyword token inside a custom section (GH-666) renders through the CELL DISPLAY
   * General ([[formatGeneral]]), so `General"A"` and the whole-code path share one definition. Not
   * [[generalText]], which is the text-conversion rule (`&`, CONCATENATE).
   */
  private[display] def generalDisplay(n: BigDecimal): String = formatGeneral(n)

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
      val body = generalDisplayUnsigned(n.bigDecimal.abs)
      if n.signum < 0 then "-" + body else body

  private def generalDisplayUnsigned(a: java.math.BigDecimal): String =
    // adjusted exponent in Long: a scale near either end of the Int range overflows Int
    val exp: Long = a.precision.toLong - a.scale.toLong - 1L
    val plain =
      if exp >= -4L && exp <= 10L then
        val decimals = if exp >= 0L then math.max(0, 9 - exp.toInt) else GeneralDisplayWidth - 2
        val text =
          a.setScale(decimals, java.math.RoundingMode.HALF_UP).stripTrailingZeros.toPlainString
        Option.when(text.length <= GeneralDisplayWidth)(text)
      else None
    plain.getOrElse(generalDisplayScientific(a, exp))

  /**
   * The E form of [[formatGeneral]]. Rounds the unscaled digits as a BigInteger rather than the
   * BigDecimal, whose scale would overflow Int at the extremes.
   */
  private def generalDisplayScientific(a: java.math.BigDecimal, exp: Long): String =
    def exponentDigits(e: Long): Int = math.max(2, math.abs(e).toString.length)
    // mantissa `d.dddd` + `E±` + exponent within the width; one digit (no point) at the least
    def significantFor(e: Long): Int = math.max(1, GeneralDisplayWidth - 3 - exponentDigits(e))
    val digits = a.unscaledValue
    val precision = a.precision
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
    numFmt match
      case NumFmt.Date | NumFmt.DateTime | NumFmt.Time =>
        // The calendar variants render straight off the LocalDateTime through their parsed
        // canonical codes (GH-410) — no serial round-trip, which truncates seconds. Bare 'h'
        // is the 24-hour clock (no AM/PM in these codes, ECMA-376 §18.8.31).
        builtInFormats.get(numFmt) match
          case Some(fmt) => FormatCodeParser.applyDateFormat(dt, fmt)
          case None => dt.toString // unreachable: the map covers the three variants

      case NumFmt.Custom(code) if isGeneralCode(code) =>
        // "General" keyword code: dates ARE numbers in Excel, so render the serial (GH-283)
        formatNumber(dateTimeSerial(dt), NumFmt.General)

      case NumFmt.Custom(code) =>
        // Route through section selection on the serial (GH-283): ';;;' hides dates,
        // numeric sections render the serial, conditional codes pick sections by serial
        FormatCodeParser.parse(code) match
          case Right(fmt) => formatCustom(dateTimeSerial(dt), Some(dt), fmt)
          case Left(_) => dt.toString // Fallback for parse errors

      case other =>
        // Dates ARE numbers in Excel: any numeric format (General included) displays
        // the underlying serial number, never ISO text (GH-283)
        formatNumber(dateTimeSerial(dt), other)

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
    dt: => Option[LocalDateTime],
    fmt: FormatCodeParser.FormatCode
  ): String =
    FormatCodeParser.selectSection(n, fmt) match
      case None => formatGeneral(n)
      case Some(section) if FormatCodeParser.hasDateTokens(section) =>
        dt match
          case Some(d) => FormatCodeParser.applyDateFormat(d, section)
          case None => "######"
      case Some(_) => FormatCodeParser.applyFormat(n, fmt)._1

  /** Exclusive upper bound of Excel's displayable date serials (9999-12-31 is 2958465). */
  private val maxDateSerialExclusive = BigDecimal(2958466)

  /**
   * Calendar view of a date serial, or None when the serial lies outside Excel's displayable range
   * (negative or on/after 10000-01-01) — Excel fills such cells with `#` (GH-283).
   */
  private def serialToDateTime(serial: BigDecimal): Option[LocalDateTime] =
    if serial < 0 || serial >= maxDateSerialExclusive then None
    else Some(excelSerialToDateTime(serial))

  /** Excel serial number (days since 1899-12-30 + day fraction) of a LocalDateTime. */
  private def dateTimeSerial(dt: LocalDateTime): BigDecimal =
    BigDecimal(CellValue.dateTimeToExcelSerial(dt))

  /**
   * Convert Excel date serial number to LocalDateTime.
   *
   * @param serial
   *   Excel date serial number (days since 1899-12-30)
   * @return
   *   LocalDateTime
   */
  private def excelSerialToDateTime(serial: BigDecimal): LocalDateTime =
    import java.time.LocalDate
    // Excel serial date: 1 = 1900-01-01 (with 1900 leap year bug)
    val baseDate = LocalDate.of(1899, 12, 30) // Adjusted for Excel's bug
    val days = serial.toLong
    val date = baseDate.plusDays(days)

    // Handle time component if present
    val timeFraction = (serial % 1).toDouble
    if timeFraction > 0 then
      val hours = (timeFraction * 24).toInt
      val minutes = ((timeFraction * 24 * 60) % 60).toInt
      val seconds = (((timeFraction * 24 * 60 * 60) % 60)).toInt
      date.atTime(hours, minutes, seconds)
    else date.atStartOfDay()

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
