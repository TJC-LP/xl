package com.tjclp.xl.display

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.styles.numfmt.NumFmt

import java.text.DecimalFormat
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

/**
 * Formats cell values according to Excel number format codes.
 *
 * Implements Excel-accurate display formatting for all NumFmt types (Currency, Percent, Date,
 * etc.).
 *
 * @since 0.2.0
 */
object NumFmtFormatter:

  /** Parsed pattern for the built-in NumFmt.Fraction (format id 12, "# ?/?"). */
  private val builtInFractionFormat: Either[String, FormatCodeParser.FormatCode] =
    FormatCodeParser.parse("# ?/?")

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
      case CellValue.Formula(expr, _) =>
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

      case NumFmt.Integer =>
        n.setScale(0, BigDecimal.RoundingMode.HALF_UP).toString

      case NumFmt.Decimal =>
        f"${n.toDouble}%.2f"

      case NumFmt.ThousandsSeparator =>
        val formatter = DecimalFormat("#,##0")
        formatter.format(n.toDouble)

      case NumFmt.ThousandsDecimal =>
        val formatter = DecimalFormat("#,##0.00")
        formatter.format(n.toDouble)

      case NumFmt.Currency =>
        val formatter = DecimalFormat("$#,##0.00")
        formatter.format(n.toDouble)

      case NumFmt.Percent =>
        // Excel stores 0.15 for 15%, multiply by 100 for display
        val pct = (n * 100).setScale(0, BigDecimal.RoundingMode.HALF_UP)
        s"$pct%"

      case NumFmt.PercentDecimal =>
        val pct = (n * 100).setScale(1, BigDecimal.RoundingMode.HALF_UP)
        s"$pct%"

      case NumFmt.Scientific =>
        f"${n.toDouble}%.2E"

      case NumFmt.Date =>
        // Interpret number as Excel date serial (days since 1900-01-01)
        formatExcelDateSerial(n, "M/d/yy")

      case NumFmt.DateTime =>
        formatExcelDateSerial(n, "M/d/yy h:mm")

      case NumFmt.Time =>
        // Interpret fractional part as time of day
        val hours = ((n % 1) * 24).toInt
        val minutes = (((n % 1) * 24 * 60) % 60).toInt
        val seconds = (((n % 1) * 24 * 60 * 60) % 60).toInt
        f"$hours%d:$minutes%02d:$seconds%02d"

      case NumFmt.Fraction =>
        // Built-in fraction format id 12: "# ?/?" (up to one denominator digit, GH-243)
        builtInFractionFormat match
          case Right(fmt) => FormatCodeParser.applyFormat(n, fmt)._1
          case Left(_) => formatGeneral(n)

      case NumFmt.Text =>
        // Text format displays numbers as text
        n.toString

      case NumFmt.Custom(code) =>
        FormatCodeParser.parse(code) match
          case Right(fmt) => formatCustom(n, serialToDateTime(n), fmt)
          case Left(_) => formatGeneral(n) // Fallback for unparseable formats

  /**
   * Format in General style (Excel's default number format).
   *
   * Rules:
   *   - Integers: No decimal point
   *   - Decimals: Up to 11 significant digits
   *   - Scientific: For very large/small numbers (>= 1e12 or < 1e-4)
   */
  private def formatGeneral(n: BigDecimal): String =
    if n.isWhole then n.toBigInt.toString
    else
      val plain = n.underlying.stripTrailingZeros.toPlainString
      val sigDigits = countSignificantDigits(plain)
      if sigDigits > 11 then
        val mc = new java.math.MathContext(11)
        val rounded = n.underlying.round(mc)
        val roundedPlain = rounded.stripTrailingZeros.toPlainString
        val abs = n.abs
        if abs >= BigDecimal("1E12") || abs < BigDecimal("1E-4") then f"${rounded.doubleValue}%.6E"
        else roundedPlain
      else plain

  private def countSignificantDigits(plain: String): Int =
    val s = if plain.startsWith("-") then plain.substring(1) else plain
    if s.contains('.') then
      val stripped = s.stripPrefix("0.").dropWhile(_ == '0')
      stripped.replace(".", "").length
    else
      val trimmed = s.reverse.dropWhile(_ == '0')
      if trimmed.isEmpty then 1 else trimmed.length

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
      case NumFmt.Date =>
        dt.toLocalDate.format(DateTimeFormatter.ofPattern("M/d/yy"))

      case NumFmt.DateTime =>
        dt.format(DateTimeFormatter.ofPattern("M/d/yy H:mm")) // 24-hour format

      case NumFmt.Time =>
        dt.toLocalTime.format(DateTimeFormatter.ofPattern("H:mm:ss")) // 24-hour format

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
   * Format Excel date serial number (days since 1900-01-01).
   *
   * @param serial
   *   Excel date serial number
   * @param pattern
   *   DateTimeFormatter pattern
   * @return
   *   Formatted date string
   */
  private def formatExcelDateSerial(serial: BigDecimal, pattern: String): String =
    excelSerialToDateTime(serial).format(DateTimeFormatter.ofPattern(pattern))

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
