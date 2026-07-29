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
  private def isGeneralCode(code: String): Boolean = code.equalsIgnoreCase("General")

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
