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
        // TODO: Implement fraction formatting (e.g., 0.5 → "1/2")
        formatGeneral(n)

      case NumFmt.Text =>
        // Text format displays numbers as text
        n.toString

      case NumFmt.Custom(code) =>
        FormatCodeParser.parse(code) match
          case Right(fmt) if FormatCodeParser.hasDateTokens(fmt) =>
            // Custom date format: convert serial number to LocalDateTime
            val dt = excelSerialToDateTime(n)
            FormatCodeParser.applyDateFormat(dt, fmt)
          case Right(fmt) =>
            FormatCodeParser.applyFormat(n, fmt)._1
          case Left(_) =>
            formatGeneral(n) // Fallback for unparseable formats

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
        FormatCodeParser.parse(code) match
          case Right(fmt) if FormatCodeParser.hasDateTokens(fmt) =>
            FormatCodeParser.applyDateFormat(dt, fmt)
          case _ => dt.toString // Fallback for non-date formats or parse errors

      case _ =>
        dt.toString

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
