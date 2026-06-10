package com.tjclp.xl.cells

import java.time.LocalDateTime
import com.tjclp.xl.richtext.{RichText => Rt}

/** Cell value types supported by Excel */
enum CellValue:
  /** Text/string value */
  case Text(value: String)

  /** Rich text with multiple formatting runs (inline string with formatting) */
  case RichText(value: Rt)

  /** Numeric value (Excel uses double internally, but we preserve precision) */
  case Number(value: BigDecimal)

  /** Boolean value */
  case Bool(value: Boolean)

  /** Date/time value */
  case DateTime(value: LocalDateTime)

  /**
   * Formula expression with optional cached result value.
   *
   * @param expression
   *   The formula string (e.g., "=A1+B1"). Must be non-empty.
   * @param cachedValue
   *   Optional cached result value from Excel (preserved during roundtrip). This is the last
   *   calculated value stored in the XLSX file. Can be Number, Text, Bool, Error, or Empty. Must
   *   never be another Formula.
   * @note
   *   Use `CellValue.formula()` for validated construction.
   */
  case Formula(expression: String, cachedValue: Option[CellValue] = None)

  /** Empty cell */
  case Empty

  /** Error value */
  case Error(error: CellError)

object CellValue:
  /** Smart constructor from Any */
  def from(value: Any): CellValue = value match
    case cv: CellValue => cv // Already a CellValue, return as-is
    case s: String => Text(s)
    case i: Int => Number(BigDecimal(i))
    case l: Long => Number(BigDecimal(l))
    case d: Double => Number(BigDecimal(d))
    case bd: BigDecimal => Number(bd)
    case b: Boolean => Bool(b)
    case dt: LocalDateTime => DateTime(dt)
    case _ => Text(value.toString)

  /**
   * Validated constructor for Formula values.
   *
   * @param expression
   *   The formula string (e.g., "=A1+B1"). Must be non-empty.
   * @param cachedValue
   *   Optional cached result value. Must not be a Formula.
   * @throws IllegalArgumentException
   *   if expression is empty or cachedValue contains a Formula
   */
  def formula(expression: String, cachedValue: Option[CellValue] = None): Formula =
    require(expression.nonEmpty, "Formula expression cannot be empty")
    require(
      !cachedValue.exists { case _: Formula => true; case _ => false },
      "Cached value cannot be a Formula"
    )
    Formula(expression, cachedValue)

  // Excel epoch for the 1900 date system: December 30, 1899 (not Jan 1, 1900, to account for
  // Excel's 1900 leap-year bug). Epoch for the 1904 date system (legacy Mac Excel): January 1,
  // 1904 = serial 0; that system has NO phantom leap day.
  private val epoch1900 = LocalDateTime.of(1899, 12, 30, 0, 0, 0)
  private val epoch1904 = LocalDateTime.of(1904, 1, 1, 0, 0, 0)

  /**
   * Convert LocalDateTime to Excel serial number in the default 1900 date system.
   *
   * Excel represents dates as the number of days since December 30, 1899, with fractional days
   * representing time. Excel has a bug where it treats 1900 as a leap year (it wasn't); this
   * implementation reproduces that quirk so serials match Excel exactly: 1900-01-01 = 1, 1900-02-28 =
   * 59, the phantom 1900-02-29 = 60, 1900-03-01 = 61.
   *
   * @param dt
   *   The LocalDateTime to convert
   * @return
   *   Excel serial number (days since 1899-12-30 + fractional time)
   */
  def dateTimeToExcelSerial(dt: LocalDateTime): Double =
    dateTimeToExcelSerial(dt, date1904 = false)

  /**
   * Convert LocalDateTime to Excel serial number in the given date system (GH-243).
   *
   * In the 1904 date system (`<workbookPr date1904="1"/>`, legacy Mac Excel) serials count days
   * since 1904-01-01 (= serial 0). The 1900 system's phantom 1900-02-29 does not exist in the 1904
   * system, so no leap-day adjustment is applied. For the same date on/after 1900-03-01 the two
   * systems differ by exactly 1462 days.
   *
   * @param dt
   *   The LocalDateTime to convert
   * @param date1904
   *   true for the 1904 date system, false for the default 1900 system
   * @return
   *   Excel serial number (days since the system epoch + fractional time)
   */
  def dateTimeToExcelSerial(dt: LocalDateTime, date1904: Boolean): Double =
    import java.time.temporal.ChronoUnit

    val days =
      if date1904 then ChronoUnit.DAYS.between(epoch1904, dt)
      else
        val rawDays = ChronoUnit.DAYS.between(epoch1900, dt)
        // Excel 1900 leap-year bug: it counts a phantom 1900-02-29 (serial 60), so real dates before
        // 1900-03-01 (rawDays < 61 from the 1899-12-30 epoch) are one higher than Excel. Subtract it
        // back so 1900-01-01 -> 1 and 1900-02-28 -> 59. Dates on/after 1900-03-01 are already correct.
        if rawDays < 61 then rawDays - 1 else rawDays

    // Calculate fractional day for time component
    val dayStart = dt.toLocalDate.atStartOfDay
    val secondsInDay = ChronoUnit.SECONDS.between(dayStart, dt)
    val fractionOfDay = secondsInDay.toDouble / 86400.0

    days.toDouble + fractionOfDay

  /**
   * Convert Excel serial number to LocalDateTime in the default 1900 date system.
   *
   * Excel represents dates as the number of days since December 30, 1899, with fractional days
   * representing time. This is the inverse of dateTimeToExcelSerial.
   *
   * @param serial
   *   Excel serial number (days since 1899-12-30 + fractional time)
   * @return
   *   LocalDateTime corresponding to the serial number
   */
  def excelSerialToDateTime(serial: Double): LocalDateTime =
    excelSerialToDateTime(serial, date1904 = false)

  /**
   * Convert Excel serial number to LocalDateTime in the given date system (GH-243).
   *
   * In the 1904 date system serial 0 = 1904-01-01 and no phantom-leap-day inverse adjustment is
   * applied. This is the inverse of `dateTimeToExcelSerial(dt, date1904)`.
   *
   * @param serial
   *   Excel serial number (days since the system epoch + fractional time)
   * @param date1904
   *   true for the 1904 date system, false for the default 1900 system
   * @return
   *   LocalDateTime corresponding to the serial number
   */
  def excelSerialToDateTime(serial: Double, date1904: Boolean): LocalDateTime =
    // Extract whole days and fractional day
    val wholeDays = serial.toLong
    val fractionOfDay = serial - wholeDays
    val seconds = (fractionOfDay * 86400.0).toLong

    if date1904 then epoch1904.plusDays(wholeDays).plusSeconds(seconds)
    else
      // Inverse of the 1900 leap-year adjustment: serials below 60 are shifted one day forward;
      // serial 60 is Excel's phantom 1900-02-29, mapped here to 1900-02-28. Serials >= 61 are exact.
      val dayShift = if wholeDays < 60 then 1L else 0L
      epoch1900.plusDays(wholeDays + dayShift).plusSeconds(seconds)

  /**
   * Characters that trigger formula injection when at the start of a cell value.
   *
   * When a cell value starts with these characters, Excel (and LibreOffice, Google Sheets) may
   * interpret it as a formula or command, potentially leading to:
   *   - `=` - Formula evaluation
   *   - `+` - Formula evaluation (alternate prefix)
   *   - `-` - Formula evaluation (alternate prefix)
   *   - `@` - External data linking / potential command injection
   *
   * When writing untrusted data to Excel, use `escape()` to prefix with a single quote.
   */
  private val formulaInjectionChars: Set[Char] = Set('=', '+', '-', '@')

  /**
   * Check if a string could be interpreted as a formula by Excel.
   *
   * @param text
   *   The text to check
   * @return
   *   true if the text starts with a formula injection character
   */
  def couldBeFormula(text: String): Boolean =
    text.nonEmpty && formulaInjectionChars.contains(text.charAt(0))

  /**
   * Escape text to prevent formula injection.
   *
   * When writing untrusted data to Excel files, text starting with `=`, `+`, `-`, or `@` could be
   * interpreted as formulas. This function prefixes such text with a single quote (`'`), which
   * tells Excel to treat the value as literal text.
   *
   * This is idempotent: already-escaped text (starting with `'`) is returned unchanged.
   *
   * @param text
   *   The text to escape
   * @return
   *   The escaped text (prefixed with `'` if needed)
   *
   * @example
   *   {{{
   * CellValue.escape("=SUM(A1)")    // => "'=SUM(A1)"
   * CellValue.escape("+1234")       // => "'+1234"
   * CellValue.escape("Hello")       // => "Hello" (no change)
   * CellValue.escape("'=already")   // => "'=already" (already escaped)
   *   }}}
   */
  def escape(text: String): String =
    if text.isEmpty then text
    else if text.charAt(0) == '\'' then text // Already escaped
    else if couldBeFormula(text) then "'" + text
    else text

  /**
   * Unescape text that was escaped for formula injection.
   *
   * Removes the leading single quote if it was added by `escape()`. This is useful when reading
   * back data that was escaped during writing.
   *
   * @param text
   *   The text to unescape
   * @return
   *   The unescaped text
   *
   * @example
   *   {{{
   * CellValue.unescape("'=SUM(A1)")  // => "=SUM(A1)"
   * CellValue.unescape("Hello")      // => "Hello" (no change)
   * CellValue.unescape("''quoted")   // => "'quoted" (single level unescape)
   *   }}}
   */
  def unescape(text: String): String =
    if text.length >= 2 && text.charAt(0) == '\'' && couldBeFormula(text.substring(1)) then
      text.substring(1)
    else text
