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

  /**
   * Convert LocalDateTime to Excel serial number.
   *
   * Excel represents dates as the number of days since December 30, 1899, with fractional days
   * representing time. Note: Excel has a bug where it treats 1900 as a leap year (it wasn't), so
   * dates before March 1, 1900 are off by one day. This implementation matches Excel's behavior.
   *
   * @param dt
   *   The LocalDateTime to convert
   * @return
   *   Excel serial number (days since 1899-12-30 + fractional time)
   */
  def dateTimeToExcelSerial(dt: LocalDateTime): Double =
    import java.time.temporal.ChronoUnit

    // Excel epoch: December 30, 1899 (not Jan 1, 1900, to account for 1900 leap year bug)
    val epoch1900 = LocalDateTime.of(1899, 12, 30, 0, 0, 0)

    // Calculate days since epoch
    val days = ChronoUnit.DAYS.between(epoch1900, dt)

    // Calculate fractional day for time component
    val dayStart = dt.toLocalDate.atStartOfDay
    val secondsInDay = ChronoUnit.SECONDS.between(dayStart, dt)
    val fractionOfDay = secondsInDay.toDouble / 86400.0

    days.toDouble + fractionOfDay

  /**
   * Convert Excel serial number to LocalDateTime.
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
    import java.time.temporal.ChronoUnit

    // Excel epoch: December 30, 1899
    val epoch1900 = LocalDateTime.of(1899, 12, 30, 0, 0, 0)

    // Extract whole days and fractional day
    val wholeDays = serial.toLong
    val fractionOfDay = serial - wholeDays
    val seconds = (fractionOfDay * 86400.0).toLong

    // Add days and seconds to epoch
    epoch1900.plusDays(wholeDays).plusSeconds(seconds)

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
