package com.tjclp.xl.cell

import java.time.LocalDateTime

/** Cell value types supported by Excel */
enum CellValue:
  /** Text/string value */
  case Text(value: String)

  /** Rich text with multiple formatting runs (inline string with formatting) */
  case RichText(value: com.tjclp.xl.RichText)

  /** Numeric value (Excel uses double internally, but we preserve precision) */
  case Number(value: BigDecimal)

  /** Boolean value */
  case Bool(value: Boolean)

  /** Date/time value */
  case DateTime(value: LocalDateTime)

  /** Formula expression (stored as string, can be typed later) */
  case Formula(expression: String)

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
