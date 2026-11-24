package com.tjclp.xl.formula.eval

import java.time.LocalDate
import java.time.temporal.ChronoUnit
import scala.math.BigDecimal

/**
 * #953: the date a text spells, as Excel's 1900-system serial, in Excel's US forms. The one parser
 * behind DATEVALUE, VALUE, every single-value text-to-number coercion
 * ([[ScalarCoercion.parseValueText]]) and the XIRR/XNPV dates range (#956).
 *
 * Accepted, trimmed and case-insensitive (`mmm` is a three-letter month or the full name):
 *   - `m/d/yyyy`, `m-d-yyyy`, `m/d/yy` (the separator repeats; a two-digit year 00-29 is 20yy,
 *     30-99 is 19yy, Excel's default cutoff)
 *   - `yyyy-m-d`, `yyyy/m/d`
 *   - `d-mmm-yyyy`, `d mmm yyyy`, `d-mmm-yy`
 *   - `mmm d, yyyy`, `mmm d yyyy`
 *   - `mmm-yyyy`, `mmm yyyy` (the first of the month)
 *
 * The date must exist and lie in 1900-01-01..9999-12-31; Excel's phantom 1900-02-29 is serial 60.
 * Not read (None, so `#VALUE!`): a time part, forms without a year (Excel takes the current year,
 * which would make a result depend on the clock), and day-first numeric order.
 */
private[formula] object DateText:

  private val Months: Map[String, Int] =
    val names = List(
      "january",
      "february",
      "march",
      "april",
      "may",
      "june",
      "july",
      "august",
      "september",
      "october",
      "november",
      "december"
    )
    names.zipWithIndex.flatMap { case (name, i) =>
      List(name -> (i + 1), name.take(3) -> (i + 1))
    }.toMap

  private val MonthDayYear = """(\d{1,2})([/-])(\d{1,2})\2(\d{1,2}|\d{4})""".r
  private val YearMonthDay = """(\d{4})([/-])(\d{1,2})\2(\d{1,2})""".r
  private val DayMonYear = """(\d{1,2})([- ])([a-z]+)\2(\d{2}|\d{4})""".r
  private val MonDayYear = """([a-z]+) (\d{1,2}),? (\d{4})""".r
  private val MonYear = """([a-z]+)[- ](\d{4})""".r

  private val Epoch = LocalDate.of(1899, 12, 30)
  private val FirstDate = LocalDate.of(1900, 1, 1)
  private val LastDate = LocalDate.of(9999, 12, 31)
  private val PhantomLeapDaySerial = BigDecimal(60)

  /** The serial of the date `s` spells, or None when it spells none of the accepted forms. */
  def serial(s: String): Option[BigDecimal] =
    val t = s.trim.toLowerCase(java.util.Locale.ROOT).replaceAll("\\s+", " ")
    t match
      case MonthDayYear(m, _, d, y) => toSerial(year(y), m.toInt, d.toInt)
      case YearMonthDay(y, _, m, d) => toSerial(y.toInt, m.toInt, d.toInt)
      case DayMonYear(d, _, mon, y) => Months.get(mon).flatMap(toSerial(year(y), _, d.toInt))
      case MonDayYear(mon, d, y) => Months.get(mon).flatMap(toSerial(y.toInt, _, d.toInt))
      case MonYear(mon, y) => Months.get(mon).flatMap(toSerial(y.toInt, _, 1))
      case _ => None

  /** A one- or two-digit year by Excel's default cutoff; a four-digit year as written. */
  private def year(y: String): Int =
    val n = y.toInt
    if y.length > 2 then n else if n < 30 then 2000 + n else 1900 + n

  private def toSerial(y: Int, m: Int, d: Int): Option[BigDecimal] =
    if y == 1900 && m == 2 && d == 29 then Some(PhantomLeapDaySerial)
    else
      scala.util
        .Try(LocalDate.of(y, m, d))
        .toOption
        .filter(date => !date.isBefore(FirstDate) && !date.isAfter(LastDate))
        .map { date =>
          val days = ChronoUnit.DAYS.between(Epoch, date)
          // Excel counts the phantom 1900-02-29, so dates before 1900-03-01 are one lower
          BigDecimal(if days < 61 then days - 1 else days)
        }
