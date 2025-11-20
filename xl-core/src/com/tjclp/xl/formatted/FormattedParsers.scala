package com.tjclp.xl.formatted

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError
import com.tjclp.xl.styles.numfmt.NumFmt
import java.time.LocalDate
import scala.util.{Try, Success, Failure}

/**
 * Pure runtime parsers for formatted literal strings.
 *
 * These mirror the compile-time macro parsing logic but operate at runtime. All parsers are total
 * functions returning Either[XLError, Formatted].
 */
object FormattedParsers:

  /** Excel's maximum cell content length per ECMA-376 specification */
  private val ExcelCellLimit = 32767

  /**
   * Parse money format: $1,234.56
   *
   * Accepts:
   *   - With/without dollar sign: "$1234.56" or "1234.56"
   *   - With/without commas: "$1,234.56" or "$1234.56"
   *   - With/without spaces: "$ 1,234.56" or "$1,234.56"
   *
   * Rejects:
   *   - Non-numeric: "$ABC"
   *   - Multiple decimals: "$1.2.3"
   *   - Empty: "" or "$"
   */
  def parseMoney(s: String): Either[XLError, Formatted] =
    if s.length > ExcelCellLimit then
      Left(
        XLError.MoneyFormatError(
          s.take(50) + "...",
          s"Input exceeds Excel cell limit ($ExcelCellLimit chars)"
        )
      )
    else
      Try {
        val cleaned = s.replaceAll("[\\$,\\s]", "")
        if cleaned.isEmpty then
          throw new NumberFormatException("Empty value after removing $ and ,")
        val num = BigDecimal(cleaned)
        Formatted(CellValue.Number(num), NumFmt.Currency)
      }.toEither.left.map { err =>
        XLError.MoneyFormatError(s, err.getMessage)
      }

  /**
   * Parse percent format: 45.5%
   *
   * Accepts:
   *   - With percent sign: "45.5%"
   *   - Without percent sign: "45.5" (treated as 45.5%, converted to 0.455)
   *   - Integer: "50%"
   *
   * Rejects:
   *   - Non-numeric: "ABC%"
   *   - Multiple percent signs: "50%%"
   *   - Empty: "" or "%"
   */
  def parsePercent(s: String): Either[XLError, Formatted] =
    if s.length > ExcelCellLimit then
      Left(
        XLError.PercentFormatError(
          s.take(50) + "...",
          s"Input exceeds Excel cell limit ($ExcelCellLimit chars)"
        )
      )
    else
      Try {
        // Check for multiple percent signs before removing them
        if s.count(_ == '%') > 1 then throw new NumberFormatException("Multiple percent signs")
        val cleaned = s.replace("%", "").trim
        if cleaned.isEmpty then throw new NumberFormatException("Empty value after removing %")
        val num = BigDecimal(cleaned) / 100
        Formatted(CellValue.Number(num), NumFmt.Percent)
      }.toEither.left.map { err =>
        XLError.PercentFormatError(s, err.getMessage)
      }

  /**
   * Parse ISO date format: 2025-11-10
   *
   * Accepts:
   *   - ISO 8601 format only: "YYYY-MM-DD"
   *   - Valid dates only (no Feb 30, etc.)
   *
   * Rejects:
   *   - US format: "11/10/2025"
   *   - European format: "10/11/2025"
   *   - Invalid dates: "2025-02-30"
   *   - Non-dates: "not-a-date"
   *
   * Note: Other formats will be added in Phase 2 (compile-time optimization)
   */
  def parseDate(s: String): Either[XLError, Formatted] =
    if s.length > ExcelCellLimit then
      Left(
        XLError.DateFormatError(
          s.take(50) + "...",
          s"Input exceeds Excel cell limit ($ExcelCellLimit chars)"
        )
      )
    else
      Try {
        val localDate = LocalDate.parse(s) // ISO format: YYYY-MM-DD
        val dateTime = localDate.atStartOfDay()
        Formatted(CellValue.DateTime(dateTime), NumFmt.Date)
      }.toEither.left.map { err =>
        XLError.DateFormatError(
          s,
          s"Expected ISO format (YYYY-MM-DD): ${err.getMessage}"
        )
      }

  /**
   * Parse accounting format: ($123.45) or $123.45
   *
   * Accepts:
   *   - Positive: "$123.45"
   *   - Negative with parens: "($123.45)"
   *   - With/without commas: "$1,234.56" or "$1234.56"
   *
   * Rejects:
   *   - Negative without parens: "-$123.45" (use parseMoney for this)
   *   - Non-numeric: "$ABC"
   *   - Empty: "" or "$()"
   */
  def parseAccounting(s: String): Either[XLError, Formatted] =
    if s.length > ExcelCellLimit then
      Left(
        XLError.AccountingFormatError(
          s.take(50) + "...",
          s"Input exceeds Excel cell limit ($ExcelCellLimit chars)"
        )
      )
    else
      Try {
        val isNegative = s.contains("(") && s.contains(")")
        val cleaned = s.replaceAll("[\\$,()\\s]", "")
        if cleaned.isEmpty then
          throw new NumberFormatException("Empty value after removing $ , ( )")
        val num = if isNegative then -BigDecimal(cleaned) else BigDecimal(cleaned)
        Formatted(CellValue.Number(num), NumFmt.Currency)
      }.toEither.left.map { err =>
        XLError.AccountingFormatError(s, err.getMessage)
      }
