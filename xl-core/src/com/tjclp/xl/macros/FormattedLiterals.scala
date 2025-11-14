package com.tjclp.xl.macros

import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.formatted.Formatted
import com.tjclp.xl.style.numfmt.NumFmt

import java.time.LocalDate
import scala.quoted.*

/** Formatted literal macros for compile-time parsing */
object FormattedLiterals:

  extension (inline sc: StringContext)
    /** Money literal: money"$1,234.56" → Formatted(Number(1234.56), NumFmt.Currency) */
    transparent inline def money(): Formatted =
      ${ moneyImpl('sc) }

    /** Percent literal: percent"45.5%" → Formatted(Number(0.455), NumFmt.Percent) */
    transparent inline def percent(): Formatted =
      ${ percentImpl('sc) }

    /** Date literal: date"2025-11-10" → Formatted(DateTime(...), NumFmt.Date) */
    transparent inline def date(): Formatted =
      ${ dateImpl('sc) }

    /** Accounting literal: accounting"($123.45)" → Formatted(Number(-123.45), NumFmt.Currency) */
    transparent inline def accounting(): Formatted =
      ${ accountingImpl('sc) }

  // Helper to extract literal from StringContext
  private def getLiteral(sc: Expr[StringContext])(using Quotes): String =
    val parts = sc.valueOrAbort.parts
    if parts.lengthCompare(1) != 0 then
      quotes.reflect.report.errorAndAbort("literal must be a single part")
    parts(0) // Safe: length == 1 verified above

  private def moneyImpl(sc: Expr[StringContext])(using Quotes): Expr[Formatted] =
    import quotes.reflect.report
    val s = getLiteral(sc)

    try
      // Parse money format: strip $, commas, parse number
      val cleaned = s.replaceAll("[\\$,]", "")
      val numStr = cleaned // Keep as string for runtime parsing
      '{
        Formatted(
          CellValue.Number(BigDecimal(${ Expr(numStr) })),
          NumFmt.Currency
        )
      }
    catch
      case e: Exception =>
        report.errorAndAbort(s"Invalid money literal '$s': ${e.getMessage}")

  private def percentImpl(sc: Expr[StringContext])(using Quotes): Expr[Formatted] =
    import quotes.reflect.report
    val s = getLiteral(sc)

    try
      // Parse percent: strip %, divide by 100
      val cleaned = s.replace("%", "")
      val num = BigDecimal(cleaned) / 100
      val numStr = num.toString
      '{
        Formatted(
          CellValue.Number(BigDecimal(${ Expr(numStr) })),
          NumFmt.Percent
        )
      }
    catch
      case e: Exception =>
        report.errorAndAbort(s"Invalid percent literal '$s': ${e.getMessage}")

  private def dateImpl(sc: Expr[StringContext])(using Quotes): Expr[Formatted] =
    import quotes.reflect.report
    val s = getLiteral(sc)

    try
      // Parse ISO date format: 2025-11-10 and convert to string for runtime parsing
      val localDate = LocalDate.parse(s)
      val dateStr = localDate.toString // ISO format
      '{
        Formatted(
          CellValue.DateTime(java.time.LocalDate.parse(${ Expr(dateStr) }).atStartOfDay()),
          NumFmt.Date
        )
      }
    catch
      case e: Exception =>
        report.errorAndAbort(
          s"Invalid date literal '$s': expected ISO format (YYYY-MM-DD), got error: ${e.getMessage}"
        )

  private def accountingImpl(sc: Expr[StringContext])(using Quotes): Expr[Formatted] =
    import quotes.reflect.report
    val s = getLiteral(sc)

    try
      // Parse accounting format: ($123.45) or $123.45
      val isNegative = s.contains("(") && s.contains(")")
      val cleaned = s.replaceAll("[\\$,()\\s]", "")
      val num = if isNegative then -BigDecimal(cleaned) else BigDecimal(cleaned)
      val numStr = num.toString
      '{
        Formatted(
          CellValue.Number(BigDecimal(${ Expr(numStr) })),
          NumFmt.Currency
        )
      }
    catch
      case e: Exception =>
        report.errorAndAbort(s"Invalid accounting literal '$s': ${e.getMessage}")

end FormattedLiterals

/** Export formatted literals for import */
export FormattedLiterals.*
