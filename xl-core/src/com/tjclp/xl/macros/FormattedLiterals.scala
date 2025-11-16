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
    transparent inline def money(
      inline args: Any*
    ): Formatted | Either[com.tjclp.xl.error.XLError, Formatted] =
      ${ moneyImplN('sc, 'args) }

    /** Percent literal: percent"45.5%" → Formatted(Number(0.455), NumFmt.Percent) */
    transparent inline def percent(
      inline args: Any*
    ): Formatted | Either[com.tjclp.xl.error.XLError, Formatted] =
      ${ percentImplN('sc, 'args) }

    /** Date literal: date"2025-11-10" → Formatted(DateTime(...), NumFmt.Date) */
    transparent inline def date(
      inline args: Any*
    ): Formatted | Either[com.tjclp.xl.error.XLError, Formatted] =
      ${ dateImplN('sc, 'args) }

    /** Accounting literal: accounting"($123.45)" → Formatted(Number(-123.45), NumFmt.Currency) */
    transparent inline def accounting(
      inline args: Any*
    ): Formatted | Either[com.tjclp.xl.error.XLError, Formatted] =
      ${ accountingImplN('sc, 'args) }

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

  // ===== Runtime Implementations =====

  private def moneyImplN(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[Formatted | Either[com.tjclp.xl.error.XLError, Formatted]] =
    args match
      case Varargs(exprs) if exprs.isEmpty =>
        // Branch 1: No interpolation (Phase 1)
        moneyImpl(sc)
      case Varargs(_) =>
        MacroUtil.allLiterals(args) match
          case Some(literals) =>
            // Branch 2: All compile-time constants - OPTIMIZE (Phase 2)
            moneyCompileTimeOptimized(sc, literals)
          case None =>
            // Branch 3: Has runtime variables (Phase 1)
            moneyRuntimePath(sc, args)

  private def moneyCompileTimeOptimized(
    sc: Expr[StringContext],
    literals: Seq[Any]
  )(using Quotes): Expr[Formatted] =
    import quotes.reflect.report
    val parts = sc.valueOrAbort.parts
    val fullString = MacroUtil.reconstructString(parts, literals)

    try
      val cleaned = fullString.replaceAll("[\\$,]", "")
      val numStr = cleaned
      '{
        Formatted(CellValue.Number(BigDecimal(${ Expr(numStr) })), NumFmt.Currency)
      }
    catch
      case e: Exception =>
        report.errorAndAbort(MacroUtil.formatCompileError("money", fullString, e.getMessage))

  private def moneyRuntimePath(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[Either[com.tjclp.xl.error.XLError, Formatted]] =
    '{
      com.tjclp.xl.formatted.FormattedParsers.parseMoney($sc.s($args*))
    }.asExprOf[Either[com.tjclp.xl.error.XLError, Formatted]]

  private def percentImplN(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[Formatted | Either[com.tjclp.xl.error.XLError, Formatted]] =
    args match
      case Varargs(exprs) if exprs.isEmpty =>
        percentImpl(sc)
      case Varargs(_) =>
        MacroUtil.allLiterals(args) match
          case Some(literals) => percentCompileTimeOptimized(sc, literals)
          case None => percentRuntimePath(sc, args)

  private def percentCompileTimeOptimized(
    sc: Expr[StringContext],
    literals: Seq[Any]
  )(using Quotes): Expr[Formatted] =
    import quotes.reflect.report
    val parts = sc.valueOrAbort.parts
    val fullString = MacroUtil.reconstructString(parts, literals)

    try
      val cleaned = fullString.replace("%", "")
      val num = BigDecimal(cleaned) / 100
      val numStr = num.toString
      '{
        Formatted(CellValue.Number(BigDecimal(${ Expr(numStr) })), NumFmt.Percent)
      }
    catch
      case e: Exception =>
        report.errorAndAbort(MacroUtil.formatCompileError("percent", fullString, e.getMessage))

  private def percentRuntimePath(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[Either[com.tjclp.xl.error.XLError, Formatted]] =
    '{
      com.tjclp.xl.formatted.FormattedParsers.parsePercent($sc.s($args*))
    }.asExprOf[Either[com.tjclp.xl.error.XLError, Formatted]]

  private def dateImplN(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[Formatted | Either[com.tjclp.xl.error.XLError, Formatted]] =
    args match
      case Varargs(exprs) if exprs.isEmpty =>
        dateImpl(sc)
      case Varargs(_) =>
        MacroUtil.allLiterals(args) match
          case Some(literals) => dateCompileTimeOptimized(sc, literals)
          case None => dateRuntimePath(sc, args)

  private def dateCompileTimeOptimized(
    sc: Expr[StringContext],
    literals: Seq[Any]
  )(using Quotes): Expr[Formatted] =
    import quotes.reflect.report
    val parts = sc.valueOrAbort.parts
    val fullString = MacroUtil.reconstructString(parts, literals)

    try
      val localDate = LocalDate.parse(fullString)
      val dateStr = localDate.toString
      '{
        Formatted(
          CellValue.DateTime(java.time.LocalDate.parse(${ Expr(dateStr) }).atStartOfDay()),
          NumFmt.Date
        )
      }
    catch
      case e: Exception =>
        report.errorAndAbort(MacroUtil.formatCompileError("date", fullString, e.getMessage))

  private def dateRuntimePath(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[Either[com.tjclp.xl.error.XLError, Formatted]] =
    '{
      com.tjclp.xl.formatted.FormattedParsers.parseDate($sc.s($args*))
    }.asExprOf[Either[com.tjclp.xl.error.XLError, Formatted]]

  private def accountingImplN(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[Formatted | Either[com.tjclp.xl.error.XLError, Formatted]] =
    args match
      case Varargs(exprs) if exprs.isEmpty =>
        accountingImpl(sc)
      case Varargs(_) =>
        MacroUtil.allLiterals(args) match
          case Some(literals) => accountingCompileTimeOptimized(sc, literals)
          case None => accountingRuntimePath(sc, args)

  private def accountingCompileTimeOptimized(
    sc: Expr[StringContext],
    literals: Seq[Any]
  )(using Quotes): Expr[Formatted] =
    import quotes.reflect.report
    val parts = sc.valueOrAbort.parts
    val fullString = MacroUtil.reconstructString(parts, literals)

    try
      val isNegative = fullString.contains("(") && fullString.contains(")")
      val cleaned = fullString.replaceAll("[\\$,()\\s]", "")
      val num = if isNegative then -BigDecimal(cleaned) else BigDecimal(cleaned)
      val numStr = num.toString
      '{
        Formatted(CellValue.Number(BigDecimal(${ Expr(numStr) })), NumFmt.Currency)
      }
    catch
      case e: Exception =>
        report.errorAndAbort(MacroUtil.formatCompileError("accounting", fullString, e.getMessage))

  private def accountingRuntimePath(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[Either[com.tjclp.xl.error.XLError, Formatted]] =
    '{
      com.tjclp.xl.formatted.FormattedParsers.parseAccounting($sc.s($args*))
    }.asExprOf[Either[com.tjclp.xl.error.XLError, Formatted]]

end FormattedLiterals

/** Export formatted literals for import */
export FormattedLiterals.*
