package com.tjclp.xl

import scala.quoted.*

/**
 * Compile-time validated cell and range literals.
 *
 * Usage: cell"A1" // ARef validated at compile time range"A1:B2" // CellRange validated at compile
 * time
 */
object macros:

  // -------- Public API (0-arg fast path) --------
  extension (inline sc: StringContext)
    transparent inline def cell(): ARef = ${ cellImpl0('sc) }
    transparent inline def range(): CellRange = ${ rangeImpl0('sc) }
    transparent inline def fx(): CellValue = ${ fxImpl0('sc) }

    inline def cell(inline args: Any*): ARef =
      ${ errorNoInterpolation('sc, 'args, "cell") }
    inline def range(inline args: Any*): CellRange =
      ${ errorNoInterpolation('sc, 'args, "range") }
    inline def fx(inline args: Any*): CellValue =
      ${ errorNoInterpolation('sc, 'args, "fx") }

  // -------- Implementations --------
  private def cellImpl0(sc: Expr[StringContext])(using Quotes): Expr[ARef] =
    import quotes.reflect.report
    val s = literal(sc)
    try {
      val (c0, r0) = parseCellLit(s)
      constARef(c0, r0)
    } catch {
      case e: IllegalArgumentException =>
        report.errorAndAbort(s"Invalid cell literal '$s': ${e.getMessage}")
    }

  private def rangeImpl0(sc: Expr[StringContext])(using Quotes): Expr[CellRange] =
    import quotes.reflect.report
    val s = literal(sc)
    try {
      val ((cs, rs), (ce, re)) = parseRangeLit(s)
      constRange(cs, rs, ce, re)
    } catch {
      case e: IllegalArgumentException =>
        report.errorAndAbort(s"Invalid range literal '$s': ${e.getMessage}")
    }

  private def fxImpl0(sc: Expr[StringContext])(using Quotes): Expr[CellValue] =
    import quotes.reflect.report
    val s = literal(sc)

    // Minimal validation: no empty formulas, basic character check
    if (s.isEmpty) {
      report.errorAndAbort("Formula literal cannot be empty")
    }

    // Check for balanced parentheses (simple validation)
    var depth = 0
    for (c <- s) {
      if (c == '(') depth += 1
      else if (c == ')') depth -= 1
      if (depth < 0) {
        report.errorAndAbort(s"Formula literal has unbalanced parentheses: '$s'")
      }
    }
    if (depth != 0) {
      report.errorAndAbort(s"Formula literal has unbalanced parentheses: '$s'")
    }

    // Emit CellValue.Formula with the validated string
    '{ CellValue.Formula(${ Expr(s) }) }

  private def errorNoInterpolation(sc: Expr[StringContext], args: Expr[Seq[Any]], kind: String)(
    using Quotes
  ): Expr[Nothing] =
    import quotes.reflect.*
    args match
      case Varargs(Nil) => report.errorAndAbort(s"""Use $kind"...": no interpolation supported""")
      case _ => report.errorAndAbort(s"""$kind"...": interpolation not supported""")

  private def literal(sc: Expr[StringContext])(using Quotes): String =
    val parts = sc.valueOrAbort.parts
    if (parts.lengthCompare(1) != 0)
      quotes.reflect.report.errorAndAbort("literal must be a single part")
    parts.head

  // Opaque-constant emitters
  private def constARef(col0: Int, row0: Int)(using Quotes): Expr[ARef] =
    val packed = (row0.toLong << 32) | (col0.toLong & 0xffffffffL)
    '{ (${ Expr(packed) }: Long).asInstanceOf[ARef] }

  private def constRange(cs: Int, rs: Int, ce: Int, re: Int)(using Quotes): Expr[CellRange] =
    '{ CellRange(${ constARef(cs, rs) }, ${ constARef(ce, re) }) }

  // Micro parsers (no regex, no allocations)
  private def parseCellLit(s: String): (Int, Int) = {
    var i = 0; val n = s.length
    if (n == 0) fail("empty")
    var col = 0
    var parsing = true
    while (i < n && parsing) {
      val ch = s.charAt(i)
      if (ch >= 'A' && ch <= 'Z') { col = col * 26 + (ch - 'A' + 1); i += 1 }
      else if (ch >= 'a' && ch <= 'z') { col = col * 26 + (ch - 'a' + 1); i += 1 }
      else { parsing = false }
    }
    if (col == 0) fail("missing column letters")
    var row = 0; var sawDigit = false
    while (i < n && s.charAt(i).isDigit) {
      row = row * 10 + (s.charAt(i) - '0'); i += 1; sawDigit = true
    }
    if (!sawDigit) fail("missing row digits")
    if (i != n) fail("trailing junk")
    if (row < 1) fail("row must be ≥ 1")
    (col - 1, row - 1)
  }

  private def parseRangeLit(s: String): ((Int, Int), (Int, Int)) = {
    val i = s.indexOf(':'); if (i <= 0 || i >= s.length - 1) fail("use A1:B2 form")
    val (a, b) = (s.substring(0, i), s.substring(i + 1))
    val (c1, r1) = parseCellLit(a); val (c2, r2) = parseCellLit(b)
    val (cs, rs) = (math.min(c1, c2), math.min(r1, r2))
    val (ce, re) = (math.max(c1, c2), math.max(r1, r2))
    ((cs, rs), (ce, re))
  }

  private def fail(msg: String): Nothing = throw new IllegalArgumentException(msg)

end macros

/** Import to use cell and range literals */
export macros.{cell, range}

/** Batch put macro for elegant multi-cell updates */
object putMacro:
  import scala.quoted.*

  /**
   * Batch put with automatic CellValue conversion
   *
   * Usage:
   * {{{
   * import com.tjclp.xl.putMacro.put
   *
   * sheet.put(
   *   cell"A1" -> "Hello",
   *   cell"B1" -> 42,
   *   cell"C1" -> true
   * )
   * }}}
   *
   * Expands to chained .put() calls with zero intermediate allocations.
   */
  extension (sheet: com.tjclp.xl.Sheet)
    transparent inline def put(inline pairs: (ARef, Any)*): com.tjclp.xl.Sheet =
      ${ putImpl('sheet, 'pairs) }

  private def putImpl(sheet: Expr[com.tjclp.xl.Sheet], pairs: Expr[Seq[(ARef, Any)]])(using
    Quotes
  ): Expr[com.tjclp.xl.Sheet] =
    import quotes.reflect.*

    pairs match
      case Varargs(pairExprs) =>
        // Build chained put calls
        pairExprs.foldLeft(sheet) { (sheetExpr, pairExpr) =>
          // Extract ArrowAssoc(ref).->(value) components
          val (ref, value) = pairExpr.asTerm match
            // Match: ref -> value (which becomes ArrowAssoc(ref).->(value))
            case Apply(TypeApply(Select(Apply(_, List(r)), "->"), _), List(v)) =>
              (r.asExprOf[ARef], v.asExpr)
            // Also try inlined version
            case Inlined(_, _, Apply(TypeApply(Select(Apply(_, List(r)), "->"), _), List(v))) =>
              (r.asExprOf[ARef], v.asExpr)
            case other =>
              report.errorAndAbort(s"Batch put requires tuple pairs: cell\"A1\" -> value")

          // Generate CellValue based on runtime value (use CellValue.from)
          '{ $sheetExpr.put($ref, com.tjclp.xl.CellValue.from($value)) }
        }

      case _ =>
        report.errorAndAbort("Batch put requires literal tuple arguments")

end putMacro

/** Formatted literal macros for compile-time parsing */
object formattedLiterals:
  import scala.quoted.*
  import java.time.LocalDate

  extension (inline sc: StringContext)
    /** Money literal: money"$1,234.56" → Formatted(Number(1234.56), NumFmt.Currency) */
    transparent inline def money(): com.tjclp.xl.Formatted =
      ${ moneyImpl('sc) }

    /** Percent literal: percent"45.5%" → Formatted(Number(0.455), NumFmt.Percent) */
    transparent inline def percent(): com.tjclp.xl.Formatted =
      ${ percentImpl('sc) }

    /** Date literal: date"2025-11-10" → Formatted(DateTime(...), NumFmt.Date) */
    transparent inline def date(): com.tjclp.xl.Formatted =
      ${ dateImpl('sc) }

    /** Accounting literal: accounting"($123.45)" → Formatted(Number(-123.45), NumFmt.Currency) */
    transparent inline def accounting(): com.tjclp.xl.Formatted =
      ${ accountingImpl('sc) }

  // Helper to extract literal from StringContext
  private def getLiteral(sc: Expr[StringContext])(using Quotes): String =
    val parts = sc.valueOrAbort.parts
    if (parts.lengthCompare(1) != 0)
      quotes.reflect.report.errorAndAbort("literal must be a single part")
    parts.head

  private def moneyImpl(sc: Expr[StringContext])(using Quotes): Expr[com.tjclp.xl.Formatted] =
    import quotes.reflect.report
    val s = getLiteral(sc)

    try
      // Parse money format: strip $, commas, parse number
      val cleaned = s.replaceAll("[\\$,]", "")
      val numStr = cleaned // Keep as string for runtime parsing
      '{
        com.tjclp.xl.Formatted(
          com.tjclp.xl.CellValue.Number(BigDecimal(${ Expr(numStr) })),
          com.tjclp.xl.NumFmt.Currency
        )
      }
    catch
      case e: Exception =>
        report.errorAndAbort(s"Invalid money literal '$s': ${e.getMessage}")

  private def percentImpl(sc: Expr[StringContext])(using Quotes): Expr[com.tjclp.xl.Formatted] =
    import quotes.reflect.report
    val s = getLiteral(sc)

    try
      // Parse percent: strip %, divide by 100
      val cleaned = s.replace("%", "")
      val num = BigDecimal(cleaned) / 100
      val numStr = num.toString
      '{
        com.tjclp.xl.Formatted(
          com.tjclp.xl.CellValue.Number(BigDecimal(${ Expr(numStr) })),
          com.tjclp.xl.NumFmt.Percent
        )
      }
    catch
      case e: Exception =>
        report.errorAndAbort(s"Invalid percent literal '$s': ${e.getMessage}")

  private def dateImpl(sc: Expr[StringContext])(using Quotes): Expr[com.tjclp.xl.Formatted] =
    import quotes.reflect.report
    val s = getLiteral(sc)

    try
      // Parse ISO date format: 2025-11-10 and convert to string for runtime parsing
      val localDate = LocalDate.parse(s)
      val dateStr = localDate.toString // ISO format
      '{
        com.tjclp.xl.Formatted(
          com.tjclp.xl.CellValue
            .DateTime(java.time.LocalDate.parse(${ Expr(dateStr) }).atStartOfDay()),
          com.tjclp.xl.NumFmt.Date
        )
      }
    catch
      case e: Exception =>
        report.errorAndAbort(
          s"Invalid date literal '$s': expected ISO format (YYYY-MM-DD), got error: ${e.getMessage}"
        )

  private def accountingImpl(sc: Expr[StringContext])(using Quotes): Expr[com.tjclp.xl.Formatted] =
    import quotes.reflect.report
    val s = getLiteral(sc)

    try
      // Parse accounting format: ($123.45) or $123.45
      val isNegative = s.contains("(") && s.contains(")")
      val cleaned = s.replaceAll("[\\$,()\\s]", "")
      val num = if isNegative then -BigDecimal(cleaned) else BigDecimal(cleaned)
      val numStr = num.toString
      '{
        com.tjclp.xl.Formatted(
          com.tjclp.xl.CellValue.Number(BigDecimal(${ Expr(numStr) })),
          com.tjclp.xl.NumFmt.Currency
        )
      }
    catch
      case e: Exception =>
        report.errorAndAbort(s"Invalid accounting literal '$s': ${e.getMessage}")

end formattedLiterals

/** Export formatted literals for import */
export formattedLiterals.{money, percent, date, accounting}
