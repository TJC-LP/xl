package com.tjclp.xl

import scala.quoted.*

/**
 * Compile-time validated cell and range literals.
 *
 * Usage:
 * {{{
 * cell"A1" // ARef validated at compile time
 * range"A1:B2" // CellRange validated at compile time
 * }}}
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
    try
      val (c0, r0) = parseCellLit(s)
      constARef(c0, r0)
    catch
      case e: IllegalArgumentException =>
        report.errorAndAbort(s"Invalid cell literal '$s': ${e.getMessage}")

  private def rangeImpl0(sc: Expr[StringContext])(using Quotes): Expr[CellRange] =
    import quotes.reflect.report
    val s = literal(sc)
    try
      val ((cs, rs), (ce, re)) = parseRangeLit(s)
      constRange(cs, rs, ce, re)
    catch
      case e: IllegalArgumentException =>
        report.errorAndAbort(s"Invalid range literal '$s': ${e.getMessage}")

  private def fxImpl0(sc: Expr[StringContext])(using Quotes): Expr[CellValue] =
    import quotes.reflect.report
    val s = literal(sc)

    // Minimal validation: no empty formulas, basic character check
    if s.isEmpty then report.errorAndAbort("Formula literal cannot be empty")

    // Check for balanced parentheses (simple validation)
    var depth = 0
    for c <- s do
      if c == '(' then depth += 1
      else if c == ')' then depth -= 1
      if depth < 0 then report.errorAndAbort(s"Formula literal has unbalanced parentheses: '$s'")
    if depth != 0 then report.errorAndAbort(s"Formula literal has unbalanced parentheses: '$s'")

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
    if parts.lengthCompare(1) != 0 then
      quotes.reflect.report.errorAndAbort("literal must be a single part")
    parts.head

  // Opaque-constant emitters
  private def constARef(col0: Int, row0: Int)(using Quotes): Expr[ARef] =
    val packed = (row0.toLong << 32) | (col0.toLong & 0xffffffffL)
    '{ (${ Expr(packed) }: Long).asInstanceOf[ARef] }

  private def constRange(cs: Int, rs: Int, ce: Int, re: Int)(using Quotes): Expr[CellRange] =
    '{ CellRange(${ constARef(cs, rs) }, ${ constARef(ce, re) }) }

  // Micro parsers (no regex, no allocations)
  private def parseCellLit(s: String): (Int, Int) =
    var i = 0; val n = s.length
    if n == 0 then fail("empty")
    var col = 0
    var parsing = true
    while i < n && parsing do
      val ch = s.charAt(i)
      if ch >= 'A' && ch <= 'Z' then { col = col * 26 + (ch - 'A' + 1); i += 1 }
      else if ch >= 'a' && ch <= 'z' then { col = col * 26 + (ch - 'a' + 1); i += 1 }
      else parsing = false
    if col == 0 then fail("missing column letters")
    var row = 0; var sawDigit = false
    while i < n && s.charAt(i).isDigit do
      row = row * 10 + (s.charAt(i) - '0'); i += 1; sawDigit = true
    if !sawDigit then fail("missing row digits")
    if i != n then fail("trailing junk")
    if row < 1 then fail("row must be ≥ 1")
    (col - 1, row - 1)

  private def parseRangeLit(s: String): ((Int, Int), (Int, Int)) =
    val i = s.indexOf(':'); if i <= 0 || i >= s.length - 1 then fail("use A1:B2 form")
    val (a, b) = (s.substring(0, i), s.substring(i + 1))
    val (c1, r1) = parseCellLit(a); val (c2, r2) = parseCellLit(b)
    val (cs, rs) = (math.min(c1, c2), math.min(r1, r2))
    val (ce, re) = (math.max(c1, c2), math.max(r1, r2))
    ((cs, rs), (ce, re))

  private def fail(msg: String): Nothing = throw new IllegalArgumentException(msg)

end macros

/** Export compile-time literals */
export macros.{cell, range}
