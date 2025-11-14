package com.tjclp.xl.macros

import com.tjclp.xl.addressing.{ARef, CellRange, RefType, SheetName}

import scala.quoted.*
import scala.util.boundary, boundary.break

/**
 * Compile-time validated unified reference literal.
 *
 * Supports all Excel reference formats with transparent return types:
 *   - `ref"A1"` → ARef (unwrapped)
 *   - `ref"A1:B10"` → CellRange (unwrapped)
 *   - `ref"Sales!A1"` → RefType.QualifiedCell
 *   - `ref"'Q1 Sales'!A1:B10"` → RefType.QualifiedRange
 *   - `ref"'It''s Q1'!A1"` → RefType.QualifiedCell (escaped quotes: '' → ')
 *
 * Uses zero-allocation parsing at compile time for maximum performance.
 *
 * Return type is transparent: simple refs return unwrapped ARef/CellRange for backwards
 * compatibility, qualified refs return RefType for sheet information.
 */
@SuppressWarnings(
  Array(
    "org.wartremover.warts.Var",
    "org.wartremover.warts.While"
  )
)
object RefLiteral:

  // -------- Public API --------
  extension (inline sc: StringContext)
    transparent inline def ref(): ARef | CellRange | RefType = ${ refImpl0('sc) }

    inline def ref(inline args: Any*): ARef | CellRange | RefType =
      ${ errorNoInterpolation('sc, 'args, "ref") }

  // -------- Implementation --------
  private def refImpl0(sc: Expr[StringContext])(using Quotes): Expr[ARef | CellRange | RefType] =
    import quotes.reflect.report
    val s = literal(sc)

    try
      // Check for sheet qualifier (!)
      val bangIdx = findUnquotedBang(s)
      if bangIdx < 0 then
        // No sheet qualifier - return unwrapped ARef or CellRange
        if s.contains(':') then
          val ((cs, rs), (ce, re)) = parseRangeLit(s)
          constCellRange(cs, rs, ce, re)
        else
          val (c0, r0) = parseCellLit(s)
          constARef(c0, r0)
      else
        // Has sheet qualifier - return RefType wrapper
        val sheetPart = s.substring(0, bangIdx)
        val refPart = s.substring(bangIdx + 1)

        if refPart.isEmpty then report.errorAndAbort(s"Missing reference after '!' in: $s")

        // Parse sheet name (handle quotes and escaping)
        val sheetName = if sheetPart.startsWith("'") then
          if !sheetPart.endsWith("'") then
            report.errorAndAbort(
              s"Unbalanced quotes in sheet name: $sheetPart (missing closing quote)"
            )
          val quoted = sheetPart.substring(1, sheetPart.length - 1)
          if quoted.isEmpty then report.errorAndAbort("Empty sheet name in quotes")
          // Unescape '' → ' (Excel convention)
          val unescaped = quoted.replace("''", "'")
          validateSheetName(unescaped)
          unescaped
        else
          if sheetPart.contains("'") then
            report.errorAndAbort(
              s"Misplaced quote in sheet name: $sheetPart (quotes must wrap entire name)"
            )
          validateSheetName(sheetPart)
          sheetPart

        // Parse ref part as cell or range
        if refPart.contains(':') then
          val ((cs, rs), (ce, re)) = parseRangeLit(refPart)
          constQualifiedRange(sheetName, cs, rs, ce, re)
        else
          val (c0, r0) = parseCellLit(refPart)
          constQualifiedCell(sheetName, c0, r0)

    catch
      case e: IllegalArgumentException =>
        report.errorAndAbort(s"Invalid ref literal '$s': ${e.getMessage}")

  private def errorNoInterpolation(sc: Expr[StringContext], args: Expr[Seq[Any]], kind: String)(
    using Quotes
  ): Expr[Nothing] =
    import quotes.reflect.report
    args match
      case Varargs(Nil) => report.errorAndAbort(s"""Use $kind"...": no interpolation supported""")
      case _ => report.errorAndAbort(s"""$kind"...": interpolation not supported""")

  private def literal(sc: Expr[StringContext])(using Quotes): String =
    val parts = sc.valueOrAbort.parts
    if parts.lengthCompare(1) != 0 then
      quotes.reflect.report.errorAndAbort("literal must be a single part")
    parts(0) // Safe: length == 1 verified above

  // -------- Const Emitters --------

  /**
   * Emit unwrapped ARef constant (for simple refs).
   *
   * Uses ARef.from0 public API for encapsulation. The inline function expands to the packed Long
   * representation at compile time with zero runtime overhead.
   */
  private def constARef(col0: Int, row0: Int)(using Quotes): Expr[ARef] =
    '{ ARef.from0(${ Expr(col0) }, ${ Expr(row0) }) }

  /** Emit unwrapped CellRange constant (for simple ranges) */
  private def constCellRange(cs: Int, rs: Int, ce: Int, re: Int)(using Quotes): Expr[CellRange] =
    '{ CellRange(${ constARef(cs, rs) }, ${ constARef(ce, re) }) }

  /** Emit RefType.QualifiedCell (for qualified refs) */
  private def constQualifiedCell(sheetName: String, col0: Int, row0: Int)(using
    Quotes
  ): Expr[RefType] =
    '{
      RefType.QualifiedCell(
        SheetName.unsafe(${ Expr(sheetName) }),
        ${ constARef(col0, row0) }
      )
    }

  /** Emit RefType.QualifiedRange (for qualified ranges) */
  private def constQualifiedRange(
    sheetName: String,
    cs: Int,
    rs: Int,
    ce: Int,
    re: Int
  )(using Quotes): Expr[RefType] =
    '{
      RefType.QualifiedRange(
        SheetName.unsafe(${ Expr(sheetName) }),
        CellRange(${ constARef(cs, rs) }, ${ constARef(ce, re) })
      )
    }

  // -------- Parsers (Zero-Allocation) --------

  /**
   * Find index of unquoted '!' (not inside 'quotes').
   *
   * Uses a toggle approach: each ' flips the inQuote state. This handles escaped quotes ('')
   * correctly because Excel's escaping convention uses two consecutive quotes to represent a single
   * literal quote.
   *
   * Examples:
   *   - "Sales!A1" → returns 5 (unquoted)
   *   - "'Sales!A1'!B1" → returns 11 (second ! is unquoted)
   *   - "'It''s Q1'!A1" → returns 10 ('' toggles twice, staying inside quotes)
   *
   * Returns -1 if no unquoted bang found.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def findUnquotedBang(s: String): Int =
    boundary:
      var i = 0
      var inQuote = false
      while i < s.length do
        val c = s.charAt(i)
        if c == '\'' then inQuote = !inQuote
        else if c == '!' && !inQuote then break(i)
        i += 1
      -1

  /**
   * Validate sheet name according to Excel rules. Max 31 chars, no `: \ / ? * [ ]`
   */
  private def validateSheetName(name: String): Unit =
    if name.isEmpty then fail("sheet name cannot be empty")
    if name.length > 31 then fail("sheet name max length is 31 chars")
    val invalid = Set(':', '\\', '/', '?', '*', '[', ']')
    name.foreach { c =>
      if invalid.contains(c) then fail(s"sheet name cannot contain: $c")
    }

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

end RefLiteral

/** Export unified ref literal */
export RefLiteral.*
