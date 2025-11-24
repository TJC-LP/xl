package com.tjclp.xl.macros

import com.tjclp.xl.cells.CellValue

import scala.quoted.*

/**
 * Formula literal macro.
 *
 * Usage:
 * {{{
 * fx"=SUM(A1:A10)" // CellValue.Formula validated at compile time
 * }}}
 *
 * Note: For cell and range references, use the unified `ref` macro instead.
 */
@SuppressWarnings(Array("org.wartremover.warts.Var"))
object CellRangeLiterals:

  // -------- Public API --------
  extension (inline sc: StringContext)
    transparent inline def fx(
      inline args: Any*
    ): CellValue | Either[com.tjclp.xl.error.XLError, CellValue] =
      ${ fxImplN('sc, 'args) }

  // -------- Implementations --------
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
    import quotes.reflect.report
    args match
      case Varargs(Nil) => report.errorAndAbort(s"""Use $kind"...": no interpolation supported""")
      case _ => report.errorAndAbort(s"""$kind"...": interpolation not supported""")

  /**
   * Macro implementation supporting compile-time literals, compile-time optimization, and runtime
   * interpolation.
   */
  private def fxImplN(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[CellValue | Either[com.tjclp.xl.error.XLError, CellValue]] =
    import quotes.reflect.*

    args match
      case Varargs(exprs) if exprs.isEmpty =>
        // Branch 1: No interpolation (Phase 1)
        fxImpl0(sc)

      case Varargs(_) =>
        MacroUtil.allLiterals(args) match
          case Some(literals) =>
            // Branch 2: All compile-time constants - OPTIMIZE (Phase 2)
            fxCompileTimeOptimized(sc, literals)
          case None =>
            // Branch 3: Has runtime variables (Phase 1)
            fxRuntimePath(sc, args)

  private def fxCompileTimeOptimized(
    sc: Expr[StringContext],
    literals: Seq[Any]
  )(using Quotes): Expr[CellValue] =
    import quotes.reflect.report
    val parts = sc.valueOrAbort.parts
    // Process escape sequences before reconstruction
    val processedParts = parts.map(processEscapes)
    val fullString = MacroUtil.reconstructString(processedParts, literals)

    // Call runtime parser at compile-time to ensure validation consistency
    com.tjclp.xl.cells.FormulaParser.parse(fullString) match
      case Right(cellValue) =>
        // Valid - emit constant
        cellValue match
          case CellValue.Formula(expr, _) =>
            '{ CellValue.Formula(${ Expr(expr) }) }
          case _ =>
            report.errorAndAbort("Unexpected cell value type in formula literal")
      case Left(error) =>
        // Invalid - compile error
        report.errorAndAbort(error.message)

  private def fxRuntimePath(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[Either[com.tjclp.xl.error.XLError, CellValue]] =
    '{
      com.tjclp.xl.cells.FormulaParser.parse($sc.s($args*))
    }.asExprOf[Either[com.tjclp.xl.error.XLError, CellValue]]

  private def literal(sc: Expr[StringContext])(using Quotes): String =
    val parts = sc.valueOrAbort.parts
    if parts.lengthCompare(1) != 0 then
      quotes.reflect.report.errorAndAbort("literal must be a single part")
    // Process escape sequences (custom interpolators receive raw escapes)
    processEscapes(parts(0))

  /**
   * Process standard string escape sequences.
   *
   * Custom string interpolators in Scala 3 receive the raw source string including escape sequences
   * like `\"`, `\\`, `\n`, etc. This function converts them to their actual character values.
   */
  @SuppressWarnings(Array("org.wartremover.warts.While", "org.wartremover.warts.Var"))
  private def processEscapes(s: String): String =
    val sb = new StringBuilder(s.length)
    var i = 0
    while i < s.length do
      if s.charAt(i) == '\\' && i + 1 < s.length then
        s.charAt(i + 1) match
          case 'n' => sb.append('\n'); i += 2
          case 'r' => sb.append('\r'); i += 2
          case 't' => sb.append('\t'); i += 2
          case '\\' => sb.append('\\'); i += 2
          case '"' => sb.append('"'); i += 2
          case '\'' => sb.append('\''); i += 2
          case _ => sb.append(s.charAt(i)); i += 1 // Keep unknown escapes as-is
      else
        sb.append(s.charAt(i))
        i += 1
    sb.toString

end CellRangeLiterals

/** Export formula literal */
export CellRangeLiterals.fx
