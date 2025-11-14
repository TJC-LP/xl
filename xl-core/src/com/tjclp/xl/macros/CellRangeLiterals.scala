package com.tjclp.xl.macros

import com.tjclp.xl.cell.CellValue

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
    transparent inline def fx(): CellValue = ${ fxImpl0('sc) }

    inline def fx(inline args: Any*): CellValue =
      ${ errorNoInterpolation('sc, 'args, "fx") }

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

  private def literal(sc: Expr[StringContext])(using Quotes): String =
    val parts = sc.valueOrAbort.parts
    if parts.lengthCompare(1) != 0 then
      quotes.reflect.report.errorAndAbort("literal must be a single part")
    parts(0) // Safe: length == 1 verified above

end CellRangeLiterals

/** Export formula literal */
export CellRangeLiterals.fx
