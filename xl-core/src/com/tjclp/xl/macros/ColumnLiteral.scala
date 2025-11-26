package com.tjclp.xl.macros

import com.tjclp.xl.addressing.Column

import scala.quoted.*

/**
 * Compile-time validated column literal.
 *
 * Supports Excel column letter notation with compile-time validation:
 *   - `col"A"` → Column (index 0)
 *   - `col"Z"` → Column (index 25)
 *   - `col"AA"` → Column (index 26)
 *   - `col"XFD"` → Column (index 16383, max Excel column)
 *
 * Invalid columns are rejected at compile time:
 *   - `col"XFE"` → compile error (out of range)
 *   - `col"1"` → compile error (not a valid column letter)
 *   - `col""` → compile error (empty)
 */
object ColumnLiteral:

  // -------- Public API --------
  extension (inline sc: StringContext)
    inline def col(): Column = ${ colImpl0('sc) }

    inline def col(inline args: Any*): Column = ${ colImplN('sc, 'args) }

  // -------- Implementation --------
  private def colImpl0(sc: Expr[StringContext])(using Quotes): Expr[Column] =
    import quotes.reflect.report
    val s = literal(sc)

    Column.fromLetter(s) match
      case Right(col) => '{ Column.from0(${ Expr(col.index0) }) }
      case Left(err) => report.errorAndAbort(s"Invalid column literal '$s': $err")

  private def colImplN(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[Column] =
    import quotes.reflect.*

    args match
      case Varargs(exprs) if exprs.isEmpty =>
        // No interpolation - compile-time literal
        colImpl0(sc)

      case Varargs(exprs) =>
        // Check if all arguments are compile-time constants
        MacroUtil.allLiterals(args) match
          case Some(literals) =>
            // All compile-time constants - optimize
            val parts = sc.valueOrAbort.parts
            val fullString = MacroUtil.reconstructString(parts, literals)
            Column.fromLetter(fullString) match
              case Right(col) => '{ Column.from0(${ Expr(col.index0) }) }
              case Left(err) =>
                report.errorAndAbort(MacroUtil.formatCompileError("col", fullString, err))

          case None =>
            // Has runtime variables - runtime parsing
            '{
              val str = $sc.s($args*)
              Column.fromLetter(str) match
                case Right(col) => col
                case Left(err) =>
                  throw new IllegalArgumentException(s"Invalid column: $str - $err")
            }

  private def literal(sc: Expr[StringContext])(using Quotes): String =
    val parts = sc.valueOrAbort.parts
    if parts.lengthCompare(1) != 0 then
      quotes.reflect.report.errorAndAbort("col literal must be a single part without interpolation")
    parts(0) // Safe: length == 1 verified above

end ColumnLiteral

/** Export column literal */
export ColumnLiteral.*
