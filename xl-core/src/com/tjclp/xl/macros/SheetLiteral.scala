package com.tjclp.xl.macros

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet

import scala.quoted.*

/**
 * Compile-time validated sheet name for Sheet.apply.
 *
 * When Sheet("name") is called with a string literal, the name is validated at compile time and
 * returns `Sheet` directly. When called with a runtime variable, validation occurs at runtime and
 * returns `XLResult[Sheet]`.
 *
 * Validation rules (Excel sheet name constraints):
 *   - Cannot be empty
 *   - Maximum 31 characters
 *   - Cannot contain: : \ / ? * [ ]
 *
 * Examples:
 *   - `Sheet("Sales")` → compiles, returns `Sheet` (valid name)
 *   - `Sheet("Q1:Q2")` → compile error (contains ':')
 *   - `Sheet(userInput)` → compiles, returns `XLResult[Sheet]` (validated at runtime)
 */
object SheetLiteral:

  /**
   * Create a Sheet with compile-time validation when possible.
   *
   * If the name is a string literal, validates at compile time and returns `Sheet`. If the name is
   * a runtime expression, validates at runtime and returns `XLResult[Sheet]`.
   */
  transparent inline def sheet(inline name: String): Sheet | XLResult[Sheet] = ${ sheetImpl('name) }

  def sheetImpl(name: Expr[String])(using Quotes): Expr[Sheet | XLResult[Sheet]] =
    import quotes.reflect.*

    name.value match
      case Some(literalName) =>
        // Compile-time literal - validate now
        SheetName(literalName) match
          case Right(_) =>
            // Valid - create sheet with unsafe (we just validated)
            '{ Sheet(SheetName.unsafe($name)) }
          case Left(err) =>
            report.errorAndAbort(s"Invalid sheet name '$literalName': $err")

      case None =>
        // Runtime expression - validate at runtime (returns XLResult[Sheet])
        '{
          val result: XLResult[Sheet] = SheetName($name) match
            case Right(validName) => Right(Sheet(validName))
            case Left(err) => Left(XLError.InvalidSheetName($name, err))
          result
        }

end SheetLiteral

export SheetLiteral.*
