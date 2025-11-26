package com.tjclp.xl.macros

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.sheets.Sheet

import scala.quoted.*

/**
 * Compile-time validated sheet name for Sheet.apply.
 *
 * When Sheet("name") is called with a string literal, the name is validated at compile time. When
 * called with a runtime variable, validation is deferred to the OOXML write boundary.
 *
 * Validation rules (Excel sheet name constraints):
 *   - Cannot be empty
 *   - Maximum 31 characters
 *   - Cannot contain: : \ / ? * [ ]
 *
 * Examples:
 *   - `Sheet("Sales")` → compiles, valid name
 *   - `Sheet("Q1:Q2")` → compile error (contains ':')
 *   - `Sheet(userInput)` → compiles, validated at write time
 */
object SheetLiteral:

  /**
   * Create a Sheet with compile-time validation when possible.
   *
   * If the name is a string literal, validates at compile time. If the name is a runtime
   * expression, defers validation to write boundary.
   */
  inline def sheet(inline name: String): Sheet = ${ sheetImpl('name) }

  def sheetImpl(name: Expr[String])(using Quotes): Expr[Sheet] =
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
        // Runtime expression - defer validation to write boundary
        '{ Sheet(SheetName.unsafe($name)) }

end SheetLiteral

export SheetLiteral.*
