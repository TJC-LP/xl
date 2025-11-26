package com.tjclp.xl.macros

import com.tjclp.xl.addressing.{ARef, CellRange, RefParser}
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.sheets.Sheet

import scala.quoted.*

/**
 * Compile-time validated string refs for cell lookup methods.
 *
 * When `sheet.cell("A1")` is called with a string literal, the reference is validated at compile
 * time. Invalid literals like `"INVALID"` fail to compile. Runtime strings parse at runtime and
 * return None/empty for invalid refs.
 *
 * This enables the "if it compiles, it's correct" philosophy for read operations.
 */
object CellLookupMacros:

  /**
   * Macro impl for cell(String) - validates literal cell refs at compile time.
   *
   * Returns `Option[Cell]` in all cases:
   *   - Literal refs: compile-time validated, runtime lookup
   *   - Runtime strings: runtime parsing, None if invalid
   */
  def cellImpl(
    sheet: Expr[Sheet],
    cellRef: Expr[String]
  )(using Quotes): Expr[Option[Cell]] =
    import quotes.reflect.*

    cellRef.value match
      case Some(literal) =>
        // Compile-time literal - validate now
        RefParser.parse(literal) match
          case Right(RefParser.ParsedRef.Cell(None, col0, row0)) =>
            // Valid cell ref - emit direct lookup
            '{ $sheet.cells.get(ARef.from0(${ Expr(col0) }, ${ Expr(row0) })) }
          case Right(RefParser.ParsedRef.Cell(Some(_), _, _)) =>
            report.errorAndAbort(s"Sheet-qualified refs not supported in cell lookup: '$literal'")
          case Right(RefParser.ParsedRef.Range(_, _, _, _, _)) =>
            report.errorAndAbort(
              s"Range ref not supported in cell() - use range() instead: '$literal'"
            )
          case Left(err) =>
            report.errorAndAbort(s"Invalid cell reference '$literal': $err")

      case None =>
        // Runtime expression - defer to runtime parsing
        '{ ARef.parse($cellRef).toOption.flatMap(ref => $sheet.cells.get(ref)) }

  /**
   * Macro impl for range(String) - validates literal range refs at compile time.
   *
   * Returns `List[Cell]` in all cases:
   *   - Literal refs: compile-time validated, runtime lookup
   *   - Runtime strings: runtime parsing, empty list if invalid
   */
  def rangeImpl(
    sheet: Expr[Sheet],
    rangeRef: Expr[String]
  )(using Quotes): Expr[List[Cell]] =
    import quotes.reflect.*

    rangeRef.value match
      case Some(literal) =>
        // Compile-time literal - validate now
        RefParser.parse(literal) match
          case Right(RefParser.ParsedRef.Range(None, cs, rs, ce, re)) =>
            // Valid range ref - emit direct lookup
            '{
              val range = CellRange(
                ARef.from0(${ Expr(cs) }, ${ Expr(rs) }),
                ARef.from0(${ Expr(ce) }, ${ Expr(re) })
              )
              range.cells.flatMap($sheet.cells.get).toList
            }
          case Right(RefParser.ParsedRef.Range(Some(_), _, _, _, _)) =>
            report.errorAndAbort(s"Sheet-qualified refs not supported in range lookup: '$literal'")
          case Right(RefParser.ParsedRef.Cell(_, _, _)) =>
            report.errorAndAbort(
              s"Cell ref not supported in range() - use cell() instead: '$literal'"
            )
          case Left(err) =>
            report.errorAndAbort(s"Invalid range reference '$literal': $err")

      case None =>
        // Runtime expression - defer to runtime parsing
        '{
          CellRange
            .parse($rangeRef)
            .toOption
            .map(r => r.cells.flatMap($sheet.cells.get).toList)
            .getOrElse(List.empty)
        }

  /**
   * Macro impl for get(String) - auto-detects cell vs range at compile time.
   *
   * Returns `List[Cell]` in all cases:
   *   - Literal cell ref: validated, returns single-element list or empty
   *   - Literal range ref: validated, returns cells in range
   *   - Runtime string: auto-detects at runtime
   */
  def getImpl(
    sheet: Expr[Sheet],
    ref: Expr[String]
  )(using Quotes): Expr[List[Cell]] =
    import quotes.reflect.*

    ref.value match
      case Some(literal) =>
        // Compile-time literal - detect type and validate
        RefParser.parse(literal) match
          case Right(RefParser.ParsedRef.Cell(None, col0, row0)) =>
            // Valid cell ref - return as single-element list
            '{ $sheet.cells.get(ARef.from0(${ Expr(col0) }, ${ Expr(row0) })).toList }
          case Right(RefParser.ParsedRef.Cell(Some(_), _, _)) =>
            report.errorAndAbort(s"Sheet-qualified refs not supported in get(): '$literal'")
          case Right(RefParser.ParsedRef.Range(None, cs, rs, ce, re)) =>
            // Valid range ref - return cells in range
            '{
              val range = CellRange(
                ARef.from0(${ Expr(cs) }, ${ Expr(rs) }),
                ARef.from0(${ Expr(ce) }, ${ Expr(re) })
              )
              range.cells.flatMap($sheet.cells.get).toList
            }
          case Right(RefParser.ParsedRef.Range(Some(_), _, _, _, _)) =>
            report.errorAndAbort(s"Sheet-qualified refs not supported in get(): '$literal'")
          case Left(err) =>
            report.errorAndAbort(s"Invalid reference '$literal': $err")

      case None =>
        // Runtime expression - auto-detect at runtime
        '{
          if $ref.contains(":") then
            CellRange
              .parse($ref)
              .toOption
              .map(r => r.cells.flatMap($sheet.cells.get).toList)
              .getOrElse(List.empty)
          else ARef.parse($ref).toOption.flatMap(r => $sheet.cells.get(r)).toList
        }

end CellLookupMacros
