package com.tjclp.xl.macros

import com.tjclp.xl.addressing.{ARef, CellRange, RefParser}
import com.tjclp.xl.codec.CellWriter
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle

import scala.quoted.*

/**
 * Compile-time validated string refs for Sheet.put methods.
 *
 * When `sheet.put("A1", value)` is called with a string literal, the reference is validated at
 * compile time and returns `Sheet` directly. When called with a runtime string, validation is
 * deferred and returns `XLResult[Sheet]`.
 *
 * This enables the "if it compiles, it's correct" philosophy while maintaining backwards
 * compatibility for runtime strings.
 */
object PutLiteral:

  /**
   * Macro impl for put(String, A) - validates literal refs at compile time.
   *
   * Returns `Sheet` for literal refs (compile-time validated), `XLResult[Sheet]` for runtime refs.
   */
  def putImpl[A: Type](
    sheet: Expr[Sheet],
    ref: Expr[String],
    value: Expr[A],
    cw: Expr[CellWriter[A]]
  )(using Quotes): Expr[Sheet | XLResult[Sheet]] =
    import quotes.reflect.*

    ref.value match
      case Some(literalRef) =>
        // Compile-time literal - validate now
        RefParser.parse(literalRef) match
          case Right(RefParser.ParsedRef.Cell(None, col0, row0)) =>
            // Valid cell ref - emit direct put call (returns Sheet)
            '{ $sheet.put(ARef.from0(${ Expr(col0) }, ${ Expr(row0) }), $value)(using $cw) }
          case Right(RefParser.ParsedRef.Cell(Some(_), _, _)) =>
            report.errorAndAbort(s"Sheet-qualified refs not supported in put: '$literalRef'")
          case Right(RefParser.ParsedRef.Range(_, _, _, _, _)) =>
            report.errorAndAbort(s"Range refs not supported in put: '$literalRef'")
          case Left(err) =>
            report.errorAndAbort(s"Invalid cell reference '$literalRef': $err")

      case None =>
        // Runtime expression - defer to runtime parsing (returns XLResult[Sheet])
        '{
          ARef.parse($ref) match
            case Left(err) => Left(XLError.InvalidCellRef($ref, err))
            case Right(aref) => Right($sheet.put(aref, $value)(using $cw))
        }

  /**
   * Macro impl for put(String, A, CellStyle) - validates literal refs at compile time.
   *
   * Returns `Sheet` for literal refs (compile-time validated), `XLResult[Sheet]` for runtime refs.
   */
  def putStyledImpl[A: Type](
    sheet: Expr[Sheet],
    ref: Expr[String],
    value: Expr[A],
    style: Expr[CellStyle],
    cw: Expr[CellWriter[A]]
  )(using Quotes): Expr[Sheet | XLResult[Sheet]] =
    import quotes.reflect.*

    ref.value match
      case Some(literalRef) =>
        RefParser.parse(literalRef) match
          case Right(RefParser.ParsedRef.Cell(None, col0, row0)) =>
            '{ $sheet.put(ARef.from0(${ Expr(col0) }, ${ Expr(row0) }), $value, $style)(using $cw) }
          case Right(RefParser.ParsedRef.Cell(Some(_), _, _)) =>
            report.errorAndAbort(s"Sheet-qualified refs not supported in put: '$literalRef'")
          case Right(RefParser.ParsedRef.Range(_, _, _, _, _)) =>
            report.errorAndAbort(s"Range refs not supported in put: '$literalRef'")
          case Left(err) =>
            report.errorAndAbort(s"Invalid cell reference '$literalRef': $err")

      case None =>
        '{
          ARef.parse($ref) match
            case Left(err) => Left(XLError.InvalidCellRef($ref, err))
            case Right(aref) => Right($sheet.put(aref, $value, $style)(using $cw))
        }

  /**
   * Macro impl for style(String, CellStyle) - validates literal refs at compile time.
   *
   * Handles both cell refs ("A1") and range refs ("A1:B10"). Returns `Sheet` for literal refs
   * (compile-time validated), `XLResult[Sheet]` for runtime refs.
   */
  def styleImpl(
    sheet: Expr[Sheet],
    ref: Expr[String],
    style: Expr[CellStyle]
  )(using Quotes): Expr[Sheet | XLResult[Sheet]] =
    import quotes.reflect.*

    // Import extension methods for use in generated code
    import com.tjclp.xl.sheets.styleSyntax.*

    ref.value match
      case Some(literalRef) =>
        RefParser.parse(literalRef) match
          case Right(RefParser.ParsedRef.Cell(None, col0, row0)) =>
            // Valid cell ref - emit direct withCellStyle call (returns Sheet)
            '{
              import com.tjclp.xl.sheets.styleSyntax.*
              $sheet.withCellStyle(ARef.from0(${ Expr(col0) }, ${ Expr(row0) }), $style)
            }
          case Right(RefParser.ParsedRef.Cell(Some(_), _, _)) =>
            report.errorAndAbort(s"Sheet-qualified refs not supported in style: '$literalRef'")
          case Right(RefParser.ParsedRef.Range(None, cs, rs, ce, re)) =>
            // Valid range ref - emit direct withRangeStyle call (returns Sheet)
            '{
              import com.tjclp.xl.sheets.styleSyntax.*
              $sheet.withRangeStyle(
                CellRange(
                  ARef.from0(${ Expr(cs) }, ${ Expr(rs) }),
                  ARef.from0(${ Expr(ce) }, ${ Expr(re) })
                ),
                $style
              )
            }
          case Right(RefParser.ParsedRef.Range(Some(_), _, _, _, _)) =>
            report.errorAndAbort(s"Sheet-qualified refs not supported in style: '$literalRef'")
          case Left(err) =>
            report.errorAndAbort(s"Invalid reference '$literalRef': $err")

      case None =>
        // Runtime expression - defer to runtime parsing (returns XLResult[Sheet])
        '{
          import com.tjclp.xl.sheets.styleSyntax.*
          if $ref.contains(":") then
            CellRange.parse($ref) match
              case Left(err) => Left(XLError.InvalidCellRef($ref, s"Invalid range: $err"))
              case Right(range) => Right($sheet.withRangeStyle(range, $style))
          else
            ARef.parse($ref) match
              case Left(err) => Left(XLError.InvalidCellRef($ref, s"Invalid cell reference: $err"))
              case Right(aref) => Right($sheet.withCellStyle(aref, $style))
        }

  /**
   * Macro impl for merge(String) - validates literal range refs at compile time.
   *
   * Only accepts range refs ("A1:B10"). Returns `Sheet` for literal refs (compile-time validated),
   * `XLResult[Sheet]` for runtime refs.
   */
  def mergeImpl(
    sheet: Expr[Sheet],
    rangeRef: Expr[String]
  )(using Quotes): Expr[Sheet | XLResult[Sheet]] =
    import quotes.reflect.*

    rangeRef.value match
      case Some(literalRef) =>
        RefParser.parse(literalRef) match
          case Right(RefParser.ParsedRef.Range(None, cs, rs, ce, re)) =>
            // Valid range ref - emit direct merge call (returns Sheet)
            '{
              $sheet.merge(
                CellRange(
                  ARef.from0(${ Expr(cs) }, ${ Expr(rs) }),
                  ARef.from0(${ Expr(ce) }, ${ Expr(re) })
                )
              )
            }
          case Right(RefParser.ParsedRef.Range(Some(_), _, _, _, _)) =>
            report.errorAndAbort(s"Sheet-qualified refs not supported in merge: '$literalRef'")
          case Right(RefParser.ParsedRef.Cell(_, _, _)) =>
            report.errorAndAbort(
              s"merge() requires a range like 'A1:B1', got cell ref: '$literalRef'"
            )
          case Left(err) =>
            report.errorAndAbort(s"Invalid range reference '$literalRef': $err")

      case None =>
        // Runtime expression - defer to runtime parsing (returns XLResult[Sheet])
        '{
          CellRange.parse($rangeRef) match
            case Left(err) => Left(XLError.InvalidCellRef($rangeRef, s"Invalid range: $err"))
            case Right(range) => Right($sheet.merge(range))
        }

  /**
   * Macro impl for comment(String, Comment) - validates literal cell refs at compile time.
   *
   * Returns `Sheet` for literal refs (compile-time validated), `XLResult[Sheet]` for runtime refs.
   */
  def commentImpl(
    sheet: Expr[Sheet],
    ref: Expr[String],
    cmt: Expr[com.tjclp.xl.cells.Comment]
  )(using Quotes): Expr[Sheet | XLResult[Sheet]] =
    import quotes.reflect.*

    ref.value match
      case Some(literalRef) =>
        // Compile-time literal - validate now
        RefParser.parse(literalRef) match
          case Right(RefParser.ParsedRef.Cell(None, col0, row0)) =>
            // Valid cell ref - emit direct comment call (returns Sheet)
            '{ $sheet.comment(ARef.from0(${ Expr(col0) }, ${ Expr(row0) }), $cmt) }
          case Right(RefParser.ParsedRef.Cell(Some(_), _, _)) =>
            report.errorAndAbort(s"Sheet-qualified refs not supported in comment: '$literalRef'")
          case Right(RefParser.ParsedRef.Range(_, _, _, _, _)) =>
            report.errorAndAbort(s"Range refs not supported in comment: '$literalRef'")
          case Left(err) =>
            report.errorAndAbort(s"Invalid cell reference '$literalRef': $err")

      case None =>
        // Runtime expression - defer to runtime parsing (returns XLResult[Sheet])
        '{
          ARef.parse($ref) match
            case Left(err) => Left(XLError.InvalidCellRef($ref, err))
            case Right(aref) => Right($sheet.comment(aref, $cmt))
        }

end PutLiteral
