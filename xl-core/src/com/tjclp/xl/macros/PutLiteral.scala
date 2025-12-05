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

  /**
   * Macro impl for put((String, A)*) - validates all literal refs at compile time.
   *
   * When all refs are string literals, validates them at compile time and returns `Sheet`. When any
   * ref is a runtime expression, falls back to runtime parsing and returns `XLResult[Sheet]`.
   */
  def putTuplesImpl[A: Type](
    sheet: Expr[Sheet],
    updates: Expr[Seq[(String, A)]],
    cw: Expr[CellWriter[A]]
  )(using Quotes): Expr[Sheet | XLResult[Sheet]] =
    import quotes.reflect.*

    // Extract individual tuples from the varargs Seq
    val tupleExprs: List[(Expr[String], Expr[A])] = updates match
      case Varargs(args) =>
        args.toList.map { tupleExpr =>
          tupleExpr match
            case '{ ($ref: String, $value: A) } => (ref, value)
            case '{ ($ref: String) -> ($value: A) } => (ref, value)
            case '{ ArrowAssoc($ref: String).->($value: A) } => (ref, value)
            case other =>
              report.errorAndAbort(
                s"Expected (String, value) tuple, got: ${other.show}"
              )
        }
      case other =>
        report.errorAndAbort(s"Expected varargs, got: ${other.show}")

    // Check if all refs are compile-time literals
    val parsedRefs: List[Either[Expr[String], (Int, Int, Expr[A])]] =
      tupleExprs.map { (refExpr, valueExpr) =>
        refExpr.value match
          case Some(literalRef) =>
            RefParser.parse(literalRef) match
              case Right(RefParser.ParsedRef.Cell(None, col0, row0)) =>
                Right((col0, row0, valueExpr))
              case Right(RefParser.ParsedRef.Cell(Some(_), _, _)) =>
                report.errorAndAbort(
                  s"Sheet-qualified refs not supported in put: '$literalRef'"
                )
              case Right(RefParser.ParsedRef.Range(_, _, _, _, _)) =>
                report.errorAndAbort(s"Range refs not supported in put: '$literalRef'")
              case Left(err) =>
                report.errorAndAbort(s"Invalid cell reference '$literalRef': $err")
          case None =>
            Left(refExpr) // Runtime expression
      }

    // If all refs are literals, emit direct put calls (returns Sheet)
    val allLiterals = parsedRefs.forall(_.isRight)

    if allLiterals then
      // Build chained put calls: sheet.put(aref1, v1).put(aref2, v2)...
      val literals = parsedRefs.collect { case Right(t) => t }
      literals.foldLeft(sheet) { case (accSheet, (col0, row0, valueExpr)) =>
        '{ $accSheet.put(ARef.from0(${ Expr(col0) }, ${ Expr(row0) }), $valueExpr)(using $cw) }
      }
    else
      // Fall back to runtime parsing (returns XLResult[Sheet])
      '{
        val refValuePairs: Seq[(String, A)] = $updates
        refValuePairs.foldLeft[XLResult[Sheet]](Right($sheet)) { (accResult, pair) =>
          accResult.flatMap { s =>
            ARef.parse(pair._1) match
              case Left(err) => Left(XLError.InvalidCellRef(pair._1, err))
              case Right(aref) => Right(s.put(aref, pair._2)(using $cw))
          }
        }
      }

  /**
   * Macro impl for style((String, CellStyle)*) - validates all literal refs at compile time.
   *
   * When all refs are string literals, validates them at compile time and returns `Sheet`. When any
   * ref is a runtime expression, falls back to runtime parsing and returns `XLResult[Sheet]`.
   */
  def styleTuplesImpl(
    sheet: Expr[Sheet],
    updates: Expr[Seq[(String, CellStyle)]]
  )(using Quotes): Expr[Sheet | XLResult[Sheet]] =
    import quotes.reflect.*

    // Import extension methods for use in generated code
    import com.tjclp.xl.sheets.styleSyntax.*

    // Extract individual tuples from the varargs Seq
    val tupleExprs: List[(Expr[String], Expr[CellStyle])] = updates match
      case Varargs(args) =>
        args.toList.map { tupleExpr =>
          tupleExpr match
            case '{ ($ref: String, $style: CellStyle) } => (ref, style)
            case '{ ($ref: String) -> ($style: CellStyle) } => (ref, style)
            case '{ ArrowAssoc($ref: String).->($style: CellStyle) } => (ref, style)
            case other =>
              report.errorAndAbort(
                s"Expected (String, CellStyle) tuple, got: ${other.show}"
              )
        }
      case other =>
        report.errorAndAbort(s"Expected varargs, got: ${other.show}")

    // Parsed refs: Right = literal cell (col0, row0, style), Left = (refExpr, styleExpr) for runtime
    // For ranges we need a different representation
    sealed trait ParsedStyleRef
    case class CellRef(col0: Int, row0: Int, style: Expr[CellStyle]) extends ParsedStyleRef
    case class RangeRef(cs: Int, rs: Int, ce: Int, re: Int, style: Expr[CellStyle])
        extends ParsedStyleRef
    case class RuntimeRef(ref: Expr[String], style: Expr[CellStyle]) extends ParsedStyleRef

    val parsedRefs: List[ParsedStyleRef] =
      tupleExprs.map { (refExpr, styleExpr) =>
        refExpr.value match
          case Some(literalRef) =>
            RefParser.parse(literalRef) match
              case Right(RefParser.ParsedRef.Cell(None, col0, row0)) =>
                CellRef(col0, row0, styleExpr)
              case Right(RefParser.ParsedRef.Cell(Some(_), _, _)) =>
                report.errorAndAbort(
                  s"Sheet-qualified refs not supported in style: '$literalRef'"
                )
              case Right(RefParser.ParsedRef.Range(None, cs, rs, ce, re)) =>
                RangeRef(cs, rs, ce, re, styleExpr)
              case Right(RefParser.ParsedRef.Range(Some(_), _, _, _, _)) =>
                report.errorAndAbort(
                  s"Sheet-qualified refs not supported in style: '$literalRef'"
                )
              case Left(err) =>
                report.errorAndAbort(s"Invalid reference '$literalRef': $err")
          case None =>
            RuntimeRef(refExpr, styleExpr) // Runtime expression
      }

    // If all refs are literals, emit direct style calls (returns Sheet)
    val allLiterals = parsedRefs.forall(!_.isInstanceOf[RuntimeRef])

    if allLiterals then
      // Build chained style calls: sheet.withCellStyle(...).withRangeStyle(...)
      parsedRefs.foldLeft(sheet) {
        case (accSheet, CellRef(col0, row0, styleExpr)) =>
          '{
            import com.tjclp.xl.sheets.styleSyntax.*
            $accSheet.withCellStyle(ARef.from0(${ Expr(col0) }, ${ Expr(row0) }), $styleExpr)
          }
        case (accSheet, RangeRef(cs, rs, ce, re, styleExpr)) =>
          '{
            import com.tjclp.xl.sheets.styleSyntax.*
            $accSheet.withRangeStyle(
              CellRange(
                ARef.from0(${ Expr(cs) }, ${ Expr(rs) }),
                ARef.from0(${ Expr(ce) }, ${ Expr(re) })
              ),
              $styleExpr
            )
          }
        case (accSheet, _: RuntimeRef) =>
          accSheet // Won't happen due to allLiterals check
      }
    else
      // Fall back to runtime parsing (returns XLResult[Sheet])
      '{
        import com.tjclp.xl.sheets.styleSyntax.*
        val refStylePairs: Seq[(String, CellStyle)] = $updates
        refStylePairs.foldLeft[XLResult[Sheet]](Right($sheet)) { (accResult, pair) =>
          accResult.flatMap { s =>
            val ref = pair._1
            val style = pair._2
            if ref.contains(":") then
              CellRange.parse(ref) match
                case Left(err) => Left(XLError.InvalidCellRef(ref, s"Invalid range: $err"))
                case Right(range) => Right(s.withRangeStyle(range, style))
            else
              ARef.parse(ref) match
                case Left(err) => Left(XLError.InvalidCellRef(ref, s"Invalid cell reference: $err"))
                case Right(aref) => Right(s.withCellStyle(aref, style))
          }
        }
      }

end PutLiteral
