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
   *
   * Performance: Uses batch put internally (O(N) single Sheet copy) rather than chained single-item
   * puts (O(N²) intermediate copies).
   *
   * Duplicate cell references are detected and rejected:
   *   - For literals: compile-time error
   *   - For runtime strings: returns `Left(XLError.DuplicateCellRef(...))`
   *
   * Note: For runtime strings, validation is fail-fast. If the 3rd ref is invalid, refs 4+ are not
   * validated.
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

    // Check for duplicate literal refs at compile time
    val literalRefs = parsedRefs.collect { case Right((col, row, _)) => (col, row) }
    val duplicates = literalRefs.groupBy(identity).filter(_._2.size > 1).keys.toList
    if duplicates.nonEmpty then
      val dupStr = duplicates.map { case (c, r) => ARef.from0(c, r).toA1 }.mkString(", ")
      report.errorAndAbort(s"Duplicate cell references not allowed: $dupStr")

    // If all refs are literals, emit direct put calls (returns Sheet)
    val allLiterals = parsedRefs.forall(_.isRight)

    if allLiterals then
      // Build single batch put call: sheet.put((aref1, v1), (aref2, v2), ...)
      // This is O(N) vs O(N²) for chained single puts
      val literals = parsedRefs.collect { case Right(t) => t }
      val pairExprs = literals.map { case (col0, row0, valueExpr) =>
        '{ (ARef.from0(${ Expr(col0) }, ${ Expr(row0) }), $valueExpr) }
      }
      '{ $sheet.put(${ Varargs(pairExprs) }*)(using $cw) }
    else
      // Fall back to runtime parsing (returns XLResult[Sheet])
      '{
        val refValuePairs: Seq[(String, A)] = $updates
        // Check for duplicate refs at runtime
        val refStrings = refValuePairs.map(_._1)
        val duplicates = refStrings.groupBy(identity).filter(_._2.size > 1).keys.toList
        if duplicates.nonEmpty then Left(XLError.DuplicateCellRef(duplicates.mkString(", ")))
        else
          // Parse all refs first, then batch put if all valid
          val parsed = refValuePairs.map { case (refStr, value) =>
            ARef.parse(refStr).map(aref => (aref, value))
          }
          parsed.collectFirst { case Left(err) => err } match
            case Some(err) =>
              val badRef = refValuePairs(parsed.indexWhere(_.isLeft))._1
              Left(XLError.InvalidCellRef(badRef, err))
            case None =>
              val validPairs = parsed.collect { case Right(pair) => pair }
              Right($sheet.put(validPairs*)(using $cw))
      }

  /**
   * Macro impl for style((String, CellStyle)*) - validates all literal refs at compile time.
   *
   * When all refs are string literals, validates them at compile time and returns `Sheet`. When any
   * ref is a runtime expression, falls back to runtime parsing and returns `XLResult[Sheet]`.
   *
   * Duplicate references (same string) are detected and rejected:
   *   - For literals: compile-time error
   *   - For runtime strings: returns `Left(XLError.DuplicateCellRef(...))`
   *
   * Note: For runtime strings, validation is fail-fast. If the 3rd ref is invalid, refs 4+ are not
   * validated.
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

    // Check for duplicate literal refs at compile time (by original string)
    val literalRefStrings = tupleExprs.collect { (refExpr, _) =>
      refExpr.value
    }.flatten
    val duplicates = literalRefStrings.groupBy(identity).filter(_._2.size > 1).keys.toList
    if duplicates.nonEmpty then
      report.errorAndAbort(s"Duplicate references not allowed: ${duplicates.mkString(", ")}")

    // If all refs are literals, emit direct style calls (returns Sheet)
    val allLiterals = parsedRefs.forall {
      case _: RuntimeRef => false
      case _ => true
    }

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
        // Check for duplicate refs at runtime
        val refStrings = refStylePairs.map(_._1)
        val duplicates = refStrings.groupBy(identity).filter(_._2.size > 1).keys.toList
        if duplicates.nonEmpty then Left(XLError.DuplicateCellRef(duplicates.mkString(", ")))
        else
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
                  case Left(err) =>
                    Left(XLError.InvalidCellRef(ref, s"Invalid cell reference: $err"))
                  case Right(aref) => Right(s.withCellStyle(aref, style))
            }
          }
      }

end PutLiteral
