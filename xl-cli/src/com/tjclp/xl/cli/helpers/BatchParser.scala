package com.tjclp.xl.cli.helpers

import cats.effect.{IO, Resource}
import cats.syntax.traverse.*
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, RefType, SheetName}
import com.tjclp.xl.cells.CellValue
import upickle.default.*

/**
 * Batch operation parsing and execution for CLI.
 *
 * Handles parsing JSON batch input and applying operations to workbooks. Uses uPickle for robust
 * JSON parsing that handles all edge cases (nested braces, escaping, unicode).
 */
object BatchParser:

  /**
   * Batch operation ADT.
   */
  enum BatchOp:
    case Put(ref: String, value: String)
    case PutFormula(ref: String, formula: String)

  /** JSON representation for parsing */
  private case class BatchOpJson(op: String, ref: String, value: ujson.Value) derives ReadWriter

  /**
   * Read batch input from file or stdin.
   *
   * @param source
   *   File path or "-" for stdin
   * @return
   *   IO containing input string
   */
  def readBatchInput(source: String): IO[String] =
    if source == "-" then IO.blocking(scala.io.Source.stdin.mkString)
    else
      Resource
        .fromAutoCloseable(IO.blocking(scala.io.Source.fromFile(source)))
        .use(src => IO.blocking(src.mkString))

  /**
   * Parse batch JSON input. Expects format:
   * {{{
   * [
   *   {"op": "put", "ref": "A1", "value": "Hello"},
   *   {"op": "putf", "ref": "B1", "value": "=A1*2"}
   * ]
   * }}}
   *
   * @param input
   *   JSON string
   * @return
   *   IO containing parsed operations
   */
  def parseBatchOperations(input: String): IO[Vector[BatchOp]] =
    IO.fromEither {
      val trimmed = input.trim
      if !trimmed.startsWith("[") then Left(new Exception("Batch input must be a JSON array"))
      else parseBatchJson(trimmed)
    }

  /**
   * Parse batch JSON using uPickle.
   *
   * Handles all JSON edge cases: nested braces, escaping, unicode, multi-line values. Format:
   * [{"op":"put"|"putf", "ref":"A1", "value": ...}]
   *
   * Values can be strings, numbers, booleans, or null - all converted to string for cell values.
   */
  def parseBatchJson(json: String): Either[Exception, Vector[BatchOp]] =
    try
      val parsed = ujson.read(json)
      val arr = parsed.arrOpt.getOrElse(
        throw new Exception("Batch input must be a JSON array")
      )

      val ops = arr.value.toVector.zipWithIndex.map { case (obj, idx) =>
        val objMap = obj.objOpt.getOrElse(
          throw new Exception(
            s"Object ${idx + 1}: Expected JSON object, got ${obj.getClass.getSimpleName}"
          )
        )

        val op = objMap
          .get("op")
          .flatMap(_.strOpt)
          .getOrElse(
            throw new Exception(
              s"Object ${idx + 1}: Missing or invalid 'op' field. Expected \"op\": \"put\" or \"op\": \"putf\""
            )
          )

        val ref = objMap
          .get("ref")
          .flatMap(_.strOpt)
          .getOrElse(
            throw new Exception(
              s"Object ${idx + 1}: Missing or invalid 'ref' field. Expected \"ref\": \"A1\""
            )
          )

        val value = objMap
          .get("value")
          .map {
            case v if v.strOpt.isDefined => v.str
            case v if v.numOpt.isDefined => v.num.toString
            case v if v.boolOpt.isDefined => v.bool.toString
            case v if v.isNull => ""
            case v => throw new Exception(s"Object ${idx + 1}: Unsupported value type for 'value'")
          }
          .getOrElse(
            throw new Exception(
              s"Object ${idx + 1}: Missing 'value' field. Expected \"value\": \"text\" or \"value\": 123"
            )
          )

        op match
          case "put" => BatchOp.Put(ref, value)
          case "putf" => BatchOp.PutFormula(ref, value)
          case other =>
            throw new Exception(s"Object ${idx + 1}: Unknown operation '$other'. Valid: put, putf")
      }

      Right(ops)
    catch
      case e: ujson.ParseException =>
        Left(new Exception(s"JSON parse error: ${e.getMessage}"))
      case e: Exception =>
        Left(e)

  /**
   * Resolved operation with parsed ref and target sheet name.
   */
  private case class ResolvedOp(sheetName: SheetName, ref: ARef, value: CellValue)

  /**
   * Apply batch operations to a workbook.
   *
   * Optimized: Groups operations by sheet and uses batch put for O(N) instead of O(N²). Uses
   * Sheet.put((ARef, CellValue)*) varargs for single-pass accumulation with style deduplication.
   *
   * @param wb
   *   Workbook to modify
   * @param defaultSheetOpt
   *   Default sheet for unqualified refs
   * @param ops
   *   Operations to apply
   * @return
   *   IO containing updated workbook
   */
  def applyBatchOperations(
    wb: Workbook,
    defaultSheetOpt: Option[Sheet],
    ops: Vector[BatchOp]
  ): IO[Workbook] =
    val defaultSheetName = defaultSheetOpt.map(_.name)

    // Phase 1: Resolve all refs upfront, validating and extracting sheet names
    val resolvedOpsIO: IO[Vector[ResolvedOp]] = ops.traverse { op =>
      val (refStr, value) = op match
        case BatchOp.Put(r, v) =>
          (r, ValueParser.parseValue(v))
        case BatchOp.PutFormula(r, f) =>
          val formula = if f.startsWith("=") then f.drop(1) else f
          (r, CellValue.Formula(formula, None))

      IO.fromEither(RefType.parse(refStr).left.map(e => new Exception(e))).flatMap {
        case RefType.Cell(ref) =>
          // Unqualified ref - use default sheet
          defaultSheetName match
            case Some(name) => IO.pure(ResolvedOp(name, ref, value))
            case None =>
              IO.raiseError(new Exception(s"batch requires --sheet for unqualified ref '$refStr'"))
        case RefType.QualifiedCell(sheetName, ref) =>
          // Qualified ref - use specified sheet
          IO.pure(ResolvedOp(sheetName, ref, value))
        case RefType.Range(_) | RefType.QualifiedRange(_, _) =>
          IO.raiseError(new Exception(s"batch put requires single cell ref, not range: $refStr"))
      }
    }

    resolvedOpsIO.flatMap { resolvedOps =>
      // Phase 2: Group by sheet name
      val bySheet: Map[SheetName, Vector[ResolvedOp]] = resolvedOps.groupBy(_.sheetName)

      // Phase 3: Apply batch puts per sheet (O(1) workbook update per sheet)
      bySheet.foldLeft(IO.pure(wb)) { case (wbIO, (sheetName, sheetOps)) =>
        wbIO.flatMap { currentWb =>
          currentWb.sheets.find(_.name == sheetName) match
            case None =>
              IO.raiseError(
                new Exception(
                  s"Sheet '${sheetName.value}' not found. " +
                    s"Available: ${currentWb.sheetNames.map(_.value).mkString(", ")}"
                )
              )
            case Some(sheet) =>
              // Build (ARef, CellValue)* for batch put
              val updates: Seq[(ARef, CellValue)] = sheetOps.map(op => (op.ref, op.value))
              // Single batch put - O(N) with style deduplication
              val updatedSheet = sheet.put(updates*)
              IO.pure(currentWb.put(updatedSheet))
        }
      }
    }
