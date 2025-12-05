package com.tjclp.xl.cli.helpers

import cats.effect.IO
import com.tjclp.xl.{*, given}

/**
 * Batch operation parsing and execution for CLI.
 *
 * Handles parsing JSON batch input and applying operations to workbooks.
 */
object BatchParser:

  /**
   * Batch operation ADT.
   */
  enum BatchOp:
    case Put(ref: String, value: String)
    case PutFormula(ref: String, formula: String)

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
    else IO.blocking(scala.io.Source.fromFile(source).mkString)

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
   * Simple JSON parser for batch operations.
   *
   * Handles: [{"op":"put"|"putf", "ref":"A1", "value":"..."}]
   */
  def parseBatchJson(json: String): Either[Exception, Vector[BatchOp]] =
    // Very simple JSON parsing - handles the specific format we need
    val objPattern = """\{[^}]+\}""".r
    val opPattern = """"op"\s*:\s*"(\w+)"""".r
    val refPattern = """"ref"\s*:\s*"([^"]+)"""".r
    val valuePattern = """"value"\s*:\s*"((?:[^"\\]|\\.)*)"""".r

    val ops = objPattern.findAllIn(json).toVector.map { obj =>
      val op = opPattern.findFirstMatchIn(obj).map(_.group(1))
      val ref = refPattern.findFirstMatchIn(obj).map(_.group(1))
      val value = valuePattern
        .findFirstMatchIn(obj)
        .map(_.group(1))
        .map(_.replace("\\\"", "\"").replace("\\n", "\n").replace("\\\\", "\\"))

      (op, ref, value) match
        case (Some("put"), Some(r), Some(v)) => Right(BatchOp.Put(r, v))
        case (Some("putf"), Some(r), Some(v)) => Right(BatchOp.PutFormula(r, v))
        case (Some(unknown), _, _) => Left(new Exception(s"Unknown operation: $unknown"))
        case _ => Left(new Exception(s"Invalid batch operation: $obj"))
    }

    val errors = ops.collect { case Left(e) => e }
    errors.headOption match
      case Some(e) => Left(e)
      case None => Right(ops.collect { case Right(op) => op })

  /**
   * Apply batch operations to a workbook.
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
    ops.foldLeft(IO.pure(wb)) { (wbIO, op) =>
      wbIO.flatMap { currentWb =>
        op match
          case BatchOp.Put(refStr, value) =>
            SheetResolver.resolveRef(currentWb, defaultSheetOpt, refStr, "batch put").map {
              case (targetSheet, Left(ref)) =>
                val cellValue = ValueParser.parseValue(value)
                val updatedSheet = targetSheet.put(ref, cellValue)
                currentWb.put(updatedSheet)
              case (_, Right(_)) =>
                // This shouldn't happen for single refs, but handle gracefully
                currentWb
            }
          case BatchOp.PutFormula(refStr, formula) =>
            SheetResolver.resolveRef(currentWb, defaultSheetOpt, refStr, "batch putf").map {
              case (targetSheet, Left(ref)) =>
                val normalizedFormula = if formula.startsWith("=") then formula.drop(1) else formula
                val updatedSheet = targetSheet.put(ref, CellValue.Formula(normalizedFormula, None))
                currentWb.put(updatedSheet)
              case (_, Right(_)) =>
                currentWb
            }
      }
    }
