package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cli.helpers.SheetResolver
import com.tjclp.xl.io.ExcelIO

/**
 * Cell-level command handlers.
 *
 * Commands for merge/unmerge operations.
 */
object CellCommands:

  /**
   * Merge cells in range.
   */
  def merge(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    rangeStr: String,
    outputPath: Path
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, rangeStr, "merge")
      (targetSheet, refOrRange) = resolved
      range <- refOrRange match
        case Right(r) => IO.pure(r)
        case Left(_) =>
          IO.raiseError(
            new Exception("merge requires a range (e.g., A1:C1), not a single cell")
          )
      updatedSheet = targetSheet.merge(range)
      updatedWb = wb.put(updatedSheet)
      _ <- ExcelIO.instance[IO].write(updatedWb, outputPath)
    yield s"Merged: ${range.toA1}\nSaved: $outputPath"

  /**
   * Unmerge cells in range.
   */
  def unmerge(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    rangeStr: String,
    outputPath: Path
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, rangeStr, "unmerge")
      (targetSheet, refOrRange) = resolved
      range <- refOrRange match
        case Right(r) => IO.pure(r)
        case Left(_) =>
          IO.raiseError(
            new Exception("unmerge requires a range (e.g., A1:C1), not a single cell")
          )
      updatedSheet = targetSheet.unmerge(range)
      updatedWb = wb.put(updatedSheet)
      _ <- ExcelIO.instance[IO].write(updatedWb, outputPath)
    yield s"Unmerged: ${range.toA1}\nSaved: $outputPath"
