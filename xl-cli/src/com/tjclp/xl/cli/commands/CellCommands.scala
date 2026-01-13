package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cli.helpers.SheetResolver
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

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
    outputPath: Path,
    config: WriterConfig
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
      _ <- ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config)
    yield s"Merged: ${range.toA1}\nSaved: $outputPath"

  /**
   * Unmerge cells in range.
   */
  def unmerge(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    rangeStr: String,
    outputPath: Path,
    config: WriterConfig
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
      _ <- ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config)
    yield s"Unmerged: ${range.toA1}\nSaved: $outputPath"

  /**
   * Clear cell contents, styles, or comments from range.
   *
   * Modes:
   *   - Default (no flags): Clear contents only (remove cells from range)
   *   - --all: Clear contents, styles, and comments
   *   - --styles: Clear styles only (reset styleId to None)
   *   - --comments: Clear comments only
   *
   * Flags can be combined (e.g., --styles --comments clears both).
   */
  def clear(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    rangeStr: String,
    all: Boolean,
    styles: Boolean,
    comments: Boolean,
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, rangeStr, "clear")
      (targetSheet, refOrRange) = resolved
      range = refOrRange match
        case Right(r) => r
        case Left(ref) => CellRange(ref, ref) // Single cell as 1x1 range

      // Determine what to clear
      // --all means clear everything
      // if only specific flags, clear only those
      // if no flags at all, clear contents (default behavior)
      clearContents = all || (!styles && !comments)
      clearStyles = all || styles
      clearComments = all || comments

      // Apply clearing operations
      sheet1 = if clearContents then targetSheet.removeRange(range) else targetSheet
      sheet2 = if clearStyles then sheet1.clearStylesInRange(range) else sheet1
      sheet3 = if clearComments then sheet2.clearCommentsInRange(range) else sheet2

      // Unmerge any merged regions that overlap with the cleared range
      // This prevents leaving orphaned merge regions that reference cleared cells
      sheet4 =
        if clearContents then
          val overlappingMerges = sheet3.mergedRanges.filter(_.intersects(range))
          overlappingMerges.foldLeft(sheet3)((s, mr) => s.unmerge(mr))
        else sheet3

      updatedWb = wb.put(sheet4)
      _ <- ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config)

      // Build description of what was cleared
      clearedItems = Vector(
        if clearContents then Some("contents") else None,
        if clearStyles then Some("styles") else None,
        if clearComments then Some("comments") else None
      ).flatten
      clearedDesc = clearedItems.mkString(", ")
    yield s"Cleared $clearedDesc from ${range.toA1}\nSaved: $outputPath"
