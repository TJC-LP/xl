package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cli.contract.Warning
import com.tjclp.xl.cli.helpers.SheetResolver
import com.tjclp.xl.cli.MemoryGuard
import com.tjclp.xl.cli.output.Format
import com.tjclp.xl.formula.eval.DynamicArrayAuthoring
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Cell-level command handlers.
 *
 * Commands for merge/unmerge operations. All methods accept a `stream` parameter to use the
 * SAX/StAX workbook writer.
 */
object CellCommands:

  /** Write workbook using the standard or SAX/StAX backend based on mode */
  private def writeWorkbook(
    wb: Workbook,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean
  ): IO[Unit] =
    val excel = MemoryGuard.writer
    if stream then excel.writeWorkbookStream(wb, outputPath, config)
    else excel.writeWith(wb, outputPath, config)

  /**
   * Merge cells in range.
   *
   * @param stream
   *   If true, uses the SAX/StAX workbook writer
   */
  def merge(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    rangeStr: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
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
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
    yield s"Merged: ${range.toA1}\n${Format.saveSuffix(outputPath, stream)}"

  /**
   * Unmerge cells in range.
   *
   * @param stream
   *   If true, uses the SAX/StAX workbook writer
   */
  def unmerge(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    rangeStr: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
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
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
    yield s"Unmerged: ${range.toA1}\n${Format.saveSuffix(outputPath, stream)}"

  /**
   * Clear cell contents, styles, or comments from range.
   *
   * Modes:
   *   - Default (no flags): Clear contents only, as Excel's Clear Contents: values, formulas and
   *     hyperlinks go; number formats, fonts, fills, borders and merges stay (GH-886)
   *   - --all: Clear contents, styles, and comments; removes the cells and unmerges any merged
   *     region the range touches
   *   - --styles: Clear styles only (reset styleId to None) and unmerge any merged region the range
   *     touches, as Excel's Clear Formats (GH-894)
   *   - --comments: Clear comments only
   *
   * Flags can be combined (e.g., --styles --comments clears both).
   *
   * @param stream
   *   If true, uses the SAX/StAX workbook writer
   */
  def clear(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    rangeStr: String,
    all: Boolean,
    styles: Boolean,
    comments: Boolean,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false,
    warn: Warning => IO[Unit] = _ => IO.unit
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

      // GH-714: clearing a dynamic anchor clears its whole spill, as in Excel; the cells of a live
      // spill whose anchor stays are left as they were (Excel keeps them) and the rest is cleared,
      // with SPILL_CELLS_SKIPPED. Styles and comments do not touch a spill: they clear everywhere.
      // GH-886: a contents clear keeps each cell's formatting, as Excel's Clear Contents does;
      // only --all drops the cell records.
      contents = Option.when(clearContents)(
        DynamicArrayAuthoring.clearContents(
          targetSheet,
          DynamicArrayAuthoring.spillsOf(targetSheet),
          Vector(range),
          keepFormatting = !all
        )
      )

      // Apply clearing operations
      sheet1 = contents.fold(targetSheet)(_.sheet)
      sheet2 = if clearStyles then sheet1.clearStylesInRange(range) else sheet1
      sheet3 = if clearComments then sheet2.clearCommentsInRange(range) else sheet2

      // GH-886/GH-894: a merge is formatting. A contents or comments clear keeps it (Excel's
      // Clear Contents does); --styles and --all unmerge every merged region the range touches,
      // as Excel's Clear Formats and Clear All do.
      sheet4 =
        if clearStyles then
          val overlappingMerges = sheet3.mergedRanges.filter(_.intersects(range))
          overlappingMerges.foldLeft(sheet3)((s, mr) => s.unmerge(mr))
        else sheet3

      updatedWb = wb.put(sheet4)
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
      kept = contents.fold(Vector.empty)(_.kept).map(targetSheet.name -> _)
      _ <- WriteCommands.spillsKeptWarning(kept).fold(IO.unit)(warn)

      // Build description of what was cleared
      clearedItems = Vector(
        if clearContents then Some("contents") else None,
        if clearStyles then Some("styles") else None,
        if clearComments then Some("comments") else None
      ).flatten
      clearedDesc = clearedItems.mkString(", ")
    yield s"Cleared $clearedDesc from ${range.toA1}\n${Format.saveSuffix(outputPath, stream)}"
