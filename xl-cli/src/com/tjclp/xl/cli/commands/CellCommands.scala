package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cli.contract.{CliError, CliException, Warning}
import com.tjclp.xl.cli.helpers.SheetResolver
import com.tjclp.xl.cli.{MemoryGuard, WritePolicy}
import com.tjclp.xl.cli.output.Format
import com.tjclp.xl.formula.eval.{CellMover, DynamicArrayAuthoring}
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.sheets.SheetEdits

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
   * Merge cells in range, as Excel's Merge Cells does (GH-935): a non-empty upper-left cell keeps
   * its value; an empty one receives the first non-empty cell of the range (row-major), cut and
   * pasted — formula verbatim, format, and every reader pointed at the upper-left cell
   * ([[CellMover.merge]]). Every other cell is cleared like a contents `clear` (value, formula and
   * hyperlink go; its own formatting and comment stay). A hidden dynamic-array anchor releases its
   * spill first, as `clear` does. The formulas that read a cleared cell are recalculated before the
   * write (`--no-recalc` keeps every cache), so `SUM` over the merged area sees only the anchor.
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
    stream: Boolean = false,
    policy: WritePolicy = WritePolicy.default,
    warn: Warning => IO[Unit] = _ => IO.unit
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
      // GH-935: an empty upper-left cell receives the first non-empty cell, its readers following
      merged <- IO.fromEither(
        CellMover
          .merge(wb, targetSheet.name, range)
          .left
          .map(e => CliException(CliError.fromXLError(e, None)))
      )
      updatedSheet <- IO.fromEither(
        merged(targetSheet.name).left.map(e => CliException(CliError.fromXLError(e, None)))
      )
      // the seeds of the dependent refresh: every cell whose content the merge changed
      // GH-935 review: a dynamic array whose spill range the merge now covers is blocked (#SPILL!),
      // so its anchor is re-evaluated with the cells the merge changed
      blockedAnchors = updatedSheet.cells.iterator.collect {
        case Cell(ref, CellValue.Formula(_, _, k: FormulaKind.ArrayFormula), _, _)
            if k.mode != ArrayMode.Legacy &&
              updatedSheet.mergedRanges.exists(_.intersects(k.ref)) =>
          ref
      }.toSet
      modified = (targetSheet.cells.refs.toSet ++ updatedSheet.cells.refs.toSet).filter { r =>
        targetSheet.cells.get(r).map(_.value) != updatedSheet.cells.get(r).map(_.value)
      } ++ blockedAnchors
      // the range actually merged: grown over the merged areas it touched (absorbed)
      message = s"Merged: ${SheetEdits.mergeTarget(targetSheet, range).toA1}"
      // A merge that hides nothing (the usual title merge) changes no value: no refresh to run
      result <-
        if modified.isEmpty then
          writeWorkbook(merged, outputPath, config, stream)
            .as(s"$message\n${Format.saveSuffix(outputPath, stream)}")
        else
          WriteCommands.writeAfterRefresh(
            merged,
            targetSheet.name,
            modified,
            message,
            outputPath,
            config,
            stream,
            policy,
            warn
          )
    yield result

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
    warn: Warning => IO[Unit] = _ => IO.unit,
    policy: WritePolicy = WritePolicy.default
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
      // GH-886/GH-894: a merge is formatting. A contents or comments clear keeps it (Excel's
      // Clear Contents does); --styles and --all unmerge every merged region the range touches,
      // as Excel's Clear Formats and Clear All do, and reset the styles of the whole merge (Excel
      // expands a selection to the merges it touches), so no unmerged cell keeps the old format.
      overlappingMerges = sheet1.mergedRanges.filter(_.intersects(range)).toVector
      sheet2 =
        if clearStyles then
          // A merge inside the range is already covered; only those crossing its edge need a pass
          val crossing =
            overlappingMerges.filterNot(m => range.contains(m.start) && range.contains(m.end))
          (range +: crossing).foldLeft(sheet1)((s, r) => s.clearStylesInRange(r))
        else sheet1
      sheet3 = if clearComments then sheet2.clearCommentsInRange(range) else sheet2
      sheet4 =
        if clearStyles then overlappingMerges.foldLeft(sheet3)((s, mr) => s.unmerge(mr))
        else sheet3

      updatedWb = wb.put(sheet4)

      // Build description of what was cleared
      clearedItems = Vector(
        if clearContents then Some("contents") else None,
        if clearStyles then Some("styles") else None,
        if clearComments then Some("comments") else None
      ).flatten
      clearedDesc = clearedItems.mkString(", ")
      message = s"Cleared $clearedDesc from ${range.toA1}"

      // Found during #935: a contents clear changes what formulas read, so their caches are
      // refreshed before the write like any other content edit (--no-recalc keeps them). The
      // seeds are the range's cells and the released spill cells whose value changed.
      candidates = targetSheet.cells.inRange(range).map(_.ref).toSet ++
        contents.fold(Set.empty[ARef])(_.released)
      modified = candidates.filter(r =>
        targetSheet.cells.get(r).map(_.value) != sheet4.cells.get(r).map(_.value)
      )
      result <-
        if modified.isEmpty then
          writeWorkbook(updatedWb, outputPath, config, stream)
            .as(s"$message\n${Format.saveSuffix(outputPath, stream)}")
        else
          WriteCommands.writeAfterRefresh(
            updatedWb,
            targetSheet.name,
            modified,
            message,
            outputPath,
            config,
            stream,
            policy,
            warn
          )
      kept = contents.fold(Vector.empty)(_.kept).map(targetSheet.name -> _)
      _ <- WriteCommands.spillsKeptWarning(kept).fold(IO.unit)(warn)
    yield result
