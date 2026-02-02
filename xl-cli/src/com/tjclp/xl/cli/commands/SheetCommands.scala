package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import cats.implicits.*
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.error.XLError
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Sheet management command handlers.
 *
 * Commands for add, remove, rename, move, copy sheet operations. All methods accept a `stream`
 * parameter for O(1) output memory.
 */
object SheetCommands:

  /** Write workbook using standard or streaming writer based on mode */
  private def writeWorkbook(
    wb: Workbook,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean
  ): IO[Unit] =
    val excel = ExcelIO.instance[IO]
    if stream then excel.writeWorkbookStream(wb, outputPath, config)
    else excel.writeWith(wb, outputPath, config)

  /** Build save message suffix based on write mode */
  private def saveSuffix(outputPath: Path, stream: Boolean): String =
    if stream then s"Saved (streaming): $outputPath"
    else s"Saved: $outputPath"

  /**
   * Add a new empty sheet to workbook.
   *
   * @param stream
   *   If true, uses streaming writer for O(1) output memory
   */
  def addSheet(
    wb: Workbook,
    name: String,
    afterOpt: Option[String],
    beforeOpt: Option[String],
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      sheetName <- IO.fromEither(SheetName(name).left.map(e => new Exception(e)))
      _ <- IO
        .raiseError(
          new Exception(
            s"Sheet '$name' already exists. Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
          )
        )
        .whenA(wb.sheets.exists(_.name == sheetName))
      newSheet = Sheet(sheetName)
      updatedWb <- (afterOpt, beforeOpt) match
        case (Some(after), _) =>
          // Insert after specified sheet
          for
            afterName <- IO.fromEither(SheetName(after).left.map(e => new Exception(e)))
            idx = wb.sheets.indexWhere(_.name == afterName)
            _ <- IO
              .raiseError(
                new Exception(
                  s"Sheet '$after' not found. Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
                )
              )
              .whenA(idx < 0)
            result <- IO.fromEither(
              wb.insertAt(idx + 1, newSheet).left.map(e => new Exception(e.message))
            )
          yield result
        case (_, Some(before)) =>
          // Insert before specified sheet
          for
            beforeName <- IO.fromEither(SheetName(before).left.map(e => new Exception(e)))
            idx = wb.sheets.indexWhere(_.name == beforeName)
            _ <- IO
              .raiseError(
                new Exception(
                  s"Sheet '$before' not found. Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
                )
              )
              .whenA(idx < 0)
            result <- IO.fromEither(
              wb.insertAt(idx, newSheet).left.map(e => new Exception(e.message))
            )
          yield result
        case (None, None) =>
          // Append at end (use put for simplicity - adds if not exists)
          IO.pure(wb.put(newSheet))
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
      position = (afterOpt, beforeOpt) match
        case (Some(after), _) => s" (after '$after')"
        case (_, Some(before)) => s" (before '$before')"
        case _ => " (at end)"
    yield s"Added sheet: $name$position\n${saveSuffix(outputPath, stream)}"

  /**
   * Remove sheet from workbook.
   *
   * @param stream
   *   If true, uses streaming writer for O(1) output memory
   */
  def removeSheet(
    wb: Workbook,
    name: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      sheetName <- IO.fromEither(SheetName(name).left.map(e => new Exception(e)))
      updatedWb <- IO.fromEither(wb.remove(sheetName).left.map {
        case XLError.SheetNotFound(_) =>
          new Exception(
            s"Sheet '$name' not found. Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
          )
        case XLError.InvalidWorkbook(reason) => new Exception(reason)
        case e => new Exception(e.message)
      })
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
    yield s"Removed sheet: $name\n${saveSuffix(outputPath, stream)}"

  /**
   * Rename a sheet.
   *
   * @param stream
   *   If true, uses streaming writer for O(1) output memory
   */
  def renameSheet(
    wb: Workbook,
    oldName: String,
    newName: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      oldSheetName <- IO.fromEither(SheetName(oldName).left.map(e => new Exception(e)))
      newSheetName <- IO.fromEither(SheetName(newName).left.map(e => new Exception(e)))
      updatedWb <- IO.fromEither(wb.rename(oldSheetName, newSheetName).left.map {
        case XLError.SheetNotFound(_) =>
          new Exception(
            s"Sheet '$oldName' not found. Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
          )
        case XLError.DuplicateSheet(_) =>
          new Exception(s"Sheet '$newName' already exists")
        case e => new Exception(e.message)
      })
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
    yield s"Renamed: $oldName → $newName\n${saveSuffix(outputPath, stream)}"

  /**
   * Move sheet to new position.
   *
   * @param stream
   *   If true, uses streaming writer for O(1) output memory
   */
  def moveSheet(
    wb: Workbook,
    name: String,
    toIndexOpt: Option[Int],
    afterOpt: Option[String],
    beforeOpt: Option[String],
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      sheetName <- IO.fromEither(SheetName(name).left.map(e => new Exception(e)))
      currentIdx = wb.sheets.indexWhere(_.name == sheetName)
      _ <- IO
        .raiseError(
          new Exception(
            s"Sheet '$name' not found. Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
          )
        )
        .whenA(currentIdx < 0)
      targetIdx <- (toIndexOpt, afterOpt, beforeOpt) match
        case (Some(idx), _, _) => IO.pure(idx)
        case (_, Some(after), _) =>
          for
            afterName <- IO.fromEither(SheetName(after).left.map(e => new Exception(e)))
            afterIdx = wb.sheets.indexWhere(_.name == afterName)
            _ <- IO
              .raiseError(
                new Exception(
                  s"Sheet '$after' not found. Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
                )
              )
              .whenA(afterIdx < 0)
          yield afterIdx + 1
        case (_, _, Some(before)) =>
          for
            beforeName <- IO.fromEither(SheetName(before).left.map(e => new Exception(e)))
            beforeIdx = wb.sheets.indexWhere(_.name == beforeName)
            _ <- IO
              .raiseError(
                new Exception(
                  s"Sheet '$before' not found. Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
                )
              )
              .whenA(beforeIdx < 0)
          yield beforeIdx
        case (None, None, None) =>
          IO.raiseError(
            new Exception("move-sheet requires --to, --after, or --before option")
          )
      // Build new order: remove sheet from current position, insert at target
      currentNames = wb.sheetNames.toVector
      withoutSheet = currentNames.patch(currentIdx, Nil, 1)
      // Adjust target index if we removed from before the target
      adjustedIdx = if currentIdx < targetIdx then targetIdx - 1 else targetIdx
      clampedIdx = adjustedIdx.max(0).min(withoutSheet.size)
      newOrder = withoutSheet.patch(clampedIdx, Vector(sheetName), 0)
      updatedWb <- IO.fromEither(wb.reorder(newOrder).left.map(e => new Exception(e.message)))
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
      position = s"to position $clampedIdx"
    yield s"Moved: $name $position\n${saveSuffix(outputPath, stream)}"

  /**
   * Copy sheet to new name.
   *
   * @param stream
   *   If true, uses streaming writer for O(1) output memory
   */
  def copySheet(
    wb: Workbook,
    sourceName: String,
    targetName: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      sourceSheetName <- IO.fromEither(SheetName(sourceName).left.map(e => new Exception(e)))
      targetSheetName <- IO.fromEither(SheetName(targetName).left.map(e => new Exception(e)))
      sourceSheet <- IO.fromOption(wb.sheets.find(_.name == sourceSheetName))(
        new Exception(
          s"Sheet '$sourceName' not found. Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
        )
      )
      _ <- IO
        .raiseError(new Exception(s"Sheet '$targetName' already exists"))
        .whenA(wb.sheets.exists(_.name == targetSheetName))
      copiedSheet = sourceSheet.copy(name = targetSheetName)
      updatedWb = wb.put(copiedSheet)
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
    yield s"Copied: $sourceName → $targetName\n${saveSuffix(outputPath, stream)}"

  /**
   * Hide a sheet from the sheet tabs.
   *
   * @param veryHide
   *   If true, uses "veryHidden" state (not accessible from Excel UI, only via VBA)
   * @param stream
   *   If true, uses streaming writer for O(1) output memory
   */
  def hideSheet(
    wb: Workbook,
    name: String,
    veryHide: Boolean,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      sheetName <- IO.fromEither(SheetName(name).left.map(e => new Exception(e)))
      state = if veryHide then Some("veryHidden") else Some("hidden")
      updatedWb <- IO.fromEither(wb.setSheetState(sheetName, state).left.map {
        case XLError.SheetNotFound(_) =>
          new Exception(
            s"Sheet '$name' not found. Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
          )
        case XLError.InvalidWorkbook(reason) => new Exception(reason)
        case e => new Exception(e.message)
      })
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
      stateDesc = if veryHide then "very hidden" else "hidden"
    yield s"Sheet '$name' is now $stateDesc\n${saveSuffix(outputPath, stream)}"

  /**
   * Show a hidden sheet (make it visible).
   *
   * @param stream
   *   If true, uses streaming writer for O(1) output memory
   */
  def showSheet(
    wb: Workbook,
    name: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      sheetName <- IO.fromEither(SheetName(name).left.map(e => new Exception(e)))
      updatedWb <- IO.fromEither(wb.setSheetState(sheetName, None).left.map {
        case XLError.SheetNotFound(_) =>
          new Exception(
            s"Sheet '$name' not found. Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
          )
        case e => new Exception(e.message)
      })
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
    yield s"Sheet '$name' is now visible\n${saveSuffix(outputPath, stream)}"
