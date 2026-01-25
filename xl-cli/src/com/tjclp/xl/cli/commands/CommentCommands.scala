package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.Comment
import com.tjclp.xl.cli.helpers.SheetResolver
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Command handlers for cell comment operations.
 *
 * Supports adding and removing comments from cells. All methods accept a `stream` parameter for
 * O(1) output memory.
 */
object CommentCommands:

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
   * Add or update comment on a cell.
   *
   * @param wb
   *   Workbook containing the sheet
   * @param sheetOpt
   *   Target sheet (required)
   * @param refStr
   *   Cell reference (e.g., "A1", "B5")
   * @param text
   *   Comment text
   * @param author
   *   Optional author name
   * @param outputPath
   *   Output file path
   * @param config
   *   Writer configuration
   * @param stream
   *   If true, uses streaming writer for O(1) output memory
   * @return
   *   Success message
   */
  def addComment(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    refStr: String,
    text: String,
    author: Option[String],
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, refStr, "comment")
      (targetSheet, refOrRange) = resolved
      ref <- refOrRange match
        case Left(r) => IO.pure(r)
        case Right(_) =>
          IO.raiseError(new Exception("comment command requires single cell, not range"))

      // Create comment
      comment = Comment.plainText(text, author)

      // Add comment to sheet
      updatedSheet = targetSheet.comment(ref, comment)
      updatedWb = wb.put(updatedSheet)

      // Write to output
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)

      authorStr = author.map(a => s" (author: $a)").getOrElse("")
    yield s"""Added comment to ${ref.toA1}: "$text"$authorStr
${saveSuffix(outputPath, stream)}"""

  /**
   * Remove comment from a cell.
   *
   * @param wb
   *   Workbook containing the sheet
   * @param sheetOpt
   *   Target sheet (required)
   * @param refStr
   *   Cell reference
   * @param outputPath
   *   Output file path
   * @param config
   *   Writer configuration
   * @param stream
   *   If true, uses streaming writer for O(1) output memory
   * @return
   *   Success message
   */
  def removeComment(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    refStr: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, refStr, "remove-comment")
      (targetSheet, refOrRange) = resolved
      ref <- refOrRange match
        case Left(r) => IO.pure(r)
        case Right(_) =>
          IO.raiseError(new Exception("remove-comment command requires single cell, not range"))

      // Check if comment exists
      hadComment = targetSheet.hasComment(ref)

      // Remove comment from sheet
      updatedSheet = targetSheet.removeComment(ref)
      updatedWb = wb.put(updatedSheet)

      // Write to output
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)

      status = if hadComment then "Removed comment from" else "No comment found at"
    yield s"""$status ${ref.toA1}
${saveSuffix(outputPath, stream)}"""
