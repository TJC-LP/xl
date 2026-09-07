package com.tjclp.xl.cli.read

import java.nio.file.{Files, Path}

import cats.effect.IO

import com.tjclp.xl.Workbook
import com.tjclp.xl.cli.{FilterFormat, ViewFormat}
import com.tjclp.xl.cli.contract.{Outcome, OutputMode, Payload, Render}
import com.tjclp.xl.io.ExcelIO

/** Runs read queries over both strategies and reads their outcomes back as text. */
object ReadTestKit:

  val excel: ExcelIO[IO] = ExcelIO.instance[IO]

  def view(
    range: Option[String],
    format: ViewFormat = ViewFormat.Markdown,
    showFormulas: Boolean = false,
    evalFormulas: Boolean = false,
    strict: Boolean = false,
    limit: Int = 50,
    offset: Int = 0,
    maxCols: Int = 0,
    showLabels: Boolean = false,
    skipEmpty: Boolean = false,
    headerRow: Option[Int] = None,
    skipHidden: Boolean = false,
    rasterOutput: Option[Path] = None
  ): ReadQuery.View =
    ReadQuery.View(
      range,
      showFormulas,
      evalFormulas,
      strict,
      limit,
      offset,
      maxCols,
      format,
      printScale = false,
      showGridlines = false,
      showLabels,
      dpi = 48,
      quality = 90,
      rasterOutput,
      skipEmpty,
      headerRow,
      rasterizer = None,
      skipHidden
    )

  def filter(
    where: String,
    columns: Option[String] = None,
    limit: Int = 50,
    format: FilterFormat = FilterFormat.Markdown,
    header: Boolean = false
  ): ReadQuery.Filter = ReadQuery.Filter(where, columns, limit, format, header)

  def inMemory(
    wb: Workbook,
    sheet: Option[String],
    query: ReadQuery,
    mode: OutputMode = OutputMode.Text
  ): IO[Outcome] =
    Reads.outcome(query, SheetSource.inMemory(wb), sheet, mode)

  def streaming(
    path: Path,
    sheet: Option[String],
    query: ReadQuery,
    mode: OutputMode = OutputMode.Text
  ): IO[Outcome] =
    Reads.outcome(query, SheetSource.streaming(path, excel), sheet, mode)

  /** What text mode prints on stdout for the outcome. */
  def stdout(outcome: Outcome): String = Render.text(outcome).stdout

  /** The payload's text, failing the test on a failed outcome. */
  def text(outcome: Outcome): String =
    outcome.payload match
      case Some(Payload.Text(text, _, _)) => text
      case Some(Payload.Raw(json)) => json
      case Some(Payload.Json(value)) => ujson.write(value, indent = 2)
      case None =>
        throw new AssertionError(s"read failed: ${outcome.error.fold("?")(_.message)}")

  /** The in-memory read's text over a loaded workbook. */
  def readText(wb: Workbook, sheet: Option[String], query: ReadQuery): IO[String] =
    inMemory(wb, sheet, query).map(text)

  /** Write the workbook to a fresh temp file and hand its path to the test. */
  def withTempWorkbook[A](wb: Workbook)(test: Path => IO[A]): IO[A] =
    IO.blocking {
      val tempFile = Files.createTempFile("xl-cli-read-", ".xlsx")
      tempFile.toFile.deleteOnExit()
      tempFile
    }.flatMap(tempFile => excel.write(wb, tempFile) *> test(tempFile))
