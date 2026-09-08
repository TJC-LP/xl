package com.tjclp.xl.cli.read

import java.nio.charset.StandardCharsets.UTF_8
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.jdk.CollectionConverters.*

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

  /**
   * Copy the zip at `path` onto itself entry by entry, each entry's bytes passed through `f` (entry
   * name, bytes): how a test hands the readers a file another producer would have written.
   */
  def rewriteZip(path: Path)(f: (String, Array[Byte]) => Array[Byte]): IO[Unit] =
    IO.blocking {
      val zip = new ZipFile(path.toFile)
      val entries =
        try
          zip.entries().asScala.toVector.map { entry =>
            val in = zip.getInputStream(entry)
            try entry.getName -> in.readAllBytes()
            finally in.close()
          }
        finally zip.close()
      val out = new ZipOutputStream(Files.newOutputStream(path))
      try
        entries.foreach { (name, bytes) =>
          out.putNextEntry(new ZipEntry(name))
          out.write(f(name, bytes))
          out.closeEntry()
        }
      finally out.close()
    }

  /** One zip entry's text. */
  def zipEntry(path: Path, name: String): IO[String] =
    IO.blocking {
      val zip = new ZipFile(path.toFile)
      try
        val in = zip.getInputStream(zip.getEntry(name))
        try new String(in.readAllBytes(), UTF_8)
        finally in.close()
      finally zip.close()
    }

  /**
   * openpyxl's workbook rels: worksheet Targets package-absolute (`/xl/worksheets/sheet1.xml`)
   * where the library's writer, like Excel, writes them relative to xl/ (`worksheets/sheet1.xml`).
   */
  val openpyxlTargets: (String, Array[Byte]) => Array[Byte] = (name, bytes) =>
    if name == "xl/_rels/workbook.xml.rels" then
      new String(bytes, UTF_8)
        .replace("Target=\"worksheets/", "Target=\"/xl/worksheets/")
        .getBytes(UTF_8)
    else bytes

  /**
   * Excel's empty worksheet: `<dimension ref="A1"/>` on every worksheet the library's writer left
   * without one (it records no `<dimension>` for a sheet with no cells; Excel always writes `A1`).
   */
  val excelEmptySheetDimension: (String, Array[Byte]) => Array[Byte] = (name, bytes) =>
    if name.startsWith("xl/worksheets/sheet") && name.endsWith(".xml") then
      val xml = new String(bytes, UTF_8)
      if xml.contains("<dimension") then bytes
      else
        // CT_Worksheet order: sheetPr, dimension, sheetViews, sheetFormatPr, cols, sheetData
        val at = Vector("<sheetViews", "<sheetFormatPr", "<cols", "<sheetData")
          .map(xml.indexOf)
          .filter(_ >= 0)
          .minOption
          .getOrElse(xml.length)
        (xml.substring(0, at) + "<dimension ref=\"A1\"/>" + xml.substring(at)).getBytes(UTF_8)
    else bytes
