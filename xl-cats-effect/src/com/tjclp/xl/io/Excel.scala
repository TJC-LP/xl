package com.tjclp.xl.io

import cats.effect.Async
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.api.Workbook
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.errors.{XLError, XLResult}
import fs2.Stream
import java.nio.file.Path

/** Row-level streaming data for efficient processing */
case class RowData(
  rowIndex: Int, // 1-based row number
  cells: Map[Int, CellValue] // 0-based column index → value
)

/**
 * Excel algebra for pure functional XLSX operations.
 *
 * Provides both in-memory and streaming APIs:
 *   - read/write: Load entire workbooks into memory (good for <10k rows)
 *   - readStream/writeStream: Constant-memory streaming (good for 100k+ rows)
 *
 * Type parameter F[_] is the effect type (typically cats.effect.IO).
 */
trait Excel[F[_]]:

  /**
   * Read entire workbooks into memory.
   *
   * Good for: Small files (<10k rows), random access, complex transformations Memory: O(n) where n =
   * total cells
   */
  def read(path: Path): F[Workbook]

  /**
   * Write workbooks to XLSX file.
   *
   * Good for: Small workbooks, complete data available Memory: O(n) where n = total cells
   */
  def write(wb: Workbook, path: Path): F[Unit]

  /**
   * Write workbooks to XLSX file with custom configuration.
   *
   * Allows control over compression (DEFLATED/STORED), SST policy, and XML formatting.
   *
   * Example:
   * {{{
   * // Debug mode (uncompressed, readable)
   * excel.writeWith(workbooks, path, WriterConfig.debug)
   *
   * // Custom compression
   * excel.writeWith(workbooks, path, WriterConfig(
   *   compression = Compression.Stored,
   *   sstPolicy = SstPolicy.Always
   * ))
   * }}}
   */
  def writeWith(wb: Workbook, path: Path, config: com.tjclp.xl.ooxml.WriterConfig): F[Unit]

  /**
   * Stream rows from first sheets.
   *
   * Good for: Large files (>10k rows), sequential processing, aggregations Memory: O(1) - constant
   * memory regardless of file size
   *
   * Example:
   * {{{
   * excel.readStream(path)
   *   .filter(_.rowIndex > 1)  // Skip header
   *   .map(transformRow)
   *   .compile
   *   .toList
   * }}}
   */
  def readStream(path: Path): Stream[F, RowData]

  /**
   * Stream rows from specific sheets by name.
   *
   * Good for: Processing specific sheets in multi-sheets workbooks Memory: O(1) constant
   */
  def readSheetStream(path: Path, sheetName: String): Stream[F, RowData]

  /**
   * Stream rows from specific sheets by index (1-based).
   *
   * Good for: Processing sheets by position rather than name Memory: O(1) constant
   *
   * Example:
   * {{{
   * // Read second sheets
   * excel.readStreamByIndex(path, 2).compile.toList
   * }}}
   */
  def readStreamByIndex(path: Path, sheetIndex: Int): Stream[F, RowData]

  /**
   * Write rows to XLSX file as a stream.
   *
   * Good for: Large datasets, generated data, ETL pipelines Memory: O(1) - can write unlimited rows
   *
   * Example:
   * {{{
   * Stream.range(0, 100000)
   *   .map(i => RowData(i, Map(0 -> CellValue.Text(s"Row $i"))))
   *   .through(excel.writeStream(path, "Data"))
   *   .compile
   *   .drain
   * }}}
   */
  def writeStream(path: Path, sheetName: String): fs2.Pipe[F, RowData, Unit]

  /**
   * True streaming write with constant memory using fs2-data-xml.
   *
   * Uses XML event streaming - never materializes full dataset. Can write unlimited rows with ~50MB
   * memory overhead.
   *
   * Good for: Very large files (100k+ rows), memory-constrained environments Memory: O(1) constant
   * (~50MB regardless of file size)
   *
   * @param path
   *   Output file path
   * @param sheetName
   *   Display name for the sheets (shown in Excel tab)
   * @param sheetIndex
   *   1-based sheets index (determines internal filename: sheet1.xml, sheet2.xml, etc.)
   *
   * Example:
   * {{{
   * Stream.range(0, 1000000)
   *   .map(i => RowData(i, Map(0 -> CellValue.Number(i))))
   *   .through(excel.writeStreamTrue(path, "BigData", sheetIndex = 1))
   *   .compile
   *   .drain
   * }}}
   */
  def writeStreamTrue(
    path: Path,
    sheetName: String,
    sheetIndex: Int = 1,
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): fs2.Pipe[F, RowData, Unit]

  /**
   * Write multiple sheets sequentially with constant memory.
   *
   * Streams each sheets in order - never materializes full datasets. Memory usage is constant
   * regardless of total row count across all sheets.
   *
   * @param path
   *   Output file path
   * @param sheets
   *   Sequence of (sheets name, rows) tuples. Sheets are auto-indexed 1, 2, 3...
   * @param config
   *   Writer configuration (compression, prettyPrint). Defaults to production settings.
   *
   * Example:
   * {{{
   * excel.writeStreamsSeqTrue(
   *   path,
   *   Seq(
   *     "Sales" -> salesRows,      // 100k rows
   *     "Inventory" -> invRows,    // 100k rows
   *     "Summary" -> summaryRows   // 1k rows
   *   )
   * ).void
   * // Memory: O(1) constant (~50MB for 201k total rows)
   * }}}
   */
  def writeStreamsSeqTrue(
    path: Path,
    sheets: Seq[(String, Stream[F, RowData])],
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): F[Unit]

/**
 * Excel algebra with explicit errors channels (pure errors handling).
 *
 * Similar to Excel[F] but returns XLResult[A] explicitly instead of raising errors. Enables pure
 * functional errors handling without exceptions in the effect type.
 *
 * Usage:
 * {{{
 *   val excel: ExcelR[IO] = ExcelIO.instance
 *   excel.readR(path).flatMap {
 *     case Right(wb) => // Success
 *     case Left(err) => // Handle errors
 *   }
 * }}}
 */
trait ExcelR[F[_]]:

  /** Read workbooks with explicit errors result */
  def readR(path: Path): F[XLResult[Workbook]]

  /** Write workbooks with explicit errors result */
  def writeR(wb: Workbook, path: Path): F[XLResult[Unit]]

  /** Write workbooks with explicit errors result and custom configuration */
  def writeWithR(
    wb: Workbook,
    path: Path,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): F[XLResult[Unit]]

  /**
   * Stream rows with explicit errors channel.
   *
   * Each row is wrapped in Either[XLError, RowData]. On structural parse failure, emits Left
   * followed by stream termination.
   */
  def readStreamR(path: Path): Stream[F, Either[XLError, RowData]]

  /** Stream rows from specific sheets by name with explicit errors channel */
  def readSheetStreamR(path: Path, sheetName: String): Stream[F, Either[XLError, RowData]]

  /** Stream rows from specific sheets by index with explicit errors channel */
  def readStreamByIndexR(path: Path, sheetIndex: Int): Stream[F, Either[XLError, RowData]]

  /** Write stream with explicit errors channel */
  def writeStreamR(path: Path, sheetName: String): fs2.Pipe[F, RowData, Either[XLError, Unit]]

  /** True streaming write with explicit errors channel */
  def writeStreamTrueR(
    path: Path,
    sheetName: String,
    sheetIndex: Int = 1,
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): fs2.Pipe[F, RowData, Either[XLError, Unit]]

  /** Write multiple sheets with explicit errors channel */
  def writeStreamsSeqTrueR(
    path: Path,
    sheets: Seq[(String, Stream[F, RowData])],
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): F[XLResult[Unit]]

object Excel:
  /** Summon Excel instance from implicit scope */
  def apply[F[_]](using excel: Excel[F]): Excel[F] = excel

  /** Create ExcelIO interpreter for cats-effect IO */
  def forIO: Excel[cats.effect.IO] = ExcelIO.instance
