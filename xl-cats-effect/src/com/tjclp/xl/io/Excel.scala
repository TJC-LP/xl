package com.tjclp.xl.io

import cats.effect.Async
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.api.Workbook
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.error.{XLError, XLResult}
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
 *   - read/write: Load entire workbook into memory (good for <10k rows)
 *   - readStream/writeStream: Constant-memory streaming (good for 100k+ rows)
 *
 * Type parameter F[_] is the effect type (typically cats.effect.IO).
 */
trait Excel[F[_]]:

  /**
   * Read entire workbook into memory.
   *
   * Good for: Small files (<10k rows), random access, complex transformations Memory: O(n) where n =
   * total cells
   */
  def read(path: Path): F[Workbook]

  /**
   * Write workbook to XLSX file.
   *
   * Good for: Small workbooks, complete data available Memory: O(n) where n = total cells
   */
  def write(wb: Workbook, path: Path): F[Unit]

  /**
   * Write workbook to XLSX file with custom configuration.
   *
   * Allows control over compression (DEFLATED/STORED), SST policy, and XML formatting.
   *
   * Example:
   * {{{
   * // Debug mode (uncompressed, readable)
   * excel.writeWith(workbook, path, WriterConfig.debug)
   *
   * // Custom compression
   * excel.writeWith(workbook, path, WriterConfig(
   *   compression = Compression.Stored,
   *   sstPolicy = SstPolicy.Always
   * ))
   * }}}
   */
  def writeWith(wb: Workbook, path: Path, config: com.tjclp.xl.ooxml.WriterConfig): F[Unit]

  /**
   * Stream rows from first sheet.
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
   * Stream rows from specific sheet by name.
   *
   * Good for: Processing specific sheet in multi-sheet workbook Memory: O(1) constant
   */
  def readSheetStream(path: Path, sheetName: String): Stream[F, RowData]

  /**
   * Stream rows from specific sheet by index (1-based).
   *
   * Good for: Processing sheets by position rather than name Memory: O(1) constant
   *
   * Example:
   * {{{
   * // Read second sheet
   * excel.readStreamByIndex(path, 2).compile.toList
   * }}}
   */
  def readStreamByIndex(path: Path, sheetIndex: Int): Stream[F, RowData]

  /**
   * Streaming write with constant O(1) memory using fs2-data-xml.
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
   *   Display name for the sheet (shown in Excel tab)
   * @param sheetIndex
   *   1-based sheet index (determines internal filename: sheet1.xml, sheet2.xml, etc.)
   *
   * Example:
   * {{{
   * Stream.range(0, 1000000)
   *   .map(i => RowData(i, Map(0 -> CellValue.Number(i))))
   *   .through(excel.writeStream(path, "BigData", sheetIndex = 1))
   *   .compile
   *   .drain
   * }}}
   */
  def writeStream(
    path: Path,
    sheetName: String,
    sheetIndex: Int = 1,
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): fs2.Pipe[F, RowData, Unit]

  /**
   * Write rows to XLSX file, materializing all rows first.
   *
   * @deprecated
   *   Use writeStream instead for true O(1) streaming. This method materializes all rows in memory
   *   before writing, defeating the purpose of streaming.
   */
  @deprecated("Use writeStream for true O(1) streaming", "0.8.0")
  def writeStreamMaterialized(path: Path, sheetName: String): fs2.Pipe[F, RowData, Unit]

  /**
   * Write multiple sheets sequentially with constant memory.
   *
   * Streams each sheet in order - never materializes full datasets. Memory usage is constant
   * regardless of total row count across all sheets.
   *
   * @param path
   *   Output file path
   * @param sheets
   *   Sequence of (sheet name, rows) tuples. Sheets are auto-indexed 1, 2, 3...
   * @param config
   *   Writer configuration (compression, prettyPrint). Defaults to production settings.
   *
   * Example:
   * {{{
   * excel.writeStreamsSeq(
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
  def writeStreamsSeq(
    path: Path,
    sheets: Seq[(String, Stream[F, RowData])],
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): F[Unit]

/**
 * Excel algebra with explicit error channels (pure error handling).
 *
 * Similar to Excel[F] but returns XLResult[A] explicitly instead of raising errors. Enables pure
 * functional error handling without exceptions in the effect type.
 *
 * Usage:
 * {{{
 *   val excel: ExcelR[IO] = ExcelIO.instance
 *   excel.readR(path).flatMap {
 *     case Right(wb) => // Success
 *     case Left(err) => // Handle error
 *   }
 * }}}
 */
trait ExcelR[F[_]]:

  /** Read workbook with explicit error result */
  def readR(path: Path): F[XLResult[Workbook]]

  /** Write workbook with explicit error result */
  def writeR(wb: Workbook, path: Path): F[XLResult[Unit]]

  /** Write workbook with explicit error result and custom configuration */
  def writeWithR(
    wb: Workbook,
    path: Path,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): F[XLResult[Unit]]

  /**
   * Stream rows with explicit error channel.
   *
   * Each row is wrapped in Either[XLError, RowData]. On structural parse failure, emits Left
   * followed by stream termination.
   */
  def readStreamR(path: Path): Stream[F, Either[XLError, RowData]]

  /** Stream rows from specific sheet by name with explicit error channel */
  def readSheetStreamR(path: Path, sheetName: String): Stream[F, Either[XLError, RowData]]

  /** Stream rows from specific sheet by index with explicit error channel */
  def readStreamByIndexR(path: Path, sheetIndex: Int): Stream[F, Either[XLError, RowData]]

  /** Streaming write with explicit error channel */
  def writeStreamR(
    path: Path,
    sheetName: String,
    sheetIndex: Int = 1,
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): fs2.Pipe[F, RowData, Either[XLError, Unit]]

  /** Write stream (materializing) with explicit error channel */
  @deprecated("Use writeStreamR for true O(1) streaming", "0.8.0")
  def writeStreamMaterializedR(
    path: Path,
    sheetName: String
  ): fs2.Pipe[F, RowData, Either[XLError, Unit]]

  /** Write multiple sheets with explicit error channel */
  def writeStreamsSeqR(
    path: Path,
    sheets: Seq[(String, Stream[F, RowData])],
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): F[XLResult[Unit]]

object Excel:
  /** Summon Excel instance from implicit scope */
  def apply[F[_]](using excel: Excel[F]): Excel[F] = excel

  /** Create ExcelIO interpreter for cats-effect IO */
  def forIO: Excel[cats.effect.IO] = ExcelIO.instance

  // ===== Easy Mode API (synchronous, for scripts/REPL) =====

  import cats.effect.{IO, Sync}
  import cats.effect.unsafe.implicits.global
  import com.tjclp.xl.error.XLException
  import java.nio.file.{Files, Paths, StandardCopyOption}

  private lazy val excel = ExcelIO.instance[IO]

  /**
   * Read workbook from file path (Easy Mode).
   *
   * @param path
   *   File path (string)
   * @return
   *   Workbook
   * @throws XLException
   *   if workbook cannot be parsed
   * @throws java.io.IOException
   *   if file cannot be read
   */
  def read(path: String): Workbook =
    excel.read(Paths.get(path)).unsafeRunSync()

  /**
   * Write workbook to file path (Easy Mode).
   *
   * @param workbook
   *   Workbook to write
   * @param path
   *   File path (string)
   * @throws java.io.IOException
   *   if file cannot be written
   */
  def write(workbook: Workbook, path: String): Unit =
    excel.write(workbook, Paths.get(path)).unsafeRunSync()

  /**
   * Write workbook result to file path (Easy Mode).
   *
   * Convenience overload that handles XLResult[Workbook] directly.
   *
   * @param result
   *   XLResult[Workbook] to write (calls .unsafe internally)
   * @param path
   *   File path (string)
   * @throws XLException
   *   if result is Left (contains error)
   * @throws java.io.IOException
   *   if file cannot be written
   */
  def write(result: XLResult[Workbook], path: String): Unit =
    write(result.fold(e => throw XLException(e), identity), path)

  /**
   * Modify workbook in-place (Easy Mode: read → transform → write).
   *
   * Uses atomic file replacement to avoid ZIP corruption.
   *
   * @param path
   *   File path (string)
   * @param f
   *   Transformation function
   * @throws XLException
   *   if workbook cannot be parsed
   * @throws java.io.IOException
   *   if file cannot be read/written
   */
  def modify(path: String)(f: Workbook => Workbook): Unit =
    val targetPath = Paths.get(path)
    val result = for
      wb <- excel.read(targetPath)
      modified = f(wb)
      parent = Option(targetPath.getParent).getOrElse(Paths.get("."))
      tempFile <- Sync[IO].delay(Files.createTempFile(parent, ".xl-modify-", ".tmp"))
      writeResult <- excel.write(modified, tempFile).attempt
      _ <- writeResult match
        case scala.util.Right(_) =>
          Sync[IO].delay(
            Files.move(
              tempFile,
              targetPath,
              StandardCopyOption.REPLACE_EXISTING,
              StandardCopyOption.ATOMIC_MOVE
            )
          )
        case scala.util.Left(err) =>
          Sync[IO].delay(Files.deleteIfExists(tempFile)) >>
            IO.raiseError(err)
    yield ()
    result.unsafeRunSync()
