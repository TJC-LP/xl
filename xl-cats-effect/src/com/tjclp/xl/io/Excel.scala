package com.tjclp.xl.io

import cats.effect.Async
import fs2.Stream
import java.nio.file.Path
import com.tjclp.xl.{Workbook, Sheet, Cell, CellValue, ARef}

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
   *   Display name for the sheet (shown in Excel tab)
   * @param sheetIndex
   *   1-based sheet index (determines internal filename: sheet1.xml, sheet2.xml, etc.)
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
    sheetIndex: Int = 1
  ): fs2.Pipe[F, RowData, Unit]

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
    sheets: Seq[(String, Stream[F, RowData])]
  ): F[Unit]

object Excel:
  /** Summon Excel instance from implicit scope */
  def apply[F[_]](using excel: Excel[F]): Excel[F] = excel

  /** Create ExcelIO interpreter for cats-effect IO */
  def forIO: Excel[cats.effect.IO] = ExcelIO.instance
