package com.tjclp.xl.io

import cats.effect.{IO, Sync}
import cats.effect.unsafe.implicits.global
import com.tjclp.xl.error.XLException
import com.tjclp.xl.workbooks.Workbook

import java.nio.file.{Files, Path, Paths, StandardCopyOption}

/**
 * Simplified Excel IO for Easy Mode API.
 *
 * Provides synchronous read/write operations for scripts and REPL usage. Wraps [[ExcelIO]] with
 * string-based paths and immediate execution.
 *
 * '''Example: Basic usage'''
 * {{{
 * import com.tjclp.xl.easy.*
 *
 * val wb = Excel.read("data.xlsx")
 * // ... modify workbook
 * Excel.write(wb, "output.xlsx")
 * }}}
 *
 * '''Example: Modify pattern'''
 * {{{
 * Excel.modify("data.xlsx") { wb =>
 *   wb.sheet("Sales")
 *     .put("A1", "Updated")
 * }
 * }}}
 *
 * @note
 *   These methods use `unsafeRunSync` and throw [[XLException]] or `IOException` on errors. For
 *   pure functional code, use [[ExcelIO]] directly.
 * @since 0.3.0
 */
object EasyExcel:
  private val excel = ExcelIO.instance[IO]

  /**
   * Read workbook from file path.
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
   * Write workbook to file path.
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
   * Modify workbook in-place (read → transform → write).
   *
   * Uses atomic file replacement to avoid ZIP corruption when writing to the source file. Writes to
   * a temporary file first, then atomically moves it to the target path.
   *
   * '''Example:'''
   * {{{
   * Excel.modify("data.xlsx") { wb =>
   *   val updated = wb.sheet("Sales")
   *     .put("A1", "New Value")
   *   wb.put(updated)
   * }
   * }}}
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
      // Write to temp file to avoid reading from file being written (ZIP corruption)
      // Atomic file replacement ensures transaction safety:
      // - Original file preserved if write fails
      // - No intermediate corrupt state visible to readers
      // - ATOMIC_MOVE provides filesystem-level atomicity
      parent = Option(targetPath.getParent).getOrElse(Paths.get("."))
      tempFile <- Sync[IO].delay(Files.createTempFile(parent, ".xl-modify-", ".tmp"))
      writeResult <- excel.write(modified, tempFile).attempt
      _ <- writeResult match
        case Right(_) =>
          // Success: atomically replace original (preserves original if move fails)
          Sync[IO].delay(
            Files.move(
              tempFile,
              targetPath,
              StandardCopyOption.REPLACE_EXISTING,
              StandardCopyOption.ATOMIC_MOVE
            )
          )
        case Left(err) =>
          // Failure: clean up temp and propagate error (original unchanged)
          Sync[IO].delay(Files.deleteIfExists(tempFile)) >>
            IO.raiseError(err)
    yield ()
    result.unsafeRunSync()
