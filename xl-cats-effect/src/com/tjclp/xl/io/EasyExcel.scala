package com.tjclp.xl.io

import cats.effect.IO
import cats.effect.unsafe.implicits.global
import com.tjclp.xl.errors.XLException
import com.tjclp.xl.workbooks.Workbook

import java.nio.file.{Path, Paths}

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
 * // ... modify workbooks
 * Excel.write(wb, "output.xlsx")
 * }}}
 *
 * '''Example: Modify pattern'''
 * {{{
 * Excel.modify("data.xlsx") { wb =>
 *   wb.sheets("Sales")
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
   * Read workbooks from file path.
   *
   * @param path
   *   File path (string)
   * @return
   *   Workbook
   * @throws XLException
   *   if workbooks cannot be parsed
   * @throws java.io.IOException
   *   if file cannot be read
   */
  def read(path: String): Workbook =
    excel.read(Paths.get(path)).unsafeRunSync()

  /**
   * Write workbooks to file path.
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
   * Modify workbooks in-place (read → transform → write).
   *
   * '''Example:'''
   * {{{
   * Excel.modify("data.xlsx") { wb =>
   *   val updated = wb.sheets("Sales")
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
   *   if workbooks cannot be parsed
   * @throws java.io.IOException
   *   if file cannot be read/written
   */
  def modify(path: String)(f: Workbook => Workbook): Unit =
    val p = Paths.get(path)
    val result = for
      wb <- excel.read(p)
      modified = f(wb)
      _ <- excel.write(modified, p)
    yield ()
    result.unsafeRunSync()
