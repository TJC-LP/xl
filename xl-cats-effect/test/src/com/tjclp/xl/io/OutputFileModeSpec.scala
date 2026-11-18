package com.tjclp.xl.io

import java.nio.file.{FileSystems, Files, Path}
import java.nio.file.attribute.PosixFilePermissions

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.streaming.ZipTransformer
import com.tjclp.xl.macros.ref

/**
 * GH-868: the streaming transforms and `Excel.modify` stage their output in an owner-only temp file
 * and rename it into place, so the output used to end up 0600 whatever the umask. A new output now
 * gets the mode any new file gets in its directory; a replaced one keeps its mode, its group bits
 * capped at its others bits.
 */
class OutputFileModeSpec extends CatsEffectSuite:

  private val posix = FileSystems.getDefault.supportedFileAttributeViews.contains("posix")
  private val excel = ExcelIO.instance[IO]

  private def mode(p: Path): String =
    PosixFilePermissions.toString(Files.getPosixFilePermissions(p))

  private def setMode(p: Path, m: String): Unit =
    Files.setPosixFilePermissions(p, PosixFilePermissions.fromString(m)): Unit

  /** The mode this process's umask gives a new file in `dir`. */
  private def freshMode(dir: Path): String =
    val probe = Files.createFile(dir.resolve("umask-probe"))
    try mode(probe)
    finally Files.delete(probe)

  private def book(dir: Path): IO[Path] =
    val source = dir.resolve("source.xlsx")
    excel
      .write(Workbook(Vector(Sheet("Data").put(ref"A1", CellValue.Number(1)))), source)
      .as(source)

  private def streamPut(source: Path, output: Path): IO[Unit] =
    ZipTransformer
      .transformValues[IO](
        source,
        output,
        "xl/worksheets/sheet1.xml",
        Map(ref"A1" -> CellValue.Number(2))
      )
      .void

  test("streaming transform: a new output gets the umask's mode (GH-868)") {
    assume(posix, "POSIX permissions only")
    for
      dir <- IO.blocking(Files.createTempDirectory("xl-gh868-stream-new-"))
      source <- book(dir)
      output = dir.resolve("out.xlsx")
      _ <- streamPut(source, output)
    yield assertEquals(mode(output), freshMode(dir))
  }

  test("streaming transform: a replaced output keeps its mode (GH-868)") {
    assume(posix, "POSIX permissions only")
    for
      dir <- IO.blocking(Files.createTempDirectory("xl-gh868-stream-keep-"))
      source <- book(dir)
      _ <- IO.blocking(setMode(source, "rw-r--r--"))
      _ <- streamPut(source, source)
    yield assertEquals(mode(source), "rw-r--r--")
  }

  test("Excel.modify keeps the file's mode (GH-868)") {
    assume(posix, "POSIX permissions only")
    for
      dir <- IO.blocking(Files.createTempDirectory("xl-gh868-modify-"))
      source <- book(dir)
      _ <- IO.blocking(setMode(source, "rw-r--r--"))
      _ <- IO.blocking(Excel.modify(source.toString)(wb => wb))
    yield assertEquals(mode(source), "rw-r--r--")
  }
