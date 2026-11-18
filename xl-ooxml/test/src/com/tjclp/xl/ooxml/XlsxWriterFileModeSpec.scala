package com.tjclp.xl.ooxml

import java.nio.file.{FileSystems, Files, Path}
import java.nio.file.attribute.PosixFilePermissions

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * GH-868: the surgical write stages the output in an owner-only temp file (`createTempFile` makes
 * it 0600) and renames it into place, so the output used to keep 0600 whatever the umask. A new
 * output now gets the mode any new file gets in its directory; an overwritten one keeps its mode —
 * including an owner-only mode, which is never widened to the umask's.
 */
class XlsxWriterFileModeSpec extends FunSuite:

  private val posix = FileSystems.getDefault.supportedFileAttributeViews.contains("posix")

  private def mode(p: Path): String =
    PosixFilePermissions.toString(Files.getPosixFilePermissions(p))

  private def setMode(p: Path, m: String): Unit =
    Files.setPosixFilePermissions(p, PosixFilePermissions.fromString(m)): Unit

  /** The mode this process's umask gives a new file in `dir`. */
  private def freshMode(dir: Path): String =
    val probe = Files.createFile(dir.resolve("umask-probe"))
    try mode(probe)
    finally Files.delete(probe)

  /**
   * A workbook read back from disk and edited: it carries a source, so it takes the atomic path.
   */
  private def edited(dir: Path): Workbook =
    val source = dir.resolve("source.xlsx")
    assertEquals(
      XlsxWriter.writeWith(Workbook(Vector(Sheet("Data").put(ref"A1", 1))), source),
      Right(())
    )
    val read = XlsxReader.read(source).fold(e => fail(e.message), identity)
    read.put(read.sheets(0).put(ref"A1", 2))

  test("a new output gets the umask's mode, not 0600 (GH-868)") {
    assume(posix, "POSIX permissions only")
    val dir = Files.createTempDirectory("xl-gh868-new-")
    val dest = dir.resolve("dest.xlsx")
    assertEquals(XlsxWriter.writeWith(edited(dir), dest), Right(()))
    assertEquals(mode(dest), freshMode(dir))
  }

  test("an overwritten output keeps its mode (GH-868)") {
    assume(posix, "POSIX permissions only")
    val dir = Files.createTempDirectory("xl-gh868-keep-")
    val dest = dir.resolve("dest.xlsx")
    Files.writeString(dest, "old")
    setMode(dest, "rw-r-----")
    assertEquals(XlsxWriter.writeWith(edited(dir), dest), Right(()))
    assertEquals(mode(dest), "rw-r-----")
  }

  test("an owner-only output stays owner-only when overwritten (GH-868)") {
    assume(posix, "POSIX permissions only")
    val dir = Files.createTempDirectory("xl-gh868-private-")
    val dest = dir.resolve("dest.xlsx")
    Files.writeString(dest, "old")
    setMode(dest, "rw-------")
    assertEquals(XlsxWriter.writeWith(edited(dir), dest), Right(()))
    assertEquals(mode(dest), "rw-------")
  }
