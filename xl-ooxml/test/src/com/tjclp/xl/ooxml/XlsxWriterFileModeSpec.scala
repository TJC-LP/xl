package com.tjclp.xl.ooxml

import java.nio.file.{FileSystems, Files, Path}
import java.nio.file.attribute.{PosixFileAttributes, PosixFilePermissions}

import scala.jdk.CollectionConverters.*

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * GH-868: the surgical write stages the output in an owner-only temp file (`createTempFile` makes
 * it 0600) and renames it into place, so the output used to keep 0600 whatever the umask. A new
 * output now gets the mode any new file gets in its directory; an overwritten one keeps its mode
 * exactly (#930: its group bits used to be capped at its others bits) — including an owner-only
 * mode, which is never widened to the umask's.
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
    setMode(dest, "rw-r--r--")
    assertEquals(XlsxWriter.writeWith(edited(dir), dest), Right(()))
    assertEquals(mode(dest), "rw-r--r--")
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

  private def group(p: Path): String =
    Files.readAttributes(p, classOf[PosixFileAttributes]).group.getName

  /** Move `p` to another group the current user belongs to, if there is one. */
  private def regroup(p: Path): Option[String] =
    val proc = new ProcessBuilder("id", "-Gn").redirectErrorStream(true).start()
    val names = new String(proc.getInputStream.readAllBytes()).trim.split("\\s+").toList
    proc.waitFor(): Unit
    val lookup = p.getFileSystem.getUserPrincipalLookupService
    names.filter(n => n.nonEmpty && n != group(p)).find { name =>
      try
        Files.setAttribute(p, "posix:group", lookup.lookupPrincipalByGroupName(name))
        true
      catch case _: Exception => false
    }

  test("an overwritten output keeps its group and its exact mode (GH-868 review, #930)") {
    assume(posix, "POSIX permissions only")
    val dir = Files.createTempDirectory("xl-gh868-group-")
    val dest = dir.resolve("dest.xlsx")
    Files.writeString(dest, "old")
    setMode(dest, "rw-rw----")
    val other = regroup(dest)
    assume(other.isDefined, "needs a second group the user belongs to")
    assertEquals(XlsxWriter.writeWith(edited(dir), dest), Right(()))
    assertEquals(Some(group(dest)), other)
    assertEquals(mode(dest), "rw-rw----")
  }

  test("group bits are capped at the others bits when the group cannot be kept (GH-868 review)") {
    def cap(m: String): String =
      PosixFilePermissions.toString(
        OutputFileMode.groupCappedAtOthers(PosixFilePermissions.fromString(m))
      )
    assertEquals(cap("rw-rw----"), "rw-------")
    assertEquals(cap("rw-rw-r--"), "rw-r--r--")
    assertEquals(cap("rw-r--r--"), "rw-r--r--")
    assertEquals(cap("rw-------"), "rw-------")
  }

  test("a group-shared output keeps its group bits (#930)") {
    assume(posix, "POSIX permissions only")
    List("rw-rw-r--", "rw-rw----", "rw-r-----").foreach { m =>
      val dir = Files.createTempDirectory("xl-930-shared-")
      val dest = dir.resolve("dest.xlsx")
      Files.writeString(dest, "old")
      setMode(dest, m)
      assertEquals(XlsxWriter.writeWith(edited(dir), dest), Right(()))
      assertEquals(mode(dest), m)
    }
  }

  /** Run a command, true when it exits 0. */
  private def run(cmd: String*): Boolean =
    try
      val proc = new ProcessBuilder(cmd*).redirectErrorStream(true).start()
      proc.getInputStream.readAllBytes(): Unit
      proc.waitFor() == 0
    catch case _: java.io.IOException => false

  /**
   * Give `p` a POSIX.1e extended ACL entry for the current user (Linux `setfacl`). Only these ACLs
   * turn the group bits into a mask; a macOS ACL adds to the mode bits and never widens them.
   */
  private def addAcl(p: Path): Boolean =
    val user = System.getProperty("user.name")
    val added = run("setfacl", "-m", s"u:$user:r", p.toString)
    val proc = new ProcessBuilder("ls", "-ld", p.toString).start()
    val listing = new String(proc.getInputStream.readAllBytes())
    proc.waitFor(): Unit
    added && listing.lift(10).contains('+')

  test("an overwritten file with an extended ACL keeps its mode bits, not its ACL (#930)") {
    assume(posix, "POSIX permissions only")
    val dir = Files.createTempDirectory("xl-gh868-acl-")
    val dest = dir.resolve("dest.xlsx")
    Files.writeString(dest, "old")
    setMode(dest, "rw-rw----")
    assume(addAcl(dest), "needs Linux setfacl")
    assertEquals(XlsxWriter.writeWith(edited(dir), dest), Right(()))
    // the documented residual: the JVM can neither read nor copy the ACL, so the replaced file
    // carries none, and its group bits (the old mask) now apply to the owning group
    assertEquals(mode(dest), "rw-rw----")
  }
