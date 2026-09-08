package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path}

import scala.jdk.CollectionConverters.*

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * GH-636: `writeAtomically` stages the output as `.xl-<random>.tmp` beside the destination. Its
 * cleanup used to live in `catch case e: Exception`, so an `Error` raised while serialising — an
 * `OutOfMemoryError`, which the CLI now reports as a typed failure instead of dying — left the
 * partial temp file next to the input. The cleanup is in `finally` now.
 */
class XlsxWriterAtomicCleanupSpec extends FunSuite:

  /** The `.xl-*.tmp` files in `dir`. */
  private def tempFiles(dir: Path): Vector[String] =
    val listing = Files.list(dir)
    try
      listing.iterator.asScala
        .map(_.getFileName.toString)
        .filter(n => n.startsWith(".xl-") && n.endsWith(".tmp"))
        .toVector
    finally listing.close()

  /**
   * A cell map that raises an `OutOfMemoryError` the moment serialisation iterates it WHILE the
   * writer's temp file exists — the only way to raise a real `Error` inside the atomic write, and
   * self-verifying: `sawTemp` records that the path under test was reached.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var")) // a test probe
  final class SerialisationTrap(dir: Path)
      extends scala.collection.immutable.AbstractMap[ARef, Cell]:
    @volatile var sawTemp: Boolean = false
    def iterator: Iterator[(ARef, Cell)] =
      if tempFiles(dir).nonEmpty then
        sawTemp = true
        throw new OutOfMemoryError("test: serialisation")
      else Iterator.empty
    def get(key: ARef): Option[Cell] = None
    def removed(key: ARef): Map[ARef, Cell] = this
    def updated[V1 >: Cell](key: ARef, value: V1): Map[ARef, V1] = this

  test("an Error while serialising leaves no .xl-*.tmp beside the destination (GH-636)") {
    val dir = Files.createTempDirectory("xl-atomic-cleanup-")
    val source = dir.resolve("source.xlsx")
    val dest = dir.resolve("dest.xlsx")
    assertEquals(
      XlsxWriter.writeWith(Workbook(Vector(Sheet("Data").put(ref"A1", 1))), source),
      Right(())
    )
    // Read back WITH a source context and modify the sheet: the surgical, atomic write path
    val read = XlsxReader.read(source).fold(e => fail(e.message), identity)
    val trap = new SerialisationTrap(dir)
    val trapped = read.put(read.sheets(0).copy(cells = trap))
    // munit's intercept lets a fatal Error through: catch it by hand
    val thrown =
      try
        XlsxWriter.writeWith(trapped, dest)
        None
      catch case e: OutOfMemoryError => Some(e)
    assert(
      thrown.isDefined,
      "the Error must propagate out of the writer (only Exceptions are mapped)"
    )
    assert(trap.sawTemp, "the Error must have been raised with the writer's temp file present")
    assertEquals(tempFiles(dir), Vector.empty, "no partial temp file may outlive the failure")
    assert(!Files.exists(dest), "nothing committed")
  }
