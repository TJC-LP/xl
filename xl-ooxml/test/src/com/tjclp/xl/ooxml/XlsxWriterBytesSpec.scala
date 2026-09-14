package com.tjclp.xl.ooxml

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path, Paths}
import java.util.zip.ZipInputStream

import scala.util.Try

import munit.ScalaCheckSuite
import org.scalacheck.Prop.forAll

import com.tjclp.xl.Generators
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.error.XLResult
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * The workbooks the GH-516 tmpdir pin writes — built identically in the parent test JVM and in the
 * child (the writer is deterministic, so both sides can compare bytes).
 */
object WriteToBytesNoTmpdirMain:
  def fresh(): Workbook =
    Workbook(
      Vector(
        Sheet("Data").put(ref"A1", "hello").put(ref"B1", 42).put(ref"A2", "shared"),
        Sheet("Other").put(ref"A1", "shared")
      )
    )

  /** A read-back workbook with an edit: the surgical (source-context) path into a stream. */
  def edited(): Either[String, Workbook] =
    XlsxWriter
      .writeToBytes(fresh())
      .flatMap(XlsxReader.readFromBytes(_))
      .map(read => read.put(read.sheets(0).put(ref"C3", CellValue.Text("edited"))))
      .left
      .map(_.message)

  /**
   * Runs in a CHILD JVM whose `java.io.tmpdir` names a directory that does not exist: proves the
   * temp directory is unusable (exit 3 otherwise — the pin would be toothless), then writes the
   * fresh and the edited workbook with `writeToBytes` into args(0) / args(1) (exit 1 on a Left).
   */
  def main(args: Array[String]): Unit =
    if Try(Files.createTempFile("xl-516-probe-", ".tmp")).isSuccess then
      System.err.println(s"java.io.tmpdir=${System.getProperty("java.io.tmpdir")} is usable")
      sys.exit(3)
    val results = for
      freshBytes <- XlsxWriter.writeToBytes(fresh()).left.map(_.message)
      editedWb <- edited()
      editedBytes <- XlsxWriter.writeToBytes(editedWb).left.map(_.message)
    yield (freshBytes, editedBytes)
    results match
      case Left(message) =>
        System.err.println(message)
        sys.exit(1)
      case Right((freshBytes, editedBytes)) =>
        Files.write(Paths.get(args(0)), freshBytes)
        Files.write(Paths.get(args(1)), editedBytes)

/**
 * GH-516: `writeToBytes` serialises in memory, through the same strategy dispatch as `write` — no
 * scratch file in `java.io.tmpdir`. The temp directory cannot be re-pointed inside one JVM
 * (`Files.createTempFile` reads the property once), so the absence of a scratch file is pinned from
 * a CHILD JVM started with a `java.io.tmpdir` that does not exist (the OrExitSmokeTest pattern):
 * the write must succeed there and produce what `write(path)` produces here. The byte-level
 * contracts:
 *   - what lands in the array is what `write` puts on disk;
 *   - a CLEAN read-back workbook yields its source bytes verbatim (the `readFromBytes` contract),
 *     from an in-memory source and from an on-disk one;
 *   - a source that changed underneath the context is refused as `Left`, never thrown;
 *   - the configuration overload is honoured.
 */
class XlsxWriterBytesSpec extends ScalaCheckSuite:

  private val tempDir: Path = Files.createTempDirectory("xl-bytes-")

  override def afterAll(): Unit =
    Files.walk(tempDir).sorted(java.util.Comparator.reverseOrder()).forEach(Files.delete)

  private def orFail[A](result: XLResult[A]): A = result.fold(e => fail(e.message), identity)

  /**
   * Every XML part of the archive, concatenated: where a text lives (sheet or SST) is not the
   * point.
   */
  private def partsText(bytes: Array[Byte]): String =
    val zin = new ZipInputStream(new ByteArrayInputStream(bytes))
    try
      Iterator
        .continually(Option(zin.getNextEntry))
        .takeWhile(_.isDefined)
        .flatten
        .filter(_.getName.endsWith(".xml"))
        .map(_ => new String(zin.readAllBytes(), StandardCharsets.UTF_8))
        .mkString("\n")
    finally zin.close()

  private def book(): Workbook =
    Workbook(
      Vector(
        Sheet("Data").put(ref"A1", "hello").put(ref"B1", 42).put(ref"A2", "shared"),
        Sheet("Other").put(ref"A1", "shared")
      )
    )

  property("writeToBytes(wb) is byte-identical to what write(wb, path) puts on disk") {
    forAll(Generators.genWorkbook) { (wb: Workbook) =>
      val path = Files.createTempFile(tempDir, "gen-", ".xlsx")
      try
        orFail(XlsxWriter.write(wb, path))
        val onDisk = Files.readAllBytes(path)
        val inMemory = orFail(XlsxWriter.writeToBytes(wb))
        assert(inMemory.sameElements(onDisk), "the stream target and the file target must agree")
        true
      finally Files.deleteIfExists(path)
    }
  }

  test("a clean read-back workbook from bytes serialises to its source bytes verbatim") {
    val source = orFail(XlsxWriter.writeToBytes(book()))
    val read = orFail(XlsxReader.readFromBytes(source))
    val again = orFail(XlsxWriter.writeToBytes(read))
    assert(again.sameElements(source), "writeToBytes(readFromBytes(bytes)) must be bytes")
  }

  test("a clean read-back workbook from disk serialises to its source bytes verbatim") {
    val path = tempDir.resolve("on-disk-source.xlsx")
    orFail(XlsxWriter.write(book(), path))
    val read = orFail(XlsxReader.read(path))
    assert(orFail(XlsxWriter.writeToBytes(read)).sameElements(Files.readAllBytes(path)))
  }

  test("a clean Excel-authored workbook serialises to its source bytes verbatim") {
    // Excel's own [Content_Types].xml, rels and calcChain: regenerating any of them would show
    val path = TestFixtures.copyToTemp("datatable-excel.xlsx")
    val read = orFail(XlsxReader.read(path))
    assert(read.sourceContext.exists(_.isClean), "a fresh read is clean")
    assert(orFail(XlsxWriter.writeToBytes(read)).sameElements(Files.readAllBytes(path)))
  }

  test("a source that changed underneath the context is refused as Left, not thrown") {
    val path = tempDir.resolve("changed-source.xlsx")
    orFail(XlsxWriter.write(book(), path))
    val read = orFail(XlsxReader.read(path))
    val original = Files.readAllBytes(path)

    // Same size, different bytes: the digest catches it
    val flipped = original.clone()
    flipped(flipped.length / 2) = (flipped(flipped.length / 2) ^ 0xff).toByte
    Files.write(path, flipped)
    XlsxWriter.writeToBytes(read) match
      case Left(err) => assert(err.message.contains("changed since read"), err.message)
      case Right(_) => fail("a rewritten source must not be copied verbatim")

    // Different size: the cheap check catches it
    Files.write(path, original ++ Array[Byte](0))
    XlsxWriter.writeToBytes(read) match
      case Left(err) => assert(err.message.contains("changed size since read"), err.message)
      case Right(_) => fail("a grown source must not be copied verbatim")
  }

  test("writeToBytes(wb, config) honours the configuration: secure escapes a formula-shaped text") {
    val wb = Workbook(Vector(Sheet("Data").put(ref"A1", CellValue.Text("=1+1"))))
    val plain = partsText(orFail(XlsxWriter.writeToBytes(wb)))
    val secure = partsText(orFail(XlsxWriter.writeToBytes(wb, WriterConfig.secure)))
    val escaped = (s: String) => s.contains("'=1+1") || s.contains("&apos;=1+1")
    assert(plain.contains("=1+1") && !escaped(plain), plain)
    assert(escaped(secure), secure)
    assertEquals(
      orFail(XlsxWriter.writeToBytes(wb, WriterConfig.default)).toVector,
      orFail(XlsxWriter.writeToBytes(wb)).toVector,
      "the one-argument form is the default configuration"
    )
  }

  test("a modified read-back workbook regenerates in memory") {
    val source = orFail(XlsxWriter.writeToBytes(book()))
    val read = orFail(XlsxReader.readFromBytes(source))
    val edited = read.put(read.sheets(0).put(ref"C3", CellValue.Text("edited")))
    val bytes = orFail(XlsxWriter.writeToBytes(edited))
    assert(!bytes.sameElements(source), "an edit must not be copied verbatim")
    val back = orFail(XlsxReader.readFromBytes(bytes))
    assertEquals(back.sheets(0)(ref"C3").value, CellValue.Text("edited"))
    assertEquals(back.sheets(1)(ref"A1").value, CellValue.Text("shared"))
  }

  test(
    "GH-516: writeToBytes succeeds in a JVM whose java.io.tmpdir does not exist, byte-identical"
  ) {
    val java = Paths.get(System.getProperty("java.home"), "bin", "java").toString
    val classpath = System.getProperty("java.class.path")
    val missingTmp = tempDir.resolve("no-such-tmpdir").resolve("deeper")
    assert(!Files.exists(missingTmp))
    val childFresh = tempDir.resolve("child-fresh.xlsx")
    val childEdited = tempDir.resolve("child-edited.xlsx")
    val err = tempDir.resolve("child-stderr.txt")
    val builder = new ProcessBuilder(
      java,
      s"-Djava.io.tmpdir=$missingTmp",
      "-cp",
      classpath,
      "com.tjclp.xl.ooxml.WriteToBytesNoTmpdirMain",
      childFresh.toString,
      childEdited.toString
    )
    // a JVM that picks JAVA_TOOL_OPTIONS up announces it on stderr; the child never needs it
    builder.environment().remove("JAVA_TOOL_OPTIONS")
    val status = builder.redirectError(err.toFile).start().waitFor()
    val stderr = Try(Files.readString(err, StandardCharsets.UTF_8)).getOrElse("")
    assertEquals(status, 0, s"child JVM failed: $stderr")

    // the same workbooks written to disk from THIS JVM (a usable tmpdir) agree byte for byte
    val hereFresh = tempDir.resolve("here-fresh.xlsx")
    orFail(XlsxWriter.write(WriteToBytesNoTmpdirMain.fresh(), hereFresh))
    assert(
      Files.readAllBytes(childFresh).sameElements(Files.readAllBytes(hereFresh)),
      "fresh workbook: the no-tmpdir write must equal write(path)"
    )
    val hereEdited = tempDir.resolve("here-edited.xlsx")
    val editedWb = WriteToBytesNoTmpdirMain.edited().fold(fail(_), identity)
    orFail(XlsxWriter.write(editedWb, hereEdited))
    assert(
      Files.readAllBytes(childEdited).sameElements(Files.readAllBytes(hereEdited)),
      "edited read-back workbook: the no-tmpdir write must equal write(path)"
    )
    val back = orFail(XlsxReader.read(childEdited))
    assertEquals(back.sheets(0)(ref"C3").value, CellValue.Text("edited"))
  }

  test(
    "GH-516 follow-up: the byte sink hands back its own array when the size estimate was exact"
  ) {
    // exact estimate (a clean book's verbatim copy: the source fingerprint's size) → the buffer
    // itself, the same array on every call — no trimming copy, ~1x the archive in heap
    val exact = new XlsxWriter.SizedByteSink(5)
    exact.write(Array[Byte](1, 2, 3, 4, 5))
    assert(exact.result() eq exact.result(), "an exactly full buffer is returned as is")
    assertEquals(exact.result().toVector, Vector[Byte](1, 2, 3, 4, 5))
    // an inexact estimate → one trimmed copy of the bytes written
    val under = new XlsxWriter.SizedByteSink(8)
    under.write(Array[Byte](1, 2, 3))
    assertEquals(under.result().toVector, Vector[Byte](1, 2, 3))
    assert(under.result() ne under.result(), "a partly filled buffer is trimmed, not exposed")
    val over = new XlsxWriter.SizedByteSink(2)
    over.write(Array[Byte](1, 2, 3, 4, 5, 6, 7))
    assertEquals(over.result().toVector, Vector[Byte](1, 2, 3, 4, 5, 6, 7))
  }
