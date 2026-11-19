package com.tjclp.xl.ooxml

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path, Paths}
import java.security.MessageDigest
import java.util.zip.ZipInputStream

import scala.jdk.CollectionConverters.*

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.ooxml.writer.{SstPolicy, WriterConfig, XmlBackend}

/**
 * The bytes the writer emits for the fixture corpus, pinned (#891).
 *
 * Every workbook under `xl-ooxml/test/resources/fixtures/` is read, detached from its source (so
 * nothing is copied through: every part is regenerated from the domain model) and written with each
 * XML backend, under the default shared-string policy and with the table forced (the corpus is
 * small, so the default inlines most of its text). The SHA-256 of every worksheet part, of
 * `xl/sharedStrings.xml` and of `xl/styles.xml` must equal the digest recorded in
 * `xl-ooxml/test/resources/written-part-digests.txt`.
 *
 * The point is order: worksheets, the shared-string table and the style table are emitted by
 * walking a sheet's cells, so a change to the order cells are stored or iterated in that reaches
 * the output changes these bytes and fails here, where a read-back comparison (which sees the same
 * cells whatever their order in the XML) would pass. Any other change to what the writer emits
 * fails here too, by design.
 *
 * A digest diff is a change to written files: review it, then re-record deliberately with
 * `XL_UPDATE_DIGESTS=1 ./mill xl-ooxml.test.testOnly com.tjclp.xl.ooxml.WrittenPartDigestSpec` and
 * say in the CHANGELOG what changed in the output. The fixtures are listed from the directory, so a
 * new one fails until its digests are recorded. `malformed-workbook.xlsx` (not well-formed on
 * purpose, [[TestFixtures.malformed]]) is the one file skipped.
 */
class WrittenPartDigestSpec extends FunSuite:
  import WrittenPartDigestSpec.*

  private val update = sys.env.get("XL_UPDATE_DIGESTS").contains("1")

  private val fixtures: Vector[String] = listFixtures.filterNot(TestFixtures.malformed.contains)

  /** Every fixture's lines, computed once. */
  private lazy val actual: Map[String, Vector[String]] =
    fixtures.map(name => name -> digestLines(name)).toMap

  private lazy val recorded: Map[String, Vector[String]] =
    if !Files.isRegularFile(digestFile) then Map.empty
    else
      Files
        .readAllLines(digestFile, StandardCharsets.UTF_8)
        .asScala
        .toVector
        .filterNot(line => line.isEmpty || line.startsWith("#"))
        .groupBy(_.takeWhile(_ != ' '))

  test("the corpus is present and every fixture is classified") {
    assert(fixtures.nonEmpty, s"no *.xlsx under $fixtureDir")
    assertEquals(
      listFixtures.toSet,
      (TestFixtures.all ++ TestFixtures.malformed).toSet,
      "a fixture on disk is missing from TestFixtures (or the reverse)"
    )
  }

  test("written-part-digests.txt records exactly the fixtures on disk") {
    if update then
      val body = fixtures.flatMap(actual).mkString("", "\n", "\n")
      Files.writeString(digestFile, Header + body, StandardCharsets.UTF_8)
    else
      assertEquals(
        recorded.keySet,
        fixtures.toSet,
        s"$digestFile and the fixture directory differ (re-record with XL_UPDATE_DIGESTS=1)"
      )
  }

  fixtures.foreach { name =>
    test(s"written bytes: $name") {
      if !update then
        val expected = recorded.getOrElse(name, Vector.empty)
        val got = actual(name)
        val changed = got.filterNot(expected.contains).map(line => s"  now      $line") ++
          expected.filterNot(got.contains).map(line => s"  recorded $line")
        assert(
          changed.isEmpty,
          s"the bytes written for $name changed (review the output, then re-record deliberately " +
            s"with XL_UPDATE_DIGESTS=1 and add a CHANGELOG line):\n${changed.mkString("\n")}"
        )
    }
  }

  test("writing a fixture twice gives the same digests") {
    fixtures.take(3).foreach(name => assertEquals(digestLines(name), actual(name), name))
  }

object WrittenPartDigestSpec:

  private val Header =
    """# SHA-256 of the parts XlsxWriter emits for each fixture, read and written from scratch
      |# (no source context), per XML backend and shared-string policy. Generated: do not edit by
      |# hand. Re-record with
      |#   XL_UPDATE_DIGESTS=1 ./mill xl-ooxml.test.testOnly com.tjclp.xl.ooxml.WrittenPartDigestSpec
      |# and review the diff: a changed digest is a change to written files.
      |# <fixture> <backend>/<sstPolicy> <part> <sha256>
      |""".stripMargin

  /** The checkout root: the nearest ancestor of the working directory holding `build.mill`. */
  private val repoRoot: Path =
    val start = Paths.get("").toAbsolutePath
    Iterator
      .iterate(Option(start))(_.flatMap(p => Option(p.getParent)))
      .takeWhile(_.isDefined)
      .flatten
      .find(dir => Files.exists(dir.resolve("build.mill")))
      .getOrElse(
        throw new IllegalStateException(s"no build.mill at or above $start; cannot locate fixtures")
      )

  val fixtureDir: Path = repoRoot.resolve("xl-ooxml/test/resources/fixtures")
  val digestFile: Path = repoRoot.resolve("xl-ooxml/test/resources/written-part-digests.txt")

  /** Every `*.xlsx` in the fixture directory, sorted by name. */
  def listFixtures: Vector[String] =
    if !Files.isDirectory(fixtureDir) then Vector.empty
    else
      val listing = Files.list(fixtureDir)
      try
        listing.iterator.asScala
          .map(_.getFileName.toString)
          .filter(_.endsWith(".xlsx"))
          .toVector
          .sorted
      finally listing.close()

  private def pinned(entry: String): Boolean =
    entry.startsWith("xl/worksheets/") && entry.endsWith(".xml") && !entry.contains("_rels") ||
      entry == "xl/sharedStrings.xml" || entry == "xl/styles.xml"

  private def sha256(bytes: Array[Byte]): String =
    MessageDigest.getInstance("SHA-256").digest(bytes).map(b => f"${b & 0xff}%02x").mkString

  /** The pinned parts of a written workbook: entry name to digest, sorted by name. */
  private def partDigests(xlsx: Array[Byte]): Vector[(String, String)] =
    val zip = new ZipInputStream(new ByteArrayInputStream(xlsx))
    try
      Iterator
        .continually(zip.getNextEntry)
        .takeWhile(_ != null)
        .filter(entry => pinned(entry.getName))
        .map(entry => entry.getName -> sha256(zip.readAllBytes()))
        .toVector
        .sortBy(_._1)
    finally zip.close()

  /**
   * `<fixture> <backend>/<sstPolicy> <part> <sha256>` for every pinned part, each variant in turn.
   */
  def digestLines(name: String): Vector[String] =
    val read = XlsxReader.read(fixtureDir.resolve(name)) match
      case Right(wb) => wb.copy(sourceContext = None)
      case Left(err) => throw new IllegalStateException(s"$name failed to read: ${err.message}")
    val variants =
      for
        backend <- Vector(XmlBackend.ScalaXml, XmlBackend.SaxStax)
        policy <- Vector(SstPolicy.Auto, SstPolicy.Always)
      yield (backend, policy)
    variants.flatMap { (backend, policy) =>
      XlsxWriter.writeToBytes(read, WriterConfig(backend = backend, sstPolicy = policy)) match
        case Left(err) => throw new IllegalStateException(s"$name failed to write: ${err.message}")
        case Right(bytes) =>
          partDigests(bytes).map((part, digest) => s"$name $backend/$policy $part $digest")
    }
