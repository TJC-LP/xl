package com.tjclp.xl.io

import java.io.{FilterInputStream, InputStream}
import java.nio.file.{Files, Path}
import java.util.concurrent.atomic.AtomicInteger
import java.util.zip.{ZipEntry, ZipFile}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.error.{XLError, XLException}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{DynamicArrayFixtures, SheetMetadataPart, XlsxReader}
import com.tjclp.xl.ooxml.style.WorkbookStyles

/** Metadata loaded beside a row stream obeys the reader's ZIP inflation limits. */
class CellMetadataSecuritySpec extends CatsEffectSuite:
  private val excel = ExcelIO.instance[IO]
  private val relsPath = "xl/_rels/workbook.xml.rels"
  private val metadataPath = "xl/custom/state.xml"

  private def fixture(parts: Vector[(String, String)]): Resource[IO, Path] =
    Resource.make(IO.blocking {
      val path = Files.createTempFile("xl-metadata-security-", ".xlsx")
      Files.write(path, DynamicArrayFixtures.xlsx(parts))
      path
    })(path => IO.blocking(Files.deleteIfExists(path)).void)

  private def security[A](result: Either[Throwable, A], part: String, reason: String): Unit =
    result match
      case Left(error: XLException) =>
        error.error match
          case XLError.SecurityError(message) =>
            assert(message.contains(part), message)
            assert(message.contains(reason), message)
          case other => fail(s"expected SecurityError, got $other")
      case other => fail(s"expected typed metadata security failure, got $other")

  private final class TrackedZip(path: Path) extends ZipFile(path.toFile):
    val opened = new AtomicInteger(0)
    val closed = new AtomicInteger(0)
    override def getInputStream(entry: ZipEntry): InputStream =
      opened.incrementAndGet()
      new FilterInputStream(super.getInputStream(entry)):
        override def close(): Unit =
          try super.close()
          finally closed.incrementAndGet()

  test("a byte cap refuses oversized metadata before parsing and closes both entry streams") {
    // Invalid XML would otherwise silently become an empty index. The limit must win first.
    val parts = DynamicArrayFixtures.relocatedMetadataParts(metadataPath, "x" * 8192)
    fixture(parts).use { path =>
      IO.blocking {
        val zip = new TrackedZip(path)
        try
          val result = excel.loadCellMetadataIndexSync(
            zip,
            XlsxReader.ReaderConfig.default
              .copy(maxUncompressedSize = 1024, maxCompressionRatio = 0)
          )
          security(result, metadataPath, "1024 bytes")
          assertEquals(zip.opened.get(), 2)
          assertEquals(zip.closed.get(), zip.opened.get())
        finally zip.close()
      }
    }
  }

  test("the relationship part uses the same byte cap and closes on refusal") {
    fixture(DynamicArrayFixtures.parts()).use { path =>
      IO.blocking {
        val zip = new TrackedZip(path)
        try
          security(
            excel.loadCellMetadataIndexSync(
              zip,
              XlsxReader.ReaderConfig.default
                .copy(maxUncompressedSize = 64, maxCompressionRatio = 0)
            ),
            relsPath,
            "64 bytes"
          )
          assertEquals(zip.opened.get(), 1)
          assertEquals(zip.closed.get(), zip.opened.get())
        finally zip.close()
      }
    }
  }

  test("normal metadata still resolves under configured limits and closes its streams") {
    fixture(DynamicArrayFixtures.relocatedMetadataParts(metadataPath)).use { path =>
      IO.blocking {
        val zip = new TrackedZip(path)
        try
          val result = excel.loadCellMetadataIndexSync(
            zip,
            XlsxReader.ReaderConfig.default.copy(maxUncompressedSize = 2048)
          )
          assertEquals(result, Right(SheetMetadataPart.parse(DynamicArrayFixtures.metadataXml)))
          assertEquals(zip.opened.get(), 2)
          assertEquals(zip.closed.get(), zip.opened.get())
        finally zip.close()
      }
    }
  }

  for part <- List(metadataPath, relsPath) do
    test(s"row and single-cell entry points retain typed ratio failures for $part") {
      val parts = DynamicArrayFixtures.relocatedMetadataParts(metadataPath).map { (name, xml) =>
        if name == part then name -> (xml + (" " * 200000)) else name -> xml
      }
      fixture(parts).use { path =>
        for
          rows <- excel.readStream(path).compile.toVector.attempt
          named <- excel.readSheetStream(path, "Sheet1", None).compile.toVector.attempt
          cell <- excel.streamCellDetails(path, "Sheet1", ref"B1").attempt
          preloaded <- excel
            .streamCellDetails(path, "Sheet1", ref"B1", None, WorkbookStyles.default)
            .attempt
        yield
          security(rows, part, "Compression ratio")
          security(named, part, "Compression ratio")
          security(cell, part, "Compression ratio")
          security(preloaded, part, "Compression ratio")
      }
    }

  test("withStreamLimits carries a caller's lifted limits to the rels and metadata parts") {
    val parts = DynamicArrayFixtures.relocatedMetadataParts(metadataPath).map { (name, xml) =>
      if name == metadataPath then name -> (xml + (" " * 200000)) else name -> xml
    }
    val lifted = excel.withStreamLimits(XlsxReader.ReaderConfig.permissive)
    fixture(parts).use { path =>
      for
        refused <- excel.readSheetStream(path, "Sheet1", None).compile.toVector.attempt
        rows <- lifted.readSheetStream(path, "Sheet1", None).compile.toVector
        cell <- lifted.streamCellDetails(path, "Sheet1", ref"B1")
        dir <- IO.blocking(Files.createTempDirectory("xl-spill-"))
      yield
        security(refused, metadataPath, "Compression ratio")
        assert(rows.nonEmpty)
        // B1 resolves through the metadata part the default limits refused
        assert(cell.value.toString.contains("Dynamic"), cell.value.toString)
        // the limits survive withSpillDir, and withStreamLimits keeps the spill directory
        assertEquals(lifted.withSpillDir(dir).streamLimits, XlsxReader.ReaderConfig.permissive)
        val both = excel.withSpillDir(dir).withStreamLimits(XlsxReader.ReaderConfig.permissive)
        assertEquals(both.spillDir, Some(dir))
    }
  }

  /** `bytes` with the first byte of `name`'s compressed data set to an invalid deflate block. */
  private def corrupted(bytes: Array[Byte], name: String): Array[Byte] =
    val copy = bytes.clone()
    val needle = name.getBytes(java.nio.charset.StandardCharsets.UTF_8)
    def u16(i: Int): Int = (copy(i) & 0xff) | ((copy(i + 1) & 0xff) << 8)
    // the entry's local header: PK\u0003\u0004, the name at offset 30
    val at = (0 until copy.length - 30 - needle.length)
      .find { i =>
        copy(i) == 0x50 && copy(i + 1) == 0x4b && copy(i + 2) == 3 && copy(i + 3) == 4 &&
        java.util.Arrays.equals(copy.slice(i + 30, i + 30 + needle.length), needle)
      }
      .getOrElse(fail(s"no local header for $name"))
    // BFINAL=1, BTYPE=11: a reserved block type, which the inflater rejects
    copy(at + 30 + u16(at + 26) + u16(at + 28)) = 0xff.toByte
    copy

  test("a damaged metadata entry reads as no index, and streamed reads go on") {
    val bytes =
      corrupted(DynamicArrayFixtures.xlsx(DynamicArrayFixtures.parts()), "xl/metadata.xml")
    val written = Resource.make(IO.blocking {
      val path = Files.createTempFile("xl-metadata-corrupt-", ".xlsx")
      Files.write(path, bytes)
      path
    })(path => IO.blocking(Files.deleteIfExists(path)).void)
    written.use { path =>
      for
        // the entry really is damaged: inflating it fails
        raw <- IO.blocking {
          val zip = new ZipFile(path.toFile)
          try
            val in = zip.getInputStream(zip.getEntry("xl/metadata.xml"))
            try in.readAllBytes().length
            finally in.close()
          finally zip.close()
        }.attempt
        index <- IO.blocking {
          val zip = new ZipFile(path.toFile)
          try excel.loadCellMetadataIndexSync(zip)
          finally zip.close()
        }
        rows <- excel.readStream(path).compile.toVector
        named <- excel.readSheetStream(path, "Sheet1", None).compile.toVector
        cell <- excel.streamCellDetails(path, "Sheet1", ref"B1")
      yield
        assert(raw.left.exists(_.isInstanceOf[java.util.zip.ZipException]), raw.toString)
        assertEquals(index, Right(com.tjclp.xl.ooxml.CellMetadataIndex.empty))
        assert(rows.nonEmpty)
        assertEquals(named.size, rows.size)
        // without its index, B1's cm resolves to nothing: the legacy array it then reads as
        assert(!cell.value.toString.contains("Dynamic"), cell.value.toString)
    }
  }

  test("pinned: a metadata part over the ratio limit fails a streamed read of sound worksheets") {
    // Today's behaviour (PR #721 review): the whole read fails, though only the dynamic-array
    // index is at stake; withStreamLimits (above) is the caller's way through.
    val parts = DynamicArrayFixtures.relocatedMetadataParts(metadataPath).map { (name, xml) =>
      if name == metadataPath then name -> (xml + (" " * 200000)) else name -> xml
    }
    fixture(parts).use { path =>
      for
        ratios <- IO.blocking {
          val zip = new ZipFile(path.toFile)
          try
            Vector("xl/worksheets/sheet1.xml", metadataPath).map { name =>
              val entry = zip.getEntry(name)
              name -> entry.getSize.toDouble / entry.getCompressedSize.max(1L)
            }.toMap
          finally zip.close()
        }
        rows <- excel.readStream(path).compile.toVector.attempt
      yield
        assert(ratios("xl/worksheets/sheet1.xml") < 100, ratios.toString)
        assert(ratios(metadataPath) > 100, ratios.toString)
        security(rows, metadataPath, "Compression ratio")
    }
  }
