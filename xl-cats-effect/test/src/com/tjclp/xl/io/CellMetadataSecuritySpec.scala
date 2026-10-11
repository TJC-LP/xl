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
