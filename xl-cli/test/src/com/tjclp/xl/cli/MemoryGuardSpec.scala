package com.tjclp.xl.cli

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipInputStream, ZipOutputStream}

import scala.jdk.CollectionConverters.*

import cats.effect.{IO, Resource}
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cli.contract.{CliException, ErrorCode, Location}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.XlsxReader
import com.tjclp.xl.ooxml.XlsxReader.ReaderConfig

/**
 * GH-636: the in-memory load under the memory guard. An `OutOfMemoryError` is fatal to cats-effect
 * (the runtime halts the process with a stack trace and exit 1 — it never reaches
 * `handleErrorWith`), so the guard catches it INSIDE the thunk that raised it and re-raises the
 * typed `RESOURCE_LIMIT` failure; and before a load whose `--max-size` lifted the default, it
 * estimates the footprint from the worksheet part sizes and refuses what cannot fit.
 */
class MemoryGuardSpec extends CatsEffectSuite:

  private val dir = ResourceSuiteLocalFixture(
    "memory-guard-fixtures",
    Resource.make(IO.blocking(Files.createTempDirectory("xl-memory-guard-")))(d =>
      IO.blocking {
        Files.list(d).iterator.asScala.toVector.foreach(Files.deleteIfExists)
        Files.deleteIfExists(d)
        ()
      }
    )
  )

  override def munitFixtures = List(dir)

  private def book(): Workbook =
    Workbook(
      Vector(
        Sheet("Data").put(ref"A1", "hello").put(ref"B1", 42).put(ref"A2", "shared text"),
        Sheet("Other").put(ref"A1", "shared text")
      )
    )

  private def fixture(name: String): IO[Path] =
    val path = dir().resolve(name)
    IO.blocking(Files.exists(path)).flatMap {
      case true => IO.pure(path)
      case false => ExcelIO.instance[IO].write(book(), path).as(path)
    }

  /** `source` with `xl/styles.xml` dropped, so the reader emits `MissingStylesXml`. */
  private def withoutStyles(source: Path, target: Path): IO[Path] = IO.blocking {
    val in = new ZipInputStream(Files.newInputStream(source))
    val out = new ZipOutputStream(Files.newOutputStream(target))
    try
      Iterator
        .continually(in.getNextEntry)
        .takeWhile(_ != null)
        .filterNot(_.getName == "xl/styles.xml")
        .foreach { entry =>
          out.putNextEntry(new ZipEntry(entry.getName))
          out.write(in.readAllBytes())
          out.closeEntry()
        }
    finally
      in.close()
      out.close()
    target
  }

  private def resourceLimit(attempt: Either[Throwable, Any]): CliException = attempt match
    case Left(e: CliException) if e.error.code == ErrorCode.RESOURCE_LIMIT => e
    case other => fail(s"expected the RESOURCE_LIMIT CliException, got $other")

  // --- blocking: the OutOfMemoryError seam ------------------------------------------------------

  test("blocking: an OutOfMemoryError inside the thunk is the RESOURCE_LIMIT failure, not fatal") {
    MemoryGuard.blocking[Int](throw new OutOfMemoryError("test")).attempt.map { attempt =>
      val e = resourceLimit(attempt)
      assertEquals(e.error, MemoryGuard.exhausted)
      assert(e.error.message.contains("out of memory"), e.error.message)
      assert(
        e.error.message.contains(MemoryGuard.human(MemoryGuard.maxHeapBytes)),
        s"the message must name the heap: ${e.error.message}"
      )
      assertEquals(e.error.hint, Some(MemoryGuard.hint))
      assert(MemoryGuard.hint.contains("--stream"), MemoryGuard.hint)
      assert(MemoryGuard.hint.contains("-Xmx"), MemoryGuard.hint)
      assertEquals(e.error.exitCode.code, 3)
    }
  }

  test("blocking: a value and a non-fatal exception pass through unchanged") {
    for
      value <- MemoryGuard.blocking(41 + 1)
      plain <- MemoryGuard.blocking[Int](throw new IllegalStateException("plain")).attempt
    yield
      assertEquals(value, 42)
      plain match
        case Left(e: IllegalStateException) => assertEquals(e.getMessage, "plain")
        case other => fail(s"expected the exception unchanged, got $other")
  }

  test("exhausted is pre-built: the same instance every time, so the catch allocates nothing") {
    assert(MemoryGuard.exhausted eq MemoryGuard.exhausted)
    assertEquals(MemoryGuard.exhausted.code, ErrorCode.RESOURCE_LIMIT)
  }

  // --- the pre-load estimate --------------------------------------------------------------------

  test("raised: only a --max-size above the default (or 0 = unlimited) arms the pre-load guard") {
    val default = ReaderConfig.default
    assert(!MemoryGuard.raised(default))
    assert(MemoryGuard.raised(ReaderConfig.permissive), "0 = unlimited lifts the limit")
    assert(MemoryGuard.raised(default.copy(maxUncompressedSize = 500L * 1024 * 1024)))
    assert(!MemoryGuard.raised(default.copy(maxUncompressedSize = 50L * 1024 * 1024)))
    assert(!MemoryGuard.raised(default.copy(maxUncompressedSize = default.maxUncompressedSize)))
  }

  test("decide: 30x the worksheet XML above 70% of the heap refuses; at or below proceeds") {
    val path = Path.of("nyc1m.xlsx")
    val xml = 1_090_000_000L // the dogfood book: 1.09 GB of sheet XML needed 36-45 GB of heap
    val eightGb = 8L << 30
    val refused = MemoryGuard.decide(path, xml, eightGb)
    refused match
      case Left(err) =>
        assertEquals(err.code, ErrorCode.RESOURCE_LIMIT)
        assertEquals(err.hint, Some(MemoryGuard.hint))
        assertEquals(err.location, Some(Location.file("nyc1m.xlsx")))
        assert(err.message.contains("nyc1m.xlsx"), err.message)
        assert(err.message.contains(MemoryGuard.human(xml)), err.message)
        assert(err.message.contains(MemoryGuard.human(xml * MemoryGuard.multiplier)), err.message)
        assert(err.message.contains(MemoryGuard.human(eightGb)), err.message)
        assertEquals(err.exitCode.code, 3)
      case Right(()) => fail("1.09 GB of sheet XML must be refused on an 8 GB heap")
    // 64 GB: 32.7 GB estimated against a 44.8 GB budget passes (the documented -Xmx64g override)
    assertEquals(MemoryGuard.decide(path, xml, 64L << 30), Right(()))
    // exactly at the budget passes; one byte over refuses
    val budget = (eightGb * MemoryGuard.budget).toLong
    val atBudget = budget / MemoryGuard.multiplier
    assertEquals(MemoryGuard.decide(path, atBudget, eightGb), Right(()))
    assert(MemoryGuard.decide(path, atBudget + 1, eightGb).isLeft)
    // an unbounded heap (Runtime.maxMemory reports Long.MaxValue) never refuses
    assertEquals(MemoryGuard.decide(path, xml, Long.MaxValue), Right(()))
    // nothing to load never refuses
    assertEquals(MemoryGuard.decide(path, 0L, 1L), Right(()))
  }

  test("human: binary units with one decimal, the way -Xmx counts") {
    assertEquals(MemoryGuard.human(8L << 30), "8.0 GB")
    assertEquals(MemoryGuard.human(512L << 20), "512.0 MB")
    assertEquals(MemoryGuard.human(1_090_000_000L), "1.0 GB")
    assertEquals(MemoryGuard.human(1536L), "1.5 KB")
    assertEquals(MemoryGuard.human(0L), "0 B")
  }

  test("footprintBytes: the worksheet and sharedStrings part sizes from the central directory") {
    fixture("book.xlsx").flatMap { path =>
      val expected = IO.blocking {
        val zip = new ZipFile(path.toFile)
        try
          zip
            .entries()
            .asScala
            .filter(e => MemoryGuard.counted(e.getName))
            .map(_.getSize)
            .sum
        finally zip.close()
      }
      (expected, MemoryGuard.footprintBytes(path)).mapN { (want, got) =>
        assert(want > 0L, "the fixture must have worksheet XML")
        assertEquals(got, want)
      }
    }
  }

  test("counted: worksheets and the shared-string table, nothing else") {
    assert(MemoryGuard.counted("xl/worksheets/sheet1.xml"))
    assert(MemoryGuard.counted("xl/worksheets/sheet12.xml"))
    assert(MemoryGuard.counted("xl/sharedStrings.xml"))
    assert(!MemoryGuard.counted("xl/worksheets/_rels/sheet1.xml.rels"))
    assert(!MemoryGuard.counted("xl/styles.xml"))
    assert(!MemoryGuard.counted("xl/workbook.xml"))
    assert(!MemoryGuard.counted("xl/media/image1.png"))
    assert(!MemoryGuard.counted("docProps/app.xml"))
  }

  test("footprintBytes: an unreadable or non-zip input counts as 0 — the read reports the error") {
    val corrupt = dir().resolve("corrupt.xlsx")
    for
      _ <- IO.blocking(Files.write(corrupt, "not a zip".getBytes(StandardCharsets.UTF_8)))
      notZip <- MemoryGuard.footprintBytes(corrupt)
      missing <- MemoryGuard.footprintBytes(dir().resolve("missing.xlsx"))
    yield
      assertEquals(notZip, 0L)
      assertEquals(missing, 0L)
  }

  test("admit: inactive at the default limit, refuses a lifted load that cannot fit, else passes") {
    fixture("book.xlsx").flatMap { path =>
      for
        default <- MemoryGuard.admit(path, ReaderConfig.default, heap = 1L).attempt
        tiny <- MemoryGuard.admit(path, ReaderConfig.permissive, heap = 1L).attempt
        unbounded <- MemoryGuard.admit(path, ReaderConfig.permissive, heap = Long.MaxValue).attempt
        real <- MemoryGuard.admit(path, ReaderConfig.permissive).attempt
      yield
        assertEquals(default, Right(()), "the default limit never arms the pre-load guard")
        val e = resourceLimit(tiny)
        assertEquals(e.error.location, Some(Location.file(path.toString)))
        assertEquals(unbounded, Right(()))
        assertEquals(real, Right(()), "a few KB of XML fits any real heap")
    }
  }

  // --- the guarded ExcelIO ----------------------------------------------------------------------

  test("excel: an OutOfMemoryError while parsing is RESOURCE_LIMIT, through readWith and read") {
    fixture("book.xlsx").flatMap { path =>
      val oom = MemoryGuard.excel(_ => IO.unit, parse = (_, _) => throw new OutOfMemoryError("x"))
      for
        viaWith <- oom.readWith(path, ReaderConfig.default).attempt
        viaRead <- oom.read(path).attempt
      yield
        assertEquals(resourceLimit(viaWith).error, MemoryGuard.exhausted)
        assertEquals(resourceLimit(viaRead).error, MemoryGuard.exhausted)
    }
  }

  test("excel: a lifted load that cannot fit is refused BEFORE the parser runs") {
    fixture("book.xlsx").flatMap { path =>
      val parsed = new IllegalStateException("parsed")
      val guarded = MemoryGuard.excel(_ => IO.unit, parse = (_, _) => throw parsed, heap = 1L)
      for
        lifted <- guarded.readWith(path, ReaderConfig.permissive).attempt
        default <- guarded.readWith(path, ReaderConfig.default).attempt
      yield
        val e = resourceLimit(lifted)
        assert(e.error.message.contains(path.toString), e.error.message)
        assertEquals(default, Left(parsed), "at the default limit the parser runs")
    }
  }

  test("excel: reads the workbook and routes reader warnings to the handler, like ExcelIO") {
    for
      path <- fixture("book.xlsx")
      nostyles <- withoutStyles(path, dir().resolve("nostyles.xlsx"))
      seen <- IO.ref(Vector.empty[XlsxReader.Warning])
      guarded = MemoryGuard.excel(w => seen.update(_ :+ w))
      wb <- guarded.readWith(path, ReaderConfig.default)
      _ <- guarded.readWith(nostyles, ReaderConfig.default)
      warnings <- seen.get
    yield
      assertEquals(wb.sheets.map(_.name.value), Vector("Data", "Other"))
      assertEquals(warnings, Vector(XlsxReader.Warning.MissingStylesXml))
  }

  test("excel: a corrupt file fails with exactly the message ExcelIO.readWith produces") {
    val corrupt = dir().resolve("corrupt2.xlsx")
    for
      _ <- IO.blocking(Files.write(corrupt, "not a zip".getBytes(StandardCharsets.UTF_8)))
      guarded <- MemoryGuard.excel(_ => IO.unit).readWith(corrupt, ReaderConfig.default).attempt
      library <- ExcelIO.instance[IO].readWith(corrupt, ReaderConfig.default).attempt
    yield (guarded, library) match
      case (Left(g), Left(l)) =>
        assertEquals(g.getMessage, l.getMessage)
        assert(g.getMessage.startsWith("Failed to read XLSX: "), g.getMessage)
      case other => fail(s"both reads must fail the same way, got $other")
  }
