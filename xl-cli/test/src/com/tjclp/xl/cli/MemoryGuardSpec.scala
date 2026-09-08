package com.tjclp.xl.cli

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipInputStream, ZipOutputStream}

import scala.jdk.CollectionConverters.*

import cats.effect.{IO, Resource}
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cli.contract.{CliError, CliException, ErrorCode, Location, Warning, WarningCode}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{WriterConfig, XlsxReader}
import com.tjclp.xl.ooxml.XlsxReader.ReaderConfig

/**
 * GH-636: the in-memory load under the memory guard. An `OutOfMemoryError` is fatal to cats-effect
 * (the runtime halts the process with a stack trace and exit 1 — it never reaches
 * `handleErrorWith`), so the guard catches it INSIDE the thunk that raised it and re-raises the
 * typed `RESOURCE_LIMIT` failure; and before a load whose `--max-size` lifted the default, it sizes
 * the worksheet XML and answers in two bands — refuse what cannot fit even at the measured
 * lower-bound ratio, warn (`MEMORY_PRESSURE`) about what may not fit at the upper-bound one.
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
    // the reader treats a negative limit as none (`> 0` guards every check): so does the guard
    assert(MemoryGuard.raised(default.copy(maxUncompressedSize = -1L)), "negative = no limit")
  }

  /** Pin a fault for `(path, stage)` around `io`, clearing it however `io` ends. */
  private def withFault[A](path: Path, stage: String, fault: Throwable)(io: IO[A]): IO[A] =
    val key = (Some(path), stage)
    IO(MemoryGuard.Seams.faults.updateAndGet(_ + (key -> fault))).void
      .bracket(_ => io)(_ => IO(MemoryGuard.Seams.faults.updateAndGet(_ - key)).void)

  test("blocking(path, stage): a fault pinned for that file and stage is raised inside the guard") {
    val staged = Path.of("staged.xlsx")
    withFault(staged, "recalc", new OutOfMemoryError("test")) {
      for
        hit <- MemoryGuard.blocking(staged, "recalc")(42).attempt
        otherStage <- MemoryGuard.blocking(staged, "load")(42)
        otherFile <- MemoryGuard.blocking(Path.of("other.xlsx"), "recalc")(42)
      yield
        assertEquals(resourceLimit(hit).error, MemoryGuard.exhausted)
        assertEquals(otherStage, 42)
        assertEquals(otherFile, 42)
    } *> MemoryGuard.blocking(staged, "recalc")(42).map(v => assertEquals(v, 42, "fault cleared"))
  }

  test("heapFor: the process heap, unless a test pinned one for that file") {
    val pinned = Path.of("pinned.xlsx")
    assertEquals(MemoryGuard.heapFor(pinned), MemoryGuard.maxHeapBytes)
    MemoryGuard.Seams.heap.updateAndGet(_ + (pinned -> 123L))
    try assertEquals(MemoryGuard.heapFor(pinned), 123L)
    finally MemoryGuard.Seams.heap.updateAndGet(_ - pinned)
    assertEquals(MemoryGuard.heapFor(pinned), MemoryGuard.maxHeapBytes)
  }

  private val nycPath = Path.of("nyc1m.xlsx")

  private def refusal(v: MemoryGuard.Verdict): CliError = v match
    case MemoryGuard.Verdict.Refuse(err) => err
    case other => fail(s"expected a refusal, got $other")

  private def warning(v: MemoryGuard.Verdict): Warning = v match
    case MemoryGuard.Verdict.Warn(w) => w
    case other => fail(s"expected a warning, got $other")

  test("the bands are ordered: each lower coefficient is the measured floor, below its upper") {
    assert(MemoryGuard.sheetLower >= 1 && MemoryGuard.sheetLower < MemoryGuard.sheetUpper)
    assert(MemoryGuard.sstLower >= 1 && MemoryGuard.sstLower < MemoryGuard.sstUpper)
    assert(MemoryGuard.sstLower < MemoryGuard.sheetLower, "a string costs far less than a cell")
  }

  test("decide: lower bound over the heap refuses; upper bound over it warns; below both admits") {
    val xml = 100L << 20 // 100 MB of worksheet XML, no shared strings
    val lower = xml * MemoryGuard.sheetLower
    val upper = xml * MemoryGuard.sheetUpper
    // one byte short of the lower bound: hopeless, refused
    val err = refusal(MemoryGuard.decide(nycPath, xml, lower - 1))
    assertEquals(err.code, ErrorCode.RESOURCE_LIMIT)
    assertEquals(err.hint, Some(MemoryGuard.hint))
    assertEquals(err.location, Some(Location.file("nyc1m.xlsx")))
    assert(err.message.contains("does not fit"), err.message)
    assert(err.message.contains(MemoryGuard.human(xml)), err.message)
    assert(err.message.contains(MemoryGuard.human(lower)), err.message)
    assert(err.message.contains(MemoryGuard.human(lower - 1)), err.message)
    assertEquals(err.exitCode.code, 3)
    // exactly the lower bound: not hopeless, but the upper bound is over — a warning
    val warn = warning(MemoryGuard.decide(nycPath, xml, lower))
    assertEquals(warn.code, WarningCode.MEMORY_PRESSURE)
    assertEquals(warn.location, Some(Location.file("nyc1m.xlsx")))
    assert(warn.message.contains("may not fit"), warn.message)
    assert(warn.message.contains(MemoryGuard.human(lower)), warn.message)
    assert(warn.message.contains(MemoryGuard.human(upper)), warn.message)
    assert(warn.message.contains("RESOURCE_LIMIT") && warn.message.contains(MemoryGuard.hint))
    // one byte short of the upper bound still warns; at the upper bound the load is silent
    assertEquals(
      warning(MemoryGuard.decide(nycPath, xml, upper - 1)).code,
      WarningCode.MEMORY_PRESSURE
    )
    assertEquals(MemoryGuard.decide(nycPath, xml, upper), MemoryGuard.Verdict.Admit)
    // an unbounded heap (Runtime.maxMemory reports Long.MaxValue) and nothing to load never refuse
    assertEquals(MemoryGuard.decide(nycPath, xml, Long.MaxValue), MemoryGuard.Verdict.Admit)
    assertEquals(MemoryGuard.decide(nycPath, 0L, 1L), MemoryGuard.Verdict.Admit)
  }

  // --- the measured shapes (bytes from each book's central directory) -------------------------

  private val GiB = 1L << 30
  private val MiB = 1L << 20
  import MemoryGuard.{Footprint, Verdict}

  /** (a) 1,000,000 x 8 decimals: loads at 6144m under G1 (8192m ParallelGC), fails at 5120m. */
  private val numeric = Footprint(sheetBytes = 362_929_341L, sstBytes = 154L)

  /** (c) 1,000,000 x 8 short integers: loads at 8192m, fails at 6144m (ParallelGC). */
  private val ints = Footprint(sheetBytes = 324_120_309L, sstBytes = 154L)

  /** (b) 1,000,000 x 8 with 4 distinct-text columns: loads at 8192m, fails at 7168m. */
  private val text = Footprint(sheetBytes = 358_853_715L, sstBytes = 111_555_755L)

  /** (d) 1,000,000 x 1 distinct 300-char text: loads at 1792m, fails at 1536m (both collectors). */
  private val longText = Footprint(sheetBytes = 62_666_927L, sstBytes = 317_889_067L)

  /** nyc1m, the 0.21.0 dogfood: 1.09 GB of sheet XML, needed 36–45 GB of heap. */
  private val nyc1m = Footprint(sheetBytes = 1_090_000_000L, sstBytes = 0L)

  private def verdictOf(v: Verdict): String = v match
    case Verdict.Admit => "admit"
    case Verdict.Warn(_) => "warn"
    case Verdict.Refuse(_) => "refuse"

  private def assertVerdict(fp: Footprint, heap: Long, expected: String, why: String): Unit =
    assertEquals(
      verdictOf(MemoryGuard.decide(nyc1m0, fp, heap)),
      expected,
      s"$why at ${MemoryGuard.human(heap)}"
    )

  private val nyc1m0 = Path.of("book.xlsx")

  test(
    "manifest (a) numeric 1M x 8: never refused where it loads; warned on 8 GB; silent on 16 GB"
  ) {
    assertVerdict(numeric, 4 * GiB, "refuse", "fails at 4g under both collectors")
    assertVerdict(
      numeric,
      6 * GiB,
      "warn",
      "loads at 6144m under G1 — a refusal here would be false"
    )
    assertVerdict(numeric, 8 * GiB, "warn", "the native image's default heap, where it loads today")
    assertVerdict(numeric, 16 * GiB, "admit", "twice the upper estimate's need")
  }

  test("manifest (c) short integers 1M x 8: the same bands as the decimals") {
    assertVerdict(ints, 4 * GiB, "refuse", "fails at 4g")
    assertVerdict(ints, 8 * GiB, "warn", "loads at 8192m")
    assertVerdict(ints, 16 * GiB, "admit", "")
  }

  test(
    "manifest (b) text 1M x 8, 4 distinct-text columns: SST bytes are charged at the string rate"
  ) {
    assertVerdict(text, 4 * GiB, "refuse", "fails at 4g")
    assertVerdict(text, 8 * GiB, "warn", "loads at 8192m")
    assertVerdict(text, 16 * GiB, "admit", "")
    val warn = warning(MemoryGuard.decide(nyc1m0, text, 8 * GiB))
    assert(warn.message.contains("shared strings"), warn.message)
    assert(warn.message.contains(MemoryGuard.human(text.sstBytes)), warn.message)
  }

  test(
    "manifest (d) long text: 380 MB of XML that loads in 1792m is never refused on a 2 GB heap"
  ) {
    assertVerdict(longText, 1 * GiB, "refuse", "fails at 1g")
    assertVerdict(longText, 1792 * MiB, "warn", "the smallest heap it loads in")
    assertVerdict(longText, 2 * GiB, "warn", "")
    assertVerdict(longText, 4 * GiB, "admit", "")
    // a single sheet-XML multiplier could not do this: 380 MB × 14 would refuse on a 4 GB heap
    assert(longText.total * MemoryGuard.sheetLower > 4 * GiB, "the split model is load-bearing")
  }

  test(
    "manifest nyc1m (1.09 GB of sheet XML, needed 36–45 GB): refused on 8 GB, warned on 16, silent on 64"
  ) {
    val refused = refusal(MemoryGuard.decide(nyc1m0, nyc1m, 8 * GiB))
    assert(refused.message.contains("does not fit"), refused.message)
    assert(!refused.message.contains("shared strings"), refused.message)
    assertVerdict(nyc1m, 16 * GiB, "warn", "the lower estimate fits, the upper does not")
    assertVerdict(nyc1m, 64 * GiB, "admit", "the documented -Xmx64g override")
  }

  test("footprint: worksheet and shared-string bytes are split; total is footprintBytes") {
    fixture("book.xlsx").flatMap { path =>
      val expectedSst = IO.blocking {
        val zip = new ZipFile(path.toFile)
        try Option(zip.getEntry("xl/sharedStrings.xml")).fold(0L)(_.getSize)
        finally zip.close()
      }
      (MemoryGuard.footprint(path), MemoryGuard.footprintBytes(path), expectedSst).mapN {
        (fp, total, sst) =>
          assert(fp.sheetBytes > 0L, "the fixture has worksheet XML")
          assertEquals(fp.sstBytes, sst, "the shared-string bytes are the part's own size")
          assertEquals(fp.total, total)
      } *> MemoryGuard.footprint(Path.of("/nonexistent.xlsx")).map { missing =>
        assertEquals(missing, MemoryGuard.Footprint.empty)
      }
    }
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

  test("admit: inactive at the default limit; a lifted load is refused, warned or passed by band") {
    fixture("book.xlsx").flatMap { path =>
      MemoryGuard.footprint(path).flatMap { fp =>
        val doubtful = (fp.atLeast + fp.upTo) / 2 // between the bands: the load may not fit
        for
          seen <- IO.ref(Vector.empty[Warning])
          sink = (w: Warning) => seen.update(_ :+ w)
          default <- MemoryGuard.admit(path, ReaderConfig.default, heap = 1L, warn = sink).attempt
          tiny <- MemoryGuard.admit(path, ReaderConfig.permissive, heap = 1L, warn = sink).attempt
          afterRefusal <- seen.get
          warned <- MemoryGuard
            .admit(path, ReaderConfig.permissive, heap = doubtful, warn = sink)
            .attempt
          afterWarning <- seen.get
          unbounded <- MemoryGuard
            .admit(path, ReaderConfig.permissive, heap = Long.MaxValue, warn = sink)
            .attempt
          real <- MemoryGuard.admit(path, ReaderConfig.permissive, warn = sink).attempt
          afterAll <- seen.get
        yield
          assertEquals(default, Right(()), "the default limit never arms the pre-load guard")
          val e = resourceLimit(tiny)
          assertEquals(e.error.location, Some(Location.file(path.toString)))
          assertEquals(afterRefusal, Vector.empty, "a refusal is not also a warning")
          assertEquals(warned, Right(()), "a doubtful load proceeds")
          assertEquals(afterWarning.map(_.code), Vector(WarningCode.MEMORY_PRESSURE))
          assertEquals(unbounded, Right(()))
          assertEquals(real, Right(()), "a few KB of XML fits any real heap")
          assertEquals(afterAll.size, 1, "silent admissions warn nothing")
      }
    }
  }

  test("admitAll: books loaded together are sized as one — a verdict on the sum, naming both") {
    for
      a <- fixture("book.xlsx")
      b <- fixture("book2.xlsx")
      fa <- MemoryGuard.footprint(a)
      // each alone is below the refusal band at this heap; the two together are over it
      heap = fa.atLeast * 3 / 2
      seen <- IO.ref(Vector.empty[Warning])
      sink = (w: Warning) => seen.update(_ :+ w)
      alone <- MemoryGuard.admit(a, ReaderConfig.permissive, heap, sink).attempt
      together <- MemoryGuard.admitAll(Vector(a, b), ReaderConfig.permissive, heap, sink).attempt
      atDefault <- MemoryGuard.admitAll(Vector(a, b), ReaderConfig.default, 1L, sink).attempt
      none <- MemoryGuard.admitAll(Vector.empty, ReaderConfig.permissive, 1L, sink).attempt
      warnings <- seen.get
    yield
      assertEquals(alone, Right(()), "one book fits the lower band")
      val e = resourceLimit(together)
      assert(
        e.error.message.contains(a.toString) && e.error.message.contains(b.toString),
        e.error.message
      )
      assertEquals(e.error.location, Some(Location.file(s"$a, $b")))
      assertEquals(atDefault, Right(()), "the default limit never arms the guard")
      assertEquals(none, Right(()))
      assertEquals(
        warnings.map(_.code),
        Vector(WarningCode.MEMORY_PRESSURE),
        "alone: the upper band"
      )
  }

  test("writer: serialisation runs under the guard — writeWith, writeWorkbookStream and write") {
    val out = dir().resolve("guarded-write.xlsx")
    withFault(out, "write", new OutOfMemoryError("test")) {
      for
        viaWith <- MemoryGuard.writer.writeWith(book(), out, WriterConfig.default).attempt
        viaStream <- MemoryGuard.writer
          .writeWorkbookStream(book(), out, WriterConfig.default)
          .attempt
        viaWrite <- MemoryGuard.writer.write(book(), out).attempt
      yield
        assertEquals(resourceLimit(viaWith).error, MemoryGuard.exhausted)
        assertEquals(resourceLimit(viaStream).error, MemoryGuard.exhausted)
        assertEquals(resourceLimit(viaWrite).error, MemoryGuard.exhausted)
        assert(!Files.exists(out), "nothing written")
    } *> MemoryGuard.writer.writeWith(book(), out, WriterConfig.default).map { _ =>
      assert(Files.exists(out), "the fault cleared, the write lands")
    }
  }

  test("writer: an unwritable target fails with exactly the message ExcelIO.writeWith produces") {
    val target = dir().resolve("no-such-dir").resolve("out.xlsx")
    for
      guarded <- MemoryGuard.writer.writeWith(book(), target, WriterConfig.default).attempt
      library <- ExcelIO.instance[IO].writeWith(book(), target, WriterConfig.default).attempt
    yield (guarded, library) match
      case (Left(g), Left(l)) =>
        assertEquals(g.getMessage, l.getMessage)
        assert(g.getMessage.startsWith("Failed to write XLSX: "), g.getMessage)
      case other => fail(s"both writes must fail the same way, got $other")
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

  test(
    "excel: a lifted load that MAY not fit proceeds under MEMORY_PRESSURE, sent to the warn sink"
  ) {
    fixture("book.xlsx").flatMap { path =>
      MemoryGuard.footprint(path).flatMap { fp =>
        val doubtful = (fp.atLeast + fp.upTo) / 2 // between the bands: the load may not fit
        for
          seen <- IO.ref(Vector.empty[Warning])
          guarded = MemoryGuard.excel(
            _ => IO.unit,
            warn = w => seen.update(_ :+ w),
            heap = doubtful
          )
          wb <- guarded.readWith(path, ReaderConfig.permissive)
          warnings <- seen.get
        yield
          assertEquals(wb.sheets.map(_.name.value), Vector("Data", "Other"), "the load ran")
          assertEquals(warnings.map(_.code), Vector(WarningCode.MEMORY_PRESSURE))
          assertEquals(warnings.headOption.flatMap(_.location), Some(Location.file(path.toString)))
      }
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
