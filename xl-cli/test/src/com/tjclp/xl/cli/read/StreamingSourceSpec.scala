package com.tjclp.xl.cli.read

import java.nio.file.{Path, Paths}
import java.util.concurrent.ConcurrentLinkedQueue
import java.util.concurrent.atomic.AtomicInteger

import scala.concurrent.ExecutionContext
import scala.jdk.CollectionConverters.*

import cats.effect.IO
import cats.effect.unsafe.IORuntime
import fs2.Stream
import munit.CatsEffectSuite

import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.SharedStrings
import com.tjclp.xl.ooxml.XlsxReader.ReaderConfig
import com.tjclp.xl.ooxml.metadata.LightMetadata
import com.tjclp.xl.ooxml.style.WorkbookStyles

/**
 * The streaming source's dense projection of a window: the rows the reader never emitted are filled
 * with empty records, and the window's last row ends the pull — a row past it is never a reason to
 * read on.
 */
class StreamingSourceSpec extends CatsEffectSuite:

  private val sheet = SheetName.unsafe("S")

  List("metadata", "shared strings", "styles").foreach { part =>
    test(s"memoized $part failures are cached and reported only to their caller") {
      IO.blocking {
        val reported = new ConcurrentLinkedQueue[Throwable]()
        def report(error: Throwable): Unit =
          reported.add(error)
          ()
        // Force the memo's worker to finish before the consumer can attach its join. A normal
        // thread pool hits this ordering intermittently, leaking a handled failure to stderr.
        val immediate = new ExecutionContext:
          def execute(runnable: Runnable): Unit = runnable.run()
          def reportFailure(error: Throwable): Unit = report(error)
        val runtime = IORuntime
          .builder()
          .setCompute(immediate, () => ())
          .setFailureReporter(report)
          .build()
        val loads = new AtomicInteger()
        val error = new IllegalStateException(s"$part read failed")
        def failed[A]: IO[A] = IO(loads.incrementAndGet()) *> IO.raiseError[A](error)
        val excel = new ExcelIO[IO](_ => IO.unit):
          override def readMetadata(path: Path): IO[LightMetadata] = failed
          override def loadSharedStrings(
            path: Path,
            config: ReaderConfig
          ): IO[Option[SharedStrings]] =
            if part == "shared strings" then failed else IO.pure(None)
          override def loadStyles(path: Path): IO[WorkbookStyles] = failed
        def read(source: StreamingSource): IO[Unit] =
          if part == "metadata" then source.sheets.void
          else source.occupied(sheet).compile.drain
        val check = for
          source <- StreamingSource(Paths.get("unused.xlsx"), excel, ReaderConfig())
          initially <- IO(loads.get())
          first <- read(source).attempt
          second <- read(source).attempt
        yield
          assertEquals(initially, 0, "the parts remain lazy")
          assert(first.isLeft)
          assertEquals(first, second, "subsequent reads receive the cached failure")
          assertEquals(loads.get(), 1, "a failed part is not reloaded")
          assertEquals(reported.iterator().asScala.toVector, Vector.empty[Throwable])
        try check.unsafeRunSync()(using runtime)
        finally runtime.shutdown()
      }
    }
  }

  test("dense: gaps are filled, a row past the window ends the pull, the tail is never read") {
    val window = CellRange(ref"A1", ref"B3")
    val streamed: Stream[IO, RowData] =
      Stream.emits(
        Vector(
          RowData(1, Map(0 -> CellValue.Number(1))),
          RowData(3, Map(1 -> CellValue.Text("x"))),
          RowData(7, Map(0 -> CellValue.Number(7)))
        )
      ) ++ Stream.raiseError[IO](new IllegalStateException("pulled past the window"))
    StreamingSource.dense(streamed, sheet, window, WorkbookStyles.default).compile.toVector.map {
      rows =>
        assertEquals(
          rows.map(_.map(_.ref.toA1)),
          Vector(Vector("A1", "B1"), Vector("A2", "B2"), Vector("A3", "B3"))
        )
        assertEquals(
          rows.flatten.map(_.value),
          Vector(
            CellValue.Number(1),
            CellValue.Empty,
            CellValue.Empty,
            CellValue.Empty,
            CellValue.Empty,
            CellValue.Text("x")
          )
        )
    }
  }

  test("dense: a window the reader leaves entirely empty is every row of it, empty") {
    val window = CellRange(ref"C2", ref"D3")
    StreamingSource
      .dense(Stream.empty, sheet, window, WorkbookStyles.default)
      .compile
      .toVector
      .map { rows =>
        assertEquals(rows.map(_.map(_.ref.toA1)), Vector(Vector("C2", "D2"), Vector("C3", "D3")))
        assertEquals(rows.flatten.map(_.value), Vector.fill(4)(CellValue.Empty: CellValue))
      }
  }
