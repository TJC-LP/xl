package com.tjclp.xl.cli.read

import cats.effect.{IO, Ref}
import fs2.Stream
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{Outcome, OutputMode, Warning}

/**
 * GH-637: `search` stops scanning at `--limit`. The scan reads one match past the limit — enough to
 * know more exist — and reports `total` as a lower bound (`totalExact: false`); `--total` and
 * `--limit 0` read everything and report the exact total. The regression: a `--stream search
 * --limit 10` over a million-row sheet scanned every row for the exact total (56 s where 0.19.3
 * took 1.5 s). The assertions here count the records the query pulls from its source, so they do
 * not depend on wall time.
 */
class SearchScanSpec extends CatsEffectSuite:

  /** Sheet `name` with `n` numeric rows in column A (A1=1 .. An=n): every cell matches `\d`. */
  private def numbered(name: String, n: Int): Sheet =
    (1 to n).foldLeft(Sheet(SheetName.unsafe(name))) { (s, i) =>
      s.put(ARef.from0(0, i - 1), CellValue.Number(BigDecimal(i)))
    }

  private def query(limit: Int, exactTotal: Boolean = false): ReadQuery.Search =
    ReadQuery.Search("\\d", limit, None, exactTotal)

  /** What the source saw: the records pulled through `occupied` and the sheets it was opened on. */
  private final case class Scan(outcome: Outcome, visited: Int, opened: Vector[SheetName])

  /** A source that counts what `search` pulls from it and delegates everything to `underlying`. */
  private final class Spy(
    underlying: SheetSource,
    visited: Ref[IO, Int],
    opened: Ref[IO, Vector[SheetName]]
  ) extends SheetSource:
    val capabilities: Set[Capability] = underlying.capabilities
    def sheets: IO[Vector[SheetName]] = underlying.sheets
    def usedRange(sheet: SheetName): IO[Option[CellRange]] = underlying.usedRange(sheet)
    def rows(sheet: SheetName, window: CellRange): Stream[IO, Vector[CellRecord]] =
      underlying.rows(sheet, window)
    def hiddenLines(sheet: SheetName, window: CellRange): IO[(Set[Int], Set[Int])] =
      underlying.hiddenLines(sheet, window)
    def evaluated(
      sheet: SheetName,
      window: CellRange,
      strict: Boolean,
      warn: Warning => IO[Unit]
    ): IO[RecordGrid] = underlying.evaluated(sheet, window, strict, warn)
    def occupied(sheet: SheetName): Stream[IO, CellRecord] =
      Stream.exec(opened.update(_ :+ sheet)) ++
        underlying.occupied(sheet).evalTap(_ => visited.update(_ + 1))
    def detail(sheet: SheetName, ref: ARef, withStyle: Boolean): IO[CellDetail] =
      underlying.detail(sheet, ref, withStyle)
    def render(
      sheet: SheetName,
      window: CellRange,
      spec: RenderSpec,
      warn: Warning => IO[Unit]
    ): IO[String] = underlying.render(sheet, window, spec, warn)

  private def scan(
    source: SheetSource,
    sheet: Option[String],
    q: ReadQuery.Search,
    mode: OutputMode = OutputMode.Json
  ): IO[Scan] =
    for
      visited <- Ref.of[IO, Int](0)
      opened <- Ref.of[IO, Vector[SheetName]](Vector.empty)
      outcome <- Reads.outcome(q, Spy(source, visited, opened), sheet, mode)
      n <- visited.get
      sheets <- opened.get
    yield Scan(outcome, n, sheets)

  private def streaming(
    wb: Workbook,
    sheet: Option[String],
    q: ReadQuery.Search,
    mode: OutputMode = OutputMode.Json
  ): IO[Scan] =
    ReadTestKit.withTempWorkbook(wb) { path =>
      scan(SheetSource.streaming(path, ReadTestKit.excel), sheet, q, mode)
    }

  private def data(scan: Scan): ujson.Value = ujson.read(ReadTestKit.text(scan.outcome))

  private def refs(payload: ujson.Value): Vector[String] =
    payload("matches").arr.toVector.map(_("ref").str)

  // ---------------------------------------------------------------------------------------------
  // The default scan stops at --limit
  // ---------------------------------------------------------------------------------------------

  test("streaming: the scan stops one match past --limit and reports the total as a lower bound") {
    streaming(Workbook(Vector(numbered("Data", 100))), Some("Data"), query(limit = 10)).map { s =>
      val payload = data(s)
      assertEquals(payload("count"), ujson.Num(10))
      assertEquals(payload("total"), ujson.Num(11))
      assertEquals(payload("totalExact"), ujson.Bool(false))
      assertEquals(refs(payload), (1 to 10).map(i => s"A$i").toVector)
      // The key assertion: the 11th match is the last record read; A12..A100 are never pulled
      assertEquals(s.visited, 11, "the source was read past the match that proved the clip")
      assertEquals(s.outcome.warnings, Vector.empty)
    }
  }

  test("streaming --total: every record is visited and the total is exact") {
    streaming(Workbook(Vector(numbered("Data", 100))), Some("Data"), query(10, exactTotal = true))
      .map { s =>
        val payload = data(s)
        assertEquals(payload("count"), ujson.Num(10))
        assertEquals(payload("total"), ujson.Num(100))
        assertEquals(payload("totalExact"), ujson.Bool(true))
        assertEquals(refs(payload).size, 10)
        assertEquals(s.visited, 100)
      }
  }

  test("streaming --limit 0: no limit, every match listed, the total exact") {
    streaming(Workbook(Vector(numbered("Data", 100))), Some("Data"), query(limit = 0)).map { s =>
      val payload = data(s)
      assertEquals(payload("count"), ujson.Num(100))
      assertEquals(payload("total"), ujson.Num(100))
      assertEquals(payload("totalExact"), ujson.Bool(true))
      assertEquals(refs(payload).size, 100)
      assertEquals(s.visited, 100)
    }
  }

  test("streaming: a hit list that fits within --limit is exact without --total") {
    streaming(Workbook(Vector(numbered("Data", 10))), Some("Data"), query(limit = 10)).map { s =>
      val payload = data(s)
      assertEquals(payload("count"), ujson.Num(10))
      assertEquals(payload("total"), ujson.Num(10))
      assertEquals(payload("totalExact"), ujson.Bool(true))
      assertEquals(s.visited, 10)
    }
  }

  test("streaming: the limit reached on the first sheet leaves the second sheet unopened") {
    val wb = Workbook(Vector(numbered("Data", 100), numbered("More", 100)))
    streaming(wb, None, query(limit = 10)).map { s =>
      val payload = data(s)
      assertEquals(payload("sheets").arr.toVector.map(_.str), Vector("Data", "More"))
      assertEquals(payload("count"), ujson.Num(10))
      assertEquals(payload("total"), ujson.Num(11))
      assertEquals(payload("totalExact"), ujson.Bool(false))
      assertEquals(s.opened, Vector(SheetName.unsafe("Data")))
      assertEquals(s.visited, 11)
    }
  }

  test("streaming --total over two sheets: both sheets read, the total summed") {
    val wb = Workbook(Vector(numbered("Data", 100), numbered("More", 50)))
    streaming(wb, None, query(10, exactTotal = true)).map { s =>
      val payload = data(s)
      assertEquals(payload("total"), ujson.Num(150))
      assertEquals(payload("totalExact"), ujson.Bool(true))
      assertEquals(s.opened, Vector(SheetName.unsafe("Data"), SheetName.unsafe("More")))
      assertEquals(s.visited, 150)
    }
  }

  // ---------------------------------------------------------------------------------------------
  // Text mode says what the scan knows
  // ---------------------------------------------------------------------------------------------

  test("text: a stopped scan reports a lower bound and a trailer naming --total") {
    streaming(
      Workbook(Vector(numbered("Data", 100))),
      Some("Data"),
      query(limit = 10),
      OutputMode.Text
    ).map { s =>
      val out = ReadTestKit.text(s.outcome)
      assert(out.startsWith("Found at least 11 matches in Data:"), out)
      assert(out.contains("Data!A10"), out)
      assert(!out.contains("Data!A11"), s"the 11th match is counted, never listed:\n$out")
      val trailer = out.linesIterator.toVector.lastOption.getOrElse("")
      assert(trailer.startsWith("… showing first 10 matches; more exist"), trailer)
      assert(trailer.contains("--total"), trailer)
      assert(trailer.contains("--limit 0 = no limit"), trailer)
      assertEquals(s.visited, 11)
    }
  }

  test("text --total: the exact total and the same trailer as before") {
    streaming(
      Workbook(Vector(numbered("Data", 100))),
      Some("Data"),
      query(10, exactTotal = true),
      OutputMode.Text
    ).map { s =>
      val out = ReadTestKit.text(s.outcome)
      assert(out.startsWith("Found 100 matches in Data:"), out)
      assert(out.contains("… showing 10 of 100 matches"), out)
      assert(!out.contains("at least"), out)
      assertEquals(s.visited, 100)
    }
  }

  test("text: an exact, unclipped result has no trailer and no 'at least'") {
    streaming(
      Workbook(Vector(numbered("Data", 10))),
      Some("Data"),
      query(limit = 10),
      OutputMode.Text
    ).map { s =>
      val out = ReadTestKit.text(s.outcome)
      assert(out.startsWith("Found 10 matches in Data:"), out)
      assert(!out.contains("… showing"), out)
    }
  }

  // ---------------------------------------------------------------------------------------------
  // One contract, two sources
  // ---------------------------------------------------------------------------------------------

  test("in-memory: the same bounded scan and the same payload as streaming") {
    val wb = Workbook(Vector(numbered("Data", 100), numbered("More", 100)))
    for
      memory <- scan(SheetSource.inMemory(wb), None, query(limit = 10))
      streamed <- streaming(wb, None, query(limit = 10))
    yield
      val memoryData = data(memory)
      val streamData = data(streamed)
      // `hidden` belongs to a capability the streaming reader lacks (null there, false in memory)
      Vector(memoryData, streamData).foreach(_("matches").arr.foreach(_.obj.remove("hidden")))
      assertEquals(streamData, memoryData)
      assertEquals(memoryData("total"), ujson.Num(11))
      assertEquals(memoryData("totalExact"), ujson.Bool(false))
      assertEquals(memory.visited, 11)
      assertEquals(memory.opened, Vector(SheetName.unsafe("Data")))
  }

  test("in-memory --total: exact, every sheet opened") {
    val wb = Workbook(Vector(numbered("Data", 100), numbered("More", 100)))
    scan(SheetSource.inMemory(wb), None, query(10, exactTotal = true)).map { s =>
      val payload = data(s)
      assertEquals(payload("total"), ujson.Num(200))
      assertEquals(payload("totalExact"), ujson.Bool(true))
      assertEquals(s.visited, 200)
    }
  }
