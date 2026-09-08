package com.tjclp.xl.cli.read

import java.nio.file.Path

import cats.effect.{IO, Ref}
import fs2.Stream
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{Outcome, OutputMode, Warning}
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.ooxml.SharedStrings
import com.tjclp.xl.ooxml.style.WorkbookStyles

/**
 * GH-637: `search` stops scanning at `--limit`. The scan reads one match past the limit — enough to
 * know more exist — and reports `total` as a lower bound (`totalExact: false`); `--total` and
 * `--limit 0` read everything and report the exact total. The regression: a `--stream search
 * --limit 10` over a million-row sheet scanned every row for the exact total (56 s where 0.19.3
 * took 1.5 s).
 *
 * What the counts here prove, exactly: `pulled` is the number of records `Reads.search` pulls
 * across the [[SheetSource]] boundary — the query stops asking after `limit + 1` matches. It is NOT
 * what the reader parsed: the in-memory source sorts and emits every occupied cell in one chunk,
 * and the SAX reader parses ahead in 1024-row chunks, so a 100-row fixture is parsed whole either
 * way. The [[Probe]] on `ExcelIO` closes that gap for the sheet boundary: a sheet after the one
 * that filled the limit never has its worksheet stream constructed, so neither its parse nor its
 * styles load ever runs. On a million-row sheet those two bounds — stop pulling, never open the
 * next sheet — are what turn the full scan into a bounded one (the PR's measurements). Nothing here
 * depends on wall time.
 */
class SearchScanSpec extends CatsEffectSuite:

  /** Sheet `name` with `n` numeric rows in column A (A1=1 .. An=n): every cell matches `\d`. */
  private def numbered(name: String, n: Int): Sheet =
    (1 to n).foldLeft(Sheet(SheetName.unsafe(name))) { (s, i) =>
      s.put(ARef.from0(0, i - 1), CellValue.Number(BigDecimal(i)))
    }

  private def query(limit: Int, exactTotal: Boolean = false): ReadQuery.Search =
    ReadQuery.Search("\\d", limit, None, exactTotal)

  /**
   * What the query asked for: the records pulled across the source boundary, the sheets `occupied`
   * was called for, and — streaming only — the worksheet streams the reader constructed (one per
   * opened sheet) and the styles loads it ran (one per source: the parts are memoised, GH-640).
   */
  private final case class Scan(
    outcome: Outcome,
    pulled: Int,
    opened: Vector[SheetName],
    streamed: Vector[String],
    stylesLoads: Int
  )

  /** A source that counts what `search` pulls from it and delegates everything to `underlying`. */
  private final class Spy(
    underlying: SheetSource,
    pulled: Ref[IO, Int],
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
    // `evalTap` re-chunks: every record the query pulls passes through here one at a time, so the
    // count is the query's demand, whatever chunking the underlying source used
    def occupied(sheet: SheetName): Stream[IO, CellRecord] =
      Stream.exec(opened.update(_ :+ sheet)) ++
        underlying.occupied(sheet).evalTap(_ => pulled.update(_ + 1))
    def detail(sheet: SheetName, ref: ARef, withStyle: Boolean): IO[CellDetail] =
      underlying.detail(sheet, ref, withStyle)
    def render(
      sheet: SheetName,
      window: CellRange,
      spec: RenderSpec,
      warn: Warning => IO[Unit]
    ): IO[String] = underlying.render(sheet, window, spec, warn)

  /**
   * An `ExcelIO` that records the worksheet streams it constructs and the styles loads it runs: the
   * streaming source calls both once per sheet whose `occupied` stream is pulled, so a sheet that
   * appears in neither was never touched.
   */
  private final class Probe(streamed: Ref[IO, Vector[String]], stylesLoads: Ref[IO, Int])
      extends ExcelIO[IO](_ => IO.unit):
    override def readSheetStream(path: Path, sheetName: String): Stream[IO, RowData] =
      Stream.exec(streamed.update(_ :+ sheetName)) ++ super.readSheetStream(path, sheetName)
    // The streaming source streams with the table it loaded once (GH-640)
    override def readSheetStream(
      path: Path,
      sheetName: String,
      sst: Option[SharedStrings]
    ): Stream[IO, RowData] =
      Stream.exec(streamed.update(_ :+ sheetName)) ++ super.readSheetStream(path, sheetName, sst)
    override def loadStyles(path: Path): IO[WorkbookStyles] =
      stylesLoads.update(_ + 1) *> super.loadStyles(path)

  /**
   * Run `q` over `source` (built from a probed `ExcelIO` by [[streaming]]) and collect the counts.
   */
  private def scan(
    source: (Ref[IO, Vector[String]], Ref[IO, Int]) => IO[SheetSource],
    sheet: Option[String],
    q: ReadQuery.Search,
    mode: OutputMode
  ): IO[Scan] =
    for
      pulled <- Ref.of[IO, Int](0)
      opened <- Ref.of[IO, Vector[SheetName]](Vector.empty)
      streamed <- Ref.of[IO, Vector[String]](Vector.empty)
      stylesLoads <- Ref.of[IO, Int](0)
      built <- source(streamed, stylesLoads)
      outcome <- Reads.outcome(q, Spy(built, pulled, opened), sheet, mode)
      n <- pulled.get
      sheets <- opened.get
      constructed <- streamed.get
      loads <- stylesLoads.get
    yield Scan(outcome, n, sheets, constructed, loads)

  private def inMemory(
    wb: Workbook,
    sheet: Option[String],
    q: ReadQuery.Search,
    mode: OutputMode = OutputMode.Json
  ): IO[Scan] =
    scan((_, _) => IO.pure(SheetSource.inMemory(wb)), sheet, q, mode)

  private def streaming(
    wb: Workbook,
    sheet: Option[String],
    q: ReadQuery.Search,
    mode: OutputMode = OutputMode.Json
  ): IO[Scan] =
    ReadTestKit.withTempWorkbook(wb) { path =>
      scan(
        (streamed, loads) => SheetSource.streaming(path, Probe(streamed, loads)),
        sheet,
        q,
        mode
      )
    }

  private def data(scan: Scan): ujson.Value = ujson.read(ReadTestKit.text(scan.outcome))

  private def refs(payload: ujson.Value): Vector[String] =
    payload("matches").arr.toVector.map(_("ref").str)

  // ---------------------------------------------------------------------------------------------
  // The default scan stops at --limit
  // ---------------------------------------------------------------------------------------------

  test("streaming: the query pulls limit + 1 records and reports the total as a lower bound") {
    streaming(Workbook(Vector(numbered("Data", 100))), Some("Data"), query(limit = 10)).map { s =>
      val payload = data(s)
      assertEquals(payload("count"), ujson.Num(10))
      assertEquals(payload("total"), ujson.Num(11))
      assertEquals(payload("totalExact"), ujson.Bool(false))
      assertEquals(refs(payload), (1 to 10).map(i => s"A$i").toVector)
      // The 11th match proves the clip and is the last record the query asks for (the reader may
      // have parsed further ahead — that is its chunking, not the query's demand)
      assertEquals(s.pulled, 11, "the query pulled past the match that proved the clip")
      assertEquals(s.outcome.warnings, Vector.empty)
    }
  }

  test("streaming --total: every record is pulled and the total is exact") {
    streaming(Workbook(Vector(numbered("Data", 100))), Some("Data"), query(10, exactTotal = true))
      .map { s =>
        val payload = data(s)
        assertEquals(payload("count"), ujson.Num(10))
        assertEquals(payload("total"), ujson.Num(100))
        assertEquals(payload("totalExact"), ujson.Bool(true))
        assertEquals(refs(payload).size, 10)
        assertEquals(s.pulled, 100)
      }
  }

  test("streaming --limit 0: no limit, every match listed, the total exact") {
    streaming(Workbook(Vector(numbered("Data", 100))), Some("Data"), query(limit = 0)).map { s =>
      val payload = data(s)
      assertEquals(payload("count"), ujson.Num(100))
      assertEquals(payload("total"), ujson.Num(100))
      assertEquals(payload("totalExact"), ujson.Bool(true))
      assertEquals(refs(payload).size, 100)
      assertEquals(s.pulled, 100)
    }
  }

  test("streaming: a hit list that fits within --limit is exact without --total") {
    streaming(Workbook(Vector(numbered("Data", 10))), Some("Data"), query(limit = 10)).map { s =>
      val payload = data(s)
      assertEquals(payload("count"), ujson.Num(10))
      assertEquals(payload("total"), ujson.Num(10))
      assertEquals(payload("totalExact"), ujson.Bool(true))
      assertEquals(s.pulled, 10)
    }
  }

  test(
    "streaming: the limit reached on the first sheet never constructs the second sheet's stream"
  ) {
    val wb = Workbook(Vector(numbered("Data", 100), numbered("More", 100)))
    streaming(wb, None, query(limit = 10)).map { s =>
      val payload = data(s)
      assertEquals(payload("sheets").arr.toVector.map(_.str), Vector("Data", "More"))
      assertEquals(payload("count"), ujson.Num(10))
      assertEquals(payload("total"), ujson.Num(11))
      assertEquals(payload("totalExact"), ujson.Bool(false))
      assertEquals(s.opened, Vector(SheetName.unsafe("Data")))
      // The reader itself: one worksheet stream, one styles load — `More` was never touched
      assertEquals(s.streamed, Vector("Data"))
      assertEquals(s.stylesLoads, 1)
      assertEquals(s.pulled, 11)
    }
  }

  test("streaming --total over two sheets: both streams constructed, the total summed") {
    val wb = Workbook(Vector(numbered("Data", 100), numbered("More", 50)))
    streaming(wb, None, query(10, exactTotal = true)).map { s =>
      val payload = data(s)
      assertEquals(payload("total"), ujson.Num(150))
      assertEquals(payload("totalExact"), ujson.Bool(true))
      assertEquals(s.opened, Vector(SheetName.unsafe("Data"), SheetName.unsafe("More")))
      assertEquals(s.streamed, Vector("Data", "More"))
      // one source, one styles load: the parts are memoised across the run's reads (GH-640)
      assertEquals(s.stylesLoads, 1)
      assertEquals(s.pulled, 150)
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
      assertEquals(s.pulled, 11)
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
      assertEquals(s.pulled, 100)
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

  test("in-memory: the same limit + 1 demand and the same payload as streaming") {
    val wb = Workbook(Vector(numbered("Data", 100), numbered("More", 100)))
    for
      memory <- inMemory(wb, None, query(limit = 10))
      streamed <- streaming(wb, None, query(limit = 10))
    yield
      val memoryData = data(memory)
      val streamData = data(streamed)
      // `hidden` belongs to a capability the streaming reader lacks (null there, false in memory)
      Vector(memoryData, streamData).foreach(_("matches").arr.foreach(_.obj.remove("hidden")))
      assertEquals(streamData, memoryData)
      assertEquals(memoryData("total"), ujson.Num(11))
      assertEquals(memoryData("totalExact"), ujson.Bool(false))
      // The in-memory source sorts every occupied cell of `Data` before the first is pulled (its
      // scan cost is the sort); what the query asks of it is the same 11 records, and `More` is
      // never asked for
      assertEquals(memory.pulled, 11)
      assertEquals(memory.opened, Vector(SheetName.unsafe("Data")))
  }

  test("in-memory --total: exact, every sheet opened") {
    val wb = Workbook(Vector(numbered("Data", 100), numbered("More", 100)))
    inMemory(wb, None, query(10, exactTotal = true)).map { s =>
      val payload = data(s)
      assertEquals(payload("total"), ujson.Num(200))
      assertEquals(payload("totalExact"), ujson.Bool(true))
      assertEquals(s.pulled, 200)
    }
  }
