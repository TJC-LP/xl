package com.tjclp.xl.cli

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{Outcome, OutputMode, Render}
import com.tjclp.xl.cli.read.ReadTestKit
import com.tjclp.xl.macros.ref

/**
 * Tests for the filter command (GH-134, phase 1 — no SQL), as a [[com.tjclp.xl.cli.read.Reads]]
 * query (W2.4): the same code answers `filter` over a loaded workbook and under `--stream`.
 *
 * Covers: --where row filtering, --header name resolution, --columns projection, --limit, and
 * markdown/csv/json output formats.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class FilterCommandSpec extends CatsEffectSuite:

  // A small product table: headers in row 1, data in rows 2-5
  private val sheet = Sheet("Data")
    .put(ref"A1", CellValue.Text("Item"))
    .put(ref"B1", CellValue.Text("Price"))
    .put(ref"C1", CellValue.Text("Active"))
    .put(ref"A2", CellValue.Text("Widget"))
    .put(ref"B2", CellValue.Number(150))
    .put(ref"C2", CellValue.Bool(true))
    .put(ref"A3", CellValue.Text("Gadget"))
    .put(ref"B3", CellValue.Number(50))
    .put(ref"C3", CellValue.Bool(false))
    .put(ref"A4", CellValue.Text("Widget Pro"))
    .put(ref"B4", CellValue.Number(250))
    .put(ref"C4", CellValue.Bool(true))
    .put(ref"A5", CellValue.Text("Doohickey"))
    .put(ref"B5", CellValue.Number(99))
    .put(ref"C5", CellValue.Bool(false))

  private val wb = Workbook(Vector(sheet))

  private def outcome(
    where: String,
    columns: Option[String] = None,
    limit: Int = 50,
    format: FilterFormat = FilterFormat.Markdown,
    header: Boolean = false,
    book: Workbook = wb
  ): IO[Outcome] =
    ReadTestKit.inMemory(
      book,
      book.sheets.headOption.map(_.name.value),
      ReadTestKit.filter(where, columns, limit, format, header)
    )

  private def run(
    where: String,
    columns: Option[String] = None,
    limit: Int = 50,
    format: FilterFormat = FilterFormat.Markdown,
    header: Boolean = false
  ): IO[String] =
    outcome(where, columns, limit, format, header).map(ReadTestKit.text)

  private def runOn(book: Workbook, where: String, format: FilterFormat): IO[String] =
    outcome(where, format = format, book = book).map(ReadTestKit.text)

  test("filter: --where keeps only matching rows (markdown)") {
    run("B > 100").map { out =>
      assert(out.contains("Widget"), out)
      assert(out.contains("Widget Pro"), out)
      assert(!out.contains("Gadget"), out)
      assert(!out.contains("Doohickey"), out)
      assert(out.contains("2"), "original row numbers shown")
    }
  }

  test("filter: type mismatches (header text row) simply do not match") {
    // Row 1 has Text in column B ("Price") — must not match B > 100 and must not error
    run("B > 100").map { out =>
      assert(!out.contains("Item"), s"Header row should not match: $out")
    }
  }

  test("filter: --header resolves header names to columns and skips the header row") {
    run("Price > 100 AND Active = TRUE", header = true).map { out =>
      assert(out.contains("Widget"), out)
      assert(out.contains("Widget Pro"), out)
      assert(!out.contains("Gadget"), out)
    }
  }

  test("filter: --header names are case-insensitive") {
    run("price >= 99", header = true).map { out =>
      assert(out.contains("Doohickey"), out)
    }
  }

  test("filter: unknown column errors with available headers") {
    outcome("Cost > 100", header = true).map { result =>
      val message = result.error.fold("")(_.message)
      assert(!result.ok, "expected a failure")
      assert(message.contains("Cost"), message)
      assert(message.contains("Price"), s"Should list available headers: $message")
      assertEquals(result.error.map(_.code), Some("INVALID_REFERENCE"))
    }
  }

  test("filter: column letters beyond the used range error without --header") {
    outcome("ZZ > 100").map { result =>
      assert(result.error.exists(_.message.contains("ZZ")), result.error.toString)
    }
  }

  test("filter: invalid predicate is a usage error naming the predicate") {
    outcome("B >").map { result =>
      assert(
        result.error.exists(_.message.toLowerCase.contains("predicate")),
        result.error.toString
      )
      assertEquals(result.error.map(_.code), Some("USAGE"))
    }
  }

  test("filter: --columns projects a subset (single and range)") {
    run("B > 100", columns = Some("A,C")).map { out =>
      assert(out.contains("Widget"), out)
      assert(!out.contains("150"), s"Price column should be projected out: $out")
      assert(out.contains("TRUE"), out)
    }
  }

  test("filter: --columns supports letter ranges") {
    run("B > 100", columns = Some("B:C")).map { out =>
      assert(out.contains("150"), out)
      assert(!out.contains("Widget"), s"Column A should be projected out: $out")
    }
  }

  test("filter: --limit truncates and notes the total") {
    run("B > 0", limit = 2).map { out =>
      assert(out.contains("Widget"), out)
      assert(out.contains("Gadget"), out)
      assert(!out.contains("Doohickey"), out)
      assert(out.contains("4"), s"Should mention total match count: $out")
    }
  }

  test("filter: csv output is RFC 4180 with a label header line") {
    run("B > 100", format = FilterFormat.Csv).map { out =>
      val lines = out.split("\n").toVector
      assertEquals(lines.headOption, Some("row,A,B,C"))
      assert(lines.exists(_.startsWith("2,Widget,150,TRUE")), out)
    }
  }

  test("filter: csv escapes commas and quotes") {
    val s = Sheet("S").put(ref"A1", CellValue.Text("a,b \"c\"")).put(ref"B1", CellValue.Number(1))
    runOn(Workbook(Vector(s)), "B = 1", FilterFormat.Csv).map { out =>
      assert(out.contains("\"a,b \"\"c\"\"\""), out)
    }
  }

  test("filter: json output has row numbers and typed cells") {
    run("B > 100", format = FilterFormat.Json).map { out =>
      val json = ujson.read(out)
      // GH-639: the document carries the match total and the clip beside the rows
      assertEquals(json("matched").num.toInt, 2)
      assertEquals(json("shown").num.toInt, 2)
      assertEquals(json("truncated").bool, false)
      assertEquals(json("limit").num.toInt, 50)
      val rows = json("rows").arr.toVector
      assertEquals(rows.length, 2)
      assertEquals(rows.head("row").num.toInt, 2)
      assertEquals(rows.head("cells")("A").str, "Widget")
      assertEquals(rows.head("cells")("B").num, 150.0)
      assertEquals(rows.head("cells")("C").bool, true)
    }
  }

  test("filter: json keeps every digit of a number (never rounds through a Double)") {
    val s = Sheet("S")
      .put(ref"A1", CellValue.Number(BigDecimal("12345678901234567")))
      .put(ref"B1", CellValue.Number(1))
    runOn(Workbook(Vector(s)), "B = 1", FilterFormat.Json).map { out =>
      assert(out.contains("12345678901234567"), out)
    }
  }

  test("filter: json keys use header names with --header") {
    run("Price > 100", format = FilterFormat.Json, header = true).map { out =>
      val json = ujson.read(out)
      assertEquals(json("rows").arr.head("cells")("Price").num, 150.0)
      assertEquals(json("rows").arr.head("cells")("Item").str, "Widget")
    }
  }

  test("filter: no matches yields empty results per format") {
    for
      md <- run("B > 9999")
      csv <- run("B > 9999", format = FilterFormat.Csv)
      json <- run("B > 9999", format = FilterFormat.Json)
    yield
      assert(md.toLowerCase.contains("no rows"), md)
      assertEquals(csv, "row,A,B,C")
      val doc = ujson.read(json)
      assertEquals(doc("rows").arr.toVector, Vector.empty[ujson.Value])
      assertEquals(doc("matched").num.toInt, 0)
      assertEquals(doc("truncated").bool, false)
  }

  test("filter: empty sheet yields no matches rather than an error") {
    runOn(Workbook(Vector(Sheet("Empty"))), "A > 1", FilterFormat.Markdown)
      .map(out => assert(out.toLowerCase.contains("no rows"), out))
  }

  test("filter: IS EMPTY matches sparse rows") {
    val s = Sheet("S")
      .put(ref"A1", CellValue.Text("x"))
      .put(ref"B1", CellValue.Number(1))
      .put(ref"A2", CellValue.Text("y"))
      .put(ref"A3", CellValue.Text("z"))
      .put(ref"B3", CellValue.Number(3))
    runOn(Workbook(Vector(s)), "B IS EMPTY", FilterFormat.Markdown).map { out =>
      assert(out.contains("y"), out)
      assert(!out.contains("x"), out)
      assert(!out.contains("z"), out)
    }
  }

  test("filter --stream: the streaming source answers the same rows (W2.4)") {
    ReadTestKit.withTempWorkbook(wb) { path =>
      val queries = Vector(
        ReadTestKit.filter("B > 100"),
        ReadTestKit
          .filter("Price > 100 AND Active = TRUE", header = true, format = FilterFormat.Csv),
        ReadTestKit.filter("B > 0", limit = 2, format = FilterFormat.Json),
        ReadTestKit.filter("B > 0", columns = Some("Z"), format = FilterFormat.Json),
        ReadTestKit.filter("B > 0", columns = Some("Z,A"), format = FilterFormat.Csv),
        ReadTestKit.filter("B > 0", limit = 0)
      )
      queries.foldLeft(IO.unit) { (acc, query) =>
        acc *> (for
          memory <- ReadTestKit.inMemory(wb, Some("Data"), query).map(ReadTestKit.text)
          streamed <- ReadTestKit.streaming(path, Some("Data"), query).map(ReadTestKit.text)
        yield assertEquals(streamed, memory, s"filter parity broke for $query"))
      }
    }
  }

  // A used range that starts at B: a --columns token left of it selects no cell either
  private val offsetSheet = Sheet("Off")
    .put(ref"B2", CellValue.Text("k"))
    .put(ref"C2", CellValue.Number(1))
    .put(ref"B3", CellValue.Text("m"))
    .put(ref"C3", CellValue.Number(2))

  test("filter: a --columns token outside the used range is blank; the row number stays") {
    for
      json <- run("B > 0", columns = Some("Z"), format = FilterFormat.Json)
      md <- run("B > 0", columns = Some("Z,A"))
      csv <- run("B > 0", columns = Some("Z"), format = FilterFormat.Csv)
      left <- outcome(
        "C > 0",
        columns = Some("A"),
        format = FilterFormat.Json,
        book = Workbook(Vector(offsetSheet))
      ).map(ReadTestKit.text)
    yield
      val rows = ujson.read(json)("rows").arr.toVector
      assertEquals(rows.map(_("row").num.toInt), Vector(2, 3, 4, 5))
      rows.foreach(r => assertEquals(r("cells")("Z"), ujson.Null))
      assert(md.linesIterator.exists(_.matches("[|]\\s*2\\s*[|]\\s*[|]\\s*Widget\\s*[|]")), md)
      assertEquals(csv.linesIterator.toVector, Vector("row,Z", "2,", "3,", "4,", "5,"))
      val leftRows = ujson.read(left)("rows").arr.toVector
      assertEquals(leftRows.map(_("row").num.toInt), Vector(2, 3))
      leftRows.foreach(r => assertEquals(r("cells")("A"), ujson.Null))
  }

  test("filter --json: the envelope carries the rows when every selected column is outside") {
    val query = ReadTestKit.filter("B > 0", columns = Some("Z"), format = FilterFormat.Json)
    def rowsOf(outcome: Outcome): Vector[Int] =
      val data = ujson.read(Render.json(outcome, "test").stdout)("data")
      assert(data("rows").arrOpt.isDefined, s"filter's data carries no rows array: $data")
      data("rows").arr.toVector.map(_("row").num.toInt)
    ReadTestKit.withTempWorkbook(wb) { path =>
      for
        memory <- ReadTestKit.inMemory(wb, Some("Data"), query, OutputMode.Json)
        streamed <- ReadTestKit.streaming(path, Some("Data"), query, OutputMode.Json)
      yield
        assert(memory.ok, memory.error.toString)
        assertEquals(rowsOf(memory), Vector(2, 3, 4, 5))
        assertEquals(rowsOf(streamed), Vector(2, 3, 4, 5))
    }
  }

  test("filter: a repeated --columns token is a usage error") {
    for
      twice <- outcome("B > 0", columns = Some("A,A"))
      overlap <- outcome("B > 0", columns = Some("A:C,B"), format = FilterFormat.Json)
    yield
      assertEquals(twice.error.map(_.code), Some("USAGE"))
      assertEquals(twice.exitCode.code, 2)
      assertEquals(twice.error.map(_.message), Some("Duplicate column(s) in --columns: A"))
      assertEquals(overlap.error.map(_.message), Some("Duplicate column(s) in --columns: B"))
  }

  test("filter: json collapses a repeated header label to one key (first position, last value)") {
    val s = Sheet("S")
      .put(ref"A1", CellValue.Text("Total"))
      .put(ref"B1", CellValue.Text("Other"))
      .put(ref"C1", CellValue.Text("Total"))
      .put(ref"A2", CellValue.Number(1))
      .put(ref"B2", CellValue.Number(2))
      .put(ref"C2", CellValue.Number(3))
    outcome("Other > 0", format = FilterFormat.Json, header = true, book = Workbook(Vector(s)))
      .map(ReadTestKit.text)
      .map { out =>
        assertEquals(out.split("\"Total\"", -1).length - 1, 1, out)
        val cells = ujson.read(out)("rows").arr.head("cells").obj
        assertEquals(cells.keys.toList, List("Total", "Other"))
        assertEquals(cells("Total").num, 3.0)
      }
  }

  test("GH-639: --limit 0 is no limit, as for view and search") {
    for
      md <- run("B > 0", limit = 0)
      csv <- outcome("B > 0", limit = 0, format = FilterFormat.Csv)
      json <- run("B > 0", limit = 0, format = FilterFormat.Json)
    yield
      assert(md.contains("4 row(s) matched."), md)
      assert(md.contains("Doohickey"), md)
      assertEquals(ReadTestKit.text(csv).linesIterator.size, 5, ReadTestKit.text(csv))
      assertEquals(csv.warnings, Vector.empty)
      val doc = ujson.read(json)
      assertEquals(doc("rows").arr.size, 4)
      assertEquals(doc("matched").num.toInt, 4)
      assertEquals(doc("truncated").bool, false)
      assertEquals(doc("limit"), ujson.Null)
  }

  test("GH-639: a clipped result says so in every format — footer, fields, TRUNCATED warning") {
    for
      md <- run("B > 0", limit = 2)
      csv <- outcome("B > 0", limit = 2, format = FilterFormat.Csv)
      json <- outcome("B > 0", limit = 2, format = FilterFormat.Json)
    yield
      assert(md.contains("Matched 4 row(s); showing first 2 (--limit)."), md)
      // CSV stdout stays parseable: the clip is a warning, as `view --format csv` does
      assertEquals(ReadTestKit.text(csv).linesIterator.size, 3, ReadTestKit.text(csv))
      assertEquals(csv.warnings.map(_.code), Vector("TRUNCATED"))
      assertEquals(
        csv.warnings.map(_.message),
        Vector("… showing 2 of 4 matching rows (use --limit to raise; --limit 0 = no limit)")
      )
      // JSON carries the clip in its own fields and raises no warning
      val doc = ujson.read(ReadTestKit.text(json))
      assertEquals(doc("matched").num.toInt, 4)
      assertEquals(doc("shown").num.toInt, 2)
      assertEquals(doc("truncated").bool, true)
      assertEquals(doc("limit").num.toInt, 2)
      assertEquals(doc("rows").arr.map(_("row").num.toInt).toVector, Vector(2, 3))
      assertEquals(json.warnings, Vector.empty)
  }
