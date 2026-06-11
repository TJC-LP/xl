package com.tjclp.xl.cli

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.FilterCommands
import com.tjclp.xl.macros.ref

/**
 * Tests for the filter command (GH-134, phase 1 — no SQL).
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

  private def run(
    where: String,
    columns: Option[String] = None,
    limit: Int = 50,
    format: FilterFormat = FilterFormat.Markdown,
    header: Boolean = false
  ): IO[String] =
    FilterCommands.filter(wb, Some(sheet), where, columns, limit, format, header)

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
    run("Cost > 100", header = true).attempt.map {
      case Left(err) =>
        assert(err.getMessage.contains("Cost"), err.getMessage)
        assert(
          err.getMessage.contains("Price"),
          s"Should list available headers: ${err.getMessage}"
        )
      case Right(out) => fail(s"Expected error, got: $out")
    }
  }

  test("filter: column letters beyond the used range error without --header") {
    run("ZZ > 100").attempt.map {
      case Left(err) => assert(err.getMessage.contains("ZZ"), err.getMessage)
      case Right(out) => fail(s"Expected error, got: $out")
    }
  }

  test("filter: invalid predicate reports a parse error") {
    run("B >").attempt.map {
      case Left(err) =>
        assert(err.getMessage.toLowerCase.contains("predicate"), err.getMessage)
      case Right(out) => fail(s"Expected error, got: $out")
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
    FilterCommands
      .filter(Workbook(Vector(s)), Some(s), "B = 1", None, 50, FilterFormat.Csv, false)
      .map { out =>
        assert(out.contains("\"a,b \"\"c\"\"\""), out)
      }
  }

  test("filter: json output has row numbers and typed cells") {
    run("B > 100", format = FilterFormat.Json).map { out =>
      val json = ujson.read(out)
      val rows = json.arr.toVector
      assertEquals(rows.length, 2)
      assertEquals(rows.head("row").num.toInt, 2)
      assertEquals(rows.head("cells")("A").str, "Widget")
      assertEquals(rows.head("cells")("B").num, 150.0)
      assertEquals(rows.head("cells")("C").bool, true)
    }
  }

  test("filter: json keys use header names with --header") {
    run("Price > 100", format = FilterFormat.Json, header = true).map { out =>
      val json = ujson.read(out)
      assertEquals(json.arr.head("cells")("Price").num, 150.0)
      assertEquals(json.arr.head("cells")("Item").str, "Widget")
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
      assertEquals(ujson.read(json).arr.toVector, Vector.empty[ujson.Value])
  }

  test("filter: empty sheet yields no matches rather than an error") {
    val empty = Sheet("Empty")
    FilterCommands
      .filter(Workbook(Vector(empty)), Some(empty), "A > 1", None, 50, FilterFormat.Markdown, false)
      .map(out => assert(out.toLowerCase.contains("no rows"), out))
  }

  test("filter: IS EMPTY matches sparse rows") {
    val s = Sheet("S")
      .put(ref"A1", CellValue.Text("x"))
      .put(ref"B1", CellValue.Number(1))
      .put(ref"A2", CellValue.Text("y"))
      .put(ref"A3", CellValue.Text("z"))
      .put(ref"B3", CellValue.Number(3))
    FilterCommands
      .filter(Workbook(Vector(s)), Some(s), "B IS EMPTY", None, 50, FilterFormat.Markdown, false)
      .map { out =>
        assert(out.contains("y"), out)
        assert(!out.contains("x"), out)
        assert(!out.contains("z"), out)
      }
  }
