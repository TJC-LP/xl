package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.{ARef, Column}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.extensions.style
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.sheets.ColumnProperties
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * GH-869: `put`, `putf` and `batch` widen a never-sized column when a formatted number written to
 * it would not fit — what Excel does when such a value is typed — so the saved book does not open
 * showing `#####`. An explicit width, text, General numbers and values that fit are left alone.
 */
class PutWidensColumnSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  private val config: WriterConfig = WriterConfig.default

  private def withOut[A](body: Path => A): A =
    val path = Files.createTempFile("put-widens", ".xlsx")
    try body(path)
    finally Files.deleteIfExists(path)

  private def col(letter: String): Column =
    Column.fromLetter(letter).fold(e => fail(e), identity)

  private def widthOf(path: Path, letter: String): Option[Double] =
    val wb = ExcelIO.instance[IO].read(path).unsafeRunSync()
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    sheet.getColumnProperties(col(letter)).width

  private def batch(json: String, out: Path): String =
    val ops = Files.createTempFile("put-widens-ops", ".json")
    try
      Files.writeString(ops, json)
      val wb = Workbook(Sheet("Data"))
      WriteCommands.batch(wb, wb.sheets.headOption, ops.toString, out, config).unsafeRunSync()
    finally Files.deleteIfExists(ops)

  private def put(sheet: Sheet, ref: String, value: String, out: Path): String =
    val wb = Workbook(sheet)
    WriteCommands.put(wb, wb.sheets.headOption, ref, List(value), out, config).unsafeRunSync()

  // ===== put =====

  test("put: a detected currency too wide for the default column widens it (GH-869)") {
    withOut { out =>
      val message = put(Sheet("Data"), "B2", "$1,234,567.89", out)
      val width = widthOf(out, "B").getOrElse(fail("column B was not widened"))
      assert(width > 8.43, s"width $width should exceed the default 8.43")
      // the width is what `autofit` would give the column
      val read = ExcelIO.instance[IO].read(out).unsafeRunSync().sheets.head
      assertEqualsDouble(width, read.autoFitWidth(col("B")), 0.001)
      assert(message.contains("Widened column(s) B to"), message)
    }
  }

  test("put: a currency that fits leaves the column unsized (GH-869)") {
    withOut { out =>
      val message = put(Sheet("Data"), "B2", "$5.00", out)
      assertEquals(widthOf(out, "B"), None)
      assert(!message.contains("Widened"), message)
    }
  }

  test("put: a short date that Excel shows in full leaves the column unsized (GH-869)") {
    // 3/15/26 is narrower than the 64 px an unsized column has; the fit is compared in the same
    // units autofit returns (padding included), not against the UI's 8.43
    withOut { out =>
      put(Sheet("Data"), "C2", "2026-03-15", out)
      assertEquals(widthOf(out, "C"), None)
    }
  }

  test("put: long text and long General numbers do not widen (GH-869)") {
    withOut { out =>
      put(Sheet("Data"), "B2", "A heading much wider than the column", out)
      assertEquals(widthOf(out, "B"), None)
    }
    withOut { out =>
      put(Sheet("Data"), "C2", "123456789012", out)
      assertEquals(widthOf(out, "C"), None)
    }
  }

  test("put: an explicitly sized column keeps its width (GH-869)") {
    withOut { out =>
      val sheet = Sheet("Data").setColumnProperties(col("B"), ColumnProperties(width = Some(6.0)))
      put(sheet, "B2", "$1,234,567.89", out)
      assertEquals(widthOf(out, "B"), Some(6.0))
    }
  }

  test("put: a hidden column stays as it was (GH-869)") {
    withOut { out =>
      val sheet = Sheet("Data").setColumnProperties(col("B"), ColumnProperties(hidden = true))
      put(sheet, "B2", "$1,234,567.89", out)
      assertEquals(widthOf(out, "B"), None)
    }
  }

  test("put: a range of values widens each column to its widest entry (GH-869)") {
    withOut { out =>
      val wb = Workbook(Sheet("Data"))
      WriteCommands
        .put(
          wb,
          wb.sheets.headOption,
          "B2:C3",
          List("$1.00", "$9.00", "$98,765,432.10", "$2.00"),
          out,
          config
        )
        .unsafeRunSync()
      assert(widthOf(out, "B").exists(_ > 8.43), "B holds $98,765,432.10 and should widen")
      assertEquals(widthOf(out, "C"), None)
    }
  }

  // ===== putf =====

  test("putf: a formula whose formatted result overflows widens its column (GH-869)") {
    withOut { out =>
      val sheet = Sheet("Data")
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(1234567)))
        .style(ARef.from0(1, 0), CellStyle.default.withNumFmt(NumFmt.Currency))
      val wb = Workbook(sheet)
      WriteCommands
        .putFormula(wb, wb.sheets.headOption, "B1", List("=A1*1000"), out, config)
        .unsafeRunSync()
      assert(widthOf(out, "B").exists(_ > 8.43), s"B is ${widthOf(out, "B")}")
    }
  }

  test("putf --no-recalc: the formula's own cached value still widens (GH-869)") {
    withOut { out =>
      val sheet = Sheet("Data")
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(1234567)))
        .style(ARef.from0(1, 0), CellStyle.default.withNumFmt(NumFmt.Currency))
      val wb = Workbook(sheet)
      WriteCommands
        .putFormula(
          wb,
          wb.sheets.headOption,
          "B1",
          List("=A1*1000"),
          out,
          config,
          policy = WritePolicy.default.copy(noRecalc = true)
        )
        .unsafeRunSync()
      // a single putf caches its own value under --no-recalc, so B displays it and is widened
      assert(widthOf(out, "B").exists(_ > 8.43), s"B is ${widthOf(out, "B")}")
    }
  }

  // ===== batch =====

  test("batch: put and putf with a currency format widen their columns (GH-869)") {
    withOut { out =>
      val ops =
        """[{"op":"put","ref":"A1","value":"$4,500,000.00"},
          | {"op":"put","ref":"B1","value":2500000,"format":"currency"},
          | {"op":"putf","ref":"C1","formula":"=A1+B1","format":"currency"},
          | {"op":"put","ref":"D1","value":"Short label"}]""".stripMargin
      val message = batch(ops, out)
      for letter <- List("A", "B", "C") do
        assert(widthOf(out, letter).exists(_ > 8.43), s"$letter is ${widthOf(out, letter)}")
      assertEquals(widthOf(out, "D"), None)
      assert(message.contains("Widened column(s) A to"), message)
    }
  }

  test("batch: a width set in the same batch is kept (GH-869)") {
    withOut { out =>
      val ops =
        """[{"op":"colwidth","col":"A","width":7},
          | {"op":"put","ref":"A1","value":"$4,500,000.00"}]""".stripMargin
      batch(ops, out)
      assertEquals(widthOf(out, "A"), Some(7.0))
    }
  }
