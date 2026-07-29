package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{ref, Sheet, Workbook}
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.cli.helpers.BatchParser
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.sheets.AutoFilterState

/**
 * E2E tests for standalone AutoFilter authoring (GH-432): the `xl autofilter` command and its
 * `{"op": "autofilter"}` batch twin, writing through the GH-429 lift-and-overlay tri-state
 * (`Sheet.autoFilter`). Covers authoring, active removal of a preserved source element, and the
 * structural-shift wiring end-to-end from the CLI.
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class AutoFilterCommandSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val config: WriterConfig = WriterConfig.default

  private def withTempOutput[A](f: Path => A): A =
    val out = Files.createTempFile("autofilter-test", ".xlsx")
    try f(out)
    finally Files.deleteIfExists(out)

  private def readBack(path: Path): Workbook =
    ExcelIO.instance[IO].read(path).unsafeRunSync()

  private def sheetXml(path: Path): String =
    val zip = new java.util.zip.ZipFile(path.toFile)
    try
      val entry = zip.getEntry("xl/worksheets/sheet1.xml")
      new String(
        zip.getInputStream(entry).readAllBytes(),
        java.nio.charset.StandardCharsets.UTF_8
      )
    finally zip.close()

  private def range(s: String): CellRange =
    CellRange.parse(s).toOption.get

  // ========== autofilter command ==========

  test("autofilter A1:M29: writes <autoFilter ref> and round-trips through the model") {
    withTempOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(sheet)
      val result = WriteCommands
        .autoFilter(wb, Some(sheet), Some("A1:M29"), clear = false, out, config)
        .unsafeRunSync()

      assert(result.contains("A1:M29"), result)

      val xml = sheetXml(out)
      assert(xml.contains("<autoFilter ref=\"A1:M29\"/>"), xml)

      val readSheet = readBack(out).sheets.head
      assertEquals(readSheet.autoFilter, Some(AutoFilterState.Ranged(range("A1:M29"))))
    }
  }

  test("autofilter: single-cell argument authors a 1x1 filter range") {
    withTempOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(sheet)
      WriteCommands
        .autoFilter(wb, Some(sheet), Some("B2"), clear = false, out, config)
        .unsafeRunSync()

      // A single cell promotes to a 1x1 range; the canonical A1 form is "B2:B2".
      assert(sheetXml(out).contains("<autoFilter ref=\"B2:B2\"/>"), sheetXml(out))
      assertEquals(
        readBack(out).sheets.head.autoFilter,
        Some(AutoFilterState.Ranged(range("B2:B2")))
      )
    }
  }

  test("autofilter --clear removes a source file's preserved autoFilter (active removal)") {
    withTempOutput { source =>
      withTempOutput { out =>
        // Author a filter, then re-read: the raw <autoFilter> element is preserved AND its
        // ref is lifted into the model (GH-429 lift-and-overlay).
        val wb0 = Workbook(Sheet("Data"))
        WriteCommands
          .autoFilter(wb0, Some(wb0.sheets.head), Some("A1:C5"), clear = false, source, config)
          .unsafeRunSync()
        val wb1 = readBack(source)
        assertEquals(
          wb1.sheets.head.autoFilter,
          Some(AutoFilterState.Ranged(range("A1:C5"))),
          "precondition: reader lifts the authored filter"
        )

        val result = WriteCommands
          .autoFilter(wb1, Some(wb1.sheets.head), None, clear = true, out, config)
          .unsafeRunSync()

        assert(result.contains("Removed"), result)
        val xml = sheetXml(out)
        assert(!xml.contains("<autoFilter"), s"autoFilter must be stripped: $xml")
        assertEquals(readBack(out).sheets.head.autoFilter, None)
      }
    }
  }

  test("autofilter: range and --clear are mutually exclusive") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Data"))
      val result = WriteCommands
        .autoFilter(wb, Some(wb.sheets.head), Some("A1:B2"), clear = true, out, config)
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
      assert(
        result.swap.toOption.get.getMessage.contains("mutually exclusive"),
        result.swap.toOption.get.getMessage
      )
    }
  }

  test("autofilter: neither range nor --clear yields clean error") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Data"))
      val result = WriteCommands
        .autoFilter(wb, Some(wb.sheets.head), None, clear = false, out, config)
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
      assert(
        result.swap.toOption.get.getMessage.contains("--clear"),
        result.swap.toOption.get.getMessage
      )
    }
  }

  test("autofilter: qualified range targets the named sheet without --sheet") {
    withTempOutput { out =>
      val s1 = Sheet("Sheet1")
      val s2 = Sheet("Sheet2")
      val wb = Workbook(s1).put(s2)
      WriteCommands
        .autoFilter(wb, None, Some("Sheet2!A1:D4"), clear = false, out, config)
        .unsafeRunSync()

      val readWb = readBack(out)
      assertEquals(readWb.sheets.head.autoFilter, None)
      assertEquals(
        readWb.sheets.lift(1).flatMap(_.autoFilter),
        Some(AutoFilterState.Ranged(range("A1:D4")))
      )
    }
  }

  // ========== structural shift wiring (end-to-end from the CLI) ==========

  test("insert-rows shifts an authored autoFilter (write -> read -> insert -> re-read)") {
    withTempOutput { source =>
      withTempOutput { out =>
        val sheet = Sheet("Data")
          .put(ref"A2", CellValue.Text("Region"))
          .put(ref"B2", CellValue.Text("Sales"))
          .put(ref"A3", CellValue.Text("East"))
          .put(ref"B3", CellValue.Number(100))
        val wb0 = Workbook(sheet)
        WriteCommands
          .autoFilter(wb0, Some(sheet), Some("A2:B3"), clear = false, source, config)
          .unsafeRunSync()

        val wb1 = readBack(source)
        WriteCommands
          .insertRows(wb1, Some(wb1.sheets.head), at = 1, count = 2, out, config)
          .unsafeRunSync()

        val xml = sheetXml(out)
        assert(xml.contains("<autoFilter ref=\"A4:B5\"/>"), xml)
        assertEquals(
          readBack(out).sheets.head.autoFilter,
          Some(AutoFilterState.Ranged(range("A4:B5")))
        )
      }
    }
  }

  // ========== batch op: parse ==========

  test("batch: autofilter range form parses without warnings") {
    val json = """[{"op": "autofilter", "range": "A1:M29"}]"""
    val result = BatchParser.parseBatchJson(json)
    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get.warnings, Vector.empty[String])
  }

  test("batch: autofilter clear form parses without warnings") {
    val json = """[{"op": "autofilter", "clear": true}]"""
    val result = BatchParser.parseBatchJson(json)
    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get.warnings, Vector.empty[String])
  }

  test("batch: autofilter unknown property warns") {
    val json = """[{"op": "autofilter", "range": "A1:M29", "bogus": 1}]"""
    val result = BatchParser.parseBatchJson(json)
    assert(result.isRight, s"Should parse: $result")
    val warnings = result.toOption.get.warnings
    assert(warnings.exists(_.contains("bogus")), warnings.toString)
  }

  // ========== batch op: apply ==========

  test("batch: autofilter sets the sheet-level range on the model") {
    val sheet = Sheet("Data")
    val wb = Workbook(sheet)
    val ops = BatchParser
      .parseBatchJson("""[{"op": "autofilter", "range": "A1:M29"}]""")
      .toOption
      .get
      .ops
    val updated = BatchParser.applyBatchOperations(wb, Some(sheet), ops).unsafeRunSync()
    assertEquals(
      updated.sheets.head.autoFilter,
      Some(AutoFilterState.Ranged(range("A1:M29")))
    )
  }

  test("batch: autofilter clear sets the active-removal state") {
    val sheet = Sheet("Data")
    val wb = Workbook(sheet)
    val ops = BatchParser
      .parseBatchJson("""[{"op": "autofilter", "clear": true}]""")
      .toOption
      .get
      .ops
    val updated = BatchParser.applyBatchOperations(wb, Some(sheet), ops).unsafeRunSync()
    assertEquals(updated.sheets.head.autoFilter, Some(AutoFilterState.Remove))
  }

  test("batch: autofilter accepts a qualified range targeting another sheet") {
    val s1 = Sheet("Sheet1")
    val s2 = Sheet("Sheet2")
    val wb = Workbook(s1).put(s2)
    val ops = BatchParser
      .parseBatchJson("""[{"op": "autofilter", "range": "Sheet2!A1:B3"}]""")
      .toOption
      .get
      .ops
    val updated = BatchParser.applyBatchOperations(wb, Some(s1), ops).unsafeRunSync()
    assertEquals(updated.sheets.head.autoFilter, None)
    assertEquals(
      updated.sheets.lift(1).flatMap(_.autoFilter),
      Some(AutoFilterState.Ranged(range("A1:B3")))
    )
  }

  test("batch: autofilter range and clear are mutually exclusive at apply time") {
    val sheet = Sheet("Data")
    val wb = Workbook(sheet)
    val ops = BatchParser
      .parseBatchJson("""[{"op": "autofilter", "range": "A1:M29", "clear": true}]""")
      .toOption
      .get
      .ops
    val result = BatchParser.applyBatchOperations(wb, Some(sheet), ops).attempt.unsafeRunSync()
    assert(result.isLeft)
    assert(
      result.swap.toOption.get.getMessage.contains("mutually exclusive"),
      result.swap.toOption.get.getMessage
    )
  }

  test("batch: autofilter with neither range nor clear fails cleanly") {
    val sheet = Sheet("Data")
    val wb = Workbook(sheet)
    val ops = BatchParser
      .parseBatchJson("""[{"op": "autofilter"}]""")
      .toOption
      .get
      .ops
    val result = BatchParser.applyBatchOperations(wb, Some(sheet), ops).attempt.unsafeRunSync()
    assert(result.isLeft)
    assert(
      result.swap.toOption.get.getMessage.contains("range"),
      result.swap.toOption.get.getMessage
    )
  }

  test("batch: autofilter with unqualified range requires --sheet") {
    val ops = BatchParser
      .parseBatchJson("""[{"op": "autofilter", "range": "A1:M29"}]""")
      .toOption
      .get
      .ops
    val wb = Workbook(Sheet("Data"))
    val result = BatchParser.applyBatchOperations(wb, None, ops).attempt.unsafeRunSync()
    assert(result.isLeft)
    assert(
      result.swap.toOption.get.getMessage.contains("--sheet"),
      result.swap.toOption.get.getMessage
    )
  }
