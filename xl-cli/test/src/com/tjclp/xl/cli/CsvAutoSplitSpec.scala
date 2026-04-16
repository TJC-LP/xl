package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Tests for CSV auto-split in the put command.
 *
 * When put receives a single comma-separated value targeting a range, it should auto-split and
 * distribute the values across cells if the element count matches the range size.
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class CsvAutoSplitSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val outputPath: Path = Files.createTempFile("csv-split-test", ".xlsx")
  val config: WriterConfig = WriterConfig.default

  override def afterEach(context: AfterEach): Unit =
    if Files.exists(outputPath) then Files.delete(outputPath)

  test("put: CSV auto-split distributes comma-separated values across range") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands
        .put(wb, Some(wb.sheets.head), "A1:C1", List("1,2,3"), outputPath, config)
        .unsafeRunSync()

    // Should behave like batch values mode (not fill pattern)
    assert(result.contains("Put 3 values"))
    assert(result.contains("row-major"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    assertEquals(
      sheet.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal("1")))
    ) // A1
    assertEquals(
      sheet.cells.get(ARef.from0(1, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal("2")))
    ) // B1
    assertEquals(
      sheet.cells.get(ARef.from0(2, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal("3")))
    ) // C1
  }

  test("put: CSV auto-split trims whitespace around values") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands
        .put(wb, Some(wb.sheets.head), "A1:C1", List("Q1 , Q2 , Q3"), outputPath, config)
        .unsafeRunSync()

    assert(result.contains("Put 3 values"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    assertEquals(
      sheet.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Text("Q1"))
    )
    assertEquals(
      sheet.cells.get(ARef.from0(1, 0)).map(_.value),
      Some(CellValue.Text("Q2"))
    )
    assertEquals(
      sheet.cells.get(ARef.from0(2, 0)).map(_.value),
      Some(CellValue.Text("Q3"))
    )
  }

  test("put: single cell with commas stays as literal text") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands
        .put(wb, Some(wb.sheets.head), "A1", List("hello,world"), outputPath, config)
        .unsafeRunSync()

    // Single cell mode: never split
    assert(result.contains("Put: A1"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    assertEquals(
      sheet.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Text("hello,world"))
    )
  }

  test("put: CSV auto-split count mismatch falls back to literal fill") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands
        .put(wb, Some(wb.sheets.head), "A1:D1", List("a,b,c"), outputPath, config)
        .unsafeRunSync()

    // 3 CSV parts vs 4 cells: falls back to fill pattern (literal text)
    assert(result.contains("Filled 4 cells"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    // All cells should have the literal text "a,b,c"
    (0 to 3).foreach { col =>
      assertEquals(
        sheet.cells.get(ARef.from0(col, 0)).map(_.value),
        Some(CellValue.Text("a,b,c"))
      )
    }
  }

  test("put: CSV auto-split applies smart type detection") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands
        .put(wb, Some(wb.sheets.head), "A1:C1", List("true,42,hello"), outputPath, config)
        .unsafeRunSync()

    assert(result.contains("Put 3 values"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    assertEquals(
      sheet.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Bool(true))
    ) // A1
    assertEquals(
      sheet.cells.get(ARef.from0(1, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal("42")))
    ) // B1
    assertEquals(
      sheet.cells.get(ARef.from0(2, 0)).map(_.value),
      Some(CellValue.Text("hello"))
    ) // C1
  }
