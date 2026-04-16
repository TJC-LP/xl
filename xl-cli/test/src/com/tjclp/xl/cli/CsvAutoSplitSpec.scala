package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Sheet, Workbook}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Tests for the opt-in `--csv` flag on the put command.
 *
 * The flag enables splitting a single comma-separated value across a target range. Without the
 * flag, comma-containing values are written as literal text (the safe default).
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

  test("put --csv: splits comma-separated values across range with matching count") {
    val wb = Workbook(Sheet("Test"))
    val result = WriteCommands
      .put(
        wb,
        Some(wb.sheets.head),
        "A1:C1",
        List("1,2,3"),
        outputPath,
        config,
        stream = false,
        csvSplit = true
      )
      .unsafeRunSync()

    assert(result.contains("Put 3 values"), s"Unexpected message: $result")

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    assertEquals(
      sheet.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal("1")))
    )
    assertEquals(
      sheet.cells.get(ARef.from0(1, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal("2")))
    )
    assertEquals(
      sheet.cells.get(ARef.from0(2, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal("3")))
    )
  }

  test("put --csv: trims whitespace around split values") {
    val wb = Workbook(Sheet("Test"))
    WriteCommands
      .put(
        wb,
        Some(wb.sheets.head),
        "A1:C1",
        List("Q1 , Q2 , Q3"),
        outputPath,
        config,
        stream = false,
        csvSplit = true
      )
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    assertEquals(
      sheet.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Text("Q1"))
    )
    assertEquals(
      sheet.cells.get(ARef.from0(2, 0)).map(_.value),
      Some(CellValue.Text("Q3"))
    )
  }

  test("put --csv: errors when split count does not match range size") {
    val wb = Workbook(Sheet("Test"))
    val err = WriteCommands
      .put(
        wb,
        Some(wb.sheets.head),
        "A1:D1",
        List("a,b,c"),
        outputPath,
        config,
        stream = false,
        csvSplit = true
      )
      .attempt
      .unsafeRunSync()

    assert(err.isLeft, s"Expected error but got: $err")
    val msg = err.swap.getOrElse(throw new Exception("expected left")).getMessage
    assert(msg.contains("--csv"), s"Expected error to mention --csv: $msg")
  }

  test("put --csv: rejects single-cell target") {
    val wb = Workbook(Sheet("Test"))
    val err = WriteCommands
      .put(
        wb,
        Some(wb.sheets.head),
        "A1",
        List("hello,world"),
        outputPath,
        config,
        stream = false,
        csvSplit = true
      )
      .attempt
      .unsafeRunSync()

    assert(err.isLeft, "Expected error for --csv with single-cell target")
    val msg = err.swap.getOrElse(throw new Exception("expected left")).getMessage
    assert(msg.contains("single cell"))
  }

  test("put (no --csv): comma-containing value stays as literal fill across range") {
    val wb = Workbook(Sheet("Test"))
    WriteCommands
      .put(
        wb,
        Some(wb.sheets.head),
        "A1:B1",
        List("Smith, John"),
        outputPath,
        config
      )
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    // Default behavior: fill both cells with the literal comma-containing string
    assertEquals(
      sheet.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Text("Smith, John"))
    )
    assertEquals(
      sheet.cells.get(ARef.from0(1, 0)).map(_.value),
      Some(CellValue.Text("Smith, John"))
    )
  }

  test("put (no --csv): single cell with commas is literal text") {
    val wb = Workbook(Sheet("Test"))
    WriteCommands
      .put(
        wb,
        Some(wb.sheets.head),
        "A1",
        List("hello,world"),
        outputPath,
        config
      )
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    assertEquals(
      sheet.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Text("hello,world"))
    )
  }

  test("put --csv: applies smart type detection to each split value") {
    val wb = Workbook(Sheet("Test"))
    WriteCommands
      .put(
        wb,
        Some(wb.sheets.head),
        "A1:C1",
        List("true,42,hello"),
        outputPath,
        config,
        stream = false,
        csvSplit = true
      )
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    assertEquals(
      sheet.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Bool(true))
    )
    assertEquals(
      sheet.cells.get(ARef.from0(1, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal("42")))
    )
    assertEquals(
      sheet.cells.get(ARef.from0(2, 0)).map(_.value),
      Some(CellValue.Text("hello"))
    )
  }
