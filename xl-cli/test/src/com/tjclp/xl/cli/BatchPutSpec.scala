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
 * Integration tests for batch put and fill pattern functionality.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class BatchPutSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val outputPath: Path = Files.createTempFile("test", ".xlsx")
  val config: WriterConfig = WriterConfig.default

  override def afterEach(context: AfterEach): Unit =
    if Files.exists(outputPath) then Files.delete(outputPath)

  // ========== Put Command Mode Detection ==========

  test("put: single cell mode") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands.put(wb, Some(wb.sheets.head), "A1", List("100"), outputPath, config).unsafeRunSync()

    assert(result.contains("Put: A1 = 100"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val cellValue = imported.sheets.head.cells.get(ARef.from0(0, 0)).map(_.value)
    assertEquals(cellValue, Some(CellValue.Number(BigDecimal("100"))))
  }

  test("put: fill pattern mode (1 value, range)") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands.put(wb, Some(wb.sheets.head), "A1:A5", List("TBD"), outputPath, config).unsafeRunSync()

    assert(result.contains("Filled 5 cells"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    (0 to 4).foreach { row =>
      assertEquals(sheet.cells.get(ARef.from0(0, row)).map(_.value), Some(CellValue.Text("TBD")))
    }
  }

  test("put: batch values mode (N values, N-cell range)") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands
        .put(wb, Some(wb.sheets.head), "A1:A3", List("10", "20", "30"), outputPath, config)
        .unsafeRunSync()

    assert(result.contains("Put 3 values"))
    assert(result.contains("row-major"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    assertEquals(sheet.cells.get(ARef.from0(0, 0)).map(_.value), Some(CellValue.Number(BigDecimal("10"))))
    assertEquals(sheet.cells.get(ARef.from0(0, 1)).map(_.value), Some(CellValue.Number(BigDecimal("20"))))
    assertEquals(sheet.cells.get(ARef.from0(0, 2)).map(_.value), Some(CellValue.Number(BigDecimal("30"))))
  }

  test("put: batch values 2D range (row-major)") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands
        .put(wb, Some(wb.sheets.head), "A1:B2", List("1", "2", "3", "4"), outputPath, config)
        .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    // Row-major: A1, B1, A2, B2
    assertEquals(sheet.cells.get(ARef.from0(0, 0)).map(_.value), Some(CellValue.Number(BigDecimal("1")))) // A1
    assertEquals(sheet.cells.get(ARef.from0(1, 0)).map(_.value), Some(CellValue.Number(BigDecimal("2")))) // B1
    assertEquals(sheet.cells.get(ARef.from0(0, 1)).map(_.value), Some(CellValue.Number(BigDecimal("3")))) // A2
    assertEquals(sheet.cells.get(ARef.from0(1, 1)).map(_.value), Some(CellValue.Number(BigDecimal("4")))) // B2
  }

  test("put: error - count mismatch (too few values)") {
    val wb = Workbook(Sheet("Test"))
    val result = WriteCommands
      .put(wb, Some(wb.sheets.head), "A1:A5", List("1", "2", "3"), outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    val error = result.swap.getOrElse(throw new Exception("Expected error"))
    assert(error.getMessage.contains("5 cells but 3 values"))
    assert(error.getMessage.contains("Hint"))
  }

  test("put: error - count mismatch (too many values)") {
    val wb = Workbook(Sheet("Test"))
    val result = WriteCommands
      .put(wb, Some(wb.sheets.head), "A1:A2", List("1", "2", "3", "4"), outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    val error = result.swap.getOrElse(throw new Exception("Expected error"))
    assert(error.getMessage.contains("2 cells but 4 values"))
  }

  test("put: error - multiple values to single cell") {
    val wb = Workbook(Sheet("Test"))
    val result = WriteCommands
      .put(wb, Some(wb.sheets.head), "A1", List("1", "2", "3"), outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    val error = result.swap.getOrElse(throw new Exception("Expected error"))
    assert(error.getMessage.contains("Cannot put 3 values to single cell"))
    assert(error.getMessage.contains("Use a range"))
  }

  // ========== Putf Command Mode Detection ==========

  test("putf: single cell mode") {
    val wb = Workbook(Sheet("Test").put(ARef.from0(0, 0), CellValue.Number(BigDecimal("10"))))
    val result = WriteCommands
      .putFormula(wb, Some(wb.sheets.head), "B1", List("=A1*2"), outputPath, config)
      .unsafeRunSync()

    assert(result.contains("Put: B1"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    sheet.cells.get(ARef.from0(1, 0)).map(_.value) match
      case Some(CellValue.Formula(formula, _)) =>
        assertEquals(formula, "A1*2")
      case other => fail(s"Expected Formula, got $other")
  }

  test("putf: formula dragging mode preserves $ anchors") {
    val wb = Workbook(Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Number(BigDecimal("10")))
      .put(ARef.from0(0, 1), CellValue.Number(BigDecimal("20")))
      .put(ARef.from0(0, 2), CellValue.Number(BigDecimal("30"))))

    val result = WriteCommands
      .putFormula(wb, Some(wb.sheets.head), "B1:B3", List("=SUM($A$1:A1)"), outputPath, config)
      .unsafeRunSync()

    assert(result.contains("with anchor-aware dragging"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    // Verify $ anchors work
    sheet.cells.get(ARef.from0(1, 0)).map(_.value) match
      case Some(CellValue.Formula(formula, _)) => assert(formula.contains("$A$1"))
      case other => fail(s"Expected Formula with $$A$$1, got $other")
  }

  test("putf: batch formulas mode (no dragging)") {
    val wb = Workbook(Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Number(BigDecimal("1")))
      .put(ARef.from0(1, 0), CellValue.Number(BigDecimal("2")))
      .put(ARef.from0(0, 1), CellValue.Number(BigDecimal("3")))
      .put(ARef.from0(1, 1), CellValue.Number(BigDecimal("4"))))

    val result = WriteCommands
      .putFormula(
        wb,
        Some(wb.sheets.head),
        "C1:C3",
        List("=A1+B1", "=A2*B2", "=100"),
        outputPath,
        config
      )
      .unsafeRunSync()

    assert(result.contains("explicit, no dragging"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    // Verify formulas are as-is (no dragging)
    sheet.cells.get(ARef.from0(2, 0)).map(_.value) match
      case Some(CellValue.Formula(formula, _)) => assertEquals(formula, "A1+B1")
      case other => fail(s"Expected Formula A1+B1, got $other")
    sheet.cells.get(ARef.from0(2, 1)).map(_.value) match
      case Some(CellValue.Formula(formula, _)) => assertEquals(formula, "A2*B2")
      case other => fail(s"Expected Formula A2*B2, got $other")
    sheet.cells.get(ARef.from0(2, 2)).map(_.value) match
      case Some(CellValue.Formula(formula, _)) => assertEquals(formula, "100")
      case other => fail(s"Expected Formula 100, got $other")
  }

  test("putf: error - count mismatch") {
    val wb = Workbook(Sheet("Test"))
    val result = WriteCommands
      .putFormula(wb, Some(wb.sheets.head), "B1:B5", List("=X", "=Y"), outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    val error = result.swap.getOrElse(throw new Exception("Expected error"))
    assert(error.getMessage.contains("5 cells but 2 formulas"))
    assert(error.getMessage.contains("Hint"))
  }
