package com.tjclp.xl.ooxml

import munit.FunSuite
import java.nio.file.{Files, Path}
import java.time.{LocalDate, LocalDateTime}
import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.api.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.Font

/**
 * Integration tests for writeFast (SAX/StAX backend) functionality.
 *
 * These tests verify that the streaming write backend produces valid XLSX files that can be read
 * back with all data intact. This addresses PR feedback about benchmarks not verifying output
 * correctness.
 *
 * writeFast uses XmlBackend.SaxStax for better performance on large files. These tests ensure the
 * SAX backend produces functionally equivalent output to the default (Scala XML) backend.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class WriteFastSpec extends FunSuite:

  val tempDir: Path = Files.createTempDirectory("xl-writefast-test-")

  override def afterAll(): Unit =
    // Clean up temp files
    Files.walk(tempDir)
      .sorted(java.util.Comparator.reverseOrder())
      .forEach(Files.delete)

  // Configuration for SAX/StAX backend (what writeFast uses internally)
  val saxConfig: WriterConfig = WriterConfig.default.copy(backend = XmlBackend.SaxStax)

  // ============================================================================
  // Basic Functionality Tests
  // ============================================================================

  test("writeFast produces valid xlsx readable by XlsxReader") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Header", ref"B1" -> 123, ref"C1" -> BigDecimal("99.99"))
      .put(ref"A2" -> "Data", ref"B2" -> 456, ref"C2" -> BigDecimal("50.00"))
    val wb = Workbook(Vector(sheet))

    val outputPath = tempDir.resolve("basic-fast.xlsx")
    val writeResult = XlsxWriter.writeWith(wb, outputPath, saxConfig)
    assert(writeResult.isRight, s"writeFast failed: $writeResult")

    // Read back and verify
    val readResult = XlsxReader.read(outputPath)
    assert(readResult.isRight, s"Read after writeFast failed: $readResult")

    readResult.foreach { readWb =>
      assertEquals(readWb.sheets.size, 1)
      val readSheet = readWb.sheets.head
      assertEquals(readSheet(ref"A1").value, CellValue.Text("Header"))
      assertEquals(readSheet(ref"B1").value, CellValue.Number(BigDecimal(123)))
      assertEquals(readSheet(ref"C1").value, CellValue.Number(BigDecimal("99.99")))
    }
  }

  test("writeFast and write produce functionally equivalent output") {
    val sheet = Sheet("Comparison")
      .put(ref"A1" -> "Test", ref"B1" -> 100)
      .put(ref"A2" -> 123.456, ref"B2" -> true)
      .put(ref"A3" -> "Multi\nLine", ref"B3" -> BigDecimal("1234567890.12"))
    val wb = Workbook(Vector(sheet))

    val fastPath = tempDir.resolve("comparison-fast.xlsx")
    val defaultPath = tempDir.resolve("comparison-default.xlsx")

    // Write with both backends
    val fastResult = XlsxWriter.writeWith(wb, fastPath, saxConfig)
    val defaultResult = XlsxWriter.writeWith(wb, defaultPath, WriterConfig.default)

    assert(fastResult.isRight, s"writeFast failed: $fastResult")
    assert(defaultResult.isRight, s"write failed: $defaultResult")

    // Read both back
    val fastWb = XlsxReader.read(fastPath)
    val defaultWb = XlsxReader.read(defaultPath)

    assert(fastWb.isRight, s"Read fast failed: $fastWb")
    assert(defaultWb.isRight, s"Read default failed: $defaultWb")

    // Compare cell values (not byte-identical due to different XML formatting)
    (fastWb, defaultWb) match
      case (Right(fwb), Right(dwb)) =>
        val fastSheet = fwb.sheets.head
        val defaultSheet = dwb.sheets.head
        assertEquals(fastSheet.cells.keySet, defaultSheet.cells.keySet, "Cell refs mismatch")
        fastSheet.cells.foreach { case (cellRef, cell) =>
          assertEquals(cell.value, defaultSheet(cellRef).value, s"Mismatch at $cellRef")
        }
      case _ => fail("Both reads should succeed")
  }

  // ============================================================================
  // Data Type Tests
  // ============================================================================

  test("writeFast preserves all cell value types") {
    val today = LocalDate.of(2025, 6, 15)
    val now = LocalDateTime.of(2025, 6, 15, 14, 30, 0)

    val sheet = Sheet("Types")
      .put(ref"A1" -> "Text value")
      .put(ref"A2" -> CellValue.Number(BigDecimal(42)))
      .put(ref"A3" -> CellValue.Bool(true))
      .put(ref"A4" -> CellValue.Bool(false))
      .put(ref"A5" -> CellValue.DateTime(today.atStartOfDay()))
      .put(ref"A6" -> CellValue.DateTime(now))
      .put(ref"A7" -> CellValue.Empty)
      .put(ref"A8" -> CellValue.Error(com.tjclp.xl.cells.CellError.Div0))
    val wb = Workbook(Vector(sheet))

    val outputPath = tempDir.resolve("types-fast.xlsx")
    val writeResult = XlsxWriter.writeWith(wb, outputPath, saxConfig)
    assert(writeResult.isRight, s"writeFast failed: $writeResult")

    val readResult = XlsxReader.read(outputPath)
    assert(readResult.isRight, s"Read failed: $readResult")

    readResult.foreach { readWb =>
      val readSheet = readWb.sheets.head
      assertEquals(readSheet(ref"A1").value, CellValue.Text("Text value"))
      assertEquals(readSheet(ref"A2").value, CellValue.Number(BigDecimal(42)))
      assertEquals(readSheet(ref"A3").value, CellValue.Bool(true))
      assertEquals(readSheet(ref"A4").value, CellValue.Bool(false))
      // Date/DateTime are stored as serial numbers, verify they round-trip correctly
      readSheet(ref"A5").value match
        case CellValue.Number(serial) =>
          val roundTripped = CellValue.excelSerialToDateTime(serial.toDouble).toLocalDate
          assertEquals(roundTripped, today)
        case other => fail(s"Expected Number (serial), got $other")
      readSheet(ref"A6").value match
        case CellValue.Number(serial) =>
          val roundTripped = CellValue.excelSerialToDateTime(serial.toDouble)
          // Allow up to 60 seconds difference due to floating point precision loss
          assert(
            java.time.Duration.between(now, roundTripped).abs().toSeconds < 60,
            s"Expected ~$now, got $roundTripped"
          )
        case other => fail(s"Expected Number (serial), got $other")
    }
  }

  test("writeFast preserves formulas") {
    val sheet = Sheet("Formulas")
      .put(ref"A1" -> CellValue.Number(BigDecimal(10)))
      .put(ref"A2" -> CellValue.Number(BigDecimal(20)))
      .put(ref"A3" -> CellValue.Formula("SUM(A1:A2)", Some(CellValue.Number(BigDecimal(30)))))
    val wb = Workbook(Vector(sheet))

    val outputPath = tempDir.resolve("formulas-fast.xlsx")
    val writeResult = XlsxWriter.writeWith(wb, outputPath, saxConfig)
    assert(writeResult.isRight, s"writeFast failed: $writeResult")

    val readResult = XlsxReader.read(outputPath)
    assert(readResult.isRight, s"Read failed: $readResult")

    readResult.foreach { readWb =>
      val readSheet = readWb.sheets.head
      readSheet(ref"A3").value match
        case CellValue.Formula(expr, cached) =>
          assertEquals(expr, "SUM(A1:A2)")
          assertEquals(cached, Some(CellValue.Number(BigDecimal(30))))
        case other => fail(s"Expected Formula, got $other")
    }
  }

  // ============================================================================
  // Multi-Sheet Tests
  // ============================================================================

  test("writeFast handles multiple sheets") {
    val sheet1 = Sheet("Data").put(ref"A1" -> "Sheet 1")
    val sheet2 = Sheet("Summary").put(ref"A1" -> "Sheet 2")
    val sheet3 = Sheet("Analysis").put(ref"A1" -> "Sheet 3")
    val wb = Workbook(Vector(sheet1, sheet2, sheet3))

    val outputPath = tempDir.resolve("multi-sheet-fast.xlsx")
    val writeResult = XlsxWriter.writeWith(wb, outputPath, saxConfig)
    assert(writeResult.isRight, s"writeFast failed: $writeResult")

    val readResult = XlsxReader.read(outputPath)
    assert(readResult.isRight, s"Read failed: $readResult")

    readResult.foreach { readWb =>
      assertEquals(readWb.sheets.size, 3)
      assertEquals(readWb.sheets(0).name.value, "Data")
      assertEquals(readWb.sheets(1).name.value, "Summary")
      assertEquals(readWb.sheets(2).name.value, "Analysis")
    }
  }

  // ============================================================================
  // Style Tests
  // ============================================================================

  test("writeFast preserves cell styles") {
    val boldStyle = CellStyle.default.withFont(Font("Arial", 12.0, bold = true))
    val sheet = Sheet("Styles")
      .put(ref"A1" -> "Bold text")
      .withCellStyle(ref"A1", boldStyle)
    val wb = Workbook(Vector(sheet))

    val outputPath = tempDir.resolve("styles-fast.xlsx")
    val writeResult = XlsxWriter.writeWith(wb, outputPath, saxConfig)
    assert(writeResult.isRight, s"writeFast failed: $writeResult")

    val readResult = XlsxReader.read(outputPath)
    assert(readResult.isRight, s"Read failed: $readResult")

    readResult.foreach { readWb =>
      val readSheet = readWb.sheets.head
      val cell = readSheet(ref"A1")
      assert(cell.styleId.isDefined, "Cell should have a style")
      // Style is applied - exact style ID may differ but cell has styling
    }
  }
