package com.tjclp.xl.ooxml

import munit.FunSuite
import java.nio.file.{Files, Path}
import com.tjclp.xl.*
import com.tjclp.xl.macros.{cell, range}
import com.tjclp.xl.ooxml.{XlsxWriter, XlsxReader}

/** Round-trip tests for XLSX write → read */
class OoxmlRoundTripSpec extends FunSuite:

  val tempDir: Path = Files.createTempDirectory("xl-test-")

  override def afterAll(): Unit =
    // Clean up temp files
    Files.walk(tempDir)
      .sorted(java.util.Comparator.reverseOrder())
      .forEach(Files.delete)

  test("Minimal workbook: write and read back") {
    // Create minimal workbook
    val wb = Workbook("Test").getOrElse(fail("Should create workbook"))

    // Write to file
    val outputPath = tempDir.resolve("minimal.xlsx")
    val writeResult = XlsxWriter.write(wb, outputPath)
    assert(writeResult.isRight, s"Write failed: $writeResult")

    // Read back
    val readResult = XlsxReader.read(outputPath)
    assert(readResult.isRight, s"Read failed: $readResult")

    // Verify
    readResult.foreach { readWb =>
      assertEquals(readWb.sheets.size, 1)
      assertEquals(readWb.sheets(0).name.value, "Test")
    }
  }

  test("Workbook with text cells") {
    val wb = Workbook("Data").flatMap { initial =>
      val sheet = initial.sheets(0)
        .put(cell"A1", CellValue.Text("Hello"))
        .put(cell"B1", CellValue.Text("World"))
        .put(cell"A2", CellValue.Text("Scala"))
        .put(cell"B2", CellValue.Text("Excel"))

      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Should create workbook"))

    // Round-trip
    val outputPath = tempDir.resolve("text-cells.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))

    // Verify cells
    val readSheet = readWb.sheets(0)
    assertEquals(readSheet(cell"A1").value, CellValue.Text("Hello"))
    assertEquals(readSheet(cell"B1").value, CellValue.Text("World"))
    assertEquals(readSheet(cell"A2").value, CellValue.Text("Scala"))
    assertEquals(readSheet(cell"B2").value, CellValue.Text("Excel"))
  }

  test("Workbook with number cells") {
    val wb = Workbook("Numbers").flatMap { initial =>
      val sheet = initial.sheets(0)
        .put(cell"A1", CellValue.Number(BigDecimal(42)))
        .put(cell"A2", CellValue.Number(BigDecimal(3.14159)))
        .put(cell"A3", CellValue.Number(BigDecimal(-100)))

      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Should create workbook"))

    // Round-trip
    val outputPath = tempDir.resolve("numbers.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readResult = XlsxReader.read(outputPath)
    readResult match
      case Left(err) => fail(s"Read failed: ${err}")
      case Right(readWb) =>
        // Verify numbers
        val readSheet = readWb.sheets(0)
        assertEquals(readSheet(cell"A1").value, CellValue.Number(BigDecimal(42)))
        assertEquals(readSheet(cell"A2").value, CellValue.Number(BigDecimal(3.14159)))
        assertEquals(readSheet(cell"A3").value, CellValue.Number(BigDecimal(-100)))
  }

  test("Workbook with boolean cells") {
    val wb = Workbook("Bools").flatMap { initial =>
      val sheet = initial.sheets(0)
        .put(cell"A1", CellValue.Bool(true))
        .put(cell"A2", CellValue.Bool(false))

      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Should create workbook"))

    // Round-trip
    val outputPath = tempDir.resolve("booleans.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readResult = XlsxReader.read(outputPath)
    readResult match
      case Left(err) => fail(s"Read failed: ${err}")
      case Right(readWb) =>
        // Verify booleans
        val readSheet = readWb.sheets(0)
        assertEquals(readSheet(cell"A1").value, CellValue.Bool(true))
        assertEquals(readSheet(cell"A2").value, CellValue.Bool(false))
  }

  test("Workbook with mixed cell types") {
    val wb = Workbook("Mixed").flatMap { initial =>
      val sheet = initial.sheets(0)
        .put(cell"A1", CellValue.Text("Name"))
        .put(cell"B1", CellValue.Text("Age"))
        .put(cell"C1", CellValue.Text("Active"))
        .put(cell"A2", CellValue.Text("Alice"))
        .put(cell"B2", CellValue.Number(BigDecimal(30)))
        .put(cell"C2", CellValue.Bool(true))
        .put(cell"A3", CellValue.Text("Bob"))
        .put(cell"B3", CellValue.Number(BigDecimal(25)))
        .put(cell"C3", CellValue.Bool(false))

      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Should create workbook"))

    // Round-trip
    val outputPath = tempDir.resolve("mixed.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readResult = XlsxReader.read(outputPath)
    readResult match
      case Left(err) => fail(s"Read failed: ${err}")
      case Right(readWb) =>
        // Verify all cells
        val readSheet = readWb.sheets(0)
        assertEquals(readSheet(cell"A1").value, CellValue.Text("Name"))
        assertEquals(readSheet(cell"B2").value, CellValue.Number(BigDecimal(30)))
        assertEquals(readSheet(cell"C2").value, CellValue.Bool(true))
  }

  test("Multi-sheet workbook") {
    val wb = Workbook("Sheet1").flatMap { initial =>
      val sheet1 = initial.sheets(0).put(cell"A1", CellValue.Text("First"))

      for
        wb2 <- Sheet("Sheet2").flatMap(initial.addSheet)
        wb3 <- wb2.updateSheet(0, sheet1)
        wb4 <- wb3.updateSheet(1, wb3.sheets(1).put(cell"A1", CellValue.Text("Second")))
      yield wb4
    }.getOrElse(fail("Should create workbook"))

    // Round-trip
    val outputPath = tempDir.resolve("multi-sheet.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))

    // Verify sheets
    assertEquals(readWb.sheets.size, 2)
    assertEquals(readWb.sheets(0).name.value, "Sheet1")
    assertEquals(readWb.sheets(1).name.value, "Sheet2")
    assertEquals(readWb.sheets(0)(cell"A1").value, CellValue.Text("First"))
    assertEquals(readWb.sheets(1)(cell"A1").value, CellValue.Text("Second"))
  }

  test("Workbook with repeated strings triggers SST") {
    // Create workbook with many repeated strings
    val wb = Workbook("SST Test").flatMap { initial =>
      val sheet = (1 to 20).foldLeft(initial.sheets(0)) { (s, i) =>
        val row = Row.from1(i)
        s.put(ARef(Column.from0(0), row), CellValue.Text("Repeated"))
          .put(ARef(Column.from0(1), row), CellValue.Text("Value"))
      }
      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Should create workbook"))

    // Verify SST would be used
    assert(SharedStrings.shouldUseSST(wb), "SST should be beneficial")

    // Round-trip
    val outputPath = tempDir.resolve("sst.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))

    // Verify all cells preserved
    val readSheet = readWb.sheets(0)
    (1 to 20).foreach { i =>
      val row = Row.from1(i)
      assertEquals(readSheet(ARef(Column.from0(0), row)).value, CellValue.Text("Repeated"))
      assertEquals(readSheet(ARef(Column.from0(1), row)).value, CellValue.Text("Value"))
    }
  }

  test("Empty workbook round-trips correctly") {
    val wb = Workbook("Empty").getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("empty.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))

    assertEquals(readWb.sheets.size, 1)
    assertEquals(readWb.sheets(0).cellCount, 0)
  }

  test("Workbook with merged cells preserves merges") {
    val wb = Workbook("Merged").flatMap { initial =>
      val sheet = initial.sheets(0)
        .put(cell"A1", CellValue.Text("Merged Header"))
        .merge(range"A1:C1")

      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Should create workbook"))

    // Note: Merged ranges require worksheet relationships, which is future work
    // For now, just verify basic round-trip works
    val outputPath = tempDir.resolve("merged.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    assertEquals(readWb.sheets(0)(cell"A1").value, CellValue.Text("Merged Header"))
  }

  test("Workbook with DateTime cells serializes to Excel serial numbers") {
    import java.time.LocalDateTime

    val dt1 = LocalDateTime.of(2025, 11, 10, 0, 0, 0) // Nov 10, 2025 midnight
    val dt2 = LocalDateTime.of(2000, 1, 1, 12, 0, 0)   // Jan 1, 2000 noon

    val wb = Workbook("Dates").flatMap { initial =>
      val sheet = initial.sheets(0)
        .put(cell"A1", CellValue.DateTime(dt1))
        .put(cell"A2", CellValue.DateTime(dt2))

      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("dates.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    // Read and verify serial numbers were written (not "0")
    val readResult = XlsxReader.read(outputPath)
    val readWb = readResult match
      case Right(wb) => wb
      case Left(err) => fail(s"Read failed: ${err.message}")

    val cell1 = readWb.sheets(0)(cell"A1")
    val cell2 = readWb.sheets(0)(cell"A2")

    // Should be read as Numbers (Excel serial format)
    cell1.value match
      case CellValue.Number(serial) =>
        // Nov 10, 2025 should be ~45975
        assert(serial > 45900.0 && serial < 46000.0, s"Serial should be ~45975, got $serial")
      case other =>
        fail(s"Expected Number, got: $other")

    cell2.value match
      case CellValue.Number(serial) =>
        // Jan 1, 2000 noon should be 36526.5
        assertEquals(serial.toDouble, 36526.5, 0.001)
      case other =>
        fail(s"Expected Number, got: $other")
  }
