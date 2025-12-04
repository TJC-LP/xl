package com.tjclp.xl.ooxml

import munit.FunSuite
import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.time.{Duration, LocalDate, LocalDateTime}
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.error.XLResult
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{XlsxWriter, XlsxReader}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.richtext.RichText.given
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.dsl.*
import java.util.zip.ZipInputStream

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
    val wb = Workbook("Test")

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
    val initial = Workbook("Data")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Text("Hello"))
      .put(ref"B1", CellValue.Text("World"))
      .put(ref"A2", CellValue.Text("Scala"))
      .put(ref"B2", CellValue.Text("Excel"))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    // Round-trip
    val outputPath = tempDir.resolve("text-cells.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))

    // Verify cells
    val readSheet = readWb.sheets(0)
    assertEquals(readSheet(ref"A1").value, CellValue.Text("Hello"))
    assertEquals(readSheet(ref"B1").value, CellValue.Text("World"))
    assertEquals(readSheet(ref"A2").value, CellValue.Text("Scala"))
    assertEquals(readSheet(ref"B2").value, CellValue.Text("Excel"))
  }

  test("Workbook with number cells") {
    val initial = Workbook("Numbers")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Number(BigDecimal(42)))
      .put(ref"A2", CellValue.Number(BigDecimal(3.14159)))
      .put(ref"A3", CellValue.Number(BigDecimal(-100)))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    // Round-trip
    val outputPath = tempDir.resolve("numbers.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readResult = XlsxReader.read(outputPath)
    readResult match
      case Left(err) => fail(s"Read failed: ${err}")
      case Right(readWb) =>
        // Verify numbers
        val readSheet = readWb.sheets(0)
        assertEquals(readSheet(ref"A1").value, CellValue.Number(BigDecimal(42)))
        assertEquals(readSheet(ref"A2").value, CellValue.Number(BigDecimal(3.14159)))
        assertEquals(readSheet(ref"A3").value, CellValue.Number(BigDecimal(-100)))
  }

  test("Workbook with boolean cells") {
    val initial = Workbook("Bools")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Bool(true))
      .put(ref"A2", CellValue.Bool(false))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    // Round-trip
    val outputPath = tempDir.resolve("booleans.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readResult = XlsxReader.read(outputPath)
    readResult match
      case Left(err) => fail(s"Read failed: ${err}")
      case Right(readWb) =>
        // Verify booleans
        val readSheet = readWb.sheets(0)
        assertEquals(readSheet(ref"A1").value, CellValue.Bool(true))
        assertEquals(readSheet(ref"A2").value, CellValue.Bool(false))
  }

  test("Workbook with mixed cell types") {
    val initial = Workbook("Mixed")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Text("Name"))
      .put(ref"B1", CellValue.Text("Age"))
      .put(ref"C1", CellValue.Text("Active"))
      .put(ref"A2", CellValue.Text("Alice"))
      .put(ref"B2", CellValue.Number(BigDecimal(30)))
      .put(ref"C2", CellValue.Bool(true))
      .put(ref"A3", CellValue.Text("Bob"))
      .put(ref"B3", CellValue.Number(BigDecimal(25)))
      .put(ref"C3", CellValue.Bool(false))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    // Round-trip
    val outputPath = tempDir.resolve("mixed.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readResult = XlsxReader.read(outputPath)
    readResult match
      case Left(err) => fail(s"Read failed: ${err}")
      case Right(readWb) =>
        // Verify all cells
        val readSheet = readWb.sheets(0)
        assertEquals(readSheet(ref"A1").value, CellValue.Text("Name"))
        assertEquals(readSheet(ref"B2").value, CellValue.Number(BigDecimal(30)))
        assertEquals(readSheet(ref"C2").value, CellValue.Bool(true))
  }

  test("Multi-sheet workbook") {
    val initial = Workbook("Sheet1")
    val sheet1 = initial.sheets(0).put(ref"A1", CellValue.Text("First"))

    val wb2 = initial.put(Sheet("Sheet2"))
    val wb3 = wb2.update(wb2.sheets(0).name, _ => sheet1).getOrElse(fail("Should update sheet1"))
    val wb = wb3.update(wb3.sheets(1).name, _ => wb3.sheets(1).put(ref"A1", CellValue.Text("Second"))).getOrElse(fail("Should update sheet2"))

    // Round-trip
    val outputPath = tempDir.resolve("multi-sheet.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))

    // Verify sheets
    assertEquals(readWb.sheets.size, 2)
    assertEquals(readWb.sheets(0).name.value, "Sheet1")
    assertEquals(readWb.sheets(1).name.value, "Sheet2")
    assertEquals(readWb.sheets(0)(ref"A1").value, CellValue.Text("First"))
    assertEquals(readWb.sheets(1)(ref"A1").value, CellValue.Text("Second"))
  }

  test("Workbook with repeated strings triggers SST") {
    // Create workbook with many repeated strings
    val initial = Workbook("SST Test")
    val sheet = (1 to 20).foldLeft(initial.sheets(0)) { (s, i) =>
      val row = Row.from1(i)
      s.put(ARef(Column.from0(0), row), CellValue.Text("Repeated"))
        .put(ARef(Column.from0(1), row), CellValue.Text("Value"))
    }
    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

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
    val wb = Workbook("Empty")

    val outputPath = tempDir.resolve("empty.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))

    assertEquals(readWb.sheets.size, 1)
    assertEquals(readWb.sheets(0).cellCount, 0)
  }

  // ===== Formula Roundtrip Tests =====

  test("Formula cell roundtrips expression without cached value") {
    val initial = Workbook("Formulas")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Number(BigDecimal(100)))
      .put(ref"B1", CellValue.Number(BigDecimal(200)))
      .put(ref"C1", CellValue.Formula("A1+B1", None))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("formula-no-cache.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    readSheet(ref"C1").value match
      case CellValue.Formula(expr, _) =>
        assertEquals(expr, "A1+B1")
      case other =>
        fail(s"Expected Formula, got $other")
  }

  test("Formula cell roundtrips expression with cached numeric value") {
    val initial = Workbook("Formulas")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Number(BigDecimal(100)))
      .put(ref"B1", CellValue.Number(BigDecimal(200)))
      .put(ref"C1", CellValue.Formula("A1+B1", Some(CellValue.Number(BigDecimal(300)))))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("formula-cached-number.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    readSheet(ref"C1").value match
      case CellValue.Formula(expr, Some(CellValue.Number(cached))) =>
        assertEquals(expr, "A1+B1")
        assertEquals(cached, BigDecimal(300))
      case CellValue.Formula(expr, None) =>
        fail(s"Formula $expr missing cached value")
      case other =>
        fail(s"Expected Formula with cached Number, got $other")
  }

  test("Formula cell roundtrips expression with cached text value") {
    val initial = Workbook("Formulas")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Text("Hello"))
      .put(ref"B1", CellValue.Text(" World"))
      .put(ref"C1", CellValue.Formula("CONCATENATE(A1,B1)", Some(CellValue.Text("Hello World"))))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("formula-cached-text.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    readSheet(ref"C1").value match
      case CellValue.Formula(expr, Some(CellValue.Text(cached))) =>
        assertEquals(expr, "CONCATENATE(A1,B1)")
        assertEquals(cached, "Hello World")
      case CellValue.Formula(expr, cached) =>
        fail(s"Formula $expr has unexpected cached value: $cached")
      case other =>
        fail(s"Expected Formula with cached Text, got $other")
  }

  test("Formula cell roundtrips expression with cached boolean value") {
    val initial = Workbook("Formulas")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Number(BigDecimal(100)))
      .put(ref"B1", CellValue.Number(BigDecimal(50)))
      .put(ref"C1", CellValue.Formula("A1>B1", Some(CellValue.Bool(true))))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("formula-cached-bool.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    readSheet(ref"C1").value match
      case CellValue.Formula(expr, Some(CellValue.Bool(cached))) =>
        assertEquals(expr, "A1>B1")
        assertEquals(cached, true)
      case CellValue.Formula(expr, cached) =>
        fail(s"Formula $expr has unexpected cached value: $cached")
      case other =>
        fail(s"Expected Formula with cached Bool, got $other")
  }

  test("Formula cell roundtrips expression with cached error value") {
    val initial = Workbook("Formulas")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Number(BigDecimal(100)))
      .put(ref"B1", CellValue.Number(BigDecimal(0)))
      .put(ref"C1", CellValue.Formula("A1/B1", Some(CellValue.Error(CellError.Div0))))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("formula-cached-error.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    readSheet(ref"C1").value match
      case CellValue.Formula(expr, Some(CellValue.Error(err))) =>
        assertEquals(expr, "A1/B1")
        assertEquals(err, CellError.Div0)
      case CellValue.Formula(expr, cached) =>
        fail(s"Formula $expr has unexpected cached value: $cached")
      case other =>
        fail(s"Expected Formula with cached Error, got $other")
  }

  test("Multiple formula cells roundtrip correctly") {
    val initial = Workbook("Formulas")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Number(BigDecimal(10)))
      .put(ref"A2", CellValue.Number(BigDecimal(20)))
      .put(ref"A3", CellValue.Number(BigDecimal(30)))
      .put(ref"B1", CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(20)))))
      .put(ref"B2", CellValue.Formula("A2*2", Some(CellValue.Number(BigDecimal(40)))))
      .put(ref"B3", CellValue.Formula("A3*2", Some(CellValue.Number(BigDecimal(60)))))
      .put(ref"C1", CellValue.Formula("SUM(B1:B3)", Some(CellValue.Number(BigDecimal(120)))))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("formula-multiple.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    // Verify all formulas roundtripped
    readSheet(ref"B1").value match
      case CellValue.Formula("A1*2", Some(CellValue.Number(n))) => assertEquals(n, BigDecimal(20))
      case other => fail(s"B1: Expected Formula, got $other")

    readSheet(ref"B2").value match
      case CellValue.Formula("A2*2", Some(CellValue.Number(n))) => assertEquals(n, BigDecimal(40))
      case other => fail(s"B2: Expected Formula, got $other")

    readSheet(ref"B3").value match
      case CellValue.Formula("A3*2", Some(CellValue.Number(n))) => assertEquals(n, BigDecimal(60))
      case other => fail(s"B3: Expected Formula, got $other")

    readSheet(ref"C1").value match
      case CellValue.Formula("SUM(B1:B3)", Some(CellValue.Number(n))) => assertEquals(n, BigDecimal(120))
      case other => fail(s"C1: Expected Formula, got $other")
  }

  test("Formula cached value roundtrip preserves exact value (explicit regression)") {
    // Explicit regression test per GitHub issue #42
    val formula = CellValue.Formula("=A1*2", Some(CellValue.Number(BigDecimal(42))))
    val initial = Workbook("Formulas")
    val sheet = initial.sheets(0).put(ref"B1", formula)
    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("update"))

    val outputPath = tempDir.resolve("formula-cached-explicit.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))
    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))

    readWb.sheets(0)(ref"B1").value match
      case CellValue.Formula(expr, Some(CellValue.Number(n))) =>
        assertEquals(expr, "=A1*2")
        assertEquals(n, BigDecimal(42))
      case other => fail(s"Unexpected: $other")
  }

  test("Workbook with merged cells preserves merges") {
    val initial = Workbook("Merged")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Text("Merged Header"))
      .merge(ref"A1:C1")

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("merged.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    // Verify cell value preserved
    assertEquals(readSheet(ref"A1").value, CellValue.Text("Merged Header"))

    // Verify merged ranges preserved
    assertEquals(readSheet.mergedRanges.size, 1, "Should have exactly one merged range")
    assert(
      readSheet.mergedRanges.contains(ref"A1:C1"),
      "Should preserve A1:C1 merge"
    )
  }

  test("Workbook with DateTime cells serializes to Excel serial numbers") {
    import java.time.LocalDateTime

    val dt1 = LocalDateTime.of(2025, 11, 10, 0, 0, 0) // Nov 10, 2025 midnight
    val dt2 = LocalDateTime.of(2000, 1, 1, 12, 0, 0)   // Jan 1, 2000 noon

    val initial = Workbook("Dates")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.DateTime(dt1))
      .put(ref"A2", CellValue.DateTime(dt2))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("dates.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    // Read and verify serial numbers were written (not "0")
    val readResult = XlsxReader.read(outputPath)
    val readWb = readResult match
      case Right(wb) => wb
      case Left(err) => fail(s"Read failed: ${err.message}")

    val cell1 = readWb.sheets(0)(ref"A1")
    val cell2 = readWb.sheets(0)(ref"A2")

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

  test("Full-feature workbook round-trips values, styles, merges, and bytes") {
    val original = orFail(buildComprehensiveWorkbook(), "Failed to build workbook")

    val bytes1 = orFail(XlsxWriter.writeToBytes(original), "Failed to write bytes")
    val roundTripped =
      orFail(XlsxReader.readFromBytes(bytes1), "Failed to read round-tripped workbook")
    val bytes2 = orFail(XlsxWriter.writeToBytes(roundTripped), "Failed to re-write bytes")

    assertWorkbookEquality(original, roundTripped)
    assertZipBytesEqual(bytes1, bytes2)
  }

  test("Chained writes to same file succeed (regression test for atomic write)") {
    // This test verifies that writing to the same file multiple times works correctly.
    // Previously, this would fail with "zip END header not found" because the output
    // file was opened (and truncated) before the source file was fully read.
    val outputPath = tempDir.resolve("chained-write.xlsx")

    // First write: create initial file
    val initial = Workbook("Test").sheets(0).put(ref"A1", CellValue.Text("First"))
    val wb1 = Workbook(Vector(initial))
    val write1 = XlsxWriter.write(wb1, outputPath)
    assert(write1.isRight, s"First write failed: $write1")

    // Read back the first file
    val read1 = XlsxReader.read(outputPath)
    assert(read1.isRight, s"First read failed: $read1")
    val wb1Read = read1.toOption.get

    // Second write: modify and write back to SAME file using put() for proper tracking
    val modified = wb1Read.sheets(0).put(ref"A2", CellValue.Text("Second"))
    val wb2 = wb1Read.put(modified) // Uses put() to properly track modification
    val write2 = XlsxWriter.write(wb2, outputPath)
    assert(write2.isRight, s"Second write (chained) failed: $write2")

    // Third write: another chained write to verify stability
    val read2 = XlsxReader.read(outputPath)
    assert(read2.isRight, s"Second read failed: $read2")
    val wb2Read = read2.toOption.get

    val modified2 = wb2Read.sheets(0).put(ref"A3", CellValue.Text("Third"))
    val wb3 = wb2Read.put(modified2) // Uses put() to properly track modification
    val write3 = XlsxWriter.write(wb3, outputPath)
    assert(write3.isRight, s"Third write (chained) failed: $write3")

    // Verify final contents
    val finalRead = XlsxReader.read(outputPath)
    assert(finalRead.isRight, s"Final read failed: $finalRead")
    finalRead.foreach { wb =>
      val sheet = wb.sheets(0)
      assertEquals(sheet.cells.get(ref"A1").map(_.value), Some(CellValue.Text("First")))
      assertEquals(sheet.cells.get(ref"A2").map(_.value), Some(CellValue.Text("Second")))
      assertEquals(sheet.cells.get(ref"A3").map(_.value), Some(CellValue.Text("Third")))
    }
  }

  test("Sheet elements have r namespace bound for r:id attribute (openpyxl compatibility)") {
    // This test verifies that r:id attributes on sheet elements have the "r" namespace
    // properly bound. Previously, files created by openpyxl (which use different namespace
    // declarations) would fail on re-read with "prefix 'r' not bound" error.
    val outputPath = tempDir.resolve("namespace-test.xlsx")

    // Create and write a workbook
    val wb = Workbook("Test").sheets(0).put(ref"A1", CellValue.Text("Hello"))
    val write1 = XlsxWriter.write(Workbook(Vector(wb)), outputPath)
    assert(write1.isRight, s"First write failed: $write1")

    // Read it back
    val read1 = XlsxReader.read(outputPath)
    assert(read1.isRight, s"First read failed: $read1")

    // Modify and write again (this exercises the workbook.xml regeneration)
    val modified = read1.toOption.get.sheets(0).put(ref"A2", CellValue.Text("World"))
    val wb2 = read1.toOption.get.put(modified)
    val write2 = XlsxWriter.write(wb2, outputPath)
    assert(write2.isRight, s"Second write failed: $write2")

    // Re-read the modified file - this would fail if r namespace isn't bound
    val read2 = XlsxReader.read(outputPath)
    assert(read2.isRight, s"Re-read after modification failed: $read2")

    // Verify contents
    read2.foreach { wb =>
      assertEquals(wb.sheets.size, 1)
      assertEquals(wb.sheets(0).cells.get(ref"A1").map(_.value), Some(CellValue.Text("Hello")))
      assertEquals(wb.sheets(0).cells.get(ref"A2").map(_.value), Some(CellValue.Text("World")))
    }
  }

  // ---------- Helpers ----------

  private def buildComprehensiveWorkbook(): XLResult[Workbook] =
    for
      summary <- buildSummarySheet
      detail <- buildDetailSheet
      timeline <- buildTimelineSheet
    yield Workbook(Vector(summary, detail, timeline))

  private def buildSummarySheet: XLResult[Sheet] =
    Right {
      val base = Sheet("Summary")
      val populated = base
        .put(
          ref"A1" -> "Quarterly Report",
          ref"A2" -> "Revenue",
          ref"B2" -> BigDecimal("125000.50"),
          ref"A3" -> "Expenses",
          ref"B3" -> BigDecimal("82000.10"),
          ref"A4" -> "Net Income",
          ref"B4" -> BigDecimal("43000.40"),
          ref"C2" -> true,
          ref"C3" -> false,
          ref"D2" -> LocalDate.of(2025, 11, 15),
          ref"D3" -> LocalDate.of(2025, 12, 31),
          ref"A6" -> "Styled note",
          ref"B6" -> "Padded text",
          ref"C6" -> BigDecimal("42.42"),
          ref"D6" -> LocalDateTime.of(2025, 11, 20, 9, 15),
          ref"E6" -> "dup-string",
          ref"E7" -> "dup-string"
        )
      val enriched = populated.put(ref"E2", CellValue.Error(CellError.Div0))
      val header = CellStyle.default.bold.size(16.0).white.bgBlue.center.middle
      val currency = CellStyle.default.bold.size(12.0).right
      val warning = CellStyle.default.red.bgYellow
      val rich = CellStyle.default.italic
      enriched
        .withCellStyle(ref"A1", header)
        .withRangeStyle(ref"A2:C4", currency)
        .withCellStyle(ref"E2", warning)
        .withCellStyle(ref"A6", rich)
        .merge(ref"A1:E1")
    }

  private def buildDetailSheet: XLResult[Sheet] =
    Right {
      val base = Sheet("Detail")
      val populated = base
        .put(
          ref"A1" -> "Region",
          ref"B1" -> "Units",
          ref"C1" -> "Price",
          ref"D1" -> "Total",
          ref"E1" -> "Updated",
          ref"A2" -> "North",
          ref"B2" -> 120,
          ref"C2" -> BigDecimal("12.50"),
          ref"D2" -> BigDecimal("1500.00"),
          ref"A3" -> "South",
          ref"B3" -> 85,
          ref"C3" -> BigDecimal("11.00"),
          ref"D3" -> BigDecimal("935.00"),
          ref"A4" -> "West",
          ref"B4" -> 60,
          ref"C4" -> BigDecimal("10.25"),
          ref"D4" -> BigDecimal("615.00"),
          ref"E2" -> LocalDateTime.of(2025, 1, 1, 9, 30),
          ref"E3" -> LocalDateTime.of(2025, 1, 2, 10, 0),
          ref"E4" -> LocalDateTime.of(2025, 1, 3, 11, 15),
          ref"A6" -> "Notes",
          ref"B6" -> "Use xml:space to keep double spaces"
        )
      val header = CellStyle.default.bold.bgGray.center
      val number = CellStyle.default.right
      val noteStyle = CellStyle.default.italic.wrap
      populated
        .withRangeStyle(ref"A1:E1", header)
        .withRangeStyle(ref"B2:D4", number)
        .withCellStyle(ref"B6", noteStyle)
    }

  private def buildTimelineSheet: XLResult[Sheet] =
    Right {
      val base = Sheet("Timeline")
      val populated = base
        .put(
          ref"A1" -> "Milestone",
          ref"B1" -> "Date",
          ref"C1" -> "Status",
          ref"D1" -> "Approved",
          ref"A2" -> "Kickoff",
          ref"B2" -> LocalDate.of(2025, 1, 5),
          ref"C2" -> "On Track",
          ref"D2" -> true,
          ref"A3" -> "Launch",
          ref"B3" -> LocalDate.of(2025, 3, 15),
          ref"C3" -> "At Risk",
          ref"D3" -> false
        )
      val header = CellStyle.default.bold.bgBlue.white.center
      val dateStyle = CellStyle.default.right
      val warning = CellStyle.default.bold.bgYellow
      populated
        .withRangeStyle(ref"A1:D1", header)
        .withRangeStyle(ref"B2:B3", dateStyle)
        .withCellStyle(ref"C3", warning)
    }

  private def assertWorkbookEquality(expected: Workbook, actual: Workbook): Unit =
    assertEquals(actual.sheets.size, expected.sheets.size, "Sheet count mismatch after round-trip")
    expected.sheets.zip(actual.sheets).foreach { case (expSheet, actSheet) =>
      assertEquals(
        actSheet.cells.keySet,
        expSheet.cells.keySet,
        s"Cell refs mismatch in sheet ${expSheet.name.value}"
      )
      expSheet.cells.foreach { case (ref, expCell) =>
        val actCell = actSheet.cells(ref)
        assertCellValueEquals(expSheet.name.value, ref, expCell.value, actCell.value)
        val expStyle = expCell.styleId.flatMap(expSheet.styleRegistry.get)
        val actStyle = actCell.styleId.flatMap(actSheet.styleRegistry.get)
        // Normalize styles: compare semantic meaning (numFmt) not raw ID (numFmtId)
        // Round-trip enriches with numFmtId, which is expected behavior
        val normalizedExp = expStyle.map(_.copy(numFmtId = None))
        val normalizedAct = actStyle.map(_.copy(numFmtId = None))
        assertEquals(
          normalizedAct,
          normalizedExp,
          s"Cell style mismatch at ${expSheet.name.value}!${ref.toA1}"
        )
      }
      assertEquals(
        actSheet.mergedRanges,
        expSheet.mergedRanges,
        s"Merged ranges mismatch in sheet ${expSheet.name.value}"
      )
    }

  private def orFail[A](result: XLResult[A], context: String): A =
    result match
      case Right(value) => value
      case Left(err) => fail(s"$context: ${err.message}")

  private def assertCellValueEquals(
    sheetName: String,
    ref: ARef,
    expected: CellValue,
    actual: CellValue
  ): Unit =
    (expected, actual) match
      case (CellValue.DateTime(exp), CellValue.Number(serial)) =>
        val asDate = CellValue.excelSerialToDateTime(serial.toDouble)
        assertDateApproxEquals(sheetName, ref, exp, asDate)
      case (CellValue.Number(serial), CellValue.DateTime(actualDate)) =>
        val expectedDate = CellValue.excelSerialToDateTime(serial.toDouble)
        assertDateApproxEquals(sheetName, ref, expectedDate, actualDate)
      case _ =>
        assertEquals(
          actual,
          expected,
          s"Cell value mismatch at $sheetName!${ref.toA1}"
        )

  private def assertDateApproxEquals(
    sheetName: String,
    ref: ARef,
    expected: LocalDateTime,
    actual: LocalDateTime
  ): Unit =
    val delta = Duration.between(expected, actual).abs()
    val tolerance = Duration.ofSeconds(1)
    assert(
      delta.compareTo(tolerance) <= 0,
      s"Date mismatch at $sheetName!${ref.toA1}: expected $expected, got $actual"
    )

  private def assertZipBytesEqual(expected: Array[Byte], actual: Array[Byte]): Unit =
    if !expected.sameElements(actual) then
      val expectedEntries = readZipEntries(expected)
      val actualEntries = readZipEntries(actual)
      val missing = expectedEntries.keySet.diff(actualEntries.keySet)
      val extra = actualEntries.keySet.diff(expectedEntries.keySet)
      val mismatched = expectedEntries.keySet
        .intersect(actualEntries.keySet)
        .filter(name => !expectedEntries(name).sameElements(actualEntries(name)))

      val message = new StringBuilder("Round-trip ZIP mismatch detected:\n")
      if missing.nonEmpty then message.append(s"Missing entries: ${missing.mkString(", ")}\n")
      if extra.nonEmpty then message.append(s"Extra entries: ${extra.mkString(", ")}\n")
      mismatched.headOption.foreach { sample =>
        val snippet = mismatched.toSeq.sorted.take(3).mkString(", ")
        message.append(s"Entries with differing content (first 3): $snippet\n")
        val expSnippet = new String(expectedEntries(sample), StandardCharsets.UTF_8).take(500)
        val actSnippet = new String(actualEntries(sample), StandardCharsets.UTF_8).take(500)
        message.append(s"Sample diff for $sample:\nEXPECTED: $expSnippet\nACTUAL:   $actSnippet\n")
      }
      fail(message.toString)

  private def readZipEntries(bytes: Array[Byte]): Map[String, Array[Byte]] =
    val zip = new ZipInputStream(new ByteArrayInputStream(bytes))
    try
      def loop(acc: Map[String, Array[Byte]]): Map[String, Array[Byte]] =
        Option(zip.getNextEntry) match
          case None => acc
          case Some(entry) =>
            val data = zip.readAllBytes()
            loop(acc + (entry.getName -> data))
      loop(Map.empty)
    finally zip.close()
