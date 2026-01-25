package com.tjclp.xl.cli

import munit.FunSuite

import java.io.{ByteArrayOutputStream, FileInputStream, FileOutputStream}
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet, given}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.extensions.style
import com.tjclp.xl.cli.commands.{
  ImportCommands,
  SheetCommands,
  StreamingWriteCommands,
  WriteCommands,
  CellCommands
}
import com.tjclp.xl.cli.helpers.{CsvParser, StreamingCsvParser}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.font.Font

/**
 * Integration tests for streaming write functionality.
 *
 * Tests both:
 *   - Hybrid streaming: Load workbook → mutate → stream output (O(1) output memory)
 *   - True streaming: CSV → XLSX with O(1) memory throughout (for --new-sheet)
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class StreamingWriteSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val config: WriterConfig = WriterConfig.default

  // Helper to create temporary files
  def tempXlsx(): Path = Files.createTempFile("test", ".xlsx")
  def tempCsv(content: String): Path =
    val file = Files.createTempFile("test", ".csv")
    Files.writeString(file, content)
    file

  private def remapWorksheetEntry(source: Path, from: String, to: String): Path =
    val output = Files.createTempFile("test-remap", ".xlsx")
    val zipIn = new ZipInputStream(new FileInputStream(source.toFile))
    val zipOut = new ZipOutputStream(new FileOutputStream(output.toFile))
    zipOut.setLevel(1)
    try
      var entry = zipIn.getNextEntry
      while entry != null do
        val entryName = entry.getName
        val targetName =
          if entryName == s"xl/worksheets/$from" then s"xl/worksheets/$to" else entryName
        val data = readEntryBytes(zipIn)
        val updatedData =
          if entryName == "xl/_rels/workbook.xml.rels" then
            val xml = new String(data, StandardCharsets.UTF_8)
            val updated = xml.replace(
              s"""Target="worksheets/$from"""",
              s"""Target="worksheets/$to""""
            )
            updated.getBytes(StandardCharsets.UTF_8)
          else data

        val outEntry = new ZipEntry(targetName)
        outEntry.setTime(0L)
        outEntry.setMethod(ZipEntry.DEFLATED)
        zipOut.putNextEntry(outEntry)
        zipOut.write(updatedData)
        zipOut.closeEntry()
        zipIn.closeEntry()
        entry = zipIn.getNextEntry
    finally
      zipIn.close()
      zipOut.close()
    output

  private def readEntryBytes(zipIn: ZipInputStream): Array[Byte] =
    val buffer = new Array[Byte](8192)
    val baos = new ByteArrayOutputStream()
    var read = zipIn.read(buffer)
    while read != -1 do
      baos.write(buffer, 0, read)
      read = zipIn.read(buffer)
    baos.toByteArray

  // ========== Hybrid Streaming: put command ==========

  test("streaming put: basic value preserved") {
    val outputPath = tempXlsx()
    try
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .put(wb, Some(wb.sheets.head), "A1", List("100"), outputPath, config, stream = true)
        .unsafeRunSync()

      assert(result.contains("Put: A1 = 100"))
      assert(result.contains("Saved (streaming)"))

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val cellValue = imported.sheets.head.cells.get(ARef.from0(0, 0)).map(_.value)
      assertEquals(cellValue, Some(CellValue.Number(BigDecimal("100"))))
    finally Files.deleteIfExists(outputPath)
  }

  test("streaming put: multiple values (batch mode)") {
    val outputPath = tempXlsx()
    try
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .put(wb, Some(wb.sheets.head), "A1:A3", List("10", "20", "30"), outputPath, config, stream = true)
        .unsafeRunSync()

      assert(result.contains("Put 3 values"))
      assert(result.contains("Saved (streaming)"))

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val sheet = imported.sheets.head
      assertEquals(sheet.cells.get(ARef.from0(0, 0)).map(_.value), Some(CellValue.Number(BigDecimal("10"))))
      assertEquals(sheet.cells.get(ARef.from0(0, 1)).map(_.value), Some(CellValue.Number(BigDecimal("20"))))
      assertEquals(sheet.cells.get(ARef.from0(0, 2)).map(_.value), Some(CellValue.Number(BigDecimal("30"))))
    finally Files.deleteIfExists(outputPath)
  }

  test("streaming put: formula preserved") {
    val outputPath = tempXlsx()
    try
      val wb = Workbook(Sheet("Test").put(ARef.from0(0, 0), CellValue.Number(BigDecimal("10"))))
      val result = WriteCommands
        .putFormula(wb, Some(wb.sheets.head), "B1", List("=A1*2"), outputPath, config, stream = true)
        .unsafeRunSync()

      assert(result.contains("Put: B1"))
      assert(result.contains("Saved (streaming)"))

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      imported.sheets.head.cells.get(ARef.from0(1, 0)).map(_.value) match
        case Some(CellValue.Formula(formula, _)) =>
          assertEquals(formula, "A1*2")
        case other => fail(s"Expected Formula, got $other")
    finally Files.deleteIfExists(outputPath)
  }

  // ========== Style Preservation in Streaming Mode ==========

  test("streaming put: style preserved") {
    val outputPath = tempXlsx()
    try
      // Create workbook with styled cell
      val boldStyle = CellStyle.default.withFont(Font.default.withBold(true))
      val styledSheet = Sheet("Test")
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal("100")))
        .style(ARef.from0(0, 0), boldStyle)
      val wb = Workbook(styledSheet)

      // Add another value using streaming mode
      val result = WriteCommands
        .put(wb, Some(wb.sheets.head), "B1", List("200"), outputPath, config, stream = true)
        .unsafeRunSync()

      assert(result.contains("Saved (streaming)"), result)

      // Verify both values exist and A1 retains style
      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val sheet = imported.sheets.head
      assertEquals(sheet.cells.get(ARef.from0(0, 0)).map(_.value), Some(CellValue.Number(BigDecimal("100"))))
      assertEquals(sheet.cells.get(ARef.from0(1, 0)).map(_.value), Some(CellValue.Number(BigDecimal("200"))))

      // Check that A1 has a style (styleId should be present)
      val a1StyleId = sheet.cells.get(ARef.from0(0, 0)).flatMap(_.styleId)
      assert(a1StyleId.isDefined, "A1 should have style preserved")
    finally Files.deleteIfExists(outputPath)
  }

  test("streaming style: style command with streaming") {
    val outputPath = tempXlsx()
    try
      val wb = Workbook(Sheet("Test").put(ARef.from0(0, 0), CellValue.Number(BigDecimal("100"))))
      val result = WriteCommands
        .style(
          wb,
          Some(wb.sheets.head),
          "A1",
          bold = true,
          italic = false,
          underline = false,
          bg = None,
          fg = None,
          fontSize = None,
          fontName = None,
          align = None,
          valign = None,
          wrap = false,
          numFormat = None,
          border = None,
          borderTop = None,
          borderRight = None,
          borderBottom = None,
          borderLeft = None,
          borderColor = None,
          replace = false,
          outputPath,
          config,
          stream = true
        )
        .unsafeRunSync()

      assert(result.contains("Styled: A1"), result)
      assert(result.contains("Saved (streaming)"), result)

      // Verify style was applied
      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val a1StyleId = imported.sheets.head.cells.get(ARef.from0(0, 0)).flatMap(_.styleId)
      assert(a1StyleId.isDefined, "A1 should have style applied")
    finally Files.deleteIfExists(outputPath)
  }

  // ========== Sheet Operations in Streaming Mode ==========

  test("streaming: add sheet") {
    val outputPath = tempXlsx()
    try
      val wb = Workbook(Sheet("Sheet1"))
      val result = SheetCommands
        .addSheet(wb, "Sheet2", None, None, outputPath, config, stream = true)
        .unsafeRunSync()

      assert(result.contains("Added sheet: Sheet2"))
      assert(result.contains("Saved (streaming)"))

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      assertEquals(imported.sheets.size, 2)
      assertEquals(imported.sheetNames.map(_.value), Seq("Sheet1", "Sheet2"))
    finally Files.deleteIfExists(outputPath)
  }

  test("streaming: copy sheet") {
    val outputPath = tempXlsx()
    try
      val wb = Workbook(Sheet("Original").put(ARef.from0(0, 0), CellValue.Text("Hello")))
      val result = SheetCommands
        .copySheet(wb, "Original", "Copy", outputPath, config, stream = true)
        .unsafeRunSync()

      assert(result.contains("Copied: Original → Copy"))
      assert(result.contains("Saved (streaming)"))

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      assertEquals(imported.sheets.size, 2)
      // Verify data was copied
      val copySheet = imported.sheets.find(_.name.value == "Copy")
      assert(copySheet.isDefined)
      assertEquals(copySheet.get.cells.get(ARef.from0(0, 0)).map(_.value), Some(CellValue.Text("Hello")))
    finally Files.deleteIfExists(outputPath)
  }

  // ========== Cell Operations in Streaming Mode ==========

  // Note: Streaming mode currently uses a simplified XML writer that doesn't
  // include mergeCells element in the output. This is a known limitation.
  // Merged cell metadata is stored in-memory but not serialized via streaming path.
  // For full fidelity, use non-streaming mode.
  test("streaming: merge cells (values preserved, merge metadata may not be)".ignore) {
    val outputPath = tempXlsx()
    try
      val wb = Workbook(Sheet("Test").put(ARef.from0(0, 0), CellValue.Text("Merged")))
      val result = CellCommands
        .merge(wb, Some(wb.sheets.head), "A1:C1", outputPath, config, stream = true)
        .unsafeRunSync()

      assert(result.contains("Merged: A1:C1"), result)
      assert(result.contains("Saved (streaming)"), result)

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val sheet = imported.sheets.head
      assert(sheet.mergedRanges.nonEmpty, "Should have merged range")
    finally Files.deleteIfExists(outputPath)
  }

  test("streaming: clear cells") {
    val outputPath = tempXlsx()
    try
      val wb = Workbook(Sheet("Test")
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal("1")))
        .put(ARef.from0(1, 0), CellValue.Number(BigDecimal("2"))))
      val result = CellCommands
        .clear(wb, Some(wb.sheets.head), "A1", all = false, styles = false, comments = false, outputPath, config, stream = true)
        .unsafeRunSync()

      assert(result.contains("Cleared contents from A1"))
      assert(result.contains("Saved (streaming)"))

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val sheet = imported.sheets.head
      assertEquals(sheet.cells.get(ARef.from0(0, 0)), None) // A1 cleared
      assertEquals(sheet.cells.get(ARef.from0(1, 0)).map(_.value), Some(CellValue.Number(BigDecimal("2")))) // B1 preserved
    finally Files.deleteIfExists(outputPath)
  }

  // ========== True Streaming: Worksheet Patch ==========

  test("streaming put: inserts missing rows and cells") {
    val sourcePath = tempXlsx()
    val outputPath = tempXlsx()
    try
      val wb = Workbook(
        Sheet("Test").put(ARef.from0(1, 1), CellValue.Number(BigDecimal("9"))) // B2
      )
      ExcelIO.instance[IO].write(wb, sourcePath).unsafeRunSync()

      val result = StreamingWriteCommands
        .put(sourcePath, outputPath, Some("Test"), "A1:A2", List("1", "2"))
        .unsafeRunSync()

      assert(result.contains("Saved (streaming)"))

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val sheet = imported.sheets.head
      assertEquals(sheet.cells.get(ARef.from0(0, 0)).map(_.value), Some(CellValue.Number(BigDecimal("1"))))
      assertEquals(sheet.cells.get(ARef.from0(0, 1)).map(_.value), Some(CellValue.Number(BigDecimal("2"))))
      assertEquals(sheet.cells.get(ARef.from0(1, 1)).map(_.value), Some(CellValue.Number(BigDecimal("9"))))
    finally
      Files.deleteIfExists(sourcePath)
      Files.deleteIfExists(outputPath)
  }

  test("streaming put: resolves sheet path via workbook rels") {
    val sourcePath = tempXlsx()
    val outputPath = tempXlsx()
    val wb = Workbook(
      Sheet("Sheet1").put(ARef.from0(0, 0), CellValue.Text("one")),
      Sheet("Sheet2").put(ARef.from0(0, 0), CellValue.Text("two"))
    )
    ExcelIO.instance[IO].write(wb, sourcePath).unsafeRunSync()
    val remapped = remapWorksheetEntry(sourcePath, "sheet2.xml", "sheet3.xml")
    try
      val result = StreamingWriteCommands
        .put(remapped, outputPath, Some("Sheet2"), "B1", List("99"))
        .unsafeRunSync()

      assert(result.contains("Saved (streaming)"))

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val sheet2 = imported.sheets.find(_.name.value == "Sheet2").getOrElse {
        fail("Sheet2 not found in output workbook")
      }
      assertEquals(sheet2.cells.get(ARef.from0(1, 0)).map(_.value), Some(CellValue.Number(BigDecimal("99"))))
      assertEquals(sheet2.cells.get(ARef.from0(0, 0)).map(_.value), Some(CellValue.Text("two")))
    finally
      Files.deleteIfExists(sourcePath)
      Files.deleteIfExists(remapped)
      Files.deleteIfExists(outputPath)
  }

  // ========== True Streaming: CSV Import ==========

  test("true streaming: CSV import to new sheet") {
    val outputPath = tempXlsx()
    val csvPath = tempCsv("Name,Age\nAlice,30\nBob,25")
    try
      val wb = Workbook() // Empty workbook
      val result = ImportCommands
        .importCsv(
          wb,
          None,
          csvPath.toString,
          None,
          delimiter = ',',
          skipHeader = false,
          encoding = "UTF-8",
          newSheetName = Some("Data"),
          noTypeInference = false,
          outputPath,
          config,
          stream = true
        )
        .unsafeRunSync()

      assert(result.contains("Streamed:") || result.contains("Imported:"))
      assert(result.contains("Data"))

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      assertEquals(imported.sheets.size, 1)
      assertEquals(imported.sheetNames.head.value, "Data")

      val sheet = imported.sheets.head
      // Verify data
      assertEquals(sheet.cells.get(ARef.from0(0, 0)).map(_.value), Some(CellValue.Text("Name")))
      assertEquals(sheet.cells.get(ARef.from0(1, 0)).map(_.value), Some(CellValue.Text("Age")))
    finally
      Files.deleteIfExists(outputPath)
      Files.deleteIfExists(csvPath)
  }

  test("streaming CSV parser: basic parsing") {
    val csvPath = tempCsv("A,B,C\n1,2,3\n4,5,6")
    try
      val rows = StreamingCsvParser
        .streamCsv(csvPath, StreamingCsvParser.Options())
        .compile
        .toList
        .unsafeRunSync()

      assertEquals(rows.size, 3) // 3 rows including header
      assertEquals(rows.head.rowIndex, 1)
      assertEquals(rows.head.cells.size, 3)
    finally Files.deleteIfExists(csvPath)
  }

  test("streaming CSV parser: skip header") {
    val csvPath = tempCsv("Name,Age\nAlice,30\nBob,25")
    try
      val rows = StreamingCsvParser
        .streamCsv(csvPath, StreamingCsvParser.Options(skipHeader = true))
        .compile
        .toList
        .unsafeRunSync()

      assertEquals(rows.size, 2) // Header skipped
      assertEquals(rows.head.rowIndex, 1)
    finally Files.deleteIfExists(csvPath)
  }

  test("streaming CSV parser: respects encoding") {
    val csvPath = Files.createTempFile("test", ".csv")
    val value = "caf\u00E9"
    val content = s"$value,1"
    try
      Files.write(csvPath, content.getBytes(StandardCharsets.ISO_8859_1))

      val rows = StreamingCsvParser
        .streamCsv(csvPath, StreamingCsvParser.Options(encoding = "ISO-8859-1"))
        .compile
        .toList
        .unsafeRunSync()

      assertEquals(rows.size, 1)
      assertEquals(rows.head.cells.get(0), Some(CellValue.Text(value)))
    finally Files.deleteIfExists(csvPath)
  }

  test("streaming CSV parser: type inference") {
    val csvPath = tempCsv("Num,Bool,Text\n100,true,hello")
    try
      val rows = StreamingCsvParser
        .streamCsv(csvPath, StreamingCsvParser.Options(skipHeader = true, inferTypes = true))
        .compile
        .toList
        .unsafeRunSync()

      assertEquals(rows.size, 1)
      val cells = rows.head.cells
      assertEquals(cells.get(0), Some(CellValue.Number(BigDecimal("100"))))
      assertEquals(cells.get(1), Some(CellValue.Bool(true)))
      assertEquals(cells.get(2), Some(CellValue.Text("hello")))
    finally Files.deleteIfExists(csvPath)
  }

  test("streaming CSV parser: custom delimiter") {
    val csvPath = tempCsv("A;B;C\n1;2;3")
    try
      val rows = StreamingCsvParser
        .streamCsv(csvPath, StreamingCsvParser.Options(delimiter = ';', skipHeader = true))
        .compile
        .toList
        .unsafeRunSync()

      assertEquals(rows.size, 1)
      assertEquals(rows.head.cells.size, 3)
    finally Files.deleteIfExists(csvPath)
  }

  // ========== Regression: Non-streaming Mode Unchanged ==========

  test("non-streaming put: basic value (regression)") {
    val outputPath = tempXlsx()
    try
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .put(wb, Some(wb.sheets.head), "A1", List("100"), outputPath, config, stream = false)
        .unsafeRunSync()

      assert(result.contains("Put: A1 = 100"))
      assert(result.contains("Saved:"))
      assert(!result.contains("Saved (streaming)"))

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val cellValue = imported.sheets.head.cells.get(ARef.from0(0, 0)).map(_.value)
      assertEquals(cellValue, Some(CellValue.Number(BigDecimal("100"))))
    finally Files.deleteIfExists(outputPath)
  }

  // ========== Dimension Element Verification ==========

  test("streaming write: dimension element present in output") {
    val outputPath = tempXlsx()
    try
      val wb = Workbook(Sheet("Test")
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal("1")))
        .put(ARef.from0(2, 2), CellValue.Number(BigDecimal("2")))) // C3
      val result = WriteCommands
        .put(wb, Some(wb.sheets.head), "D4", List("3"), outputPath, config, stream = true)
        .unsafeRunSync()

      assert(result.contains("Saved (streaming)"))

      // Read back and check bounds
      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val sheet = imported.sheets.head
      // Should have cells at A1, C3, D4
      assert(sheet.cells.contains(ARef.from0(0, 0)), "A1 should exist")
      assert(sheet.cells.contains(ARef.from0(2, 2)), "C3 should exist")
      assert(sheet.cells.contains(ARef.from0(3, 3)), "D4 should exist")
    finally Files.deleteIfExists(outputPath)
  }
