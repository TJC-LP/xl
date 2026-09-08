package com.tjclp.xl

import cats.effect.IO
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given // For CellWriter given instances
import com.tjclp.xl.error.{XLError, XLException}
import com.tjclp.xl.io.Excel
import com.tjclp.xl.ooxml.metadata.LightMetadata
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.unsafe.* // For .unsafe extension
import com.tjclp.xl.workbooks.Workbook
import java.io.FileNotFoundException
import java.nio.file.{Files, Paths}
import munit.CatsEffectSuite

/**
 * Tests for Excel Easy Mode synchronous IO operations.
 *
 * Covers Easy Mode methods in xl-cats-effect/src/com/tjclp/xl/io/Excel.scala:
 *   - Excel.read() - Synchronous file reading
 *   - Excel.write() - Synchronous file writing
 *   - Excel.modify() - In-place modification
 *
 * These tests verify IO boundary behavior and establish round-trip invariants.
 */
class EasyExcelSpec extends CatsEffectSuite:

  // ========== Read Operations ==========

  test("read throws Exception for missing file") {
    // Note: Excel Easy Mode wraps IO exceptions in generic Exception
    intercept[Exception] {
      Excel.read("/nonexistent/path/missing.xlsx")
    }
  }

  test("read throws for invalid path") {
    intercept[Exception] { // Could be FileNotFoundException or other IO error
      Excel.read("/dev/null/impossible.xlsx")
    }
  }

  test("read valid workbook succeeds") {
    val tempFile = Files.createTempFile("test-read", ".xlsx")
    try {
      // Write a valid workbook
      val original = Workbook.empty.put(Sheet("Sales"))
      Excel.write(original, tempFile.toString)

      // Read it back
      val loaded = Excel.read(tempFile.toString)
      assert(loaded.sheets.nonEmpty)
      assert(loaded.get("Sales").isDefined)
    } finally {
      Files.deleteIfExists(tempFile)
    }
  }

  // ========== Write Operations ==========

  test("write creates new file") {
    val tempFile = Files.createTempFile("test-write", ".xlsx")
    Files.delete(tempFile) // Delete so we test creation

    try {
      val workbook = Workbook.empty.put(Sheet("Test"))
      Excel.write(workbook, tempFile.toString)

      assert(Files.exists(tempFile))
      assert(Files.size(tempFile) > 0)
    } finally {
      Files.deleteIfExists(tempFile)
    }
  }

  test("write overwrites existing file") {
    val tempFile = Files.createTempFile("test-overwrite", ".xlsx")
    try {
      // Write initial
      val original = Workbook.empty.put(Sheet("First"))
      Excel.write(original, tempFile.toString)

      // Overwrite
      val updated = Workbook.empty.put(Sheet("Second"))
      Excel.write(updated, tempFile.toString)

      // Verify overwritten
      val loaded = Excel.read(tempFile.toString)
      assert(loaded.get("Second").isDefined)
      assert(loaded.get("First").isEmpty)
    } finally {
      Files.deleteIfExists(tempFile)
    }
  }

  // ========== Round-Trip Tests ==========

  test("write then read preserves data") {
    val tempFile = Files.createTempFile("test-roundtrip", ".xlsx")
    try {
      val original = Workbook.empty
        .put(Sheet("Sales").put("A1", "Revenue").put("B1", 1000).unsafe)

      Excel.write(original, tempFile.toString)
      val loaded = Excel.read(tempFile.toString)

      val sales = loaded.get("Sales").getOrElse(fail("Expected Sales sheet"))
      assertEquals(sales.cell("A1").map(_.value), Some(CellValue.Text("Revenue")))
      assertEquals(sales.cell("B1").map(_.value), Some(CellValue.Number(BigDecimal(1000))))
    } finally {
      Files.deleteIfExists(tempFile)
    }
  }

  test("write then read preserves styles") {
    val tempFile = Files.createTempFile("test-styles", ".xlsx")
    try {
      val boldStyle = CellStyle.default.bold
      val original = Workbook.empty
        .put(Sheet("Styled").put("A1", "Bold", boldStyle).unsafe)

      Excel.write(original, tempFile.toString)
      val loaded = Excel.read(tempFile.toString)

      val styled = loaded.get("Styled").getOrElse(fail("Expected Styled sheet"))
      val cell = styled.cell("A1").getOrElse(fail("Expected cell at A1"))
      // Verify style preserved (has styleId)
      assert(cell.styleId.isDefined)
    } finally {
      Files.deleteIfExists(tempFile)
    }
  }

  test("write then read preserves multiple sheets") {
    val tempFile = Files.createTempFile("test-multi", ".xlsx")
    try {
      val original = Workbook.empty
        .put(Sheet("First").put("A1", "1").unsafe)
        .put(Sheet("Second").put("A1", "2").unsafe)
        .put(Sheet("Third").put("A1", "3").unsafe)

      Excel.write(original, tempFile.toString)
      val loaded = Excel.read(tempFile.toString)

      assertEquals(
        loaded.get("First").flatMap(_.cell("A1").map(_.value)),
        Some(CellValue.Text("1"))
      )
      assertEquals(
        loaded.get("Second").flatMap(_.cell("A1").map(_.value)),
        Some(CellValue.Text("2"))
      )
      assertEquals(
        loaded.get("Third").flatMap(_.cell("A1").map(_.value)),
        Some(CellValue.Text("3"))
      )
    } finally {
      Files.deleteIfExists(tempFile)
    }
  }

  // ========== Modify Operations ==========

  test("modify reads, transforms, and writes") {
    val tempFile = Files.createTempFile("test-modify", ".xlsx")
    try {
      // Setup - note: Workbook.empty creates Sheet1 by default
      val initial = Workbook.empty
      Excel.write(initial, tempFile.toString)

      // Modify - update the default Sheet1
      Excel.modify(tempFile.toString) { wb =>
        wb.update("Sheet1", _.put("A1", "Modified").unsafe).unsafe
      }

      // Verify
      val loaded = Excel.read(tempFile.toString)
      val sheet1 = loaded.get("Sheet1").getOrElse(fail("Expected Sheet1"))
      assertEquals(sheet1.cell("A1").map(_.value), Some(CellValue.Text("Modified")))
    } finally {
      Files.deleteIfExists(tempFile)
    }
  }

  test("modify preserves unmodified sheets") {
    val tempFile = Files.createTempFile("test-preserve", ".xlsx")
    try {
      // Setup with 2 sheets
      val initial = Workbook.empty
        .put(Sheet("Preserve").put("A1", "Original").unsafe)

      Excel.write(initial, tempFile.toString)

      // Modify only Sheet1 (not Preserve)
      Excel.modify(tempFile.toString) { wb =>
        wb.update("Sheet1", _.put("A1", "Changed").unsafe).unsafe
      }

      // Verify "Preserve" sheet unchanged
      val loaded = Excel.read(tempFile.toString)
      val preserve = loaded.get("Preserve").getOrElse(fail("Expected Preserve sheet"))
      assertEquals(preserve.cell("A1").map(_.value), Some(CellValue.Text("Original")))
    } finally {
      Files.deleteIfExists(tempFile)
    }
  }

  // ========== GH-589: sync-facade completions (readSheet / readMetadata / modifyR) ==========

  private def withTempXlsx[A](prefix: String)(body: java.nio.file.Path => A): A =
    val tempFile = Files.createTempFile(prefix, ".xlsx")
    try body(tempFile)
    finally Files.deleteIfExists(tempFile)

  private val twoSheets: Workbook = Workbook.empty
    .put(Sheet("Data").put("A1", 1).unsafe)
    .put(Sheet("Summary").put("A1", "total").unsafe)
    .remove("Sheet1")
    .unsafe

  test("readSheet returns the named sheet") {
    withTempXlsx("test-readsheet") { tempFile =>
      Excel.write(twoSheets, tempFile.toString)
      val summary = Excel.readSheet(tempFile.toString, "Summary").unsafe
      assertEquals(summary.name.value, "Summary")
      assertEquals(summary.cell("A1").map(_.value), Some(CellValue.Text("total")))
    }
  }

  test("GH-615: readSheet on a missing sheet is Left(SheetNotFound) naming the candidates") {
    withTempXlsx("test-readsheet-missing") { tempFile =>
      Excel.write(twoSheets, tempFile.toString)
      val result = Excel.readSheet(tempFile.toString, "Sumary")
      assertEquals(result, Left(XLError.SheetNotFound("Sumary", Vector("Data", "Summary"))))
      val err = result.swap.getOrElse(fail("expected Left"))
      assertEquals(err.code, "SHEET_NOT_FOUND")
      // the candidates are the error's, not prose in the message
      assertEquals(err.candidates, Vector("Summary"))
      assertEquals(err.message, "Sheet not found: 'Sumary'. Available: Data, Summary")
      assert(err.renderDiagnostic.contains("  did you mean: Summary"), err.renderDiagnostic)
    }
  }

  test("readSheet with no near miss still lists the available sheets") {
    withTempXlsx("test-readsheet-far") { tempFile =>
      Excel.write(twoSheets, tempFile.toString)
      val err = Excel.readSheet(tempFile.toString, "Zebra").swap.getOrElse(fail("expected Left"))
      assertEquals(err, XLError.SheetNotFound("Zebra", Vector("Data", "Summary")))
      assertEquals(err.candidates, Vector.empty)
      assertEquals(err.message, "Sheet not found: 'Zebra'. Available: Data, Summary")
    }
  }

  test("GH-615/GH-621: readSheet on a missing file is Left(IOError) naming the cause, once") {
    val missing = Files.createTempDirectory("test-readsheet-nofile").resolve("nope.xlsx")
    val err = Excel.readSheet(missing.toString, "Data").swap.getOrElse(fail("expected Left"))
    assertEquals(err, XLError.IOError(s"no such file: $missing"))
    assertEquals(err.message, s"IO error: no such file: $missing")
  }

  test("GH-621: Excel.read on a missing file throws the XLException its documentation promises") {
    val missing = Files.createTempDirectory("test-read-nofile").resolve("nope.xlsx")
    val ex = intercept[XLException](Excel.read(missing.toString))
    assertEquals(ex.error, XLError.IOError(s"no such file: $missing"))
    // one prefix: the domain message, not "Failed to read XLSX: IO error: Failed to read XLSX: …"
    assertEquals(ex.getMessage, s"IO error: no such file: $missing")
  }

  test("readMetadata lists sheets without loading cells") {
    withTempXlsx("test-readmetadata") { tempFile =>
      Excel.write(twoSheets, tempFile.toString)
      val meta: LightMetadata = Excel.readMetadata(tempFile.toString)
      assertEquals(meta.sheets.map(_.name.value), Vector("Data", "Summary"))
      assertEquals(meta.definedNames, Vector.empty)
    }
  }

  test("modifyR writes the Right workbook in place") {
    withTempXlsx("test-modifyr") { tempFile =>
      Excel.write(twoSheets, tempFile.toString)
      Excel.modifyR(tempFile.toString)(_.update("Data", _.put("B1", "changed").unsafe))
      val loaded = Excel.read(tempFile.toString)
      assertEquals(
        loaded.get("Data").flatMap(_.cell("B1")).map(_.value),
        Some(CellValue.Text("changed"))
      )
    }
  }

  test("modifyR on Left throws XLException and leaves the file byte-identical") {
    // Own directory: the scratch-file scan below must see only this test's files.
    val dir = Files.createTempDirectory("test-modifyr-left")
    val tempFile = dir.resolve("target.xlsx")
    try
      Excel.write(twoSheets, tempFile.toString)
      val before = Files.readAllBytes(tempFile).toVector
      val ex = intercept[XLException] {
        Excel.modifyR(tempFile.toString)(_.update("NoSuchSheet", identity))
      }
      assertEquals(ex.error, XLError.SheetNotFound("NoSuchSheet", Vector("Data", "Summary")))
      assertEquals(Files.readAllBytes(tempFile).toVector, before)
      // No scratch file was left behind next to the target
      val siblings = Files.list(dir)
      try
        assertEquals(siblings.map(_.getFileName.toString).toList, java.util.List.of("target.xlsx"))
      finally siblings.close()
    finally
      Files.deleteIfExists(tempFile)
      Files.deleteIfExists(dir)
  }
