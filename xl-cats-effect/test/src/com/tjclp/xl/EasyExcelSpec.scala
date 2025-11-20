package com.tjclp.xl

import cats.effect.IO
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.extensions.given // For CellWriter given instances
import com.tjclp.xl.io.EasyExcel as Excel
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.unsafe.* // For .unsafe extension
import com.tjclp.xl.workbooks.Workbook
import java.io.FileNotFoundException
import java.nio.file.{Files, Paths}
import munit.CatsEffectSuite

/**
 * Tests for EasyExcel synchronous IO operations.
 *
 * Covers all methods in xl-cats-effect/src/com/tjclp/xl/io/EasyExcel.scala:
 *   - Excel.read() - Synchronous file reading
 *   - Excel.write() - Synchronous file writing
 *   - Excel.modify() - In-place modification
 *
 * These tests verify IO boundary behavior and establish round-trip invariants.
 */
class EasyExcelSpec extends CatsEffectSuite:

  // ========== Read Operations ==========

  test("read throws Exception for missing file") {
    // Note: EasyExcel wraps IO exceptions in generic Exception
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
      val original = Workbook.empty.put(Sheet("Sales").unsafe).unsafe
      Excel.write(original, tempFile.toString)

      // Read it back
      val loaded = Excel.read(tempFile.toString)
      assert(loaded.sheets.nonEmpty)
      assert(loaded.getSheet("Sales").isDefined)
    } finally {
      Files.deleteIfExists(tempFile)
    }
  }

  // ========== Write Operations ==========

  test("write creates new file") {
    val tempFile = Files.createTempFile("test-write", ".xlsx")
    Files.delete(tempFile) // Delete so we test creation

    try {
      val workbook = Workbook.empty.put(Sheet("Test").unsafe).unsafe
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
      val original = Workbook.empty.put(Sheet("First").unsafe).unsafe
      Excel.write(original, tempFile.toString)

      // Overwrite
      val updated = Workbook.empty.put(Sheet("Second").unsafe).unsafe
      Excel.write(updated, tempFile.toString)

      // Verify overwritten
      val loaded = Excel.read(tempFile.toString)
      assert(loaded.getSheet("Second").isDefined)
      assert(loaded.getSheet("First").isEmpty)
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
        .unsafe

      Excel.write(original, tempFile.toString)
      val loaded = Excel.read(tempFile.toString)

      val sales = loaded.getSheet("Sales").get
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
        .unsafe

      Excel.write(original, tempFile.toString)
      val loaded = Excel.read(tempFile.toString)

      val styled = loaded.getSheet("Styled").get
      val cell = styled.cell("A1").get
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
        .unsafe

      Excel.write(original, tempFile.toString)
      val loaded = Excel.read(tempFile.toString)

      assertEquals(loaded.getSheet("First").flatMap(_.cell("A1").map(_.value)),
                   Some(CellValue.Text("1")))
      assertEquals(loaded.getSheet("Second").flatMap(_.cell("A1").map(_.value)),
                   Some(CellValue.Text("2")))
      assertEquals(loaded.getSheet("Third").flatMap(_.cell("A1").map(_.value)),
                   Some(CellValue.Text("3")))
    } finally {
      Files.deleteIfExists(tempFile)
    }
  }

  // ========== Modify Operations ==========

  test("modify reads, transforms, and writes") {
    val tempFile = Files.createTempFile("test-modify", ".xlsx")
    try {
      // Setup - note: Workbook.empty creates Sheet1 by default
      val initial = Workbook.empty.unsafe
      Excel.write(initial, tempFile.toString)

      // Modify - update the default Sheet1
      Excel.modify(tempFile.toString) { wb =>
        wb.update("Sheet1", _.put("A1", "Modified").unsafe).unsafe
      }

      // Verify
      val loaded = Excel.read(tempFile.toString)
      val sheet1 = loaded.getSheet("Sheet1").get
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
        .unsafe

      Excel.write(initial, tempFile.toString)

      // Modify only Sheet1 (not Preserve)
      Excel.modify(tempFile.toString) { wb =>
        wb.update("Sheet1", _.put("A1", "Changed").unsafe).unsafe
      }

      // Verify "Preserve" sheet unchanged
      val loaded = Excel.read(tempFile.toString)
      val preserve = loaded.getSheet("Preserve").get
      assertEquals(preserve.cell("A1").map(_.value), Some(CellValue.Text("Original")))
    } finally {
      Files.deleteIfExists(tempFile)
    }
  }
