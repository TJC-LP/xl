package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.{ARef, Column}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Integration tests for auto-fit column width functionality.
 */
class AutoFitCommandSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val config: WriterConfig = WriterConfig.default

  // Helper to create unique temp files per test and clean up after
  private def withTempFile[A](test: Path => A): A =
    val path = Files.createTempFile("test-autofit", ".xlsx")
    try test(path)
    finally Files.deleteIfExists(path)

  // Helper to create ARef from A1 notation indices (col first!)
  private def ref(col: Int, row: Int): ARef = ARef.from0(col, row)

  // ========== Auto-Fit Width Tests ==========

  test("col --auto-fit: calculates width from text content") {
    withTempFile { outputPath =>
      // A1="Short", A2="This is a much longer text", A3="Medium"
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Text("Short"))
        .put(ref(0, 1), CellValue.Text("This is a much longer text"))
        .put(ref(0, 2), CellValue.Text("Medium"))
      val wb = Workbook(sheet)

      val result = WriteCommands
        .col(wb, Some(sheet), "A", None, hide = false, show = false, autoFit = true, outputPath, config)
        .unsafeRunSync()

      assert(result.contains("(auto)"), s"Expected '(auto)' in output: $result")
      // Width = 26 chars * 0.85 + 1.5 = 23.6
      assert(result.contains("23.60"), s"Expected '23.60' in output: $result")

      // Verify the file was saved with the column width
      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.headOption.getOrElse(fail("No sheets found"))
      val col = Column.fromLetter("A").toOption.getOrElse(fail("Invalid column"))
      val colProps = s.getColumnProperties(col)
      assertEquals(colProps.width, Some(23.6))
    }
  }

  test("col --auto-fit: calculates width from numeric content") {
    withTempFile { outputPath =>
      // B1=12345, B2=1234567890, B3=42
      val sheet = Sheet("Test")
        .put(ref(1, 0), CellValue.Number(BigDecimal(12345)))
        .put(ref(1, 1), CellValue.Number(BigDecimal(1234567890)))
        .put(ref(1, 2), CellValue.Number(BigDecimal(42)))
      val wb = Workbook(sheet)

      val result = WriteCommands
        .col(wb, Some(sheet), "B", None, hide = false, show = false, autoFit = true, outputPath, config)
        .unsafeRunSync()

      assert(result.contains("(auto)"), s"Expected '(auto)' in output: $result")
      // Width = 10 digits * 0.85 + 1.5 = 10.0
      assert(result.contains("10.00"), s"Expected '10.00' in output: $result")
    }
  }

  test("col --auto-fit: empty column uses default width") {
    withTempFile { outputPath =>
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Text("Data in A")) // Data in column A, not C
      val wb = Workbook(sheet)

      val result = WriteCommands
        .col(wb, Some(sheet), "C", None, hide = false, show = false, autoFit = true, outputPath, config)
        .unsafeRunSync()

      assert(result.contains("(auto)"), s"Expected '(auto)' in output: $result")
      assert(result.contains("8.43"), s"Expected '8.43' in output: $result") // Default Excel width
    }
  }

  test("col --auto-fit: handles boolean values") {
    withTempFile { outputPath =>
      // A1=true (4 chars), A2=false (5 chars)
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Bool(true))
        .put(ref(0, 1), CellValue.Bool(false))
      val wb = Workbook(sheet)

      val result = WriteCommands
        .col(wb, Some(sheet), "A", None, hide = false, show = false, autoFit = true, outputPath, config)
        .unsafeRunSync()

      // Width = 5 chars (FALSE) * 0.85 + 1.5 = 5.75
      assert(result.contains("5.75"), s"Expected '5.75' in output: $result")
    }
  }

  test("col --auto-fit: handles formulas with cached values") {
    withTempFile { outputPath =>
      // A1=100, B1=A1*2 with cached value 200
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Number(BigDecimal(100)))
        .put(ref(1, 0), CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(200)))))
      val wb = Workbook(sheet)

      val result = WriteCommands
        .col(wb, Some(sheet), "B", None, hide = false, show = false, autoFit = true, outputPath, config)
        .unsafeRunSync()

      // Width = 3 chars (200) * 0.85 + 1.5 = 4.05, min 5.0
      assert(result.contains("5.00"), s"Expected '5.00' in output: $result")
    }
  }

  test("col --auto-fit: handles decimal numbers") {
    withTempFile { outputPath =>
      // A1=123.456789 (10 chars when formatted)
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Number(BigDecimal("123.456789")))
      val wb = Workbook(sheet)

      val result = WriteCommands
        .col(wb, Some(sheet), "A", None, hide = false, show = false, autoFit = true, outputPath, config)
        .unsafeRunSync()

      // Width = 10 chars * 0.85 + 1.5 = 10.0
      assert(result.contains("10.00"), s"Expected '10.00' in output: $result")
    }
  }

  test("col --auto-fit: explicit width overrides auto-fit") {
    withTempFile { outputPath =>
      // If both width and auto-fit are specified, auto-fit takes precedence
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Text("Short"))
      val wb = Workbook(sheet)

      // autoFit=true should override width=100
      val result = WriteCommands
        .col(wb, Some(sheet), "A", Some(100.0), hide = false, show = false, autoFit = true, outputPath, config)
        .unsafeRunSync()

      // Auto-fit should win: 5 chars * 0.85 + 1.5 = 5.75
      assert(result.contains("5.75"), s"Expected '5.75' in output: $result")
      assert(result.contains("(auto)"), s"Expected '(auto)' in output: $result")
    }
  }

  test("col: manual width works when auto-fit is false") {
    withTempFile { outputPath =>
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Text("Short"))
      val wb = Workbook(sheet)

      val result = WriteCommands
        .col(wb, Some(sheet), "A", Some(25.0), hide = false, show = false, autoFit = false, outputPath, config)
        .unsafeRunSync()

      assert(result.contains("25.00"), s"Expected '25.00' in output: $result")
      assert(!result.contains("(auto)"), s"Should not contain '(auto)' in output: $result")
    }
  }

  // ========== Column Range Auto-Fit Tests ==========

  test("col --auto-fit: supports column range (A:C)") {
    withTempFile { outputPath =>
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Text("Short")) // A1
        .put(ref(1, 0), CellValue.Text("Medium length text")) // B1
        .put(ref(2, 0), CellValue.Text("This is the longest")) // C1
      val wb = Workbook(sheet)

      val result = WriteCommands
        .col(wb, Some(sheet), "A:C", None, hide = false, show = false, autoFit = true, outputPath, config)
        .unsafeRunSync()

      // Should output multiple columns
      assert(result.contains("Columns:"), s"Expected 'Columns:' in output: $result")
      assert(result.contains("A:"), s"Expected 'A:' in output: $result")
      assert(result.contains("B:"), s"Expected 'B:' in output: $result")
      assert(result.contains("C:"), s"Expected 'C:' in output: $result")
      assert(result.contains("(auto)"), s"Expected '(auto)' in output: $result")

      // Verify widths in saved file
      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.headOption.getOrElse(fail("No sheets found"))
      val colA = Column.fromLetter("A").toOption.getOrElse(fail("Invalid column A"))
      val colB = Column.fromLetter("B").toOption.getOrElse(fail("Invalid column B"))
      val colC = Column.fromLetter("C").toOption.getOrElse(fail("Invalid column C"))

      // A: 5 chars * 0.85 + 1.5 = 5.75
      assertEquals(s.getColumnProperties(colA).width, Some(5.75))
      // B: 18 chars * 0.85 + 1.5 = 16.8
      assertEquals(s.getColumnProperties(colB).width, Some(16.8))
      // C: 19 chars * 0.85 + 1.5 = 17.65
      assertEquals(s.getColumnProperties(colC).width, Some(17.65))
    }
  }

  // ========== Global Auto-Fit Tests ==========

  test("autofit: auto-fits all used columns") {
    withTempFile { outputPath =>
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Text("A column data")) // A1
        .put(ref(2, 0), CellValue.Text("C column data here")) // C1 (skip B)
        .put(ref(3, 1), CellValue.Number(BigDecimal(12345678))) // D2
      val wb = Workbook(sheet)

      val result = WriteCommands
        .autoFit(wb, Some(sheet), None, outputPath, config)
        .unsafeRunSync()

      // Should auto-fit columns A through D (used range)
      assert(result.contains("Auto-fit"), s"Expected 'Auto-fit' in output: $result")
      assert(result.contains("4 column"), s"Expected '4 column' in output: $result")
      assert(result.contains("A:"), s"Expected 'A:' in output: $result")
      assert(result.contains("D:"), s"Expected 'D:' in output: $result")

      // Verify widths
      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.headOption.getOrElse(fail("No sheets found"))
      val colA = Column.fromLetter("A").toOption.getOrElse(fail("Invalid column A"))
      val colD = Column.fromLetter("D").toOption.getOrElse(fail("Invalid column D"))

      // A: 13 chars * 0.85 + 1.5 = 12.55
      assertEquals(s.getColumnProperties(colA).width, Some(12.55))
      // D: 8 chars * 0.85 + 1.5 = 8.3
      assertEquals(s.getColumnProperties(colD).width, Some(8.3))
    }
  }

  test("autofit: accepts specific column range") {
    withTempFile { outputPath =>
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Text("Column A")) // A1
        .put(ref(1, 0), CellValue.Text("Column B")) // B1
        .put(ref(2, 0), CellValue.Text("Column C")) // C1
        .put(ref(3, 0), CellValue.Text("Column D should not change")) // D1
      val wb = Workbook(sheet)

      val result = WriteCommands
        .autoFit(wb, Some(sheet), Some("A:C"), outputPath, config)
        .unsafeRunSync()

      assert(result.contains("3 column"), s"Expected '3 column' in output: $result")

      // Verify D was not touched
      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.headOption.getOrElse(fail("No sheets found"))
      val colD = Column.fromLetter("D").toOption.getOrElse(fail("Invalid column D"))
      assertEquals(s.getColumnProperties(colD).width, None) // No width set
    }
  }

  test("autofit: handles empty sheet gracefully") {
    withTempFile { outputPath =>
      val sheet = Sheet("Empty")
      val wb = Workbook(sheet)

      val result = WriteCommands
        .autoFit(wb, Some(sheet), None, outputPath, config)
        .unsafeRunSync()

      assert(result.contains("No columns to auto-fit"), s"Expected empty sheet message in output: $result")
    }
  }
