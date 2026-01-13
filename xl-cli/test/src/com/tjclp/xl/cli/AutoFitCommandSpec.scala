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

      assert(result.contains("auto-fit"))
      // Width = 26 chars (longest) + 2 padding = 28
      assert(result.contains("width=28.00"))

      // Verify the file was saved with the column width
      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.headOption.getOrElse(fail("No sheets found"))
      val col = Column.fromLetter("A").toOption.getOrElse(fail("Invalid column"))
      val colProps = s.getColumnProperties(col)
      assertEquals(colProps.width, Some(28.0))
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

      assert(result.contains("auto-fit"))
      // Width = 10 digits (1234567890) + 2 padding = 12
      assert(result.contains("width=12.00"))
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

      assert(result.contains("auto-fit"))
      assert(result.contains("width=8.43")) // Default Excel width
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

      // Width = 5 chars (FALSE) + 2 padding = 7, but min is 8.43
      assert(result.contains("width=8.43"))
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

      // Width = 3 chars (200) + 2 padding = 5, but min is 8.43
      assert(result.contains("width=8.43"))
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

      // Width = 10 chars + 2 padding = 12
      assert(result.contains("width=12.00"))
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

      // Auto-fit should win: 5 chars + 2 padding = 7, min 8.43
      assert(result.contains("width=8.43"))
      assert(result.contains("auto-fit"))
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

      assert(result.contains("width=25"))
      assert(!result.contains("auto-fit"))
    }
  }
