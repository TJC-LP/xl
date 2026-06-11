package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.{ARef, Column}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.extensions.style
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

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

  // Extract the first reported width (e.g. "A: 10.86 (auto)") from command output
  private def extractReportedWidth(result: String): Double =
    """(\d+\.\d+)""".r
      .findFirstIn(result)
      .map(_.toDouble)
      .getOrElse(fail(s"No width found in output: $result"))

  // Auto-fit column A of the given sheet and return the persisted width
  private def autoFitWidthOfColumnA(sheet: Sheet, outputPath: Path): Double =
    val wb = Workbook(sheet)
    WriteCommands
      .col(
        wb,
        Some(sheet),
        "A",
        None,
        hide = false,
        show = false,
        autoFit = true,
        outputPath,
        config
      )
      .unsafeRunSync()
    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.headOption.getOrElse(fail("No sheets found"))
    val col = Column.fromLetter("A").toOption.getOrElse(fail("Invalid column"))
    s.getColumnProperties(col).width.getOrElse(fail("No width set"))

  // ========== Font-Metric-Aware Auto-Fit Tests (GH-156) ==========
  //
  // These assert RELATIVE behavior so they pass in both measurement modes:
  //   - AWT mode: RenderUtils.measureTextWidth measures real glyph advances (works under
  //     java.awt.headless=true — BufferedImage needs no display, so CI takes this path too)
  //   - estimation mode: when graphics creation fails entirely (no fontconfig), RenderUtils
  //     falls back to a char-count estimate that still scales with font size and bold
  // Exact widths are environment-dependent (installed fonts differ), so no exact pins here.

  test("col --auto-fit: larger font autofits wider than default size (GH-156)") {
    val text = "Quarterly Revenue Report"
    val plainWidth = withTempFile { outputPath =>
      autoFitWidthOfColumnA(Sheet("Test").put(ref(0, 0), CellValue.Text(text)), outputPath)
    }
    val bigFontWidth = withTempFile { outputPath =>
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Text(text))
        .style(ref(0, 0), CellStyle.default.withFont(Font.default.withSize(22.0)))
      autoFitWidthOfColumnA(sheet, outputPath)
    }
    assert(
      bigFontWidth > plainWidth,
      s"22pt text ($bigFontWidth) should autofit wider than 11pt text ($plainWidth)"
    )
  }

  test("col --auto-fit: bold text autofits wider than plain text (GH-156)") {
    val text = "Quarterly Revenue Report"
    val plainWidth = withTempFile { outputPath =>
      autoFitWidthOfColumnA(Sheet("Test").put(ref(0, 0), CellValue.Text(text)), outputPath)
    }
    val boldWidth = withTempFile { outputPath =>
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Text(text))
        .style(ref(0, 0), CellStyle.default.withFont(Font.default.withBold(true)))
      autoFitWidthOfColumnA(sheet, outputPath)
    }
    assert(
      boldWidth > plainWidth,
      s"Bold text ($boldWidth) should autofit wider than plain text ($plainWidth)"
    )
  }

  test("col --auto-fit: large styled cells autofit wider than the char-count heuristic (GH-156)") {
    // The pre-GH-156 heuristic predicted chars * boldFactor(1.1) * 0.90 + 1.5 regardless of
    // font size. Font-metric measurement must give a 22pt bold cell strictly more room.
    val text = "Quarterly Revenue Report"
    val heuristicPrediction = text.length * 1.1 * 0.90 + 1.5
    val width = withTempFile { outputPath =>
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Text(text))
        .style(
          ref(0, 0),
          CellStyle.default.withFont(Font.default.withBold(true).withSize(22.0))
        )
      autoFitWidthOfColumnA(sheet, outputPath)
    }
    assert(
      width > heuristicPrediction,
      s"22pt bold width ($width) should exceed the char-count heuristic ($heuristicPrediction)"
    )
  }

  test("col --auto-fit: width stays within sane bounds for default font (GH-156)") {
    // Catches unit errors in the px -> Excel-width conversion (e.g. forgetting MDW division):
    // 24 chars of Calibri/DejaVu/fallback land in roughly [0.5, 1.5] units per char.
    val text = "Quarterly Revenue Report" // 24 chars
    val width = withTempFile { outputPath =>
      autoFitWidthOfColumnA(Sheet("Test").put(ref(0, 0), CellValue.Text(text)), outputPath)
    }
    assert(width >= 12.0 && width <= 36.0, s"Width $width out of sane range for 24 chars")
  }

  test("col --auto-fit: currency format autofits wider than the bare number (GH-156)") {
    val value = CellValue.Number(BigDecimal(45500))
    val bareWidth = withTempFile { outputPath =>
      autoFitWidthOfColumnA(Sheet("Test").put(ref(0, 0), value), outputPath)
    }
    val currencyWidth = withTempFile { outputPath =>
      val sheet = Sheet("Test")
        .put(ref(0, 0), value)
        .style(ref(0, 0), CellStyle.default.withNumFmt(NumFmt.Currency))
      autoFitWidthOfColumnA(sheet, outputPath)
    }
    assert(
      currencyWidth > bareWidth,
      s"Currency-formatted ($currencyWidth) should autofit wider than bare number ($bareWidth)"
    )
  }

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
        .col(
          wb,
          Some(sheet),
          "A",
          None,
          hide = false,
          show = false,
          autoFit = true,
          outputPath,
          config
        )
        .unsafeRunSync()

      assert(result.contains("(auto)"), s"Expected '(auto)' in output: $result")

      // Verify the file was saved with the column width; exact value is font-metric
      // dependent (GH-156), so assert a sane range for the longest cell (26 chars)
      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.headOption.getOrElse(fail("No sheets found"))
      val col = Column.fromLetter("A").toOption.getOrElse(fail("Invalid column"))
      val width = s.getColumnProperties(col).width.getOrElse(fail("No width set"))
      assert(width >= 13.0 && width <= 39.0, s"Width $width out of range for 26-char content")
      // Reported width matches the persisted width
      assert(result.contains(f"$width%.2f"), s"Expected width $width in output: $result")
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
        .col(
          wb,
          Some(sheet),
          "B",
          None,
          hide = false,
          show = false,
          autoFit = true,
          outputPath,
          config
        )
        .unsafeRunSync()

      assert(result.contains("(auto)"), s"Expected '(auto)' in output: $result")
      // Driven by the 10-digit value; exact width is font-metric dependent (GH-156)
      val width = extractReportedWidth(result)
      assert(width >= 9.0 && width <= 15.0, s"Width $width out of range for 10 digits: $result")
    }
  }

  test("col --auto-fit: empty column uses default width") {
    withTempFile { outputPath =>
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Text("Data in A")) // Data in column A, not C
      val wb = Workbook(sheet)

      val result = WriteCommands
        .col(
          wb,
          Some(sheet),
          "C",
          None,
          hide = false,
          show = false,
          autoFit = true,
          outputPath,
          config
        )
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
        .col(
          wb,
          Some(sheet),
          "A",
          None,
          hide = false,
          show = false,
          autoFit = true,
          outputPath,
          config
        )
        .unsafeRunSync()

      // Driven by "FALSE" (5 chars); exact width is font-metric dependent (GH-156)
      val width = extractReportedWidth(result)
      assert(width >= 5.0 && width <= 9.0, s"Width $width out of range for 'FALSE': $result")
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
        .col(
          wb,
          Some(sheet),
          "B",
          None,
          hide = false,
          show = false,
          autoFit = true,
          outputPath,
          config
        )
        .unsafeRunSync()

      // "200" (3 chars) measures well below the 5.0 minimum width floor
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
        .col(
          wb,
          Some(sheet),
          "A",
          None,
          hide = false,
          show = false,
          autoFit = true,
          outputPath,
          config
        )
        .unsafeRunSync()

      // Driven by "123.456789" (10 chars); exact width is font-metric dependent (GH-156)
      val width = extractReportedWidth(result)
      assert(width >= 8.0 && width <= 15.0, s"Width $width out of range for 10 chars: $result")
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
        .col(
          wb,
          Some(sheet),
          "A",
          Some(100.0),
          hide = false,
          show = false,
          autoFit = true,
          outputPath,
          config
        )
        .unsafeRunSync()

      // Auto-fit should win over width=100: "Short" needs far less than 100 units
      val width = extractReportedWidth(result)
      assert(width < 10.0, s"Auto-fit width $width should be derived from 'Short', not 100")
      assert(result.contains("(auto)"), s"Expected '(auto)' in output: $result")
    }
  }

  test("col: manual width works when auto-fit is false") {
    withTempFile { outputPath =>
      val sheet = Sheet("Test")
        .put(ref(0, 0), CellValue.Text("Short"))
      val wb = Workbook(sheet)

      val result = WriteCommands
        .col(
          wb,
          Some(sheet),
          "A",
          Some(25.0),
          hide = false,
          show = false,
          autoFit = false,
          outputPath,
          config
        )
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
        .col(
          wb,
          Some(sheet),
          "A:C",
          None,
          hide = false,
          show = false,
          autoFit = true,
          outputPath,
          config
        )
        .unsafeRunSync()

      // Should output multiple columns
      assert(result.contains("Columns:"), s"Expected 'Columns:' in output: $result")
      assert(result.contains("A:"), s"Expected 'A:' in output: $result")
      assert(result.contains("B:"), s"Expected 'B:' in output: $result")
      assert(result.contains("C:"), s"Expected 'C:' in output: $result")
      assert(result.contains("(auto)"), s"Expected '(auto)' in output: $result")

      // Verify widths in saved file: exact values are font-metric dependent (GH-156) and
      // proportional fonts mean char count alone doesn't order widths (B's wide glyphs can
      // out-measure C's narrow ones), but "Short" is unambiguously narrower than both
      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.headOption.getOrElse(fail("No sheets found"))
      val colA = Column.fromLetter("A").toOption.getOrElse(fail("Invalid column A"))
      val colB = Column.fromLetter("B").toOption.getOrElse(fail("Invalid column B"))
      val colC = Column.fromLetter("C").toOption.getOrElse(fail("Invalid column C"))

      val widthA = s.getColumnProperties(colA).width.getOrElse(fail("No width for A"))
      val widthB = s.getColumnProperties(colB).width.getOrElse(fail("No width for B"))
      val widthC = s.getColumnProperties(colC).width.getOrElse(fail("No width for C"))
      assert(widthA < widthB, s"'Short' ($widthA) should be narrower than 18 chars ($widthB)")
      assert(widthA < widthC, s"'Short' ($widthA) should be narrower than 19 chars ($widthC)")
      assert(widthA >= 5.0 && widthA <= 9.0, s"Width $widthA out of range for 'Short'")
      assert(widthB >= 10.0 && widthB <= 29.0, s"Width $widthB out of range for 18 chars")
      assert(widthC >= 10.0 && widthC <= 29.0, s"Width $widthC out of range for 19 chars")
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

      // Verify widths: exact values are font-metric dependent (GH-156); 13 chars of
      // text need more room than an 8-digit number
      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.headOption.getOrElse(fail("No sheets found"))
      val colA = Column.fromLetter("A").toOption.getOrElse(fail("Invalid column A"))
      val colD = Column.fromLetter("D").toOption.getOrElse(fail("Invalid column D"))

      val widthA = s.getColumnProperties(colA).width.getOrElse(fail("No width for A"))
      val widthD = s.getColumnProperties(colD).width.getOrElse(fail("No width for D"))
      assert(widthA > widthD, s"13-char text ($widthA) should be wider than 8 digits ($widthD)")
      assert(widthA >= 7.0 && widthA <= 21.0, s"Width $widthA out of range for 13 chars")
      assert(widthD >= 6.0 && widthD <= 13.0, s"Width $widthD out of range for 8 digits")
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

      assert(
        result.contains("No columns to auto-fit"),
        s"Expected empty sheet message in output: $result"
      )
    }
  }
