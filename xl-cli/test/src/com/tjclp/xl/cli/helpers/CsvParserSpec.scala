package com.tjclp.xl.cli.helpers

import munit.FunSuite
import cats.effect.unsafe.implicits.global

import java.nio.file.{Files, Path}

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue

/**
 * Tests for CSV parsing and type inference.
 *
 * Note: Tests indirectly verify private methods (inferColumnTypes, parseValue) through the public
 * parseCsv API.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class CsvParserSpec extends FunSuite:

  // Helper to create temporary CSV files
  def createTempCsv(content: String): Path =
    val tempFile = Files.createTempFile("test", ".csv")
    Files.writeString(tempFile, content)
    tempFile

  val defaultOptions = CsvParser.ImportOptions()

  // ========== Type Inference Tests (via parseCsv) ==========

  test("parseCsv: type inference detects Number column") {
    val csvFile = createTempCsv("Value\n100\n200\n300")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), defaultOptions).unsafeRunSync()

    // All values should be Number type
    assertEquals(result(0)._2, CellValue.Number(BigDecimal("100")))
    assertEquals(result(1)._2, CellValue.Number(BigDecimal("200")))
    assertEquals(result(2)._2, CellValue.Number(BigDecimal("300")))
  }

  test("parseCsv: type inference detects Boolean column") {
    val csvFile = createTempCsv("Active\ntrue\nfalse\nTRUE")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), defaultOptions).unsafeRunSync()

    assertEquals(result(0)._2, CellValue.Bool(true))
    assertEquals(result(1)._2, CellValue.Bool(false))
    assertEquals(result(2)._2, CellValue.Bool(true))
  }

  test("parseCsv: type inference detects Date column") {
    val csvFile = createTempCsv("Date\n2024-01-15\n2024-02-20")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), defaultOptions).unsafeRunSync()

    result.foreach { case (_, value) =>
      value match
        case CellValue.DateTime(_) => () // Success
        case other => fail(s"Expected DateTime, got $other")
    }
  }

  test("parseCsv: mixed types use majority detection with individual fallback") {
    val csvFile = createTempCsv("Value\n100\nN/A\n200")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), defaultOptions).unsafeRunSync()

    // Column typed as Number (majority 2/3 >= 80% threshold), individual cells that can't parse fall back to Text
    assertEquals(result(0)._2, CellValue.Number(BigDecimal(100)))
    assertEquals(result(1)._2, CellValue.Text("N/A")) // Can't parse as number, falls back
    assertEquals(result(2)._2, CellValue.Number(BigDecimal(200)))
  }

  test("parseCsv: empty column defaults to Text") {
    val csvFile = createTempCsv("Name,Notes\nAlice,\nBob,")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), defaultOptions).unsafeRunSync()

    // Column B (Notes) is all empty, should default to Text type (no error)
    assertEquals(result(0), (ARef.from0(0, 0), CellValue.Text("Alice")))
    assertEquals(result(1), (ARef.from0(1, 0), CellValue.Empty))
    assertEquals(result(2), (ARef.from0(0, 1), CellValue.Text("Bob")))
    assertEquals(result(3), (ARef.from0(1, 1), CellValue.Empty))
  }

  // ========== Full Parsing Tests ==========

  test("parseCsv: basic CSV with header") {
    val csvFile = createTempCsv("Name,Age\nAlice,30\nBob,25")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), defaultOptions).unsafeRunSync()

    assertEquals(result.length, 4) // 2 rows × 2 columns
    // Check structure
    assert(result.exists(_._1 == ARef.from0(0, 0)), "Should have A1")
    assert(result.exists(_._1 == ARef.from0(1, 1)), "Should have B2")
  }

  test("parseCsv: no header mode") {
    val csvFile = createTempCsv("100,200\n300,400")
    val options = defaultOptions.copy(hasHeader = false)
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), options).unsafeRunSync()

    assertEquals(result.length, 4) // 2 rows × 2 columns (all rows treated as data)
    assertEquals(result(0), (ARef.from0(0, 0), CellValue.Number(BigDecimal("100"))))
  }

  test("parseCsv: no header mode with CSV that has header row (GH-170)") {
    // When using --no-header on a CSV with a header, type inference should still work
    // because majority-based detection ignores the single header row
    val csvFile = createTempCsv("Name,Amount\nAlice,100\nBob,200\nCharlie,300")
    val options = defaultOptions.copy(hasHeader = false)
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), options).unsafeRunSync()

    assertEquals(result.length, 8) // 4 rows × 2 columns (header + 3 data rows)

    // Column A (Name) should be Text
    assertEquals(result(0)._2, CellValue.Text("Name"))
    assertEquals(result(2)._2, CellValue.Text("Alice"))

    // Column B (Amount) should be Number for data rows, Text for header
    // The header "Amount" falls back to Text since it can't parse as number
    assertEquals(result(1)._2, CellValue.Text("Amount")) // Header can't parse as number
    assertEquals(result(3)._2, CellValue.Number(BigDecimal(100))) // Data row parsed as number
    assertEquals(result(5)._2, CellValue.Number(BigDecimal(200)))
    assertEquals(result(7)._2, CellValue.Number(BigDecimal(300)))
  }

  test("parseCsv: import at offset position (B5)") {
    val csvFile = createTempCsv("Name\nAlice")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(1, 4), defaultOptions).unsafeRunSync()

    // Should start at B5 (col=1, row=4)
    assertEquals(result.length, 1)
    assertEquals(result(0)._1, ARef.from0(1, 4)) // B5
    assertEquals(result(0)._2, CellValue.Text("Alice"))
  }

  test("parseCsv: custom delimiter (semicolon)") {
    val csvFile = createTempCsv("Name;Age\nAlice;30")
    val options = defaultOptions.copy(delimiter = ';')
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), options).unsafeRunSync()

    assertEquals(result.length, 2)
    assertEquals(result(0), (ARef.from0(0, 0), CellValue.Text("Alice")))
    assertEquals(result(1), (ARef.from0(1, 0), CellValue.Number(BigDecimal("30"))))
  }

  test("parseCsv: empty CSV file") {
    val csvFile = createTempCsv("")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), defaultOptions).unsafeRunSync()

    assertEquals(result, Vector.empty)
  }

  test("parseCsv: header-only file") {
    val csvFile = createTempCsv("Name,Age,City")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), defaultOptions).unsafeRunSync()

    assertEquals(result, Vector.empty) // No data rows after header
  }

  test("parseCsv: no type inference mode") {
    val csvFile = createTempCsv("ZIP\n01234\n98765")
    val options = defaultOptions.copy(inferTypes = false)
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), options).unsafeRunSync()

    // All values should be Text
    assertEquals(result(0)._2, CellValue.Text("01234"))
    assertEquals(result(1)._2, CellValue.Text("98765"))
  }

  test("parseCsv: bounds check - column overflow") {
    // Create CSV with columns that would exceed XFD when starting at XFC
    val largeCsv = "A,B,C,D,E,F,G,H,I,J\n" + ("1," * 10).dropRight(1)
    val csvFile = createTempCsv(largeCsv)

    // Start at XFC (col 16380), 10 columns would exceed XFD (16383)
    val startRef = ARef.from0(16380, 0)

    val result = CsvParser.parseCsv(csvFile, startRef, defaultOptions).attempt.unsafeRunSync()

    assert(result.isLeft, "Should fail with bounds error")
    val error = result.swap.getOrElse(throw new Exception("Expected Left but got Right"))
    assert(
      error.getMessage.contains("Excel column limit"),
      s"Error message should mention column limit: ${error.getMessage}"
    )
  }

  test("parseCsv: bounds check - row overflow") {
    // Create CSV that would exceed row limit
    // Header row + 2 data rows, starting at row 1048575 would exceed 1048575 max
    val csvFile = createTempCsv("A\n1\n2")

    // Start at row 1048575 (max index 0-based), CSV has 2 data rows → would go to 1048576
    val startRef = ARef.from0(0, 1048575)

    val result = CsvParser.parseCsv(csvFile, startRef, defaultOptions).attempt.unsafeRunSync()

    assert(result.isLeft, "Should fail with bounds error")
    val error = result.swap.getOrElse(throw new Exception("Expected Left but got Right"))
    assert(
      error.getMessage.contains("Excel row limit"),
      s"Error message should mention row limit: ${error.getMessage}"
    )
  }

  test("parseCsv: file not found error") {
    val nonExistentPath = Path.of("/tmp/nonexistent-file-xyz123.csv")
    val result =
      CsvParser.parseCsv(nonExistentPath, ARef.from0(0, 0), defaultOptions).attempt.unsafeRunSync()

    assert(result.isLeft, "Should fail with file not found error")
    val error = result.swap.getOrElse(throw new Exception("Expected Left but got Right"))
    assert(
      error.getMessage.contains("Failed to parse CSV"),
      s"Error should be wrapped: ${error.getMessage}"
    )
  }

  test("parseCsv: handles empty cells in middle of row") {
    val csvFile = createTempCsv("A,B,C\n1,,3\n4,5,")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), defaultOptions).unsafeRunSync()

    // Row 1: 1, Empty, 3
    assertEquals(result(0), (ARef.from0(0, 0), CellValue.Number(BigDecimal("1"))))
    assertEquals(result(1), (ARef.from0(1, 0), CellValue.Empty))
    assertEquals(result(2), (ARef.from0(2, 0), CellValue.Number(BigDecimal("3"))))

    // Row 2: 4, 5, Empty
    assertEquals(result(3), (ARef.from0(0, 1), CellValue.Number(BigDecimal("4"))))
    assertEquals(result(4), (ARef.from0(1, 1), CellValue.Number(BigDecimal("5"))))
    assertEquals(result(5), (ARef.from0(2, 1), CellValue.Empty))
  }

  test("parseCsv: type inference with decimals") {
    val csvFile = createTempCsv("Price\n29.99\n49.99\n39.99")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), defaultOptions).unsafeRunSync()

    assertEquals(result.length, 3)
    assertEquals(result(0)._2, CellValue.Number(BigDecimal("29.99")))
    assertEquals(result(1)._2, CellValue.Number(BigDecimal("49.99")))
    assertEquals(result(2)._2, CellValue.Number(BigDecimal("39.99")))
  }

  test("parseCsv: Boolean case-insensitive") {
    val csvFile = createTempCsv("Active\nTRUE\nfalse\nTrue")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), defaultOptions).unsafeRunSync()

    assertEquals(result(0)._2, CellValue.Bool(true))
    assertEquals(result(1)._2, CellValue.Bool(false))
    assertEquals(result(2)._2, CellValue.Bool(true))
  }

  test("parseCsv: type inference with negative numbers") {
    val csvFile = createTempCsv("Value\n-100\n-200.50\n300")
    val result = CsvParser.parseCsv(csvFile, ARef.from0(0, 0), defaultOptions).unsafeRunSync()

    assertEquals(result(0)._2, CellValue.Number(BigDecimal("-100")))
    assertEquals(result(1)._2, CellValue.Number(BigDecimal("-200.50")))
    assertEquals(result(2)._2, CellValue.Number(BigDecimal("300")))
  }
