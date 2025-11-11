package com.tjclp.xl.ooxml

import munit.FunSuite
import java.nio.file.{Files, Path}
import com.tjclp.xl.*
import com.tjclp.xl.macros.cell
import com.tjclp.xl.codec.{*, given}

/**
 * Tests for ZIP compression configuration.
 *
 * Verifies that WriterConfig correctly applies DEFLATED vs STORED compression
 * and that file sizes match expectations.
 */
class CompressionSpec extends FunSuite:

  val tempDir: FunFixture[Path] = FunFixture[Path](
    setup = _ => Files.createTempDirectory("xl-compression-"),
    teardown = dir =>
      Files
        .walk(dir)
        .sorted(java.util.Comparator.reverseOrder())
        .forEach(Files.delete)
  )

  tempDir.test("DEFLATED produces smaller files than STORED") { dir =>
    // Create workbook with repetitive data (compresses well)
    val wb = Workbook("Data").flatMap { initial =>
      val cells = (1 to 1000).flatMap { row =>
        val aRef = ARef(Column.from1(1), Row.from1(row))
        val bRef = ARef(Column.from1(2), Row.from1(row))
        val cRef = ARef(Column.from1(3), Row.from1(row))
        Seq(
          aRef -> "Repeated text content that should compress well",
          bRef -> BigDecimal(row),
          cRef -> s"Row $row with more repetitive content"
        )
      }
      val sheet = initial.sheets(0).putMixed(cells*)
      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Failed to create workbook"))

    // Write with DEFLATED (default)
    val deflatedPath = dir.resolve("deflated.xlsx")
    val deflatedConfig = WriterConfig(compression = Compression.Deflated, prettyPrint = false)
    XlsxWriter.writeWith(wb, deflatedPath, deflatedConfig) match
      case Left(err) => fail(s"Failed to write DEFLATED: $err")
      case Right(_) => ()

    // Write with STORED (no compression)
    val storedPath = dir.resolve("stored.xlsx")
    val storedConfig = WriterConfig(compression = Compression.Stored, prettyPrint = false)
    XlsxWriter.writeWith(wb, storedPath, storedConfig) match
      case Left(err) => fail(s"Failed to write STORED: $err")
      case Right(_) => ()

    // Compare file sizes
    val deflatedSize = Files.size(deflatedPath)
    val storedSize = Files.size(storedPath)

    // DEFLATED should be 50-93% smaller (compression ratio 2x-15x)
    assert(deflatedSize < storedSize, s"DEFLATED ($deflatedSize) should be smaller than STORED ($storedSize)")

    val ratio = storedSize.toDouble / deflatedSize.toDouble
    assert(ratio >= 2.0, s"Compression ratio ($ratio) should be at least 2x")
    assert(ratio <= 15.0, s"Compression ratio ($ratio) seems unreasonably high")
  }

  tempDir.test("prettyPrint increases file size") { dir =>
    val wb = Workbook("Simple").flatMap { initial =>
      val sheet = initial.sheets(0).put(cell"A1", CellValue.Text("Test"))
      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Failed to create workbook"))

    // Write with compact XML
    val compactPath = dir.resolve("compact.xlsx")
    val compactConfig = WriterConfig(compression = Compression.Deflated, prettyPrint = false)
    XlsxWriter.writeWith(wb, compactPath, compactConfig) match
      case Left(err) => fail(s"Failed to write compact: $err")
      case Right(_) => ()

    // Write with pretty XML
    val prettyPath = dir.resolve("pretty.xlsx")
    val prettyConfig = WriterConfig(compression = Compression.Deflated, prettyPrint = true)
    XlsxWriter.writeWith(wb, prettyPath, prettyConfig) match
      case Left(err) => fail(s"Failed to write pretty: $err")
      case Right(_) => ()

    // Compare sizes
    val compactSize = Files.size(compactPath)
    val prettySize = Files.size(prettyPath)

    // Pretty-printed XML should be larger (whitespace overhead)
    assert(prettySize > compactSize, s"Pretty ($prettySize) should be larger than compact ($compactSize)")
  }

  tempDir.test("default config produces valid XLSX with DEFLATED compression") { dir =>
    val wb = Workbook("DefaultTest").flatMap { initial =>
      val cells = (1 to 100).map { row =>
        ARef(Column.from1(1), Row.from1(row)) -> s"Row $row"
      }
      val sheet = initial.sheets(0).putMixed(cells*)
      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Failed to create workbook"))

    val path = dir.resolve("default.xlsx")

    // Write with default config (should use DEFLATED + compact)
    XlsxWriter.write(wb, path) match
      case Left(err) => fail(s"Failed to write with defaults: $err")
      case Right(_) => ()

    // Verify file exists and is readable
    assert(Files.exists(path), "File should exist")
    val size = Files.size(path)
    assert(size > 0, "File should not be empty")
    assert(size < 1_000_000, s"File size ($size) seems unreasonably large for 100 rows")

    // Verify it's readable
    XlsxReader.read(path) match
      case Right(readWb) =>
        assertEquals(readWb.sheets.size, 1)
        assertEquals(readWb.sheets(0).name.value, "DefaultTest")
        assertEquals(readWb.sheets(0)(cell"A1").value, CellValue.Text("Row 1"))
      case Left(err) =>
        fail(s"Failed to read back default file: $err")
  }

  tempDir.test("WriterConfig.debug uses STORED + prettyPrint") { dir =>
    val wb = Workbook("Debug").flatMap { initial =>
      val sheet = initial.sheets(0).put(cell"A1", CellValue.Number(42))
      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Failed to create workbook"))

    val path = dir.resolve("debug.xlsx")
    XlsxWriter.writeWith(wb, path, WriterConfig.debug) match
      case Left(err) => fail(s"Failed to write debug: $err")
      case Right(_) => ()

    // Verify file is readable
    XlsxReader.read(path) match
      case Right(readWb) =>
        assertEquals(readWb.sheets.size, 1)
        assertEquals(readWb.sheets(0)(cell"A1").value, CellValue.Number(42))
      case Left(err) =>
        fail(s"Failed to read debug file: $err")
  }
