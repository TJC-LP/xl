package com.tjclp.xl.io

import cats.effect.IO
import munit.CatsEffectSuite
import java.nio.file.{Files, Path}
import com.tjclp.xl.*
import com.tjclp.xl.macros.{cell, range}

/** Tests for Excel streaming API */
class ExcelIOSpec extends CatsEffectSuite:

  val tempDir: FunFixture[Path] = FunFixture[Path](
    setup = _ => Files.createTempDirectory("xl-streaming-"),
    teardown = dir =>
      Files.walk(dir)
        .sorted(java.util.Comparator.reverseOrder())
        .forEach(Files.delete)
  )

  tempDir.test("read: loads workbook into memory") { dir =>
    // Create test file using current writer
    val wb = Workbook("Test").flatMap { initial =>
      val sheet = initial.sheets(0).put(cell"A1", CellValue.Text("Hello"))
      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Failed to create workbook"))

    val path = dir.resolve("test.xlsx")
    IO(com.tjclp.xl.ooxml.XlsxWriter.write(wb, path)).flatMap { _ =>
      // Read with ExcelIO
      val excel = ExcelIO.instance[IO]
      excel.read(path).map { readWb =>
        assertEquals(readWb.sheets.size, 1)
        assertEquals(readWb.sheets(0)(cell"A1").value, CellValue.Text("Hello"))
      }
    }
  }

  tempDir.test("write: creates valid XLSX file") { dir =>
    val wb = Workbook("Output").flatMap { initial =>
      val sheet = initial.sheets(0)
        .put(cell"A1", CellValue.Text("Written"))
        .put(cell"B1", CellValue.Number(BigDecimal(42)))
      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Failed to create workbook"))

    val path = dir.resolve("output.xlsx")
    val excel = ExcelIO.instance[IO]

    excel.write(wb, path).flatMap { _ =>
      // Verify file exists and is valid
      IO(Files.exists(path)).map { exists =>
        assert(exists, "File should exist")
      }
    }
  }

  tempDir.test("readStream: streams rows from sheet") { dir =>
    // Create test file with multiple rows
    val wb = Workbook("Data").flatMap { initial =>
      val sheet = (1 to 10).foldLeft(initial.sheets(0)) { (s, i) =>
        s.put(ARef(Column.from0(0), Row.from1(i)), CellValue.Number(BigDecimal(i)))
      }
      initial.updateSheet(0, sheet)
    }.getOrElse(fail("Failed to create workbook"))

    val path = dir.resolve("stream-test.xlsx")
    IO(com.tjclp.xl.ooxml.XlsxWriter.write(wb, path)).flatMap { _ =>
      val excel = ExcelIO.instance[IO]

      // Stream rows
      excel.readStream(path).compile.toList.map { rows =>
        assertEquals(rows.size, 10)
        assert(rows.forall(_.cells.nonEmpty))
      }
    }
  }

  tempDir.test("writeStream: creates file from row stream") { dir =>
    val path = dir.resolve("streamed.xlsx")
    val excel = ExcelIO.instance[IO]

    // Create row stream
    val rows = fs2.Stream.range(1, 6).map { i =>
      RowData(i, Map(
        0 -> CellValue.Text(s"Row $i"),
        1 -> CellValue.Number(BigDecimal(i * 100))
      ))
    }

    // Write stream
    rows.through(excel.writeStream(path, "Generated")).compile.drain.flatMap { _ =>
      // Read back to verify
      excel.read(path).map { wb =>
        assertEquals(wb.sheets.size, 1)
        assertEquals(wb.sheets(0).cellCount, 10)  // 5 rows × 2 cols
      }
    }
  }

  test("ExcelIO.forIO provides IO instance") {
    val excel = Excel.forIO
    IO(assertEquals(excel.getClass.getSimpleName, "ExcelIO"))
  }

  tempDir.test("writeStreamTrue: creates valid XLSX with 1k rows") { dir =>
    val path = dir.resolve("streamed-1k.xlsx")
    val excel = ExcelIO.instance[IO]

    // Generate 1000 rows
    val rows = fs2.Stream.range(1, 1001).map { i =>
      RowData(i, Map(
        0 -> CellValue.Text(s"Row $i"),
        1 -> CellValue.Number(BigDecimal(i * 100)),
        2 -> CellValue.Bool(i % 2 == 0)
      ))
    }

    // Write with true streaming
    rows.through(excel.writeStreamTrue(path, "Data")).compile.drain.flatMap { _ =>
      // Verify file exists and is readable
      IO(com.tjclp.xl.ooxml.XlsxReader.read(path)).map { result =>
        assert(result.isRight, "File should be readable")
        result.foreach { wb =>
          assertEquals(wb.sheets.size, 1)
          assertEquals(wb.sheets(0).name.value, "Data")
          // Each row has 3 cells, so 1000 rows × 3 cells = 3000 cells
          assertEquals(wb.sheets(0).cellCount, 3000)
        }
      }
    }
  }

  tempDir.test("writeStreamTrue: handles 100k rows with constant memory") { dir =>
    val path = dir.resolve("streamed-100k.xlsx")
    val excel = ExcelIO.instance[IO]

    // Generate 100,000 rows (this would OOM with materialized approach)
    val rows = fs2.Stream.range(1, 100001).map { i =>
      RowData(i, Map(
        0 -> CellValue.Text(s"Row $i"),
        1 -> CellValue.Number(BigDecimal(i))
      ))
    }

    // Write with true streaming
    rows.through(excel.writeStreamTrue(path, "BigData")).compile.drain.flatMap { _ =>
      // Verify file size is reasonable (10-20MB for 100k rows)
      IO(java.nio.file.Files.size(path)).map { size =>
        assert(size > 1_000_000, s"File should be at least 1MB, got $size bytes")
        assert(size < 50_000_000, s"File should be less than 50MB, got $size bytes")
      }
    }
  }

  tempDir.test("writeStreamTrue: output readable by XlsxReader") { dir =>
    val path = dir.resolve("streamed-test.xlsx")
    val excel = ExcelIO.instance[IO]

    // Generate test data
    val rows = fs2.Stream.range(1, 101).map { i =>
      RowData(i, Map(
        0 -> CellValue.Number(BigDecimal(i)),
        1 -> CellValue.Text(s"Text $i")
      ))
    }

    rows.through(excel.writeStreamTrue(path, "Test")).compile.drain.flatMap { _ =>
      // Read back with XlsxReader
      IO(com.tjclp.xl.ooxml.XlsxReader.read(path)).map { result =>
        assert(result.isRight, s"Should read successfully, got: $result")
        result.foreach { wb =>
          assertEquals(wb.sheets.size, 1)
          val sheet = wb.sheets(0)
          assertEquals(sheet.cellCount, 200)  // 100 rows × 2 cols

          // Verify first cell
          val firstCell = sheet(cell"A1")
          assertEquals(firstCell.value, CellValue.Number(BigDecimal(1)))

          // Verify last row
          val lastCell = sheet(ARef(Column.from0(0), Row.from1(100)))
          assertEquals(lastCell.value, CellValue.Number(BigDecimal(100)))
        }
      }
    }
  }
