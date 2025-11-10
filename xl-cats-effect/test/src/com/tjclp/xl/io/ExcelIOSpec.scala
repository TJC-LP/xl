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
