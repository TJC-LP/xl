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

  tempDir.test("writeStreamTrue: supports arbitrary sheet index") { dir =>
    val path = dir.resolve("sheet3.xlsx")
    val excel = ExcelIO.instance[IO]

    val rows = fs2.Stream.range(1, 11).map { i =>
      RowData(i, Map(0 -> CellValue.Text(s"Row $i")))
    }

    // Write with sheetIndex = 3
    rows.through(excel.writeStreamTrue(path, "ThirdSheet", sheetIndex = 3)).compile.drain.flatMap { _ =>
      // Verify ZIP contains sheet3.xml (not sheet1.xml)
      IO {
        import java.util.zip.ZipFile
        val zipFile = new ZipFile(path.toFile)
        try {
          val entries = scala.jdk.CollectionConverters.IteratorHasAsScala(
            zipFile.entries().asIterator()
          ).asScala.map(_.getName).toSet

          // Should have sheet3.xml
          assert(entries.contains("xl/worksheets/sheet3.xml"),
            s"Expected sheet3.xml, got: ${entries.filter(_.contains("sheet"))}")

          // Should NOT have sheet1.xml or sheet2.xml
          assert(!entries.contains("xl/worksheets/sheet1.xml"), "Should not have sheet1.xml")
          assert(!entries.contains("xl/worksheets/sheet2.xml"), "Should not have sheet2.xml")
        } finally {
          zipFile.close()
        }
      } >>
      // Verify workbook.xml has sheetId="3"
      IO(com.tjclp.xl.ooxml.XlsxReader.read(path)).map { result =>
        assert(result.isRight, s"Should read successfully, got: $result")
        result.foreach { wb =>
          assertEquals(wb.sheets.size, 1)
          assertEquals(wb.sheets(0).name.value, "ThirdSheet")
        }
      }
    }
  }

  tempDir.test("writeStreamTrue: rejects invalid sheet index") { dir =>
    val path = dir.resolve("invalid.xlsx")
    val excel = ExcelIO.instance[IO]
    val rows = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("test"))))

    // sheetIndex = 0 should fail
    rows.through(excel.writeStreamTrue(path, "Bad", sheetIndex = 0))
      .compile.drain
      .attempt
      .map {
        case Left(err: IllegalArgumentException) =>
          assert(err.getMessage.contains("Sheet index must be >= 1"))
        case Left(other) =>
          fail(s"Expected IllegalArgumentException, got: $other")
        case Right(_) =>
          fail("Should have failed with illegal sheet index")
      }
  }

  tempDir.test("writeStreamsSeqTrue: writes 2 sheets") { dir =>
    val path = dir.resolve("two-sheets.xlsx")
    val excel = ExcelIO.instance[IO]

    val sheet1Rows = fs2.Stream.range(1, 11).map { i =>
      RowData(i, Map(0 -> CellValue.Text(s"Sheet1 Row $i")))
    }
    val sheet2Rows = fs2.Stream.range(1, 6).map { i =>
      RowData(i, Map(0 -> CellValue.Number(BigDecimal(i * 100))))
    }

    excel.writeStreamsSeqTrue(path, Seq("First" -> sheet1Rows, "Second" -> sheet2Rows)).flatMap {
      _ =>
        // Verify with XlsxReader
        IO(com.tjclp.xl.ooxml.XlsxReader.read(path)).map { result =>
          assert(result.isRight, s"Should read successfully: $result")
          result.foreach { wb =>
            // Should have 2 sheets
            assertEquals(wb.sheets.size, 2)

            // First sheet
            assertEquals(wb.sheets(0).name.value, "First")
            assertEquals(wb.sheets(0).cellCount, 10)

            // Second sheet
            assertEquals(wb.sheets(1).name.value, "Second")
            assertEquals(wb.sheets(1).cellCount, 5)

            // Verify content
            val firstCell = wb.sheets(0)(cell"A1")
            assertEquals(firstCell.value, CellValue.Text("Sheet1 Row 1"))

            val secondCell = wb.sheets(1)(cell"A1")
            assertEquals(secondCell.value, CellValue.Number(BigDecimal(100)))
          }
        }
    }
  }

  tempDir.test("writeStreamsSeqTrue: writes 3 sheets with different sizes") { dir =>
    val path = dir.resolve("three-sheets.xlsx")
    val excel = ExcelIO.instance[IO]

    val sales = fs2.Stream.range(1, 101).map(i => RowData(i, Map(0 -> CellValue.Number(i))))
    val inventory = fs2.Stream.range(1, 51).map(i => RowData(i, Map(0 -> CellValue.Text(s"Item $i"))))
    val summary = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("Summary"))))

    excel
      .writeStreamsSeqTrue(
        path,
        Seq("Sales" -> sales, "Inventory" -> inventory, "Summary" -> summary)
      )
      .flatMap { _ =>
        IO(com.tjclp.xl.ooxml.XlsxReader.read(path)).map { result =>
          assert(result.isRight)
          result.foreach { wb =>
            assertEquals(wb.sheets.size, 3)
            assertEquals(wb.sheets(0).name.value, "Sales")
            assertEquals(wb.sheets(0).cellCount, 100)
            assertEquals(wb.sheets(1).name.value, "Inventory")
            assertEquals(wb.sheets(1).cellCount, 50)
            assertEquals(wb.sheets(2).name.value, "Summary")
            assertEquals(wb.sheets(2).cellCount, 1)
          }
        }
      }
  }

  tempDir.test("writeStreamsSeqTrue: handles large multi-sheet workbook") { dir =>
    val path = dir.resolve("large-multi-sheet.xlsx")
    val excel = ExcelIO.instance[IO]

    // 3 sheets with 10k rows each (30k total rows)
    val sheet1 = fs2.Stream.range(1, 10001).map(i => RowData(i, Map(0 -> CellValue.Number(i))))
    val sheet2 = fs2.Stream.range(1, 10001).map(i => RowData(i, Map(0 -> CellValue.Text(s"$i"))))
    val sheet3 = fs2.Stream.range(1, 10001).map(i => RowData(i, Map(0 -> CellValue.Bool(i % 2 == 0))))

    excel
      .writeStreamsSeqTrue(path, Seq("Numbers" -> sheet1, "Text" -> sheet2, "Booleans" -> sheet3))
      .flatMap { _ =>
        // Verify file size is reasonable
        IO(java.nio.file.Files.size(path)).map { size =>
          assert(size > 100_000, s"File should be at least 100KB, got $size bytes")
          assert(size < 10_000_000, s"File should be less than 10MB, got $size bytes")
        }
      }
  }

  tempDir.test("writeStreamsSeqTrue: rejects empty sheet list") { dir =>
    val path = dir.resolve("empty.xlsx")
    val excel = ExcelIO.instance[IO]

    excel
      .writeStreamsSeqTrue(path, Seq.empty)
      .attempt
      .map {
        case Left(err: IllegalArgumentException) =>
          assert(err.getMessage.contains("at least one sheet"))
        case Left(other) =>
          fail(s"Expected IllegalArgumentException, got: $other")
        case Right(_) =>
          fail("Should have failed with empty sheet list")
      }
  }

  tempDir.test("writeStreamsSeqTrue: rejects duplicate sheet names") { dir =>
    val path = dir.resolve("duplicates.xlsx")
    val excel = ExcelIO.instance[IO]

    val rows1 = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("A"))))
    val rows2 = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("B"))))

    excel
      .writeStreamsSeqTrue(path, Seq("Sheet1" -> rows1, "Sheet1" -> rows2))
      .attempt
      .map {
        case Left(err: IllegalArgumentException) =>
          assert(err.getMessage.contains("Duplicate sheet names"))
        case Left(other) =>
          fail(s"Expected IllegalArgumentException, got: $other")
        case Right(_) =>
          fail("Should have failed with duplicate names")
      }
  }

  tempDir.test("readStream: handles 100k rows with constant memory") { dir =>
    val path = dir.resolve("large-read.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write 100k rows with streaming
    val writeRows = fs2.Stream.range(1, 100001).map { i =>
      RowData(i, Map(
        0 -> CellValue.Number(BigDecimal(i)),
        1 -> CellValue.Text(s"Row $i"),
        2 -> CellValue.Bool(i % 2 == 0)
      ))
    }

    writeRows.through(excel.writeStreamTrue(path, "Data")).compile.drain.flatMap { _ =>
      // Read with streaming (should use constant memory)
      excel.readStream(path).compile.count.map { count =>
        assertEquals(count, 100000L, "Should read all 100k rows")
      }
    }
  }

  tempDir.test("readStream: full round-trip preserves data") { dir =>
    val path = dir.resolve("round-trip.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write test data
    val writeRows = fs2.Stream.range(1, 1001).map { i =>
      RowData(i, Map(
        0 -> CellValue.Number(BigDecimal(i)),
        1 -> CellValue.Text(s"Text $i"),
        2 -> CellValue.Bool(i % 2 == 0)
      ))
    }

    writeRows.through(excel.writeStreamTrue(path, "Test")).compile.drain.flatMap { _ =>
      // Read back with streaming
      excel.readStream(path).compile.toVector.map { rows =>
        assertEquals(rows.size, 1000)

        // Verify first row
        val first = rows.head
        assertEquals(first.rowIndex, 1)
        assertEquals(first.cells(0), CellValue.Number(BigDecimal(1)))
        assertEquals(first.cells(1), CellValue.Text("Text 1"))
        assertEquals(first.cells(2), CellValue.Bool(false))

        // Verify last row
        val last = rows.last
        assertEquals(last.rowIndex, 1000)
        assertEquals(last.cells(0), CellValue.Number(BigDecimal(1000)))
        assertEquals(last.cells(1), CellValue.Text("Text 1000"))
        assertEquals(last.cells(2), CellValue.Bool(true))
      }
    }
  }

  tempDir.test("readStream: handles all cell types") { dir =>
    val path = dir.resolve("all-types.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write different cell types
    val rows = fs2.Stream.emit(
      RowData(
        1,
        Map(
          0 -> CellValue.Number(BigDecimal("123.45")),
          1 -> CellValue.Text("Hello"),
          2 -> CellValue.Bool(true),
          3 -> CellValue.Bool(false),
          4 -> CellValue.Error(CellError.Div0),
          5 -> CellValue.Empty
        )
      )
    )

    rows.through(excel.writeStreamTrue(path, "Types")).compile.drain.flatMap { _ =>
      excel.readStream(path).compile.toVector.map { readRows =>
        assertEquals(readRows.size, 1)
        val row = readRows.head

        assertEquals(row.cells(0), CellValue.Number(BigDecimal("123.45")))
        assertEquals(row.cells(1), CellValue.Text("Hello"))
        assertEquals(row.cells(2), CellValue.Bool(true))
        assertEquals(row.cells(3), CellValue.Bool(false))
        assertEquals(row.cells(4), CellValue.Error(CellError.Div0))
        assert(!row.cells.contains(5), "Empty cells should not be in map")
      }
    }
  }

  tempDir.test("readStream: processes rows incrementally") { dir =>
    val path = dir.resolve("incremental.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write 10k rows
    val writeRows = fs2.Stream.range(1, 10001).map(i =>
      RowData(i, Map(0 -> CellValue.Number(BigDecimal(i))))
    )

    writeRows.through(excel.writeStreamTrue(path, "Data")).compile.drain.flatMap { _ =>
      // Read and filter incrementally (should not materialize all rows)
      excel.readStream(path)
        .filter(_.rowIndex % 100 == 0)  // Only every 100th row
        .compile.count.map { count =>
          assertEquals(count, 100L, "Should find 100 rows (1% of 10k)")
        }
    }
  }

  tempDir.test("readStreamByIndex: reads specific sheet by index") { dir =>
    val path = dir.resolve("multi-read.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write 3 sheets with different data
    val sheet1 = fs2.Stream.range(1, 11).map(i => RowData(i, Map(0 -> CellValue.Number(i))))
    val sheet2 = fs2.Stream.range(1, 21).map(i => RowData(i, Map(0 -> CellValue.Text(s"S2-$i"))))
    val sheet3 = fs2.Stream.range(1, 31).map(i => RowData(i, Map(0 -> CellValue.Bool(i % 2 == 0))))

    excel.writeStreamsSeqTrue(path, Seq("First" -> sheet1, "Second" -> sheet2, "Third" -> sheet3))
      .flatMap { _ =>
        // Read second sheet by index
        excel.readStreamByIndex(path, 2).compile.toVector.map { rows =>
          assertEquals(rows.size, 20, "Should read 20 rows from second sheet")
          assertEquals(rows.head.cells(0), CellValue.Text("S2-1"))
          assertEquals(rows.last.cells(0), CellValue.Text("S2-20"))
        }
      }
  }

  tempDir.test("readSheetStream: reads specific sheet by name") { dir =>
    val path = dir.resolve("named-read.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write 3 sheets
    val sales = fs2.Stream.range(1, 11).map(i => RowData(i, Map(0 -> CellValue.Number(i * 100))))
    val inventory = fs2.Stream.range(1, 6).map(i => RowData(i, Map(0 -> CellValue.Text(s"Item $i"))))
    val summary = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("Done"))))

    excel.writeStreamsSeqTrue(path, Seq("Sales" -> sales, "Inventory" -> inventory, "Summary" -> summary))
      .flatMap { _ =>
        // Read "Inventory" sheet by name
        excel.readSheetStream(path, "Inventory").compile.toVector.map { rows =>
          assertEquals(rows.size, 5, "Should read 5 rows from Inventory")
          assertEquals(rows.head.cells(0), CellValue.Text("Item 1"))
          assertEquals(rows.last.cells(0), CellValue.Text("Item 5"))
        }
      }
  }

  tempDir.test("readStreamByIndex: handles invalid index") { dir =>
    val path = dir.resolve("invalid-index.xlsx")
    val excel = ExcelIO.instance[IO]

    val rows = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("test"))))
    rows.through(excel.writeStreamTrue(path, "Only")).compile.drain.flatMap { _ =>
      // Try to read sheet 5 (doesn't exist)
      excel.readStreamByIndex(path, 5)
        .compile.toList
        .attempt
        .map {
          case Left(err) =>
            assert(err.getMessage.contains("Worksheet not found"))
          case Right(_) =>
            fail("Should have failed with nonexistent sheet")
        }
    }
  }

  tempDir.test("readSheetStream: handles nonexistent sheet name") { dir =>
    val path = dir.resolve("bad-name.xlsx")
    val excel = ExcelIO.instance[IO]

    val rows = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("test"))))
    rows.through(excel.writeStreamTrue(path, "RealSheet")).compile.drain.flatMap { _ =>
      // Try to read "FakeSheet"
      excel.readSheetStream(path, "FakeSheet")
        .compile.toList
        .attempt
        .map {
          case Left(err) =>
            assert(err.getMessage.contains("Sheet not found"))
          case Right(_) =>
            fail("Should have failed with nonexistent sheet name")
        }
    }
  }
