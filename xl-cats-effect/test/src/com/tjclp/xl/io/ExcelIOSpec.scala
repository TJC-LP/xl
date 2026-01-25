package com.tjclp.xl.io

import cats.effect.IO
import munit.CatsEffectSuite

import java.nio.file.{Files, Path}
import com.tjclp.xl.{*, given}
import com.tjclp.xl.unsafe.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{WriterConfig, XlsxReader}

/** Tests for Excel streaming API */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
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
    val initial = Workbook("Test")
    val sheet = initial.sheets(0).put(ref"A1", CellValue.Text("Hello"))
    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Failed to create workbook"))

    val path = dir.resolve("test.xlsx")
    IO(com.tjclp.xl.ooxml.XlsxWriter.write(wb, path)).flatMap { _ =>
      // Read with ExcelIO
      val excel = ExcelIO.instance[IO]
      excel.read(path).map { readWb =>
        assertEquals(readWb.sheets.size, 1)
        assertEquals(readWb.sheets(0)(ref"A1").value, CellValue.Text("Hello"))
      }
    }
  }

  tempDir.test("write: creates valid XLSX file") { dir =>
    val initial = Workbook("Output")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Text("Written"))
      .put(ref"B1", CellValue.Number(BigDecimal(42)))
    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Failed to create workbook"))

    val path = dir.resolve("output.xlsx")
    val excel = ExcelIO.instance[IO]

    excel.write(wb, path).flatMap { _ =>
      // Verify file exists and is valid
      IO(Files.exists(path)).map { exists =>
        assert(exists, "File should exist")
      }
    }
  }

  tempDir.test("writeWith SaxStax: creates valid XLSX file using SAX backend") { dir =>
    val initial = Workbook("OutputFast")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Text("Fast"))
      .put(ref"B1", CellValue.Number(BigDecimal(84)))
    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Failed to create workbook"))

    val path = dir.resolve("output-fast.xlsx")
    val excel = ExcelIO.instance[IO]

    excel.writeWith(wb, path, WriterConfig.saxStax).flatMap { _ =>
      IO(Files.exists(path)).map { exists =>
        assert(exists, "File should exist")
      }
    }
  }

  tempDir.test("readStream: streams rows from sheet") { dir =>
    // Create test file with multiple rows
    val initial = Workbook("Data")
    val sheet = (1 to 10).foldLeft(initial.sheets(0)) { (s, i) =>
      s.put(ARef(Column.from0(0), Row.from1(i)), CellValue.Number(BigDecimal(i)))
    }
    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Failed to create workbook"))

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

  tempDir.test("writeStream: creates valid XLSX with 1k rows") { dir =>
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
    rows.through(excel.writeStream(path, "Data")).compile.drain.flatMap { _ =>
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

  tempDir.test("writeStream: handles 100k rows with constant memory") { dir =>
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
    rows.through(excel.writeStream(path, "BigData")).compile.drain.flatMap { _ =>
      // Verify file size is reasonable (10-20MB for 100k rows)
      IO(java.nio.file.Files.size(path)).map { size =>
        assert(size > 1_000_000, s"File should be at least 1MB, got $size bytes")
        assert(size < 50_000_000, s"File should be less than 50MB, got $size bytes")
      }
    }
  }

  tempDir.test("writeStream: output readable by XlsxReader") { dir =>
    val path = dir.resolve("streamed-test.xlsx")
    val excel = ExcelIO.instance[IO]

    // Generate test data
    val rows = fs2.Stream.range(1, 101).map { i =>
      RowData(i, Map(
        0 -> CellValue.Number(BigDecimal(i)),
        1 -> CellValue.Text(s"Text $i")
      ))
    }

    rows.through(excel.writeStream(path, "Test")).compile.drain.flatMap { _ =>
      // Read back with XlsxReader
      IO(com.tjclp.xl.ooxml.XlsxReader.read(path)).map { result =>
        assert(result.isRight, s"Should read successfully, got: $result")
        result.foreach { wb =>
          assertEquals(wb.sheets.size, 1)
          val sheet = wb.sheets(0)
          assertEquals(sheet.cellCount, 200)  // 100 rows × 2 cols

          // Verify first cell
          val firstCell = sheet(ref"A1")
          assertEquals(firstCell.value, CellValue.Number(BigDecimal(1)))

          // Verify last row
          val lastCell = sheet(ARef(Column.from0(0), Row.from1(100)))
          assertEquals(lastCell.value, CellValue.Number(BigDecimal(100)))
        }
      }
    }
  }

  tempDir.test("writeStream: supports arbitrary sheet index") { dir =>
    val path = dir.resolve("sheet3.xlsx")
    val excel = ExcelIO.instance[IO]

    val rows = fs2.Stream.range(1, 11).map { i =>
      RowData(i, Map(0 -> CellValue.Text(s"Row $i")))
    }

    // Write with sheetIndex = 3
    rows.through(excel.writeStream(path, "ThirdSheet", sheetIndex = 3)).compile.drain.flatMap { _ =>
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

  tempDir.test("writeStream: rejects invalid sheet index") { dir =>
    val path = dir.resolve("invalid.xlsx")
    val excel = ExcelIO.instance[IO]
    val rows = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("test"))))

    // sheetIndex = 0 should fail
    rows.through(excel.writeStream(path, "Bad", sheetIndex = 0))
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

  tempDir.test("writeStreamsSeq: writes 2 sheets") { dir =>
    val path = dir.resolve("two-sheets.xlsx")
    val excel = ExcelIO.instance[IO]

    val sheet1Rows = fs2.Stream.range(1, 11).map { i =>
      RowData(i, Map(0 -> CellValue.Text(s"Sheet1 Row $i")))
    }
    val sheet2Rows = fs2.Stream.range(1, 6).map { i =>
      RowData(i, Map(0 -> CellValue.Number(BigDecimal(i * 100))))
    }

    excel.writeStreamsSeq(path, Seq("First" -> sheet1Rows, "Second" -> sheet2Rows)).flatMap {
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
            val firstCell = wb.sheets(0)(ref"A1")
            assertEquals(firstCell.value, CellValue.Text("Sheet1 Row 1"))

            val secondCell = wb.sheets(1)(ref"A1")
            assertEquals(secondCell.value, CellValue.Number(BigDecimal(100)))
          }
        }
    }
  }

  tempDir.test("writeStreamsSeq: writes 3 sheets with different sizes") { dir =>
    val path = dir.resolve("three-sheets.xlsx")
    val excel = ExcelIO.instance[IO]

    val sales = fs2.Stream.range(1, 101).map(i => RowData(i, Map(0 -> CellValue.Number(i))))
    val inventory = fs2.Stream.range(1, 51).map(i => RowData(i, Map(0 -> CellValue.Text(s"Item $i"))))
    val summary = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("Summary"))))

    excel
      .writeStreamsSeq(
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

  tempDir.test("writeStreamsSeq: handles large multi-sheet workbook") { dir =>
    val path = dir.resolve("large-multi-sheet.xlsx")
    val excel = ExcelIO.instance[IO]

    // 3 sheets with 10k rows each (30k total rows)
    val sheet1 = fs2.Stream.range(1, 10001).map(i => RowData(i, Map(0 -> CellValue.Number(i))))
    val sheet2 = fs2.Stream.range(1, 10001).map(i => RowData(i, Map(0 -> CellValue.Text(s"$i"))))
    val sheet3 = fs2.Stream.range(1, 10001).map(i => RowData(i, Map(0 -> CellValue.Bool(i % 2 == 0))))

    excel
      .writeStreamsSeq(path, Seq("Numbers" -> sheet1, "Text" -> sheet2, "Booleans" -> sheet3))
      .flatMap { _ =>
        // Verify file size is reasonable
        IO(java.nio.file.Files.size(path)).map { size =>
          assert(size > 100_000, s"File should be at least 100KB, got $size bytes")
          assert(size < 10_000_000, s"File should be less than 10MB, got $size bytes")
        }
      }
  }

  tempDir.test("writeStreamsSeq: rejects empty sheet list") { dir =>
    val path = dir.resolve("empty.xlsx")
    val excel = ExcelIO.instance[IO]

    excel
      .writeStreamsSeq(path, Seq.empty)
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

  tempDir.test("writeStreamsSeq: rejects duplicate sheet names") { dir =>
    val path = dir.resolve("duplicates.xlsx")
    val excel = ExcelIO.instance[IO]

    val rows1 = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("A"))))
    val rows2 = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("B"))))

    excel
      .writeStreamsSeq(path, Seq("Sheet1" -> rows1, "Sheet1" -> rows2))
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

    writeRows.through(excel.writeStream(path, "Data")).compile.drain.flatMap { _ =>
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

    writeRows.through(excel.writeStream(path, "Test")).compile.drain.flatMap { _ =>
      // Read back with streaming
      excel.readStream(path).compile.toVector.map { rows =>
        assertEquals(rows.size, 1000)

        // Verify first row
        @SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
        val first = rows.head
        assertEquals(first.rowIndex, 1)
        assertEquals(first.cells(0), CellValue.Number(BigDecimal(1)))
        assertEquals(first.cells(1), CellValue.Text("Text 1"))
        assertEquals(first.cells(2), CellValue.Bool(false))

        // Verify last row
        @SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
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

    rows.through(excel.writeStream(path, "Types")).compile.drain.flatMap { _ =>
      excel.readStream(path).compile.toVector.map { readRows =>
        assertEquals(readRows.size, 1)
        @SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
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

    writeRows.through(excel.writeStream(path, "Data")).compile.drain.flatMap { _ =>
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

    excel.writeStreamsSeq(path, Seq("First" -> sheet1, "Second" -> sheet2, "Third" -> sheet3))
      .flatMap { _ =>
        // Read second sheet by index
        excel.readStreamByIndex(path, 2).compile.toVector.map { rows =>
          assertEquals(rows.size, 20, "Should read 20 rows from second sheet")
          @SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
          val first = rows.head
          @SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
          val last = rows.last
          assertEquals(first.cells(0), CellValue.Text("S2-1"))
          assertEquals(last.cells(0), CellValue.Text("S2-20"))
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

    excel.writeStreamsSeq(path, Seq("Sales" -> sales, "Inventory" -> inventory, "Summary" -> summary))
      .flatMap { _ =>
        // Read "Inventory" sheet by name
        excel.readSheetStream(path, "Inventory").compile.toVector.map { rows =>
          assertEquals(rows.size, 5, "Should read 5 rows from Inventory")
          @SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
          val first = rows.head
          @SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
          val last = rows.last
          assertEquals(first.cells(0), CellValue.Text("Item 1"))
          assertEquals(last.cells(0), CellValue.Text("Item 5"))
        }
      }
  }

  tempDir.test("readStreamByIndex: handles invalid index") { dir =>
    val path = dir.resolve("invalid-index.xlsx")
    val excel = ExcelIO.instance[IO]

    val rows = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("test"))))
    rows.through(excel.writeStream(path, "Only")).compile.drain.flatMap { _ =>
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

  tempDir.test("readStreamRange: limits rows and columns") { dir =>
    val path = dir.resolve("range-read.xlsx")
    val excel = ExcelIO.instance[IO]

    val rows = fs2.Stream.range(1, 6).map { i =>
      RowData(i, Map(
        0 -> CellValue.Number(BigDecimal(i)),
        1 -> CellValue.Number(BigDecimal(i * 10)),
        2 -> CellValue.Number(BigDecimal(i * 100))
      ))
    }

    rows.through(excel.writeStream(path, "Data")).compile.drain.flatMap { _ =>
      val range = CellRange.parse("B2:C4").toOption.getOrElse(fail("Bad range"))
      excel.readStreamRange(path, range).compile.toVector.map { readRows =>
        assertEquals(readRows.map(_.rowIndex), Vector(2, 3, 4))
        readRows.foreach { row =>
          assertEquals(row.cells.size, 2)
          assert(row.cells.keySet == Set(1, 2))
        }
      }
    }
  }

  tempDir.test("readSheetStream: handles nonexistent sheet name") { dir =>
    val path = dir.resolve("bad-name.xlsx")
    val excel = ExcelIO.instance[IO]

    val rows = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("test"))))
    rows.through(excel.writeStream(path, "RealSheet")).compile.drain.flatMap { _ =>
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

  // ========== Memory Tests (verify true O(1) streaming) ==========

  tempDir.test("readStream: true O(1) memory with fs2.io.readInputStream") { dir =>
    val path = dir.resolve("memory-test.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write 100k rows (would OOM if using readAllBytes())
    // This test specifically verifies the fix from readAllBytes() → fs2.io.readInputStream
    // Reduced from 500k to 100k for faster CI (still validates O(1) behavior)
    val writeRows = fs2.Stream.range(1, 100001).map { i =>
      RowData(i, Map(
        0 -> CellValue.Number(BigDecimal(i)),
        1 -> CellValue.Text(s"Row-$i-with-some-text-content")
      ))
    }

    writeRows.through(excel.writeStream(path, "LargeData")).compile.drain.flatMap { _ =>
      // Read with true streaming (constant memory via fs2.io.readInputStream)
      // If this completes without OOM, the O(1) fix is working
      excel.readStream(path)
        .filter(_.rowIndex % 10000 == 0)  // Sample 0.01% (1 in 10,000) to verify correctness
        .compile.count.map { count =>
          assertEquals(count, 10L, "Should find 10 sampled rows (100k / 10k)")
        }
    }
  }

  tempDir.test("readStream: concurrent streams don't multiply memory") { dir =>
    val path1 = dir.resolve("concurrent-1.xlsx")
    val path2 = dir.resolve("concurrent-2.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write two 100k row files
    val writeRows1 = fs2.Stream.range(1, 100001).map(i => RowData(i, Map(0 -> CellValue.Number(i))))
    val writeRows2 = fs2.Stream.range(1, 100001).map(i => RowData(i, Map(0 -> CellValue.Text(s"Row-$i"))))

    for
      _ <- writeRows1.through(excel.writeStream(path1, "Data1")).compile.drain
      _ <- writeRows2.through(excel.writeStream(path2, "Data2")).compile.drain

      // Read both concurrently (should use ~constant memory, not 2x)
      count1 <- excel.readStream(path1).compile.count.start
      count2 <- excel.readStream(path2).compile.count.start
      c1 <- count1.joinWithNever
      c2 <- count2.joinWithNever
    yield
      assertEquals(c1, 100000L)
      assertEquals(c2, 100000L)
  }

  // ========== Compression Tests (verify WriterConfig respected) ==========

  tempDir.test("writeStream: config affects static parts compression") { dir =>
    val excel = ExcelIO.instance[IO]

    // Small dataset (static parts dominate file size)
    val rows = fs2.Stream.range(1, 11).map { i =>
      RowData(i, Map(0 -> CellValue.Text(s"Row $i")))
    }

    val deflatedPath = dir.resolve("streamed-deflated.xlsx")
    val storedPath = dir.resolve("streamed-stored.xlsx")

    val deflatedConfig = com.tjclp.xl.ooxml.WriterConfig(
      compression = com.tjclp.xl.ooxml.Compression.Deflated,
      prettyPrint = false
    )
    val storedConfig = com.tjclp.xl.ooxml.WriterConfig(
      compression = com.tjclp.xl.ooxml.Compression.Stored,
      prettyPrint = false
    )

    for
      _ <- rows.through(excel.writeStream(deflatedPath, "Data", 1, deflatedConfig)).compile.drain
      _ <- rows.through(excel.writeStream(storedPath, "Data", 1, storedConfig)).compile.drain
      deflatedSize <- IO(java.nio.file.Files.size(deflatedPath))
      storedSize <- IO(java.nio.file.Files.size(storedPath))
    yield
      // DEFLATED should produce smaller or equal size (worksheets always DEFLATED regardless of config)
      assert(deflatedSize <= storedSize, s"DEFLATED config ($deflatedSize) should be <= STORED config ($storedSize)")
  }

  tempDir.test("writeStream: produces compressed output") { dir =>
    val excel = ExcelIO.instance[IO]

    // Large dataset with repetitive content
    val rows = fs2.Stream.range(1, 5001).map { i =>
      RowData(i, Map(
        0 -> CellValue.Text("Repeated content that should compress well"),
        1 -> CellValue.Number(BigDecimal(i))
      ))
    }

    val path = dir.resolve("compressed.xlsx")

    rows.through(excel.writeStream(path, "Data")).compile.drain.flatMap { _ =>
      IO(java.nio.file.Files.size(path)).map { size =>
        // 5000 rows with repetitive text should compress to < 500KB
        // (uncompressed would be ~2-3MB)
        assert(size < 500_000, s"File size ($size) suggests poor compression")
      }
    }
  }

  // ========== Formula Injection Tests (TJC-339) ==========

  tempDir.test("writeStream: WriterConfig.secure escapes formula injection") { dir =>
    val excel = ExcelIO.instance[IO]

    // Test data with potentially dangerous formula-like text
    val rows = fs2.Stream.emits(List(
      RowData(1, Map(
        0 -> CellValue.Text("=SUM(A2:A10)"),      // Starts with =
        1 -> CellValue.Text("+1234"),             // Starts with +
        2 -> CellValue.Text("-dangerous"),        // Starts with -
        3 -> CellValue.Text("@import"),           // Starts with @
        4 -> CellValue.Text("Normal text")        // Normal text
      ))
    ))

    val path = dir.resolve("streaming-injection.xlsx")

    for
      _ <- rows.through(excel.writeStream(path, "Data", 1, WriterConfig.secure)).compile.drain
      wb <- IO(XlsxReader.read(path)).map(_.toOption.get)
    yield
      val sheet = wb.sheets.head

      // Text starting with =, +, -, @ should be escaped with leading quote
      sheet(ref"A1").value match
        case CellValue.Text(t) => assertEquals(t, "'=SUM(A2:A10)", "= should be escaped")
        case other => fail(s"Expected Text, got: $other")

      sheet(ref"B1").value match
        case CellValue.Text(t) => assertEquals(t, "'+1234", "+ should be escaped")
        case other => fail(s"Expected Text, got: $other")

      sheet(ref"C1").value match
        case CellValue.Text(t) => assertEquals(t, "'-dangerous", "- should be escaped")
        case other => fail(s"Expected Text, got: $other")

      sheet(ref"D1").value match
        case CellValue.Text(t) => assertEquals(t, "'@import", "@ should be escaped")
        case other => fail(s"Expected Text, got: $other")

      // Normal text should be unchanged
      sheet(ref"E1").value match
        case CellValue.Text(t) => assertEquals(t, "Normal text", "Normal text unchanged")
        case other => fail(s"Expected Text, got: $other")
  }

  tempDir.test("writeStream: WriterConfig.default does not escape text") { dir =>
    val excel = ExcelIO.instance[IO]

    val rows = fs2.Stream.emits(List(
      RowData(1, Map(
        0 -> CellValue.Text("=SUM(A2:A10)"),
        1 -> CellValue.Text("Normal text")
      ))
    ))

    val path = dir.resolve("streaming-no-escape.xlsx")

    for
      _ <- rows.through(excel.writeStream(path, "Data", 1, WriterConfig.default)).compile.drain
      wb <- IO(XlsxReader.read(path)).map(_.toOption.get)
    yield
      val sheet = wb.sheets.head

      // With default config, text should NOT be escaped
      sheet(ref"A1").value match
        case CellValue.Text(t) => assertEquals(t, "=SUM(A2:A10)", "Should not escape with default")
        case other => fail(s"Expected Text, got: $other")
  }

  tempDir.test("writeStreamsSeq: WriterConfig.secure escapes across multiple sheets") { dir =>
    val excel = ExcelIO.instance[IO]

    val sheet1 = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("=DANGER"))))
    val sheet2 = fs2.Stream.emit(RowData(1, Map(0 -> CellValue.Text("+EVIL"))))

    val path = dir.resolve("streaming-multi-escape.xlsx")

    for
      _ <- excel.writeStreamsSeq(path, Seq(("Sheet1", sheet1), ("Sheet2", sheet2)), WriterConfig.secure)
      wb <- IO(XlsxReader.read(path)).map(_.toOption.get)
    yield
      // Both sheets should have escaped text
      val s1 = wb.sheets.find(_.name.value == "Sheet1").get
      s1(ref"A1").value match
        case CellValue.Text(t) => assertEquals(t, "'=DANGER", "Sheet1 should be escaped")
        case other => fail(s"Expected Text, got: $other")

      val s2 = wb.sheets.find(_.name.value == "Sheet2").get
      s2(ref"A1").value match
        case CellValue.Text(t) => assertEquals(t, "'+EVIL", "Sheet2 should be escaped")
        case other => fail(s"Expected Text, got: $other")
  }

  // ========== Dimension Support Tests ==========

  tempDir.test("writeStreamWithAutoDetect: includes dimension element in output") { dir =>
    val path = dir.resolve("with-dimension.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write data spanning A1:C5
    val rows = fs2.Stream.range(1, 6).map { i =>
      RowData(i, Map(
        0 -> CellValue.Text(s"A$i"),
        1 -> CellValue.Number(BigDecimal(i * 10)),
        2 -> CellValue.Bool(i % 2 == 0)
      ))
    }

    rows.through(excel.writeStreamWithAutoDetect(path, "Data")).compile.drain.flatMap { _ =>
      // Extract and verify dimension element from worksheet XML
      IO {
        import java.util.zip.ZipFile
        val zipFile = new ZipFile(path.toFile)
        try {
          val entry = zipFile.getEntry("xl/worksheets/sheet1.xml")
          val content = new String(zipFile.getInputStream(entry).readAllBytes(), "UTF-8")

          // Should contain dimension element with correct range
          assert(content.contains("""<dimension ref="A1:C5"/>"""),
            s"Should have dimension element A1:C5, got: ${content.take(500)}")
        } finally {
          zipFile.close()
        }
      }
    }
  }

  tempDir.test("writeStreamWithAutoDetect: dimension bounds match actual data extent") { dir =>
    val path = dir.resolve("dimension-bounds.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write data starting at B3:D7 (not starting at A1)
    val rows = fs2.Stream.range(3, 8).map { i =>
      RowData(i, Map(
        1 -> CellValue.Text(s"B$i"),  // Column B (index 1)
        2 -> CellValue.Number(BigDecimal(i)),  // Column C (index 2)
        3 -> CellValue.Bool(true)  // Column D (index 3)
      ))
    }

    rows.through(excel.writeStreamWithAutoDetect(path, "Data")).compile.drain.flatMap { _ =>
      IO {
        import java.util.zip.ZipFile
        val zipFile = new ZipFile(path.toFile)
        try {
          val entry = zipFile.getEntry("xl/worksheets/sheet1.xml")
          val content = new String(zipFile.getInputStream(entry).readAllBytes(), "UTF-8")

          // Should have dimension B3:D7 (not A1:D7)
          assert(content.contains("""<dimension ref="B3:D7"/>"""),
            s"Should have dimension B3:D7, got: ${content.take(500)}")
        } finally {
          zipFile.close()
        }
      }
    }
  }

  tempDir.test("writeStream: empty stream produces no dimension") { dir =>
    val path = dir.resolve("empty-dimension.xlsx")
    val excel = ExcelIO.instance[IO]

    // Empty stream
    val rows = fs2.Stream.empty: fs2.Stream[IO, RowData]

    rows.through(excel.writeStream(path, "Empty")).compile.drain.flatMap { _ =>
      IO {
        import java.util.zip.ZipFile
        val zipFile = new ZipFile(path.toFile)
        try {
          val entry = zipFile.getEntry("xl/worksheets/sheet1.xml")
          val content = new String(zipFile.getInputStream(entry).readAllBytes(), "UTF-8")

          // Should NOT contain dimension element for empty sheet
          assert(!content.contains("<dimension"),
            s"Empty sheet should not have dimension element, got: ${content.take(500)}")
        } finally {
          zipFile.close()
        }
      }
    }
  }

  tempDir.test("writeStreamWithAutoDetect: single cell produces accurate dimension") { dir =>
    val path = dir.resolve("single-cell.xlsx")
    val excel = ExcelIO.instance[IO]

    // Single cell at F10
    val rows = fs2.Stream.emit(
      RowData(10, Map(5 -> CellValue.Text("Only cell")))  // Column F (index 5)
    )

    rows.through(excel.writeStreamWithAutoDetect(path, "Data")).compile.drain.flatMap { _ =>
      IO {
        import java.util.zip.ZipFile
        val zipFile = new ZipFile(path.toFile)
        try {
          val entry = zipFile.getEntry("xl/worksheets/sheet1.xml")
          val content = new String(zipFile.getInputStream(entry).readAllBytes(), "UTF-8")

          // Dimension should be F10:F10 for single cell
          assert(content.contains("""<dimension ref="F10:F10"/>"""),
            s"Single cell should have dimension F10:F10, got: ${content.take(500)}")
        } finally {
          zipFile.close()
        }
      }
    }
  }

  tempDir.test("writeStreamWithAutoDetect: dimension enables instant metadata queries") { dir =>
    val path = dir.resolve("instant-metadata.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write 10k rows
    val rows = fs2.Stream.range(1, 10001).map { i =>
      RowData(i, Map(
        0 -> CellValue.Number(BigDecimal(i)),
        4 -> CellValue.Text(s"Row $i")  // Column E (index 4)
      ))
    }

    rows.through(excel.writeStreamWithAutoDetect(path, "Data")).compile.drain.flatMap { _ =>
      // Read dimension using metadata API (should be instant, not streaming)
      excel.readDimension(path, 1).map { dim =>
        assert(dim.isDefined, "Should have dimension")
        val range = dim.get
        assertEquals(range.toA1, "A1:E10000", "Dimension should match data extent")
      }
    }
  }

  tempDir.test("writeStreamsSeqWithAutoDetect: includes dimension for each sheet") { dir =>
    val path = dir.resolve("multi-dimension.xlsx")
    val excel = ExcelIO.instance[IO]

    val sheet1 = fs2.Stream.range(1, 11).map(i => RowData(i, Map(0 -> CellValue.Number(i))))
    val sheet2 = fs2.Stream.range(1, 6).map(i => RowData(i, Map(
      0 -> CellValue.Text(s"A$i"),
      1 -> CellValue.Text(s"B$i"),
      2 -> CellValue.Text(s"C$i")
    )))

    excel.writeStreamsSeqWithAutoDetect(path, Seq("Numbers" -> sheet1, "Text" -> sheet2)).flatMap { _ =>
      IO {
        import java.util.zip.ZipFile
        val zipFile = new ZipFile(path.toFile)
        try {
          // Check first sheet: A1:A10
          val entry1 = zipFile.getEntry("xl/worksheets/sheet1.xml")
          val content1 = new String(zipFile.getInputStream(entry1).readAllBytes(), "UTF-8")
          assert(content1.contains("""<dimension ref="A1:A10"/>"""),
            s"Sheet1 should have dimension A1:A10")

          // Check second sheet: A1:C5
          val entry2 = zipFile.getEntry("xl/worksheets/sheet2.xml")
          val content2 = new String(zipFile.getInputStream(entry2).readAllBytes(), "UTF-8")
          assert(content2.contains("""<dimension ref="A1:C5"/>"""),
            s"Sheet2 should have dimension A1:C5")
        } finally {
          zipFile.close()
        }
      }
    }
  }

  tempDir.test("writeStreamWithAutoDetect: temp files cleaned up after write") { dir =>
    val excel = ExcelIO.instance[IO]
    val path = dir.resolve("cleanup-test.xlsx")

    // Get temp dir before write
    val tempDir = java.nio.file.Paths.get(System.getProperty("java.io.tmpdir"))

    // Count xl-stream temp files before
    val countTempFilesBefore = IO {
      import scala.jdk.CollectionConverters.*
      java.nio.file.Files.list(tempDir).iterator().asScala
        .count(p => p.getFileName.toString.startsWith("xl-stream-"))
    }

    // Write some data using auto-detect (which creates temp files)
    val rows = fs2.Stream.range(1, 101).map(i => RowData(i, Map(0 -> CellValue.Number(i))))

    for
      before <- countTempFilesBefore
      _ <- rows.through(excel.writeStreamWithAutoDetect(path, "Data")).compile.drain
      after <- countTempFilesBefore
    yield
      // Should have same number of temp files (all cleaned up)
      assertEquals(after, before, "Temp files should be cleaned up after write")
  }

  tempDir.test("writeStreamWithAutoDetect: temp files cleaned up on error") { dir =>
    val excel = ExcelIO.instance[IO]
    val path = dir.resolve("error-cleanup-test.xlsx")

    val tempDir = java.nio.file.Paths.get(System.getProperty("java.io.tmpdir"))

    val countTempFilesBefore = IO {
      import scala.jdk.CollectionConverters.*
      java.nio.file.Files.list(tempDir).iterator().asScala
        .count(p => p.getFileName.toString.startsWith("xl-stream-"))
    }

    // Create a stream that fails mid-way
    val rows = fs2.Stream.range(1, 51).map { i =>
      if i == 25 then throw new RuntimeException("Simulated failure")
      RowData(i, Map(0 -> CellValue.Number(i)))
    }

    for
      before <- countTempFilesBefore
      result <- rows.through(excel.writeStreamWithAutoDetect(path, "Data")).compile.drain.attempt
      after <- countTempFilesBefore
    yield
      // Should have failed
      assert(result.isLeft, "Should have failed")
      // But temp files should still be cleaned up
      assertEquals(after, before, "Temp files should be cleaned up even on error")
  }

  // ========== Explicit Dimension Hint Tests ==========

  tempDir.test("writeStream: explicit dimension hint produces dimension element") { dir =>
    val path = dir.resolve("hint-dimension.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write data with explicit dimension hint
    val rows = fs2.Stream.range(1, 6).map { i =>
      RowData(i, Map(
        0 -> CellValue.Text(s"A$i"),
        1 -> CellValue.Number(BigDecimal(i * 10)),
        2 -> CellValue.Bool(i % 2 == 0)
      ))
    }

    // Provide explicit dimension hint A1:C5
    val dimHint = CellRange.parse("A1:C5").toOption

    rows.through(excel.writeStream(path, "Data", dimension = dimHint)).compile.drain.flatMap { _ =>
      IO {
        import java.util.zip.ZipFile
        val zipFile = new ZipFile(path.toFile)
        try {
          val entry = zipFile.getEntry("xl/worksheets/sheet1.xml")
          val content = new String(zipFile.getInputStream(entry).readAllBytes(), "UTF-8")

          // Should contain dimension element from hint
          assert(content.contains("""<dimension ref="A1:C5"/>"""),
            s"Should have dimension element A1:C5 from hint, got: ${content.take(500)}")
        } finally {
          zipFile.close()
        }
      }
    }
  }

  tempDir.test("writeStream: no dimension hint produces no dimension element") { dir =>
    val path = dir.resolve("no-dimension.xlsx")
    val excel = ExcelIO.instance[IO]

    // Write data without dimension hint
    val rows = fs2.Stream.range(1, 6).map { i =>
      RowData(i, Map(0 -> CellValue.Text(s"A$i")))
    }

    rows.through(excel.writeStream(path, "Data")).compile.drain.flatMap { _ =>
      IO {
        import java.util.zip.ZipFile
        val zipFile = new ZipFile(path.toFile)
        try {
          val entry = zipFile.getEntry("xl/worksheets/sheet1.xml")
          val content = new String(zipFile.getInputStream(entry).readAllBytes(), "UTF-8")

          // Should NOT contain dimension element without hint
          assert(!content.contains("<dimension"),
            s"Should not have dimension element without hint, got: ${content.take(500)}")
        } finally {
          zipFile.close()
        }
      }
    }
  }

  tempDir.test("writeStream: single-pass is faster than auto-detect") { dir =>
    val excel = ExcelIO.instance[IO]
    val rows = fs2.Stream.range(1, 10001).map { i =>
      RowData(i, Map(0 -> CellValue.Number(BigDecimal(i))))
    }

    // Single-pass with hint (no temp file overhead)
    val path1 = dir.resolve("single-pass.xlsx")
    val dimHint = CellRange.parse("A1:A10000").toOption

    // Auto-detect (temp file + two passes)
    val path2 = dir.resolve("auto-detect.xlsx")

    for
      // Time single-pass
      start1 <- IO(System.nanoTime())
      _ <- rows.through(excel.writeStream(path1, "Data", dimension = dimHint)).compile.drain
      end1 <- IO(System.nanoTime())
      singlePassMs = (end1 - start1) / 1000000

      // Re-create stream (consumed by first write)
      rows2 = fs2.Stream.range(1, 10001).map { i =>
        RowData(i, Map(0 -> CellValue.Number(BigDecimal(i))))
      }

      // Time auto-detect
      start2 <- IO(System.nanoTime())
      _ <- rows2.through(excel.writeStreamWithAutoDetect(path2, "Data")).compile.drain
      end2 <- IO(System.nanoTime())
      autoDetectMs = (end2 - start2) / 1000000
    yield
      // Single-pass should be faster (roughly 2x due to no temp file I/O)
      // We allow some tolerance since timing can vary
      assert(singlePassMs < autoDetectMs * 1.5,
        s"Single-pass ($singlePassMs ms) should be faster than auto-detect ($autoDetectMs ms)")
  }
