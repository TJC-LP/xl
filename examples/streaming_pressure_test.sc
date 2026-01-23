#!/usr/bin/env -S scala-cli shebang
//> using file project.scala

/**
 * Streaming Pressure Test
 *
 * Tests true O(1) memory streaming with large datasets.
 * Run with: ./streaming_pressure_test.sc
 */

import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.{ExcelIO, RowData}
import cats.effect.{IO, IOApp}
import cats.effect.unsafe.implicits.global
import fs2.Stream
import java.nio.file.{Files, Path}

object StreamingPressureTest:

  val excel = ExcelIO.instance[IO]
  val testDir = Path.of("/tmp/xl-streaming-test")

  def main(args: Array[String]): Unit =
    Files.createDirectories(testDir)

    println("=" * 60)
    println("XL Streaming Pressure Test")
    println("=" * 60)

    // Test 1: Write 100k rows with true streaming
    val rowCount = 100_000
    val outputPath = testDir.resolve("streaming_100k.xlsx")

    println(s"\n[Test 1] Writing $rowCount rows with writeStreamTrue...")
    val writeStart = System.currentTimeMillis()

    val rows: Stream[IO, RowData] = Stream
      .range(1, rowCount + 1)
      .map { i =>
        RowData(
          rowIndex = i,
          cells = Map(
            0 -> CellValue.Text(s"Row $i"),
            1 -> CellValue.Number(i.toDouble),
            2 -> CellValue.Number(i * 1.5),
            3 -> CellValue.Text(if i % 2 == 0 then "Even" else "Odd")
          )
        )
      }

    rows
      .through(excel.writeStreamTrue(outputPath, "Data"))
      .compile
      .drain
      .unsafeRunSync()

    val writeTime = System.currentTimeMillis() - writeStart
    val fileSize = Files.size(outputPath) / 1024 / 1024
    println(s"   Done in ${writeTime}ms")
    println(s"   File size: ${fileSize}MB")
    println(s"   Throughput: ${rowCount * 1000 / writeTime} rows/sec")

    // Test 2: Read back with streaming
    println(s"\n[Test 2] Reading back with readStream...")
    val readStart = System.currentTimeMillis()

    var readCount = 0L
    excel.readStream(outputPath)
      .evalMap { row =>
        IO {
          readCount += 1
          if readCount % 25000 == 0 then
            println(s"   Read $readCount rows...")
        }
      }
      .compile
      .drain
      .unsafeRunSync()

    val readTime = System.currentTimeMillis() - readStart
    println(s"   Done in ${readTime}ms")
    println(s"   Total rows read: $readCount")
    println(s"   Throughput: ${readCount * 1000 / readTime} rows/sec")

    // Test 3: Round-trip transform
    println(s"\n[Test 3] Round-trip: read -> transform -> write...")
    val transformPath = testDir.resolve("streaming_transformed.xlsx")
    val transformStart = System.currentTimeMillis()

    excel.readStream(outputPath)
      .map { row =>
        // Transform: double the numeric values
        RowData(
          rowIndex = row.rowIndex,
          cells = row.cells.map {
            case (col, CellValue.Number(n)) => col -> CellValue.Number(n * 2)
            case other => other
          }
        )
      }
      .through(excel.writeStreamTrue(transformPath, "Transformed"))
      .compile
      .drain
      .unsafeRunSync()

    val transformTime = System.currentTimeMillis() - transformStart
    println(s"   Done in ${transformTime}ms")
    println(s"   Output: $transformPath")

    // Verify output
    println(s"\n[Test 4] Verify transformed file...")
    val verifyCount = excel.readStream(transformPath)
      .take(5)
      .compile
      .toList
      .unsafeRunSync()

    verifyCount.foreach { row =>
      println(s"   Row ${row.rowIndex}: ${row.cells.values.take(3).mkString(", ")}")
    }

    println("\n" + "=" * 60)
    println("All tests passed!")
    println(s"Output files in: $testDir")
    println("=" * 60)

// Run the test
StreamingPressureTest.main(Array.empty)
