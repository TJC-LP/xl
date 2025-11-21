package com.tjclp.xl.benchmarks

import cats.effect.IO
import cats.effect.unsafe.implicits.global
import com.tjclp.xl.io.{Excel, ExcelIO}
import com.tjclp.xl.workbooks.Workbook
import org.openjdk.jmh.annotations.*
import java.util.concurrent.TimeUnit
import java.nio.file.{Files, Path}
import scala.compiletime.uninitialized

/**
 * Benchmarks for Excel file read/write operations.
 *
 * Tests throughput for various row counts (1k, 10k, 100k) to validate performance claims.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Array(Mode.AverageTime))
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1, timeUnit = TimeUnit.SECONDS)
@Measurement(iterations = 5, time = 2, timeUnit = TimeUnit.SECONDS)
@Fork(value = 1)
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.Null"))
class ReadWriteBenchmark {

  @Param(Array("1000", "10000", "100000"))
  var rows: Int = uninitialized

  var workbook: Workbook = uninitialized
  var testFilePath: Path = uninitialized
  var excel: Excel[IO] = uninitialized

  @Setup(Level.Trial)
  def setup(): Unit = {
    // Generate workbook once per trial (shared across iterations)
    workbook = BenchmarkUtils.generateWorkbook(rows, styled = false)

    // Create temp file for write/read tests
    testFilePath = Files.createTempFile("xl-benchmark-", ".xlsx")

    // Initialize Excel IO
    excel = ExcelIO.instance[IO]
  }

  @TearDown(Level.Trial)
  def teardown(): Unit = {
    // Clean up temp file
    if (testFilePath != null && Files.exists(testFilePath)) {
      Files.delete(testFilePath)
    }
  }

  @Benchmark
  def writeWorkbook(): Unit = {
    // Measure write throughput
    excel.write(workbook, testFilePath).unsafeRunSync()
  }

  @Benchmark
  def readWorkbook(): Workbook = {
    // First ensure file exists (setup writes it)
    excel.write(workbook, testFilePath).unsafeRunSync()

    // Measure read throughput
    excel.read(testFilePath).unsafeRunSync()
  }

  @Benchmark
  def roundTrip(): Workbook = {
    // Measure full round-trip: write then read
    excel
      .write(workbook, testFilePath)
      .flatMap { _ =>
        excel.read(testFilePath)
      }
      .unsafeRunSync()
  }
}
