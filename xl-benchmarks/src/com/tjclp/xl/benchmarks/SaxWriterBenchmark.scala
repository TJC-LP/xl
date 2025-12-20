package com.tjclp.xl.benchmarks

import cats.effect.IO
import cats.effect.unsafe.implicits.global
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.{WriterConfig, XmlBackend}
import com.tjclp.xl.workbooks.Workbook
import org.openjdk.jmh.annotations.*
import java.nio.file.{Files, Path}
import java.util.concurrent.TimeUnit
import scala.compiletime.uninitialized

/**
 * Benchmarks comparing XL write backends: ScalaXml vs SaxStax.
 *
 * Goal: Quantify performance difference between backends to validate optimization claims and
 * identify bottlenecks in the write path.
 *
 * Expected: SaxStax should be faster due to streaming output, but both currently build intermediate
 * Elem trees via toXml() before serialization.
 *
 * Run with: ./mill xl-benchmarks.runJmh -- SaxWriterBenchmark
 */
@State(Scope.Benchmark)
@BenchmarkMode(Array(Mode.AverageTime))
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1, timeUnit = TimeUnit.SECONDS)
@Measurement(iterations = 5, time = 2, timeUnit = TimeUnit.SECONDS)
@Fork(value = 1)
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.Null"))
class SaxWriterBenchmark {

  @Param(Array("1000", "10000", "100000"))
  var rows: Int = uninitialized

  var workbook: Workbook = uninitialized
  var scalaXmlWriteFile: Path = uninitialized
  var saxStaxWriteFile: Path = uninitialized

  val excel: ExcelIO[IO] = ExcelIO.instance[IO]

  // Pre-configured WriterConfigs
  val scalaXmlConfig: WriterConfig = WriterConfig.default.copy(backend = XmlBackend.ScalaXml)
  val saxStaxConfig: WriterConfig = WriterConfig.default.copy(backend = XmlBackend.SaxStax)

  @Setup(Level.Trial)
  def setup(): Unit = {
    // Generate workbook once per trial (shared across iterations)
    // Use verifiable workbook for consistency with other benchmarks
    workbook = BenchmarkUtils.createVerifiableWorkbook(rows)

    // Create dedicated temp files for each backend
    scalaXmlWriteFile = Files.createTempFile("xl-scalaxml-write-", ".xlsx")
    saxStaxWriteFile = Files.createTempFile("xl-saxstax-write-", ".xlsx")
  }

  @TearDown(Level.Trial)
  def teardown(): Unit = {
    Seq(scalaXmlWriteFile, saxStaxWriteFile).foreach { path =>
      if (path != null && Files.exists(path)) Files.delete(path)
    }
  }

  // ===== ScalaXml Backend (default) =====

  @Benchmark
  def writeScalaXml(): Unit = {
    // Write using ScalaXml backend (builds Elem tree, serializes to string)
    excel.writeWith(workbook, scalaXmlWriteFile, scalaXmlConfig).unsafeRunSync()
  }

  // ===== SaxStax Backend =====

  @Benchmark
  def writeSaxStax(): Unit = {
    // Write using SaxStax backend (builds Elem tree, converts to SAX events)
    excel.writeWith(workbook, saxStaxWriteFile, saxStaxConfig).unsafeRunSync()
  }

  // ===== writeFast() convenience method =====

  @Benchmark
  def writeFast(): Unit = {
    // Use the convenience method (internally uses SaxStax)
    excel.writeFast(workbook, saxStaxWriteFile).unsafeRunSync()
  }
}
