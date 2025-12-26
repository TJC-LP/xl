package com.tjclp.xl.benchmarks

import cats.effect.IO
import cats.effect.unsafe.implicits.global
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.{OoxmlWorksheet, SharedStrings, WriterConfig, XmlBackend, XmlUtil}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import org.openjdk.jmh.annotations.*
import java.nio.charset.StandardCharsets.UTF_8
import java.nio.file.{Files, Path}
import java.util.concurrent.TimeUnit
import scala.compiletime.uninitialized
import scala.xml.Elem

/**
 * Benchmarks comparing XL write backends: ScalaXml vs SaxStax.
 *
 * Goal: Quantify performance difference between backends to validate optimization claims and
 * identify bottlenecks in the write path.
 *
 * SaxStax is the default backend since 0.5.0 due to:
 *   - Direct streaming to output (no intermediate String allocation)
 *   - Better memory efficiency for large workbooks
 *   - Surgical modification support (modify existing OOXML in place)
 *
 * Phase 3 isolation benchmarks: Separate toXml() from serialization to identify true bottleneck.
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

  // For isolation benchmarks
  var sheet: Sheet = uninitialized
  var preBuiltElem: Elem = uninitialized

  val excel: ExcelIO[IO] = ExcelIO.instance[IO]

  // Pre-configured WriterConfigs
  val scalaXmlConfig: WriterConfig = WriterConfig.default.copy(backend = XmlBackend.ScalaXml)
  val saxStaxConfig: WriterConfig = WriterConfig.default.copy(backend = XmlBackend.SaxStax)

  @Setup(Level.Trial)
  def setup(): Unit = {
    // Generate workbook once per trial (shared across iterations)
    // Use verifiable workbook for consistency with other benchmarks
    workbook = BenchmarkUtils.createVerifiableWorkbook(rows)

    // Extract sheet for isolation benchmarks
    sheet = workbook.sheets.find(_.name.value == "Data").get

    // Pre-build Elem for serialization-only benchmark
    val ooxmlWorksheet = OoxmlWorksheet.fromDomainWithSST(sheet, None)
    preBuiltElem = ooxmlWorksheet.toXml

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

  // ===== SaxStax Backend (default since 0.5.0) =====

  @Benchmark
  def writeSaxStax(): Unit = {
    // Write using SaxStax backend (default since 0.5.0)
    excel.writeWith(workbook, saxStaxWriteFile, saxStaxConfig).unsafeRunSync()
  }

  // ===== Isolation Benchmarks (Phase 3) =====
  // Separate toXml() from serialization to identify true bottleneck

  @Benchmark
  def toXmlOnly(): Elem = {
    // Measure ONLY Elem tree construction from domain model
    // This is the suspected bottleneck based on Phase 2 results
    val ooxmlWorksheet = OoxmlWorksheet.fromDomainWithSST(sheet, None)
    ooxmlWorksheet.toXml
  }

  @Benchmark
  def serializeOnly(): Array[Byte] = {
    // Measure ONLY XML serialization (Elem pre-built in setup)
    // Uses same serialization path as ScalaXml backend
    XmlUtil.compact(preBuiltElem).getBytes(UTF_8)
  }
}
