#!/usr/bin/env -S scala-cli shebang
//> using file project.scala
//> using javaOpt -Xmx512m

import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.addressing.CellRange
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import fs2.Stream
import java.nio.file.{Files, Path}

val excel = ExcelIO.instance[IO]

case class BenchResult(
  mode: String,
  rows: Int,
  timeMs: Long,
  fileSizeMB: Double,
  hasDimension: Boolean
)

def generateRows(count: Int): Stream[IO, RowData] =
  Stream.range(1, count + 1).map { i =>
    RowData(
      rowIndex = i,
      cells = Map(
        0 -> CellValue.Text(s"Row-$i"),
        1 -> CellValue.Number(BigDecimal(i * 100)),
        2 -> CellValue.Text(s"data-${i % 1000}"),
        3 -> CellValue.Number(BigDecimal(i * 0.01)),
        4 -> CellValue.Bool(i % 2 == 0)
      )
    )
  }

def benchmarkSinglePass(rows: Int, path: Path): BenchResult =
  val dimension = CellRange.parse(s"A1:E$rows").toOption
  val start = System.nanoTime()

  generateRows(rows)
    .through(excel.writeStream(path, "Data", dimension = dimension))
    .compile
    .drain
    .unsafeRunSync()

  val elapsed = (System.nanoTime() - start) / 1_000_000
  val size = Files.size(path) / 1024.0 / 1024.0
  val hasDim = excel.readDimension(path, 1).unsafeRunSync().isDefined

  BenchResult("single-pass", rows, elapsed, size, hasDim)

def benchmarkTwoPass(rows: Int, path: Path): BenchResult =
  val start = System.nanoTime()

  generateRows(rows)
    .through(excel.writeStreamWithAutoDetect(path, "Data"))
    .compile
    .drain
    .unsafeRunSync()

  val elapsed = (System.nanoTime() - start) / 1_000_000
  val size = Files.size(path) / 1024.0 / 1024.0
  val hasDim = excel.readDimension(path, 1).unsafeRunSync().isDefined

  BenchResult("two-pass", rows, elapsed, size, hasDim)

def benchmarkNoDimension(rows: Int, path: Path): BenchResult =
  val start = System.nanoTime()

  generateRows(rows)
    .through(excel.writeStream(path, "Data"))  // No dimension hint
    .compile
    .drain
    .unsafeRunSync()

  val elapsed = (System.nanoTime() - start) / 1_000_000
  val size = Files.size(path) / 1024.0 / 1024.0
  val hasDim = excel.readDimension(path, 1).unsafeRunSync().isDefined

  BenchResult("no-dimension", rows, elapsed, size, hasDim)

println("=" * 80)
println("XL Streaming Write Benchmark: Single-Pass vs Two-Pass")
println("=" * 80)
println()
println("Testing: 5 columns (text, number, text, number, boolean)")
println("JVM heap: 512 MB")
println()

val rowCounts = List(10_000, 50_000, 100_000, 500_000)
val tempDir = Files.createTempDirectory("xl-bench-")

println(f"${"Rows"}%10s | ${"Mode"}%-14s | ${"Time (ms)"}%10s | ${"Size (MB)"}%10s | ${"Dimension"}%10s | ${"Rows/sec"}%12s")
println("-" * 80)

val results = rowCounts.flatMap { rowCount =>
  val path1 = tempDir.resolve(s"single-$rowCount.xlsx")
  val path2 = tempDir.resolve(s"twopass-$rowCount.xlsx")
  val path3 = tempDir.resolve(s"nodim-$rowCount.xlsx")

  // Warmup for first iteration
  if rowCount == rowCounts.head then
    generateRows(1000).through(excel.writeStream(tempDir.resolve("warmup.xlsx"), "W")).compile.drain.unsafeRunSync()

  val r1 = benchmarkSinglePass(rowCount, path1)
  val r2 = benchmarkTwoPass(rowCount, path2)
  val r3 = benchmarkNoDimension(rowCount, path3)

  List(r1, r2, r3).foreach { r =>
    val throughput = r.rows * 1000.0 / r.timeMs
    val dimStr = if r.hasDimension then "✓" else "✗"
    println(f"${r.rows}%,10d | ${r.mode}%-14s | ${r.timeMs}%,10d | ${r.fileSizeMB}%10.2f | ${dimStr}%10s | ${throughput}%,12.0f")
  }
  println("-" * 80)

  List(r1, r2, r3)
}

// Summary
println()
println("SUMMARY")
println("=" * 80)

val grouped = results.groupBy(_.rows)
grouped.toSeq.sortBy(_._1).foreach { case (rows, rs) =>
  val single = rs.find(_.mode == "single-pass").get
  val twoPass = rs.find(_.mode == "two-pass").get
  val noDim = rs.find(_.mode == "no-dimension").get

  val overhead = ((twoPass.timeMs - single.timeMs).toDouble / single.timeMs * 100)
  val noDimVsSingle = ((single.timeMs - noDim.timeMs).toDouble / noDim.timeMs * 100)

  println(f"$rows%,d rows:")
  println(f"  - Single-pass (with hint):  ${single.timeMs}%,d ms")
  println(f"  - Two-pass (auto-detect):   ${twoPass.timeMs}%,d ms (+${overhead}%.1f%% overhead)")
  println(f"  - No dimension:             ${noDim.timeMs}%,d ms")
  println()
}

// Cleanup
Files.walk(tempDir).sorted(java.util.Comparator.reverseOrder()).forEach(Files.delete)

println("Conclusion:")
println("  - Single-pass with hint ≈ No-dimension (minimal overhead for writing dimension)")
println("  - Two-pass adds ~2x I/O (temp file write + final assembly)")
println("  - Use hints when bounds known, auto-detect when unknown")
