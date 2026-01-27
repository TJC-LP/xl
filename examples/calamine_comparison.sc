#!/usr/bin/env -S scala-cli shebang
//> using file project.scala
//> using javaOpt -Xmx512m

import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.{ExcelIO, RowData}
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import java.nio.file.{Files, Path}

val excel = ExcelIO.instance[IO]
val inputPath = Path.of(sys.env.getOrElse("HOME", "/Users/rcaputo3") + "/Downloads/NYC_311_SR_2010-2020-sample-1M.xlsx")

println("=" * 70)
println("XL vs Calamine Comparison (Read-Only)")
println("=" * 70)
println()
println("Dataset: NYC 311 Service Requests 2010-2020 (1M rows, 41 cols)")
println()

// Calamine benchmark results (from their README):
println("Calamine (Rust) benchmarks from README:")
println("  Time:       25.278 s ± 0.424 s")
println("  Cells/sec:  1,122,279")
println()

// Our read-only benchmark
println("XL (Scala) streaming read benchmark...")
val startTime = System.currentTimeMillis()

var rowCount = 0L
var cellCount = 0L

excel.readStream(inputPath)
  .evalMap { row =>
    IO {
      rowCount += 1
      cellCount += row.cells.size
      if rowCount % 200000 == 0 then
        val elapsed = (System.currentTimeMillis() - startTime) / 1000.0
        val cellsPerSec = cellCount / elapsed
        println(f"  $rowCount%,d rows, $cellCount%,d cells (${cellsPerSec}%,.0f cells/sec)")
    }
  }
  .compile
  .drain
  .unsafeRunSync()

val totalTime = (System.currentTimeMillis() - startTime) / 1000.0
val cellsPerSec = cellCount / totalTime
val rowsPerSec = rowCount / totalTime

println()
println("=" * 70)
println("RESULTS")
println("=" * 70)
println(f"  Rows:       $rowCount%,d")
println(f"  Cells:      $cellCount%,d")
println(f"  Time:       $totalTime%.2f s")
println(f"  Cells/sec:  ${cellsPerSec}%,.0f")
println(f"  Rows/sec:   ${rowsPerSec}%,.0f")
println()
println("COMPARISON:")
println("=" * 70)
println(f"  Calamine (Rust):  25.28 s  |  1,122,279 cells/sec")
println(f"  XL (Scala/JVM):   $totalTime%.2f s  |  ${cellsPerSec}%,.0f cells/sec")
val ratio = 25.28 / totalTime
val cellRatio = cellsPerSec / 1122279.0
println()
if ratio > 1 then
  println(f"  XL is ${ratio}%.2fx FASTER than Calamine")
else
  println(f"  Calamine is ${1/ratio}%.2fx faster than XL")
println()
println("Note: XL also supports WRITING (calamine is read-only)")
println("=" * 70)
