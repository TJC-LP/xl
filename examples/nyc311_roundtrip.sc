#!/usr/bin/env -S scala-cli shebang
//> using file project.scala
//> using javaOpt -Xmx512m

import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.addressing.CellRange
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import java.nio.file.{Files, Path}

val excel = ExcelIO.instance[IO]
val inputPath = Path.of(sys.env.getOrElse("HOME", "/Users/rcaputo3") + "/Downloads/NYC_311_SR_2010-2020-sample-1M.xlsx")
val outputPath = Path.of("/tmp/nyc311-roundtrip.xlsx")

println("=" * 70)
println("NYC 311 Round-Trip Benchmark")
println("=" * 70)
println()

// Check input file
val inputSize = Files.size(inputPath) / 1024.0 / 1024.0
println(f"Input:  $inputPath")
println(f"Size:   $inputSize%.1f MB")

// Read dimension from source (instant)
val sourceDim = excel.readDimension(inputPath, 1).unsafeRunSync()
println(s"Source dimension: ${sourceDim.map(_.toA1).getOrElse("none")}")
println()

// Round-trip: stream read → stream write with auto-detect
println("Starting round-trip (streaming read → streaming write with auto-detect)...")
val startTime = System.currentTimeMillis()

var rowCount = 0L
excel.readStream(inputPath)
  .evalTap { row =>
    IO {
      rowCount += 1
      if rowCount % 200000 == 0 then
        val elapsed = (System.currentTimeMillis() - startTime) / 1000.0
        println(f"  $rowCount%,d rows processed (${elapsed}%.1f s)")
    }
  }
  .through(excel.writeStreamWithAutoDetect(outputPath, "NYC_311_SR_2010-2020-sample-1M"))
  .compile
  .drain
  .unsafeRunSync()

val totalTime = (System.currentTimeMillis() - startTime) / 1000.0
val outputSize = Files.size(outputPath) / 1024.0 / 1024.0
val outputDim = excel.readDimension(outputPath, 1).unsafeRunSync()

println()
println("=" * 70)
println("RESULTS")
println("=" * 70)
println(f"  Rows processed: $rowCount%,d")
println(f"  Time:           $totalTime%.1f seconds")
println(f"  Throughput:     ${rowCount / totalTime}%,.0f rows/sec")
println(f"  Input size:     $inputSize%.1f MB")
println(f"  Output size:    $outputSize%.1f MB")
println(f"  Output dim:     ${outputDim.map(_.toA1).getOrElse("none")}")
println(f"  JVM heap:       512 MB (constant)")
println()
println(s"  Output: $outputPath")
println("=" * 70)
