#!/usr/bin/env -S scala-cli shebang
//> using file project.scala
//> using javaOpt -Xmx256m

import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.addressing.CellRange
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import fs2.Stream
import java.nio.file.{Files, Path}

val excel = ExcelIO.instance[IO]
val rowCount = 1_000_000
val colCount = 5  // Columns A-E (0-4)
val path = Path.of("/tmp/one-million-rows.xlsx")

// Dimension hint enables instant metadata queries (optional but recommended for known bounds)
// Format: CellRange(startRef, endRef) where refs are 0-indexed (col, row)
val dimension = CellRange.parse(s"A1:E$rowCount").toOption

println("=" * 70)
println("XL Streaming Demo: 1 MILLION ROWS with only 256MB heap")
println("=" * 70)
println()
println(s"Writing $rowCount rows × $colCount columns = ${rowCount * colCount} cells")
println(s"Dimension hint: ${dimension.map(_.toA1).getOrElse("none")} (enables instant bounds queries)")
println()

val startTime = System.currentTimeMillis()

Stream.range(1, rowCount + 1)
  .map { i =>
    RowData(
      rowIndex = i,
      cells = Map(
        0 -> CellValue.Text(s"Employee-$i"),
        1 -> CellValue.Text(s"dept-${i % 50}@company.com"),
        2 -> CellValue.Number(BigDecimal(50000 + (i % 100000))),
        3 -> CellValue.Number(BigDecimal(i * 0.05)),
        4 -> CellValue.Text(if i % 2 == 0 then "Active" else "Pending")
      )
    )
  }
  .evalTap { row =>
    IO {
      if row.rowIndex % 250000 == 0 then
        val now = System.currentTimeMillis()
        val elapsed = (now - startTime) / 1000.0
        val rate = row.rowIndex / elapsed
        println(f"   ${row.rowIndex}%,d rows written (${rate}%,.0f rows/sec)")
    }
  }
  .through(excel.writeStream(path, "Employees", dimension = dimension))
  .compile
  .drain
  .unsafeRunSync()

val totalTime = (System.currentTimeMillis() - startTime) / 1000.0
val fileSize = Files.size(path) / 1024.0 / 1024.0
val throughput = rowCount / totalTime

// Verify dimension element was written (instant metadata query)
val readDimension = excel.readDimension(path, 1).unsafeRunSync()

println()
println("=" * 70)
println("RESULTS")
println("=" * 70)
println(f"   Rows written:  $rowCount%,d")
println(f"   Total cells:   ${rowCount * colCount}%,d")
println(f"   Time:          $totalTime%.1f seconds")
println(f"   Throughput:    $throughput%,.0f rows/sec")
println(f"   File size:     $fileSize%.1f MB")
println(f"   JVM heap:      256 MB (constant)")
println(s"   Dimension:     ${readDimension.map(_.toA1).getOrElse("(not set)")} (instant)")
println()
println(s"   Output: $path")
println("=" * 70)
