#!/usr/bin/env -S scala-cli shebang
//> using file project.scala

import com.tjclp.xl.io.ExcelIO
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import java.nio.file.Path

val excel = ExcelIO.instance[IO]
val path = Path.of("/tmp/one-million-rows.xlsx")

println("Reading first 5 rows with streaming...")
excel.readStream(path)
  .take(5)
  .evalMap { row =>
    IO.println(s"Row ${row.rowIndex}: ${row.cells.values.take(3).mkString(", ")}")
  }
  .compile
  .drain
  .unsafeRunSync()

println("\nCounting total rows with streaming...")
val count = excel.readStream(path).compile.count.unsafeRunSync()
println(s"Total rows: $count")
