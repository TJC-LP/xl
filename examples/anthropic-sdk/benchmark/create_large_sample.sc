#!/usr/bin/env -S scala-cli shebang
//> using file ../../project.scala

// Creates large_sample.xlsx for benchmarking with larger datasets
// Uses streaming API for efficient 1000-row generation

import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.addressing.CellRange
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import fs2.Stream
import java.nio.file.Path

import scala.util.Random

val seed = 42
val random = new Random(seed)
val rowCount = 1000
val products = Vector("Widget A", "Widget B", "Widget C", "Gadget X", "Gadget Y", "Tool Z")
val categories = Vector("Electronics", "Hardware", "Software", "Services")

val excel = ExcelIO.instance[IO]
val outputPath = Path.of("examples/anthropic-sdk/benchmark/large_sample.xlsx")
val dimension = CellRange.parse(s"A1:G${rowCount + 1}").toOption

println(s"Creating large_sample.xlsx with $rowCount rows...")

// Header row (row 0 = Excel row 1)
val headerRow = RowData(
  rowIndex = 0,
  cells = Map(
    0 -> CellValue.Text("ID"),
    1 -> CellValue.Text("Product"),
    2 -> CellValue.Text("Category"),
    3 -> CellValue.Text("Price"),
    4 -> CellValue.Text("Quantity"),
    5 -> CellValue.Text("Revenue"),
    6 -> CellValue.Text("Date")
  )
)

// Data rows (rows 1-1000 = Excel rows 2-1001)
val dataRows = Stream.range(1, rowCount + 1).map { i =>
  val product = products(random.nextInt(products.length))
  val category = categories(random.nextInt(categories.length))
  val price = BigDecimal(10.0 + random.nextDouble() * 990.0).setScale(2, BigDecimal.RoundingMode.HALF_UP)
  val quantity = 1 + random.nextInt(100)
  val excelRow = i + 1  // 1-indexed for formula reference

  RowData(
    rowIndex = i,
    cells = Map(
      0 -> CellValue.Number(BigDecimal(i)),
      1 -> CellValue.Text(s"Row $i - $product"),
      2 -> CellValue.Text(category),
      3 -> CellValue.Number(price),
      4 -> CellValue.Number(BigDecimal(quantity)),
      5 -> CellValue.Formula(s"=D$excelRow*E$excelRow"),
      6 -> CellValue.Text(java.time.LocalDate.of(2024, 1 + (i % 12), 1 + (i % 28)).toString)
    )
  )
}

// Combine header + data and write
(Stream.emit(headerRow) ++ dataRows)
  .through(excel.writeStream(outputPath, "Data", dimension = dimension))
  .compile
  .drain
  .unsafeRunSync()

println(s"Created $outputPath with $rowCount rows of data")
println("Note: Summary sheet with cross-sheet formulas requires non-streaming write")
