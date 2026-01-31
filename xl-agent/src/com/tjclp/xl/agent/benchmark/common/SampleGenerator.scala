package com.tjclp.xl.agent.benchmark.common

import cats.effect.IO
import com.tjclp.xl.{*, given}
import com.tjclp.xl.dsl.*

import java.nio.file.{Files, Path}

/** Generates sample Excel files for benchmarks */
object SampleGenerator:

  /** Default location for generated sample file */
  val DefaultSamplePath: Path = Path.of("examples/benchmark/sample.xlsx")

  /** Ensure sample file exists, creating it if necessary */
  def ensureSampleExists(path: Path = DefaultSamplePath): IO[Path] =
    IO.blocking(Files.exists(path)).flatMap {
      case true => IO.pure(path)
      case false =>
        for
          _ <- IO.blocking(Files.createDirectories(path.getParent))
          _ <- createSample(path)
        yield path
    }

  /** Create the standard benchmark sample file */
  def createSample(outputPath: Path): IO[Unit] =
    IO.blocking {
      val workbook = Workbook(salesSheet, summarySheet)
      Excel.write(workbook, outputPath.toString)
    }

  /** Sales data sheet with quarterly revenue, formulas, and styling */
  private def salesSheet: Sheet =
    Sheet("Sales")
      .put(
        "A1" -> "Product",
        "B1" -> "Q1",
        "C1" -> "Q2",
        "D1" -> "Q3",
        "E1" -> "Q4",
        "F1" -> "Total"
      )
      .put(
        "A2" -> "Widget A",
        "B2" -> 12500.0,
        "C2" -> 14200.0,
        "D2" -> 13800.0,
        "E2" -> 15600.0
      )
      .put(
        "A3" -> "Widget B",
        "B3" -> 8900.0,
        "C3" -> 9100.0,
        "D3" -> 8700.0,
        "E3" -> 9500.0
      )
      .put(
        "A4" -> "Widget C",
        "B4" -> 21000.0,
        "C4" -> 19500.0,
        "D4" -> 22100.0,
        "E4" -> 23400.0
      )
      .put("A5" -> "Total", "A6" -> "Average")
      // Row totals
      .put(ref"F2", fx"=SUM(B2:E2)")
      .put(ref"F3", fx"=SUM(B3:E3)")
      .put(ref"F4", fx"=SUM(B4:E4)")
      // Column totals
      .put(ref"B5", fx"=SUM(B2:B4)")
      .put(ref"C5", fx"=SUM(C2:C4)")
      .put(ref"D5", fx"=SUM(D2:D4)")
      .put(ref"E5", fx"=SUM(E2:E4)")
      .put(ref"F5", fx"=SUM(F2:F4)")
      // Averages
      .put(ref"B6", fx"=AVERAGE(B2:B4)")
      .put(ref"C6", fx"=AVERAGE(C2:C4)")
      .put(ref"D6", fx"=AVERAGE(D2:D4)")
      .put(ref"E6", fx"=AVERAGE(E2:E4)")
      .put(ref"F6", fx"=AVERAGE(F2:F4)")
      // Styling - header row (bold, blue background, white text)
      .style("A1:F1", CellStyle.default.bold.bgBlue.white)
      // Currency format for data
      .style("B2:F6", CellStyle.default.currency)
      // Bold totals row
      .style("A5:F5", CellStyle.default.bold)

  /** Summary sheet with cross-sheet references */
  private def summarySheet: Sheet =
    Sheet("Summary")
      .put("A1" -> "Sales Summary", "A3" -> "Metric", "B3" -> "Value")
      .put("A4" -> "Total Revenue", "A5" -> "Average per Product", "A6" -> "Best Quarter")
      .put(ref"B4", fx"=Sales!F5")
      .put(ref"B5", fx"=Sales!F6")
      .put(ref"B6", fx"=MAX(Sales!B5:E5)")
      // Styling
      .style("A1", CellStyle.default.bold.size(14.0))
      .style("A3:B3", CellStyle.default.bold.bgBlue.white)
      .style("B4:B6", CellStyle.default.currency)
