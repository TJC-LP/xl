package com.tjclp.xl.benchmarks

import com.tjclp.xl.*
import com.tjclp.xl.addressing.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.styles.{CellStyle, Color, Font, Fill, NumFmt}
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.patch.Patch
import java.time.LocalDate
import scala.util.Random

/** Shared utilities for benchmark test data generation */
object BenchmarkUtils {

  /** Generate a workbook with N rows of mixed data types */
  def generateWorkbook(rows: Int, styled: Boolean = false): Workbook = {
    val sheet = generateSheet("Data", rows, styled)
    Workbook.empty
      .flatMap(_.put(sheet))
      .getOrElse(Workbook(Vector.empty)) // Fallback to empty workbook
  }

  /** Generate a sheet with N rows of realistic data */
  def generateSheet(name: String, rows: Int, styled: Boolean = false): Sheet = {
    val sheetName = SheetName.unsafe(name) // Safe: controlled input
    val random = new Random(42) // Fixed seed for reproducibility

    // Build sheet with columns: ID (Int), Name (String), Date (LocalDate), Amount (BigDecimal), Active (Boolean)
    var sheet: Sheet = Sheet(sheetName: SheetName) // Type ascription to resolve overload

    // Headers - use batch put for cleaner code
    sheet = sheet
      .put(
        ref"A1" -> "ID",
        ref"B1" -> "Name",
        ref"C1" -> "Date",
        ref"D1" -> "Amount",
        ref"E1" -> "Active"
      )
      .getOrElse(sheet)

    // Data rows - build list of updates then batch apply
    val dataUpdates: Seq[(ARef, Any)] = (1 to rows).flatMap { row =>
      val rowNum = row + 1 // +1 for header

      // Generate realistic data
      val id = row
      val name = s"User_${row % 100}"
      val date = LocalDate.of(2024, 1, 1).plusDays(row % 365)
      val amount =
        BigDecimal(random.nextDouble() * 10000).setScale(2, BigDecimal.RoundingMode.HALF_UP)
      val active = random.nextBoolean()

      // Return sequence of (ARef, Any) tuples
      Seq[(ARef, Any)](
        (ARef.from1(1, rowNum), id),
        (ARef.from1(2, rowNum), name),
        (ARef.from1(3, rowNum), date),
        (ARef.from1(4, rowNum), amount),
        (ARef.from1(5, rowNum), active)
      )
    }

    // Apply all data updates in one go
    sheet = sheet.put(dataUpdates*).getOrElse(sheet)

    sheet
  }

  /** Generate patches for benchmarking patch operations */
  def generatePatches(count: Int): Seq[Patch] = {
    val random = new Random(42)
    (1 to count).map { i =>
      val col = Column.from0(i % 10)
      val row = Row.from0(i)
      val ref = ARef(col, row)
      Patch.Put(ref, CellValue.Number(BigDecimal(random.nextDouble() * 100)))
    }
  }

  /** Generate styles for benchmarking style operations */
  def generateStyles(count: Int): Seq[CellStyle] = {
    val random = new Random(42)
    (1 to count).map { _ =>
      val boldFont = Font.default.copy(bold = random.nextBoolean())
      val fill = if (random.nextBoolean()) {
        Fill.Solid(Color.Rgb(random.nextInt(0xffffff)))
      } else {
        Fill.None
      }
      CellStyle.default
        .withFont(boldFont)
        .withFill(fill)
    }
  }

  /** Create a temporary file for I/O benchmarks */
  def createTempFile(prefix: String): java.io.File = {
    val file = java.io.File.createTempFile(prefix, ".xlsx")
    file.deleteOnExit()
    file
  }
}
