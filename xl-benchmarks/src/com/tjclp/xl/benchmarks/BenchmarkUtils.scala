package com.tjclp.xl.benchmarks

import com.tjclp.xl.*
import com.tjclp.xl.addressing.*
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.codec.CellCodec.given
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
    val emptySheet: Sheet = Sheet(sheetName: SheetName) // Type ascription to resolve overload

    // Headers - use batch put for cleaner code
    val headerSheet = emptySheet
      .put(
        ref"A1" -> "ID",
        ref"B1" -> "Name",
        ref"C1" -> "Date",
        ref"D1" -> "Amount",
        ref"E1" -> "Active"
      )
      .getOrElse(emptySheet)

    // Data rows - build list of updates then batch apply
    // Note: CellValue used for type-safe heterogeneous batch operations
    val dataUpdates: Seq[(ARef, CellValue)] = (1 to rows).flatMap { row =>
      val rowNum = row + 1 // +1 for header

      // Generate realistic data
      val id = row
      val name = s"User_${row % 100}"
      val date = LocalDate.of(2024, 1, 1).plusDays(row % 365)
      val amount =
        BigDecimal(random.nextDouble() * 10000).setScale(2, BigDecimal.RoundingMode.HALF_UP)
      val active = random.nextBoolean()

      // Return sequence of (ARef, CellValue) tuples
      Seq[(ARef, CellValue)](
        (ARef.from1(1, rowNum), CellValue.Number(BigDecimal(id))),
        (ARef.from1(2, rowNum), CellValue.Text(name)),
        (ARef.from1(3, rowNum), CellValue.DateTime(date.atStartOfDay)),
        (ARef.from1(4, rowNum), CellValue.Number(amount)),
        (ARef.from1(5, rowNum), CellValue.Bool(active))
      )
    }

    // Apply all data updates in one go
    headerSheet.put(dataUpdates*).getOrElse(headerSheet)
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

  /**
   * Create workbook with verifiable data for benchmark validation.
   *
   * Pattern: Cell A{i} contains value i (1, 2, 3, ..., N)
   *
   * Expected sum: N × (N+1) / 2
   *
   * This allows verification that benchmarks actually parse data correctly:
   *   - 1,000 rows: sum = 500,500
   *   - 10,000 rows: sum = 50,005,000
   *
   * Uses direct Cell construction for clarity and to avoid varargs overhead in setup. This is safe
   * because we're building CellValue.Number directly (no codec needed).
   */
  def createVerifiableWorkbook(rows: Int): Workbook = {
    val emptySheet: Sheet = Sheet(SheetName.unsafe("Data"))

    // Create rows with cell value = row number (verifiable arithmetic series)
    val updates: Seq[(ARef, Double)] = (1 to rows).map { i =>
      (ARef.from1(1, i): ARef) -> i.toDouble
    }

    val sheet = emptySheet.put(updates*) match {
      case Right(s) => s
      case Left(err) => sys.error(s"Failed to create verifiable sheet: $err")
    }
    // Direct construction with only the Data sheet (avoids Sheet1 issue)
    Workbook(Vector(sheet))
  }

  /** Compute expected sum for verifiable workbook (arithmetic series formula) */
  def expectedSum(rows: Int): Double = rows.toDouble * (rows + 1) / 2
}
