package com.tjclp.xl.ooxml

import com.tjclp.xl.*
import com.tjclp.xl.macros.{cell, range}
import java.nio.file.Paths

/** Manual test to generate a styled XLSX for visual verification in Excel.
  *
  * Run with: ./mill xl-ooxml.test.runMain ManualStyleTest
  *
  * Then open test-styles.xlsx in Excel to verify:
  * - A1 is BOLD
  * - B1 is ITALIC
  * - C1 has RED background
  * - D1 has BLUE font
  * - A2:D2 have GRAY background (header row)
  * - B3 has currency format
  */
object ManualStyleTest:
  def main(args: Array[String]): Unit =
    // Define styles
    val boldStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))

    val italicStyle = CellStyle.default.withFont(Font("Arial", 12.0, italic = true))

    val redFillStyle = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))

    val blueFontStyle = CellStyle.default.withFont(
      Font("Arial", 12.0, color = Some(Color.Rgb(0xFF0000FF)))
    )

    val headerStyle = CellStyle.default
      .withFont(Font("Arial", 12.0, bold = true))
      .withFill(Fill.Solid(Color.Rgb(0xFFCCCCCC)))
      .withAlign(Align(HAlign.Center, VAlign.Middle))

    val currencyStyle = CellStyle.default.withNumFmt(NumFmt.Currency)

    // Build worksheet
    val sheet = Sheet("StyleDemo").getOrElse(sys.error("Failed to create sheet"))
      // Row 1: Individual styles
      .put(cell"A1", CellValue.Text("BOLD"))
      .withCellStyle(cell"A1", boldStyle)
      .put(cell"B1", CellValue.Text("ITALIC"))
      .withCellStyle(cell"B1", italicStyle)
      .put(cell"C1", CellValue.Text("Red BG"))
      .withCellStyle(cell"C1", redFillStyle)
      .put(cell"D1", CellValue.Text("Blue Font"))
      .withCellStyle(cell"D1", blueFontStyle)
      // Row 2: Header row with range styling
      .put(cell"A2", CellValue.Text("Name"))
      .put(cell"B2", CellValue.Text("Price"))
      .put(cell"C2", CellValue.Text("Quantity"))
      .put(cell"D2", CellValue.Text("Total"))
      .withRangeStyle(range"A2:D2", headerStyle)
      // Row 3: Data with currency format
      .put(cell"A3", CellValue.Text("Widget"))
      .put(cell"B3", CellValue.Number(BigDecimal("19.99")))
      .withCellStyle(cell"B3", currencyStyle)
      .put(cell"C3", CellValue.Number(BigDecimal(5)))
      .put(cell"D3", CellValue.Number(BigDecimal("99.95")))
      .withCellStyle(cell"D3", currencyStyle)

    // Create workbook
    val wb = Workbook(Vector(sheet))

    // Write to file
    val outputPath = Paths.get("test-styles.xlsx")
    val result = XlsxWriter.write(wb, outputPath)

    result match
      case Right(_) =>
        println(s"✅ Successfully wrote: $outputPath")
        println(s"   Styles in registry: ${sheet.styleRegistry.size}")
        println(s"   Open in Excel to verify:")
        println(s"     - A1 is BOLD")
        println(s"     - B1 is ITALIC")
        println(s"     - C1 has RED background")
        println(s"     - D1 has BLUE font")
        println(s"     - A2:D2 have GRAY background (header)")
        println(s"     - B3, D3 show as currency ($$19.99, $$99.95)")

      case Left(err) =>
        println(s"❌ Failed to write: ${err.message}")
        sys.error("Write failed")
