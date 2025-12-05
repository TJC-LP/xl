//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.5-SNAPSHOT
//> using dep com.tjclp::xl-ooxml:0.1.5-SNAPSHOT
//> using repository ivy2Local

// Excel Table Support Demo - run with:
//   1. Publish locally: ./mill xl-core.publishLocal xl-ooxml.publishLocal
//   2. Run script: scala-cli run examples/table_demo.sc

import com.tjclp.xl.{*, given}
import com.tjclp.xl.unsafe.*
import com.tjclp.xl.tables.{TableSpec, TableAutoFilter, TableStyle}
import com.tjclp.xl.ooxml.{XlsxWriter, XlsxReader}
import java.nio.file.{Files, Path}

println("=== Excel Table Support Demo ===\n")

// ============================================================
// PART 1: Creating a Table
// ============================================================

println("## Part 1: Creating a Sales Table\n")

// Define table structure with column names
val salesTable = TableSpec.fromColumnNames(
  name = "SalesData",
  displayName = "Q4_Sales_Data",  // Excel requires no spaces (auto-sanitized to underscores)
  range = ref"A1:D11",  // Header + 10 data rows
  columnNames = Vector("Product", "Region", "Quantity", "Revenue")
).map(_.copy(
  autoFilter = Some(TableAutoFilter(enabled = true)),
  style = TableStyle.Medium(9)  // Blue table style
)).unsafe

println(s"Created table: ${salesTable.name}")
println(s"  Display name: ${salesTable.displayName}")
println(s"  Range: ${salesTable.range.toA1}")
println(s"  Columns: ${salesTable.columns.map(_.name).mkString(", ")}")
println(s"  AutoFilter: ${salesTable.autoFilter.map(_.enabled).getOrElse(false)}")
println(s"  Style: ${salesTable.style}")
println(s"  Valid: ${salesTable.isValid}\n")

// Data range (excludes header row)
println(s"Data range: ${salesTable.dataRange.toA1}")
println(s"  Width: ${salesTable.dataRange.width} columns")
println(s"  Height: ${salesTable.dataRange.height} rows\n")

// ============================================================
// PART 2: Populating Table with Data
// ============================================================

println("## Part 2: Populating Table with Data\n")

// Create sheet and add table - Sheet.apply returns Sheet directly (infallible)
val sheet = Sheet("Q4 Sales")
  .withTable(salesTable)
  // Populate header row
  .put(
    ref"A1" -> "Product",
    ref"B1" -> "Region",
    ref"C1" -> "Quantity",
    ref"D1" -> "Revenue"
  )
  // Populate data rows
  .put(
    // Row 2
    ref"A2" -> "Widget",
    ref"B2" -> "North",
    ref"C2" -> 150,
    ref"D2" -> BigDecimal("4500.00"),
    // Row 3
    ref"A3" -> "Gadget",
    ref"B3" -> "South",
    ref"C3" -> 200,
    ref"D3" -> BigDecimal("8000.00"),
    // Row 4
    ref"A4" -> "Doohickey",
    ref"B4" -> "East",
    ref"C4" -> 175,
    ref"D4" -> BigDecimal("5250.00"),
    // Row 5
    ref"A5" -> "Widget",
    ref"B5" -> "West",
    ref"C5" -> 225,
    ref"D5" -> BigDecimal("6750.00"),
    // Row 6
    ref"A6" -> "Gadget",
    ref"B6" -> "North",
    ref"C6" -> 300,
    ref"D6" -> BigDecimal("12000.00"),
    // Row 7
    ref"A7" -> "Thingamajig",
    ref"B7" -> "South",
    ref"C7" -> 125,
    ref"D7" -> BigDecimal("3125.00"),
    // Row 8
    ref"A8" -> "Widget",
    ref"B8" -> "East",
    ref"C8" -> 190,
    ref"D8" -> BigDecimal("5700.00"),
    // Row 9
    ref"A9" -> "Doohickey",
    ref"B9" -> "West",
    ref"C9" -> 210,
    ref"D9" -> BigDecimal("6300.00"),
    // Row 10
    ref"A10" -> "Gadget",
    ref"B10" -> "South",
    ref"C10" -> 275,
    ref"D10" -> BigDecimal("11000.00"),
    // Row 11
    ref"A11" -> "Thingamajig",
    ref"B11" -> "North",
    ref"C11" -> 140,
    ref"D11" -> BigDecimal("3500.00")
  )

println(s"✓ Sheet created with ${sheet.cellCount} cells")
println(s"✓ Table '${salesTable.name}' attached\n")

// ============================================================
// PART 3: Multiple Tables on One Sheet
// ============================================================

println("## Part 3: Adding a Second Table (Summary)\n")

val summaryTable = TableSpec.fromColumnNames(
  name = "RegionSummary",
  displayName = "Regional_Summary",  // Excel requires no spaces
  range = ref"F1:G5",  // Separate from main table
  columnNames = Vector("Region", "Total_Revenue")  // Column names also sanitized
).map(_.copy(
  autoFilter = None,  // No filter on summary
  style = TableStyle.Light(15)  // Green table style
)).unsafe

val multiTableSheet = sheet
  .withTable(summaryTable)
  .put(
    // Headers
    ref"F1" -> "Region",
    ref"G1" -> "Total_Revenue",
    // Data
    ref"F2" -> "North",
    ref"G2" -> BigDecimal("20000.00"),
    ref"F3" -> "South",
    ref"G3" -> BigDecimal("22125.00"),
    ref"F4" -> "East",
    ref"G4" -> BigDecimal("10950.00"),
    ref"F5" -> "West",
    ref"G5" -> BigDecimal("13050.00")
  )

println(s"✓ Added second table: ${summaryTable.name}")
println(s"✓ Sheet now has ${multiTableSheet.tables.size} tables\n")

// ============================================================
// PART 4: Writing to XLSX File
// ============================================================

println("## Part 4: Writing to Excel File\n")

val workbook = Workbook(Vector(multiTableSheet))

// Write to /tmp/ so you can open in Excel
val outputPath = Path.of("/tmp/sales-table-demo.xlsx")

XlsxWriter.write(workbook, outputPath) match
  case Right(_) =>
    println(s"✓ Wrote workbook to: $outputPath")
    println(s"  File size: ${Files.size(outputPath)} bytes")
    println(s"  → Open this file in Excel to see tables with AutoFilter!\n")
  case Left(err) =>
    println(s"✗ Write failed: $err")
    sys.exit(1)

// ============================================================
// PART 5: Reading Tables Back
// ============================================================

println("## Part 5: Reading Tables from File\n")

XlsxReader.read(outputPath) match
  case Right(readWorkbook) =>
    println(s"✓ Read workbook with ${readWorkbook.sheets.size} sheet(s)")

    val readSheet = readWorkbook.sheets.head
    println(s"  Sheet name: ${readSheet.name.value}")
    println(s"  Cell count: ${readSheet.cellCount}")
    println(s"  Table count: ${readSheet.tables.size}\n")

    // Access tables
    readSheet.allTables.foreach { table =>
      println(s"Table: ${table.name}")
      println(s"  Display name: ${table.displayName}")
      println(s"  Range: ${table.range.toA1}")
      println(s"  Columns: ${table.columns.size}")
      table.columns.foreach { col =>
        println(s"    - ${col.name} (ID: ${col.id})")
      }
      println(s"  AutoFilter: ${table.autoFilter.map(_.enabled).getOrElse(false)}")
      println(s"  Style: ${table.style}")
      println(s"  Show header: ${table.showHeaderRow}")
      println(s"  Show totals: ${table.showTotalsRow}")
      println()
    }

    // Retrieve specific table by name
    readSheet.getTable("SalesData") match
      case Some(sales) =>
        println(s"Retrieved table by name: ${sales.name}")
        println(s"  Has AutoFilter: ${sales.autoFilter.isDefined}")
      case None =>
        println("✗ Table 'SalesData' not found!")

    println()

  case Left(err) =>
    println(s"✗ Read failed: $err")

// ============================================================
// PART 6: Different Table Styles
// ============================================================

println("## Part 6: Table Style Examples\n")

val styleExamples = Vector(
  ("Light Style", TableStyle.Light(5)),
  ("Medium Style", TableStyle.Medium(2)),
  ("Dark Style", TableStyle.Dark(3)),
  ("No Style", TableStyle.None)
)

styleExamples.foreach { case (name, style) =>
  println(s"$name: $style")
}
println()

// ============================================================
// PART 7: Table Validation
// ============================================================

println("## Part 7: Table Validation\n")

// Valid table
val validTable = TableSpec.fromColumnNames(
  name = "Valid",
  displayName = "Valid_Table",  // Excel requires no spaces
  range = ref"A1:C5",
  columnNames = Vector("Col1", "Col2", "Col3")
).unsafe
println(s"Valid table (3 cols, 3-col range): ${validTable.isValid}")

// Invalid table (column count mismatch) - this will fail validation
val invalidTableResult = TableSpec.fromColumnNames(
  name = "Invalid",
  displayName = "Invalid_Table",  // Excel requires no spaces
  range = ref"A1:D5",  // 4 columns wide
  columnNames = Vector("Col1", "Col2")  // Only 2 columns defined
)
println(s"Invalid table (2 cols, 4-col range): ${invalidTableResult.isLeft} (expected: true - column mismatch)")
println()

// ============================================================
// PART 8: Table Operations
// ============================================================

println("## Part 8: Table Operations\n")

// Check if sheet has table
println(s"Has 'SalesData' table: ${multiTableSheet.hasTable("SalesData")}")
println(s"Has 'NonExistent' table: ${multiTableSheet.hasTable("NonExistent")}")

// Remove a table
val sheetWithoutSummary = multiTableSheet.removeTable("RegionSummary")
println(s"\nAfter removing 'RegionSummary':")
println(s"  Tables remaining: ${sheetWithoutSummary.tables.size}")
sheetWithoutSummary.allTables.foreach { t =>
  println(s"    - ${t.name}")
}
println()

// ============================================================
// Summary
// ============================================================

println("=== Demo Complete ===\n")
println("Summary:")
println("  ✓ Created tables with TableSpec.fromColumnNames")
println("  ✓ Applied table styles (Light, Medium, Dark)")
println("  ✓ Enabled AutoFilter for filterable tables")
println("  ✓ Populated table ranges with data")
println("  ✓ Multiple tables on one sheet")
println("  ✓ Full round-trip (write → read → verify)")
println("  ✓ Table operations (get, has, remove)")
println()
println(s"Output file: $outputPath")
println("Open in Excel to see tables with AutoFilter and styling!")
