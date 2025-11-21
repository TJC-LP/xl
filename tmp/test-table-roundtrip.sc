//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using dep com.tjclp::xl-ooxml:0.1.0-SNAPSHOT
//> using repository ivy2Local

import com.tjclp.xl.*
import com.tjclp.xl.tables.{TableSpec, TableAutoFilter, TableStyle}
import com.tjclp.xl.ooxml.{XlsxWriter, XlsxReader}
import java.nio.file.Path

println("=== Table Round-Trip Test ===\n")

// Create a table
val table = TableSpec.fromColumnNames(
  name = "TestTable",
  displayName = "Test Table",
  range = ref"A1:C6",
  columnNames = Vector("Name", "Value", "Status")
).copy(
  autoFilter = Some(TableAutoFilter(enabled = true)),
  style = TableStyle.Medium(2)
)

println(s"Created table: ${table.name}")
println(s"  Range: ${table.range.toA1}")
println(s"  Columns: ${table.columns.size}")
println(s"  AutoFilter: ${table.autoFilter.isDefined}")
println(s"  Style: ${table.style}\n")

// Create sheet with table and data
val sheetResult = for
  sheet <- Sheet("Data")
  withTable = sheet.withTable(table)
  withData <- withTable.put(
    ref"A1" -> "Name",
    ref"B1" -> "Value",
    ref"C1" -> "Status",
    ref"A2" -> "Item 1",
    ref"B2" -> 100,
    ref"C2" -> "Active",
    ref"A3" -> "Item 2",
    ref"B3" -> 200,
    ref"C3" -> "Pending",
    ref"A4" -> "Item 3",
    ref"B4" -> 150,
    ref"C4" -> "Active",
    ref"A5" -> "Item 4",
    ref"B5" -> 300,
    ref"C5" -> "Complete",
    ref"A6" -> "Item 5",
    ref"B6" -> 175,
    ref"C6" -> "Active"
  )
yield withData

val sheet = sheetResult match
  case Right(s) =>
    println(s"✓ Sheet created with ${s.cellCount} cells")
    println(s"✓ Table attached: ${s.hasTable("TestTable")}\n")
    s
  case Left(err) =>
    println(s"✗ Error creating sheet: $err")
    sys.exit(1)

// Create workbook
val workbook = Workbook(Vector(sheet))
println(s"Workbook created with ${workbook.sheets.size} sheet(s)\n")

// Write to file
val outputPath = Path.of("tmp/test-table.xlsx")
println(s"Writing to: $outputPath")

XlsxWriter.write(workbook, outputPath) match
  case Right(_) =>
    println(s"✓ Write successful")
  case Left(err) =>
    println(s"✗ Write failed: $err")
    sys.exit(1)

// Read back
println("\nReading back from file...")

XlsxReader.read(outputPath) match
  case Right(readWorkbook) =>
    println(s"✓ Read successful")
    println(s"  Sheets: ${readWorkbook.sheets.size}")

    val readSheet = readWorkbook.sheets.head
    println(s"  Sheet name: ${readSheet.name.value}")
    println(s"  Cells: ${readSheet.cellCount}")
    println(s"  Tables: ${readSheet.tables.size}\n")

    readSheet.getTable("TestTable") match
      case Some(readTable) =>
        println("✓ Table preserved!")
        println(s"  Name: ${readTable.name}")
        println(s"  Display name: ${readTable.displayName}")
        println(s"  Range: ${readTable.range.toA1}")
        println(s"  Columns: ${readTable.columns.size}")
        readTable.columns.foreach { col =>
          println(s"    - ${col.name} (ID: ${col.id})")
        }
        println(s"  AutoFilter: ${readTable.autoFilter.isDefined}")
        println(s"  Style: ${readTable.style}")

        // Verify properties match
        println("\n=== Verification ===")
        val nameMatches = readTable.name == table.name
        val rangeMatches = readTable.range == table.range
        val colCountMatches = readTable.columns.size == table.columns.size
        val autoFilterMatches = readTable.autoFilter.isDefined == table.autoFilter.isDefined
        val styleMatches = readTable.style == table.style

        println(s"  Name matches: $nameMatches")
        println(s"  Range matches: $rangeMatches")
        println(s"  Column count matches: $colCountMatches")
        println(s"  AutoFilter matches: $autoFilterMatches")
        println(s"  Style matches: $styleMatches")

        val allMatch = nameMatches && rangeMatches && colCountMatches && autoFilterMatches && styleMatches
        println(s"\n${if allMatch then "✓✓✓ ROUND-TRIP SUCCESS ✓✓✓" else "✗✗✗ MISMATCH DETECTED ✗✗✗"}")

      case None =>
        println("✗✗✗ TABLE NOT FOUND AFTER READ ✗✗✗")
        sys.exit(1)

  case Left(err) =>
    println(s"✗ Read failed: $err")
    sys.exit(1)
