//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using repository ivy2Local

// Standalone demo script - run with:
//   1. Publish locally: ./mill xl-core.publishLocal
//   2. Run script: scala-cli run examples/demo.sc

import com.tjclp.xl.*

println("=== XL - Pure Scala 3.7 Excel Library Demo ===\n")

// Compile-time validated literals (unified ref macro)
val cellRef = ref"A1"
val rangeRef = ref"A1:B10"

println(s"Cell reference: ${cellRef.toA1}")
println(s"Range: ${rangeRef.toA1}")
println(
  s"Range size: ${rangeRef.width} cols × ${rangeRef.height} rows = ${rangeRef.size} cells\n"
)

// Create a workbook
val workbookResult = for
    workbook <- Workbook("Sales")
    sheet0 <- workbook(0)
    // Add some data
    updatedSheet = sheet0
      .put(ref"A1", CellValue.Text("Product"))
      .put(ref"B1", CellValue.Text("Price"))
      .put(ref"A2", CellValue.Text("Widget"))
      .put(ref"B2", CellValue.Number(19.99))
      .put(ref"A3", CellValue.Text("Gadget"))
      .put(ref"B3", CellValue.Number(29.99))
    finalWorkbook <- workbook.updateSheet(0, updatedSheet)
  yield finalWorkbook

workbookResult match
  case Right(wb) =>
    println("Created workbook:")
    wb.sheets.foreach { sheet =>
      println(s"\nSheet: ${sheet.name.value}")
      println(s"  Cells: ${sheet.cellCount}")
      sheet.nonEmptyCells.take(6).foreach { cell =>
        println(s"    ${cell.toA1}: ${cell.value}")
      }

      // Show used range
      sheet.usedRange.foreach { range =>
        println(s"  Used range: ${range.toA1}")
      }
    }

  case Left(err) =>
    println(s"Error: ${err.message}")

// Demonstrate addressing operations
println("\n=== Addressing Examples ===")
println(s"Column A = ${Column.from0(0).toLetter}")
println(s"Column Z = ${Column.from0(25).toLetter}")
println(s"Column AA = ${Column.from0(26).toLetter}")

val cellA1 = ref"A1"
val shifted = cellA1.shift(1, 1)
println(s"\n${cellA1.toA1} shifted by (1,1) = ${shifted.toA1}")

val range1 = ref"A1:C3"
val range2 = ref"B2:D4"
println(s"\n${range1.toA1} intersects ${range2.toA1}? ${range1.intersects(range2)}")
println(s"${range1.toA1} contains B2? ${range1.contains(ref"B2")}")

// Demonstrate sheet-qualified references
println("\n=== Sheet-Qualified References ===")
val qualifiedRef = ref"Sales!A1"
println(s"Qualified ref: ${qualifiedRef.toA1}")
qualifiedRef match
  case RefType.QualifiedCell(sheetName, cellRef) =>
    println(s"  Sheet: ${sheetName.value}, Cell: ${cellRef.toA1}")
  case _ => ()

val qualifiedRange = ref"'Q1 Data'!A1:B10"
println(s"\nQualified range: ${qualifiedRange.toA1}")
qualifiedRange match
  case RefType.QualifiedRange(sheetName, range) =>
    println(s"  Sheet: '${sheetName.value}', Range: ${range.toA1} (${range.size} cells)")
  case _ => ()

// Demonstrate formatted literals
println("\n=== Formatted Literals ===")
val priceSheetResult = for
  sheet <- Sheet("Prices")
yield sheet
  .putMixed(
    ref"A1" -> "Item",
    ref"B1" -> "Price",
    ref"C1" -> "Discount",
    ref"A2" -> "Widget",
    ref"B2" -> money"$$19.99",  // Escape $ in string interpolator
    ref"C2" -> percent"15%"
  )

priceSheetResult match
  case Right(priceSheet) =>
    println(s"Created sheet with ${priceSheet.cellCount} cells")
    priceSheet.nonEmptyCells.take(6).foreach { cell =>
      println(s"  ${cell.toA1}: ${cell.value}")
    }
  case Left(err) =>
    println(s"Error: ${err.message}")

// Demonstrate RichText
println("\n=== Rich Text ===")
val richTextSheetResult = for
  sheet <- Sheet("Report")
yield sheet
  .put(ref"A1", CellValue.RichText("Error: ".red.bold + "Fix this!"))
  .put(ref"A2", CellValue.RichText("Q1 ".size(18.0).bold + "Report".italic))

richTextSheetResult match
  case Right(richTextSheet) =>
    richTextSheet.nonEmptyCells.foreach { cell =>
      cell.value match
        case CellValue.RichText(rt) =>
          println(s"  ${cell.toA1}: ${rt.runs.size} formatted runs")
        case _ => ()
    }
  case Left(err) =>
    println(s"Error: ${err.message}")

println("\n=== Demo Complete ===")
