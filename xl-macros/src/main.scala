import com.tjclp.xl.*
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.macros.{cell, range}

@main
def demo(): Unit =
  println("=== XL - Pure Scala 3.7 Excel Library Demo ===\n")

  // Compile-time validated literals
  val ref = cell"A1"
  val rng = range"A1:B10"

  println(s"Cell reference: ${ref.toA1}")
  println(s"Range: ${rng.toA1}")
  println(s"Range size: ${rng.width} cols × ${rng.height} rows = ${rng.size} cells\n")

  // Create a workbook
  val workbookResult = for
    workbook <- Workbook("Sales")
    sheet0 <- workbook(0)
    // Add some data
    updatedSheet = sheet0
      .put(cell"A1", CellValue.Text("Product"))
      .put(cell"B1", CellValue.Text("Price"))
      .put(cell"A2", CellValue.Text("Widget"))
      .put(cell"B2", CellValue.Number(19.99))
      .put(cell"A3", CellValue.Text("Gadget"))
      .put(cell"B3", CellValue.Number(29.99))
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

  val cellA1 = cell"A1"
  val shifted = cellA1.shift(1, 1)
  println(s"\n${cellA1.toA1} shifted by (1,1) = ${shifted.toA1}")

  val range1 = range"A1:C3"
  val range2 = range"B2:D4"
  println(s"\n${range1.toA1} intersects ${range2.toA1}? ${range1.intersects(range2)}")
  println(s"${range1.toA1} contains B2? ${range1.contains(cell"B2")}")

  println("\n=== Demo Complete ===")
