import com.tjclp.xl.*
import com.tjclp.xl.macros.ref

@main
def demo(): Unit =
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

  // Demonstrate new sheet-qualified references
  println("\n=== Sheet-Qualified References (New!) ===")
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

  println("\n=== Demo Complete ===")
