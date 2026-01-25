package com.tjclp.xl.ooxml.metadata

import java.nio.file.{Files, Path}

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.ooxml.XlsxWriter

class WorkbookMetadataReaderSpec extends FunSuite:

  // Helper to create a temp xlsx file with given sheets
  private def createTempWorkbook(sheets: Vector[Sheet]): Path =
    val path = Files.createTempFile("metadata-test-", ".xlsx")
    val wb = Workbook(sheets)
    XlsxWriter.write(wb, path)
    path

  test("read: extracts sheet names from workbook") {
    val path = createTempWorkbook(
      Vector(
        Sheet(SheetName.unsafe("Sales")),
        Sheet(SheetName.unsafe("Inventory")),
        Sheet(SheetName.unsafe("Summary"))
      )
    )
    try
      val result = WorkbookMetadataReader.read(path)
      assert(result.isRight, s"Expected Right, got $result")
      val meta = result.toOption.get
      assertEquals(meta.sheets.size, 3)
      assertEquals(meta.sheets.map(_.name.value), Vector("Sales", "Inventory", "Summary"))
    finally Files.deleteIfExists(path)
  }

  test("read: extracts dimension from worksheet with data") {
    val sheet = Sheet(SheetName.unsafe("Data"))
      .put("A1" -> "Header", "B1" -> "Value")
      .put("A10" -> "End", "C10" -> 100)
    val path = createTempWorkbook(Vector(sheet))
    try
      val result = WorkbookMetadataReader.read(path)
      assert(result.isRight)
      val meta = result.toOption.get
      assertEquals(meta.sheets.size, 1)
      // Dimension should cover A1:C10
      val dim = meta.sheets.head.dimension
      assert(dim.isDefined, "Expected dimension to be present")
      val range = dim.get
      assertEquals(range.start, ARef.parse("A1").toOption.get)
      assertEquals(range.end, ARef.parse("C10").toOption.get)
    finally Files.deleteIfExists(path)
  }

  test("read: handles empty workbook (single empty sheet)") {
    val path = createTempWorkbook(Vector(Sheet(SheetName.unsafe("Empty"))))
    try
      val result = WorkbookMetadataReader.read(path)
      assert(result.isRight)
      val meta = result.toOption.get
      assertEquals(meta.sheets.size, 1)
      assertEquals(meta.sheets.head.name.value, "Empty")
    finally Files.deleteIfExists(path)
  }

  test("readSheetList: returns sheets without dimensions (faster)") {
    val path = createTempWorkbook(
      Vector(
        Sheet(SheetName.unsafe("A")),
        Sheet(SheetName.unsafe("B"))
      )
    )
    try
      val result = WorkbookMetadataReader.readSheetList(path)
      assert(result.isRight)
      val sheets = result.toOption.get
      assertEquals(sheets.size, 2)
      // Dimensions should be None (not fetched)
      assert(sheets.forall(_.dimension.isEmpty))
    finally Files.deleteIfExists(path)
  }

  test("readDimension: reads dimension for specific sheet") {
    val sheet1 = Sheet(SheetName.unsafe("Small")).put("A1" -> "x", "B2" -> "y")
    val sheet2 = Sheet(SheetName.unsafe("Large")).put("A1" -> 1, "Z100" -> 100)
    val path = createTempWorkbook(Vector(sheet1, sheet2))
    try
      // Sheet 1 dimension
      val dim1 = WorkbookMetadataReader.readDimension(path, 1)
      assert(dim1.isRight)
      val range1 = dim1.toOption.get
      assert(range1.isDefined)
      assertEquals(range1.get.toA1, "A1:B2")

      // Sheet 2 dimension
      val dim2 = WorkbookMetadataReader.readDimension(path, 2)
      assert(dim2.isRight)
      val range2 = dim2.toOption.get
      assert(range2.isDefined)
      assertEquals(range2.get.toA1, "A1:Z100")
    finally Files.deleteIfExists(path)
  }

  test("read: returns error for invalid ZIP file") {
    val path = Files.createTempFile("not-xlsx-", ".xlsx")
    Files.write(path, "not a zip".getBytes)
    try
      val result = WorkbookMetadataReader.read(path)
      assert(result.isLeft, "Expected error for invalid file")
      assert(result.left.toOption.get.message.contains("Invalid ZIP file"))
    finally Files.deleteIfExists(path)
  }

  test("read: returns error for non-existent file") {
    val path = Path.of("/nonexistent/path/to/file.xlsx")
    val result = WorkbookMetadataReader.read(path)
    assert(result.isLeft, "Expected error for non-existent file")
  }
