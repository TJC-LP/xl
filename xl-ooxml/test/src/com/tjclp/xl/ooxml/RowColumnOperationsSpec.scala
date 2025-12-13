package com.tjclp.xl.ooxml

import munit.FunSuite
import java.nio.file.{Files, Path}
import com.tjclp.xl.addressing.{ARef, Column, Row}
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.dsl.*
import com.tjclp.xl.macros.{col, ref}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}

/**
 * Integration tests for row/column operations including:
 * - col"A" compile-time macro
 * - Row/Column builder DSL
 * - Round-trip serialization of row/column properties
 */
@SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
class RowColumnOperationsSpec extends FunSuite:

  val tempDir: Path = Files.createTempDirectory("xl-rowcol-test-")

  override def afterAll(): Unit =
    Files.walk(tempDir)
      .sorted(java.util.Comparator.reverseOrder())
      .forEach(Files.delete)

  // ========== Column Macro Tests ==========

  test("col macro: single letter column A") {
    val c = col"A"
    assertEquals(c.index0, 0)
  }

  test("col macro: single letter column Z") {
    val c = col"Z"
    assertEquals(c.index0, 25)
  }

  test("col macro: double letter column AA") {
    val c = col"AA"
    assertEquals(c.index0, 26)
  }

  test("col macro: double letter column AZ") {
    val c = col"AZ"
    assertEquals(c.index0, 51)
  }

  test("col macro: triple letter column XFD (max Excel column)") {
    val c = col"XFD"
    assertEquals(c.index0, 16383)
  }

  test("col macro: lowercase letters are accepted") {
    val c = col"b"
    assertEquals(c.index0, 1)
  }

  // ========== Row Builder DSL Tests ==========

  test("row builder: creates RowBuilder for 0-based index") {
    val builder = row(0)
    assertEquals(builder.row, Row.from0(0))
  }

  test("row builder: height sets height property") {
    val builder = row(0).height(30.0)
    assertEquals(builder.properties.height, Some(30.0))
  }

  test("row builder: hidden sets hidden property") {
    val builder = row(0).hidden
    assertEquals(builder.properties.hidden, true)
  }

  test("row builder: outlineLevel sets outline level") {
    val builder = row(0).outlineLevel(2)
    assertEquals(builder.properties.outlineLevel, Some(2))
  }

  test("row builder: collapsed sets collapsed property") {
    val builder = row(0).collapsed
    assertEquals(builder.properties.collapsed, true)
  }

  test("row builder: chaining multiple properties") {
    val builder = row(5).height(25.0).hidden.outlineLevel(1).collapsed
    assertEquals(builder.properties.height, Some(25.0))
    assertEquals(builder.properties.hidden, true)
    assertEquals(builder.properties.outlineLevel, Some(1))
    assertEquals(builder.properties.collapsed, true)
  }

  test("row builder: toPatch creates SetRowProperties patch") {
    val patch = row(0).height(30.0).toPatch
    patch match
      case Patch.SetRowProperties(r, props) =>
        assertEquals(r, Row.from0(0))
        assertEquals(props.height, Some(30.0))
      case _ => fail(s"Expected SetRowProperties, got $patch")
  }

  test("rows helper: creates multiple builders") {
    val builders = rows(0 to 4)
    assertEquals(builders.size, 5)
    assertEquals(builders.head.row, Row.from0(0))
    assertEquals(builders.last.row, Row.from0(4))
  }

  // ========== Column Builder DSL Tests ==========

  test("column builder: width sets width property") {
    val builder = col"A".width(20.0)
    assertEquals(builder.properties.width, Some(20.0))
  }

  test("column builder: hidden sets hidden property") {
    val builder = col"B".hidden
    assertEquals(builder.properties.hidden, true)
  }

  test("column builder: outlineLevel sets outline level") {
    val builder = col"C".outlineLevel(3)
    assertEquals(builder.properties.outlineLevel, Some(3))
  }

  test("column builder: collapsed sets collapsed property") {
    val builder = col"D".collapsed
    assertEquals(builder.properties.collapsed, true)
  }

  test("column builder: chaining multiple properties") {
    val builder = col"E".width(15.0).hidden.outlineLevel(2).collapsed
    assertEquals(builder.properties.width, Some(15.0))
    assertEquals(builder.properties.hidden, true)
    assertEquals(builder.properties.outlineLevel, Some(2))
    assertEquals(builder.properties.collapsed, true)
  }

  test("column builder: toPatch creates SetColumnProperties patch") {
    val patch = col"A".width(25.0).toPatch
    patch match
      case Patch.SetColumnProperties(c, props) =>
        assertEquals(c, Column.from0(0))
        assertEquals(props.width, Some(25.0))
      case _ => fail(s"Expected SetColumnProperties, got $patch")
  }

  // ========== Patch Composition Tests ==========

  test("row and column patches compose with ++") {
    val patch =
      row(0).height(30.0).toPatch ++
        col"A".width(20.0).toPatch ++
        row(1).outlineLevel(1).toPatch

    patch match
      case Patch.Batch(patches) =>
        assertEquals(patches.size, 3)
      case _ => fail(s"Expected Batch, got $patch")
  }

  // ========== Round-Trip Tests ==========

  test("column width round-trips through OOXML") {
    val initial = Workbook("Test")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Text("Data"))
      .setColumnProperties(col"A", ColumnProperties(width = Some(25.5)))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("col-width.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    assertEquals(readSheet.getColumnProperties(col"A").width, Some(25.5))
  }

  test("column hidden flag round-trips through OOXML") {
    val initial = Workbook("Test")
    val sheet = initial.sheets(0)
      .put(ref"B1", CellValue.Text("Hidden"))
      .setColumnProperties(col"B", ColumnProperties(hidden = true))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("col-hidden.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    assertEquals(readSheet.getColumnProperties(col"B").hidden, true)
  }

  test("column outlineLevel round-trips through OOXML") {
    val initial = Workbook("Test")
    val sheet = initial.sheets(0)
      .put(ref"C1", CellValue.Text("Outlined"))
      .setColumnProperties(col"C", ColumnProperties(outlineLevel = Some(2)))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("col-outline.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    assertEquals(readSheet.getColumnProperties(col"C").outlineLevel, Some(2))
  }

  test("column collapsed flag round-trips through OOXML") {
    val initial = Workbook("Test")
    val sheet = initial.sheets(0)
      .put(ref"D1", CellValue.Text("Collapsed"))
      .setColumnProperties(col"D", ColumnProperties(collapsed = true, outlineLevel = Some(1)))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("col-collapsed.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    assertEquals(readSheet.getColumnProperties(col"D").collapsed, true)
    assertEquals(readSheet.getColumnProperties(col"D").outlineLevel, Some(1))
  }

  test("row height round-trips through OOXML") {
    val initial = Workbook("Test")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Text("Tall row"))
      .setRowProperties(Row.from0(0), RowProperties(height = Some(40.0)))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("row-height.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    assertEquals(readSheet.getRowProperties(Row.from0(0)).height, Some(40.0))
  }

  test("row hidden flag round-trips through OOXML") {
    val initial = Workbook("Test")
    val sheet = initial.sheets(0)
      .put(ref"A2", CellValue.Text("Hidden row"))
      .setRowProperties(Row.from0(1), RowProperties(hidden = true))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("row-hidden.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    assertEquals(readSheet.getRowProperties(Row.from0(1)).hidden, true)
  }

  test("row outlineLevel round-trips through OOXML") {
    val initial = Workbook("Test")
    val sheet = initial.sheets(0)
      .put(ref"A3", CellValue.Text("Outlined row"))
      .setRowProperties(Row.from0(2), RowProperties(outlineLevel = Some(3)))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("row-outline.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    assertEquals(readSheet.getRowProperties(Row.from0(2)).outlineLevel, Some(3))
  }

  test("row collapsed flag round-trips through OOXML") {
    val initial = Workbook("Test")
    val sheet = initial.sheets(0)
      .put(ref"A4", CellValue.Text("Collapsed row"))
      .setRowProperties(Row.from0(3), RowProperties(collapsed = true, outlineLevel = Some(1)))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("row-collapsed.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    assertEquals(readSheet.getRowProperties(Row.from0(3)).collapsed, true)
    assertEquals(readSheet.getRowProperties(Row.from0(3)).outlineLevel, Some(1))
  }

  test("multiple columns with different widths round-trip") {
    val initial = Workbook("Test")
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Text("Col A"))
      .put(ref"B1", CellValue.Text("Col B"))
      .put(ref"C1", CellValue.Text("Col C"))
      .setColumnProperties(col"A", ColumnProperties(width = Some(10.0)))
      .setColumnProperties(col"B", ColumnProperties(width = Some(20.0)))
      .setColumnProperties(col"C", ColumnProperties(width = Some(30.0)))

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("multi-col-width.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    assertEquals(readSheet.getColumnProperties(col"A").width, Some(10.0))
    assertEquals(readSheet.getColumnProperties(col"B").width, Some(20.0))
    assertEquals(readSheet.getColumnProperties(col"C").width, Some(30.0))
  }

  test("consecutive columns with same properties are grouped in <cols>") {
    // This tests that columns A, B, C with same width become a single <col min="1" max="3">
    val initial = Workbook("Test")
    val sameProps = ColumnProperties(width = Some(15.0))
    val sheet = initial.sheets(0)
      .put(ref"A1", CellValue.Text("A"))
      .put(ref"B1", CellValue.Text("B"))
      .put(ref"C1", CellValue.Text("C"))
      .setColumnProperties(col"A", sameProps)
      .setColumnProperties(col"B", sameProps)
      .setColumnProperties(col"C", sameProps)

    val wb = initial.update(initial.sheets(0).name, _ => sheet).getOrElse(fail("Should create workbook"))

    val outputPath = tempDir.resolve("grouped-cols.xlsx")
    XlsxWriter.write(wb, outputPath).getOrElse(fail("Write failed"))

    val readWb = XlsxReader.read(outputPath).getOrElse(fail("Read failed"))
    val readSheet = readWb.sheets(0)

    // All three columns should have the same width after round-trip
    assertEquals(readSheet.getColumnProperties(col"A").width, Some(15.0))
    assertEquals(readSheet.getColumnProperties(col"B").width, Some(15.0))
    assertEquals(readSheet.getColumnProperties(col"C").width, Some(15.0))
  }

  test("builder DSL patches apply correctly via Sheet.put(patch)") {
    val sheet = Sheet("Test")
    val patch =
      row(0).height(30.0).toPatch ++
        col"A".width(20.0).toPatch

    val updated = sheet.put(patch)
    assertEquals(updated.getRowProperties(Row.from0(0)).height, Some(30.0))
    assertEquals(updated.getColumnProperties(col"A").width, Some(20.0))
  }

  // ========== Validation Tests ==========

  test("outlineLevel validation: level 0-7 accepted") {
    // Should not throw
    val props0 = RowProperties(outlineLevel = Some(0))
    val props7 = RowProperties(outlineLevel = Some(7))
    assertEquals(props0.outlineLevel, Some(0))
    assertEquals(props7.outlineLevel, Some(7))
  }

  test("outlineLevel validation: negative level rejected") {
    intercept[IllegalArgumentException] {
      RowProperties(outlineLevel = Some(-1))
    }
  }

  test("outlineLevel validation: level > 7 rejected") {
    intercept[IllegalArgumentException] {
      RowProperties(outlineLevel = Some(8))
    }
  }

  test("column properties persist through multiple surgical writes (regression #col-props-bug)") {
    // Regression test: Column widths were lost on subsequent write operations.
    // Bug 1: preserved.cols.orElse(generatedCols) prioritized preserved XML over domain props
    // Bug 2: XmlUtil.elem sorted attributes alphabetically, breaking attribute order test

    val path1 = tempDir.resolve("col-props-1.xlsx")
    val path2 = tempDir.resolve("col-props-2.xlsx")
    val path3 = tempDir.resolve("col-props-3.xlsx")

    // Create initial workbook with column A width
    val sheet1 = Sheet("Test")
      .put(ref"A1", CellValue.Text("Hello"))
      .setColumnProperties(col"A", ColumnProperties(width = Some(20.0)))

    val wb1 = Workbook(Vector(sheet1))
    XlsxWriter.write(wb1, path1).fold(err => fail(s"Write 1 failed: $err"), identity)

    // Read and add column B width
    val wb2 = XlsxReader.read(path1).fold(err => fail(s"Read 1 failed: $err"), identity)
    val sheet2 = wb2.sheets.head.setColumnProperties(col"B", ColumnProperties(width = Some(30.0)))
    val modifiedWb2 = wb2.put(sheet2)
    XlsxWriter.write(modifiedWb2, path2).fold(err => fail(s"Write 2 failed: $err"), identity)

    // Read and add column C width
    val wb3 = XlsxReader.read(path2).fold(err => fail(s"Read 2 failed: $err"), identity)
    val sheet3 = wb3.sheets.head.setColumnProperties(col"C", ColumnProperties(width = Some(40.0)))
    val modifiedWb3 = wb3.put(sheet3)
    XlsxWriter.write(modifiedWb3, path3).fold(err => fail(s"Write 3 failed: $err"), identity)

    // Verify final workbook has ALL column properties
    val finalWb = XlsxReader.read(path3).fold(err => fail(s"Read 3 failed: $err"), identity)
    val finalSheet = finalWb.sheets.head

    val colAProps = finalSheet.columnProperties.get(col"A")
    val colBProps = finalSheet.columnProperties.get(col"B")
    val colCProps = finalSheet.columnProperties.get(col"C")

    assert(colAProps.isDefined, "Column A properties lost after multiple writes")
    assert(colBProps.isDefined, "Column B properties lost after multiple writes")
    assert(colCProps.isDefined, "Column C properties lost after multiple writes")

    assertEquals(colAProps.get.width, Some(20.0), "Column A width changed")
    assertEquals(colBProps.get.width, Some(30.0), "Column B width changed")
    assertEquals(colCProps.get.width, Some(40.0), "Column C width changed")

    // Also verify the XML has proper attribute order (min before max)
    val zip = new java.util.zip.ZipFile(path3.toFile)
    val sheetEntry = zip.getEntry("xl/worksheets/sheet1.xml")
    val sheetXml = scala.io.Source.fromInputStream(zip.getInputStream(sheetEntry)).mkString
    zip.close()

    assert(
      sheetXml.contains("""<col min="1""""),
      "Column XML should have min attribute first (OOXML compliance)"
    )
  }

end RowColumnOperationsSpec
