package com.tjclp.xl.ooxml.metadata

import java.nio.charset.StandardCharsets.UTF_8
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.jdk.CollectionConverters.*

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.ooxml.{TestFixtures, XlsxWriter}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

// Test code uses .get/.head for brevity in assertions
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class WorkbookMetadataReaderSpec extends FunSuite:

  // Helper to create a temp xlsx file with given sheets
  private def createTempWorkbook(sheets: Vector[Sheet]): Path =
    val path = Files.createTempFile("metadata-test-", ".xlsx")
    val wb = Workbook(sheets)
    XlsxWriter.write(wb, path)
    path

  /** Copy the zip at `from` to `to`, each entry's bytes passed through `f` (entry name, bytes). */
  private def rewriteZip(from: Path, to: Path)(f: (String, Array[Byte]) => Array[Byte]): Unit =
    val zip = new ZipFile(from.toFile)
    val entries =
      try
        zip.entries().asScala.toVector.map { entry =>
          val in = zip.getInputStream(entry)
          try entry.getName -> in.readAllBytes()
          finally in.close()
        }
      finally zip.close()
    val out = new ZipOutputStream(Files.newOutputStream(to))
    try
      entries.foreach { (name, bytes) =>
        out.putNextEntry(new ZipEntry(name))
        out.write(f(name, bytes))
        out.closeEntry()
      }
    finally out.close()

  private def zipEntryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      val in = zip.getInputStream(zip.getEntry(name))
      try new String(in.readAllBytes(), UTF_8)
      finally in.close()
    finally zip.close()

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

  test("read: resolves openpyxl's package-absolute worksheet Target to the sheet's <dimension>") {
    // A1:B2 hold values and D5 is styled but empty, so the <dimension> the writer records (every
    // stored cell: A1:D5) is wider than any scan of the non-empty cells could recover — the
    // streaming used range depends on this element being found
    val sheet = Sheet(SheetName.unsafe("Data")).put("A1" -> "h", "B1" -> 1, "A2" -> "x", "B2" -> 2)
    val styled = sheet.styleAt("D5", CellStyle.default.withNumFmt(NumFmt.Percent)).getOrElse(sheet)
    val relative = createTempWorkbook(Vector(styled))
    val absolute = Files.createTempFile("metadata-test-openpyxl-", ".xlsx")
    try
      val written = WorkbookMetadataReader.read(relative).toOption.get.sheets.head.dimension
      assertEquals(written.map(_.toA1), Some("A1:D5"))
      // openpyxl writes `Target="/xl/worksheets/sheet1.xml"`; the library (like Excel) writes the
      // Target relative to xl/. Same zip otherwise.
      rewriteZip(relative, absolute) { (name, bytes) =>
        if name == "xl/_rels/workbook.xml.rels" then
          new String(bytes, UTF_8)
            .replace("Target=\"worksheets/", "Target=\"/xl/worksheets/")
            .getBytes(UTF_8)
        else bytes
      }
      val rels = zipEntryText(absolute, "xl/_rels/workbook.xml.rels")
      assert(rels.contains("Target=\"/xl/worksheets/sheet1.xml\""), rels)
      val meta = WorkbookMetadataReader.read(absolute).toOption.get
      assert(meta.sheets.head.dimension.isDefined, "absolute Target lost the <dimension>")
      assertEquals(meta.sheets.head.dimension, written)
    finally
      Files.deleteIfExists(relative)
      Files.deleteIfExists(absolute)
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

  test("read: extracts date1904 flag from workbookPr (GH-243)") {
    val sheets = Vector(Sheet(SheetName.unsafe("Dates")))
    val path = Files.createTempFile("metadata-test-1904-", ".xlsx")
    val wb = Workbook(sheets)
    XlsxWriter.write(wb.copy(metadata = wb.metadata.copy(date1904 = true)), path)
    try
      val meta = WorkbookMetadataReader.read(path).toOption.get
      assert(meta.date1904, "expected date1904 = true from written workbookPr")
    finally Files.deleteIfExists(path)
  }

  test("read: date1904 defaults to false for 1900-system workbooks (GH-243)") {
    val path = createTempWorkbook(Vector(Sheet(SheetName.unsafe("Plain"))))
    try
      val meta = WorkbookMetadataReader.read(path).toOption.get
      assert(!meta.date1904)
    finally Files.deleteIfExists(path)
  }

  test("read: malformed workbook.xml surfaces the human-readable Xerces message (GH-349)") {
    // Fixture: small-values.xlsx with the closing </workbook> tag mangled. The parse failure
    // must carry the real Xerces diagnostic (text backed by the XMLMessages resource bundle) —
    // under native-image an unregistered bundle degrades this to "Could not load any resource
    // bundle by com.sun.org.apache.xerces.internal.impl.msg.XMLMessages", masking the real
    // error. The same fixture is exercised against the native binary in release.yml.
    val path = TestFixtures.copyToTemp("malformed-workbook.xlsx")
    val result = WorkbookMetadataReader.read(path)
    assert(result.isLeft, s"Expected Left for malformed workbook, got $result")
    val msg = result.left.toOption.get.message
    // Full Xerces text (XMLMessages ETagUnterminated):
    //   The end-tag for element type "workbook" must end with a '>' delimiter.
    assert(
      msg.contains("end-tag for element type \"workbook\""),
      s"expected the Xerces well-formedness diagnostic, got: $msg"
    )
    assert(
      !msg.contains("Could not load any resource bundle"),
      s"resource-bundle lookup failure leaked into the parse error: $msg"
    )
  }
