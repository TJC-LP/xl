package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.macros.cell
import com.tjclp.xl.sheet.syntax.*
import com.tjclp.xl.style.{CellStyle, Font}
import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}
import java.nio.charset.StandardCharsets
import scala.collection.mutable

class XlsxReaderSpec extends FunSuite:

  test("XlsxReader preserves cell styles when reading") {
    val boldStyle = CellStyle.default.withFont(Font("Arial", 12.0, bold = true))

    val sheet = Sheet("Styled").getOrElse(fail("Failed to create sheet"))
      .put(cell"A1", CellValue.Text("styled"))
      .withCellStyle(cell"A1", boldStyle)

    val wb = Workbook(Vector(sheet))
    val bytes = XlsxWriter.writeToBytes(wb).getOrElse(fail("Failed to write workbook"))

    val readWb = XlsxReader.readFromBytes(bytes).getOrElse(fail("Failed to read workbook"))
    val readSheet = readWb("Styled").getOrElse(fail("Missing sheet"))

    val style = readSheet.getCellStyle(cell"A1")
    assert(style.nonEmpty, "Cell style should be present after read")
    assertEquals(style.map(_.font.bold), Some(true))
  }

  test("XlsxReader resolves worksheet paths via relationships") {
    val sheet = Sheet("Rel").getOrElse(fail("Failed to create sheet"))
      .put(cell"A1", CellValue.Text("ok"))

    val wb = Workbook(Vector(sheet))
    val bytes = XlsxWriter.writeToBytes(wb).getOrElse(fail("Failed to write workbook"))
    val mutated = rewriteWorksheetTarget(bytes, "sheet42.xml")

    val readWb = XlsxReader.readFromBytes(mutated).getOrElse(fail("Failed to read mutated workbook"))
    val readSheet = readWb("Rel").getOrElse(fail("Missing sheet"))

    assertEquals(readSheet(cell"A1").value, CellValue.Text("ok"))
  }

  private def rewriteWorksheetTarget(bytes: Array[Byte], newSheetFile: String): Array[Byte] =
    val zip = new ZipInputStream(new ByteArrayInputStream(bytes))
    val entries = mutable.LinkedHashMap.empty[String, Array[Byte]]
    var entry = zip.getNextEntry
    while entry != null do
      if !entry.isDirectory then
        entries(entry.getName) = zip.readAllBytes()
      zip.closeEntry()
      entry = zip.getNextEntry
    zip.close()

    val oldSheet = "xl/worksheets/sheet1.xml"
    val sheetBytes = entries.remove(oldSheet).getOrElse(fail(s"Missing $oldSheet in workbook"))
    val newSheetPath = s"xl/worksheets/$newSheetFile"
    entries(newSheetPath) = sheetBytes

    entries.get("xl/_rels/workbook.xml.rels").foreach { relBytes =>
      entries.update(
        "xl/_rels/workbook.xml.rels",
        replaceAll(relBytes, "worksheets/sheet1.xml", s"worksheets/$newSheetFile")
      )
    }

    entries.get("[Content_Types].xml").foreach { ctBytes =>
      entries.update(
        "[Content_Types].xml",
        replaceAll(ctBytes, "sheet1.xml", newSheetFile)
      )
    }

    val baos = new ByteArrayOutputStream()
    val zipOut = new ZipOutputStream(baos)
    entries.foreach { case (name, data) =>
      val newEntry = new ZipEntry(name)
      zipOut.putNextEntry(newEntry)
      zipOut.write(data)
      zipOut.closeEntry()
    }
    zipOut.close()
    baos.toByteArray

  private def replaceAll(bytes: Array[Byte], from: String, to: String): Array[Byte] =
    new String(bytes, StandardCharsets.UTF_8).replace(from, to).getBytes(StandardCharsets.UTF_8)
