package com.tjclp.xl.ooxml

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import munit.FunSuite

import java.nio.file.{Files, Path}
import java.util.zip.ZipInputStream
import scala.io.Source

/**
 * Regression tests for StAX namespace handling.
 *
 * These tests verify that the StAX backend correctly emits xmlns declarations and that generated
 * files are compatible with external tools (POI, openpyxl).
 *
 * Background: StAX with IS_REPAIRING_NAMESPACES=false does not automatically emit xmlns
 * declarations when calling writeStartElement with a namespace URI. This caused POI to fail with
 * "document element namespace mismatch" and openpyxl to fail looking for sharedStrings.xml.
 *
 * @see
 *   https://github.com/TJC-LP/xl/pull/145
 */
class StaxNamespaceRegressionSpec extends FunSuite:

  private def createTestWorkbook(): Workbook =
    val sheet = Sheet("Test")
      .put("A1" -> "Hello", "B1" -> "World")
      .put("A2" -> 42, "B2" -> 3.14)
    Workbook(sheet)

  private def writeToTempAndReadBytes(wb: Workbook, config: WriterConfig): Array[Byte] =
    val temp = Files.createTempFile("xl-test-", ".xlsx")
    try
      XlsxWriter.writeWith(wb, temp, config) match
        case Right(_) => Files.readAllBytes(temp)
        case Left(err) => fail(s"Write failed: $err")
    finally Files.deleteIfExists(temp)

  private def extractEntry(bytes: Array[Byte], entryName: String): Option[String] =
    val zis = new ZipInputStream(new java.io.ByteArrayInputStream(bytes))
    try
      var entry = zis.getNextEntry
      while entry != null do
        if entry.getName == entryName then
          val content = Source.fromInputStream(zis, "UTF-8").mkString
          zis.closeEntry()
          return Some(content)
        zis.closeEntry()
        entry = zis.getNextEntry
      None
    finally zis.close()

  test("StAX backend: styles.xml has xmlns declaration (POI compatibility)") {
    val wb = createTestWorkbook()
    val bytes = writeToTempAndReadBytes(wb, WriterConfig(backend = XmlBackend.SaxStax))

    val stylesXml = extractEntry(bytes, "xl/styles.xml")
      .getOrElse(fail("xl/styles.xml not found"))

    assert(
      stylesXml.contains("""xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""""),
      s"styles.xml missing xmlns declaration:\n$stylesXml"
    )
  }

  test("StAX backend: workbook.xml has xmlns declaration") {
    val wb = createTestWorkbook()
    val bytes = writeToTempAndReadBytes(wb, WriterConfig(backend = XmlBackend.SaxStax))

    val workbookXml = extractEntry(bytes, "xl/workbook.xml")
      .getOrElse(fail("xl/workbook.xml not found"))

    assert(
      workbookXml.contains("""xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""""),
      s"workbook.xml missing xmlns declaration:\n$workbookXml"
    )
  }

  test("StAX backend: worksheet has xmlns declaration") {
    val wb = createTestWorkbook()
    val bytes = writeToTempAndReadBytes(wb, WriterConfig(backend = XmlBackend.SaxStax))

    val sheetXml = extractEntry(bytes, "xl/worksheets/sheet1.xml")
      .getOrElse(fail("xl/worksheets/sheet1.xml not found"))

    assert(
      sheetXml.contains("""xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""""),
      s"worksheet missing xmlns declaration:\n$sheetXml"
    )
  }

  test("StAX backend: relationships have xmlns declaration") {
    val wb = createTestWorkbook()
    val bytes = writeToTempAndReadBytes(wb, WriterConfig(backend = XmlBackend.SaxStax))

    val relsXml = extractEntry(bytes, "xl/_rels/workbook.xml.rels")
      .getOrElse(fail("xl/_rels/workbook.xml.rels not found"))

    assert(
      relsXml.contains("""xmlns="http://schemas.openxmlformats.org/package/2006/relationships""""),
      s"workbook.xml.rels missing xmlns declaration:\n$relsXml"
    )
  }

  test("StAX backend: no sharedStrings relationship when SST not used (openpyxl compatibility)") {
    // Small workbook with unique strings - should NOT use SST
    val wb = createTestWorkbook()
    val bytes = writeToTempAndReadBytes(wb, WriterConfig(backend = XmlBackend.SaxStax))

    val relsXml = extractEntry(bytes, "xl/_rels/workbook.xml.rels")
      .getOrElse(fail("xl/_rels/workbook.xml.rels not found"))

    // Should NOT have sharedStrings relationship
    assert(
      !relsXml.contains("sharedStrings"),
      s"workbook.xml.rels should not reference sharedStrings for small workbook:\n$relsXml"
    )

    // Should NOT have sharedStrings.xml file
    val hasSst = extractEntry(bytes, "xl/sharedStrings.xml").isDefined
    assert(!hasSst, "sharedStrings.xml should not exist for small workbook")
  }

  test("StAX backend: chained writes don't produce duplicate xmlns (regression)") {
    val temp = Files.createTempFile("xl-stax-chained-", ".xlsx")
    try
      // First write
      val wb1 = Workbook(Sheet("Sheet1").put("A1" -> "First"))
      XlsxWriter.writeWith(wb1, temp, WriterConfig(backend = XmlBackend.SaxStax)) match
        case Left(err) => fail(s"First write failed: $err")
        case Right(_) => ()

      // Read back
      val read1 = XlsxReader.read(temp).fold(err => fail(s"First read failed: $err"), identity)
      assertEquals(read1.sheets.size, 1)

      // Modify and write again (chained write)
      val wb2 = read1.put(Sheet("Sheet2").put("A1" -> "Second"))
      XlsxWriter.writeWith(wb2, temp, WriterConfig(backend = XmlBackend.SaxStax)) match
        case Left(e) => fail(s"Second write failed: $e")
        case Right(_) => ()

      // Read back again - this would fail with "xmlns already specified" before the fix
      val read2 = XlsxReader.read(temp).fold(e => fail(s"Second read failed: $e"), identity)
      assertEquals(read2.sheets.size, 2)

      // Third chained write
      val wb3 = read2.put(Sheet("Sheet3").put("A1" -> "Third"))
      XlsxWriter.writeWith(wb3, temp, WriterConfig(backend = XmlBackend.SaxStax)) match
        case Left(e) => fail(s"Third write failed: $e")
        case Right(_) => ()

      // Final verification
      val read3 = XlsxReader.read(temp).fold(err => fail(s"Third read failed: $err"), identity)
      assertEquals(read3.sheets.size, 3)
      assertEquals(read3.sheets.map(_.name.value), Vector("Sheet1", "Sheet2", "Sheet3"))
    finally Files.deleteIfExists(temp)
  }

  test("StAX backend: output matches ScalaXml backend structure") {
    val wb = createTestWorkbook()

    val staxBytes = writeToTempAndReadBytes(wb, WriterConfig(backend = XmlBackend.SaxStax))
    val scalaXmlBytes = writeToTempAndReadBytes(wb, WriterConfig(backend = XmlBackend.ScalaXml))

    // Both should have the same parts
    def listEntries(bytes: Array[Byte]): Set[String] =
      val zis = new ZipInputStream(new java.io.ByteArrayInputStream(bytes))
      try
        var entries = Set.empty[String]
        var entry = zis.getNextEntry
        while entry != null do
          entries += entry.getName
          zis.closeEntry()
          entry = zis.getNextEntry
        entries
      finally zis.close()

    val staxEntries = listEntries(staxBytes)
    val scalaXmlEntries = listEntries(scalaXmlBytes)

    assertEquals(
      staxEntries,
      scalaXmlEntries,
      s"StAX and ScalaXml should produce same ZIP entries"
    )

    // Both should have xmlns in styles.xml
    val staxStyles = extractEntry(staxBytes, "xl/styles.xml").get
    val scalaXmlStyles = extractEntry(scalaXmlBytes, "xl/styles.xml").get

    assert(
      staxStyles.contains("xmlns="),
      "StAX styles.xml should have xmlns"
    )
    assert(
      scalaXmlStyles.contains("xmlns="),
      "ScalaXml styles.xml should have xmlns"
    )
  }

  test("StAX backend: round-trip preserves all data") {
    val original = Workbook(
      Sheet("Data")
        .put("A1" -> "Text", "B1" -> 123, "C1" -> 45.67, "D1" -> true)
        .put("A2" -> "More", "B2" -> -100, "C2" -> 0.001, "D2" -> false)
    )

    val temp = Files.createTempFile("xl-stax-roundtrip-", ".xlsx")
    try
      // Write with StAX
      XlsxWriter.writeWith(original, temp, WriterConfig(backend = XmlBackend.SaxStax)) match
        case Left(err) => fail(s"Write failed: $err")
        case Right(_) => ()

      // Read back
      val loaded = XlsxReader.read(temp).fold(err => fail(s"Read failed: $err"), identity)

      // Verify data
      assertEquals(loaded.sheets.size, 1)
      val sheet = loaded.sheets.head
      assertEquals(sheet.name.value, "Data")

      // Check cell values
      assertEquals(sheet.cells.get(ref"A1").map(_.value), Some(CellValue.Text("Text")))
      assertEquals(sheet.cells.get(ref"B1").map(_.value), Some(CellValue.Number(123)))
      assertEquals(sheet.cells.get(ref"D1").map(_.value), Some(CellValue.Bool(true)))
    finally Files.deleteIfExists(temp)
  }

  test("StAX backend: POI-compatible structure (all parts have correct xmlns)") {
    // This test verifies the file structure is POI-compatible
    // POI requires proper xmlns on all OOXML parts
    val wb = createTestWorkbook()
    val bytes = writeToTempAndReadBytes(wb, WriterConfig(backend = XmlBackend.SaxStax))

    // Check all critical parts have correct xmlns
    val parts = Seq(
      "xl/styles.xml" -> "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "xl/workbook.xml" -> "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "xl/worksheets/sheet1.xml" -> "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "_rels/.rels" -> "http://schemas.openxmlformats.org/package/2006/relationships",
      "xl/_rels/workbook.xml.rels" -> "http://schemas.openxmlformats.org/package/2006/relationships"
    )

    parts.foreach { case (path, expectedNs) =>
      val content = extractEntry(bytes, path).getOrElse(fail(s"$path not found"))
      assert(
        content.contains(s"""xmlns="$expectedNs""""),
        s"$path missing or has wrong xmlns. Expected $expectedNs in:\n$content"
      )
    }
  }
