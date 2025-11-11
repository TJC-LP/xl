package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.*
import java.io.ByteArrayOutputStream
import java.util.zip.{ZipOutputStream, ZipEntry}
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

/**
 * Security tests for XLSX reader
 *
 * Verifies protection against XXE (XML External Entity) attacks
 */
class SecuritySpec extends FunSuite:

  val tempDir: Path = Files.createTempDirectory("xl-sec-test-")

  override def afterAll(): Unit =
    // Clean up temp files
    Files.walk(tempDir)
      .sorted(java.util.Comparator.reverseOrder())
      .forEach(Files.delete)

  test("XlsxReader rejects malicious XLSX with DOCTYPE declaration") {
    // Create a malicious XLSX with XXE payload in workbook.xml
    val maliciousWorkbookXml = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
  </sheets>
  <value>&xxe;</value>
</workbook>"""

    val contentTypes = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""

    val rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    val workbookRels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""

    val sheet1 = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"""

    // Build malicious XLSX ZIP
    val baos = ByteArrayOutputStream()
    val zos = ZipOutputStream(baos)

    def addEntry(name: String, content: String): Unit =
      zos.putNextEntry(ZipEntry(name))
      zos.write(content.getBytes(StandardCharsets.UTF_8))
      zos.closeEntry()

    addEntry("[Content_Types].xml", contentTypes)
    addEntry("_rels/.rels", rels)
    addEntry("xl/workbook.xml", maliciousWorkbookXml) // Malicious!
    addEntry("xl/_rels/workbook.xml.rels", workbookRels)
    addEntry("xl/worksheets/sheet1.xml", sheet1)

    zos.close()
    val maliciousBytes = baos.toByteArray

    // Write to temp file
    val tempFile = tempDir.resolve("malicious-doctype.xlsx")
    Files.write(tempFile, maliciousBytes)

    // Attempt to read malicious XLSX
    val result = XlsxReader.read(tempFile)

    // Should fail with parse error (DOCTYPE rejected)
    result match
      case Left(error) =>
        val errorMsg = error.toString.toLowerCase
        assert(
          errorMsg.contains("parse") || errorMsg.contains("doctype") || errorMsg.contains("xml"),
          s"Should reject XXE with parse error, got: $error"
        )
      case Right(_) =>
        fail("XlsxReader should reject XLSX with DOCTYPE declaration (XXE vulnerability)")
  }

  test("XlsxReader rejects external entity references") {
    // Create XLSX with external entity reference in worksheet
    val maliciousSheetXml = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE foo [<!ENTITY xxe SYSTEM "http://attacker.com/evil">]>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is><t>&xxe;</t></is>
      </c>
    </row>
  </sheetData>
</worksheet>"""

    val contentTypes = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""

    val rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    val workbookXml = """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
  </sheets>
</workbook>"""

    val workbookRels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""

    // Build malicious XLSX ZIP
    val baos = ByteArrayOutputStream()
    val zos = ZipOutputStream(baos)

    def addEntry(name: String, content: String): Unit =
      zos.putNextEntry(ZipEntry(name))
      zos.write(content.getBytes(StandardCharsets.UTF_8))
      zos.closeEntry()

    addEntry("[Content_Types].xml", contentTypes)
    addEntry("_rels/.rels", rels)
    addEntry("xl/workbook.xml", workbookXml)
    addEntry("xl/_rels/workbook.xml.rels", workbookRels)
    addEntry("xl/worksheets/sheet1.xml", maliciousSheetXml) // Malicious!

    zos.close()
    val maliciousBytes = baos.toByteArray

    // Write to temp file
    val tempFile = tempDir.resolve("malicious-entity.xlsx")
    Files.write(tempFile, maliciousBytes)

    // Attempt to read malicious XLSX
    val result = XlsxReader.read(tempFile)

    // Should fail with parse error (external entities rejected)
    result match
      case Left(error) =>
        val errorMsg = error.toString.toLowerCase
        assert(
          errorMsg.contains("parse") || errorMsg.contains("xml") || errorMsg.contains("entity"),
          s"Should reject external entity reference, got: $error"
        )
      case Right(_) =>
        fail("XlsxReader should reject XLSX with external entity references (XXE vulnerability)")
  }

  test("XlsxReader successfully parses legitimate XLSX without DOCTYPE") {
    // Verify that XXE protection doesn't break legitimate files
    val legitimateWorkbookXml = """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
  </sheets>
</workbook>"""

    val contentTypes = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""

    val rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    val workbookRels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""

    val sheet1 = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is><t>Safe content</t></is>
      </c>
    </row>
  </sheetData>
</worksheet>"""

    // Build legitimate XLSX ZIP
    val baos = ByteArrayOutputStream()
    val zos = ZipOutputStream(baos)

    def addEntry(name: String, content: String): Unit =
      zos.putNextEntry(ZipEntry(name))
      zos.write(content.getBytes(StandardCharsets.UTF_8))
      zos.closeEntry()

    addEntry("[Content_Types].xml", contentTypes)
    addEntry("_rels/.rels", rels)
    addEntry("xl/workbook.xml", legitimateWorkbookXml)
    addEntry("xl/_rels/workbook.xml.rels", workbookRels)
    addEntry("xl/worksheets/sheet1.xml", sheet1)

    zos.close()
    val legitimateBytes = baos.toByteArray

    // Write to temp file
    val tempFile = tempDir.resolve("legitimate.xlsx")
    Files.write(tempFile, legitimateBytes)

    // Should successfully parse legitimate XLSX
    val result = XlsxReader.read(tempFile)

    result match
      case Right(workbook) =>
        assertEquals(workbook.sheets.size, 1, "Should have 1 sheet")
        assertEquals(workbook.sheets.head.name.value, "Sheet1", "Sheet name should be Sheet1")
        assertEquals(workbook.sheets.head.cells.size, 1, "Should have 1 cell")
      case Left(error) =>
        fail(s"Legitimate XLSX should parse successfully, got error: $error")
  }
