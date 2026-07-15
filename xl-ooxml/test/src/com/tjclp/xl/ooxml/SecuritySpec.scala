package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import java.io.ByteArrayOutputStream
import java.util.zip.{ZipOutputStream, ZipEntry}
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

/**
 * Security tests for XLSX reader
 *
 * Verifies protection against XXE (XML External Entity) attacks
 */
@SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
class SecuritySpec extends FunSuite:

  val tempDir: Path = Files.createTempDirectory("xl-sec-test-")

  override def afterAll(): Unit =
    // Clean up temp files
    Files
      .walk(tempDir)
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

    // Should fail with parse error (GH-350: the DOCTYPE itself is stripped pre-parse, so the
    // undeclared &xxe; reference is what fails — its entity definition is never honored)
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

  test("GH-350: entity defined in a stripped DOCTYPE subset is not honored") {
    val xml = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE x [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>
<x>&xxe;</x>"""
    XmlSecurity.parseSafe(xml, "test.xml") match
      case Left(err) =>
        assert(
          err.message.toLowerCase.contains("xxe") || err.message.toLowerCase.contains("entity"),
          s"Expected undeclared-entity failure, got: ${err.message}"
        )
      case Right(_) => fail("Entity defined in a stripped DOCTYPE subset must not resolve")
  }

  test("GH-350: stripLeadingDoctype removes a SYSTEM-form doctype from the prolog") {
    val in = "<?xml version=\"1.0\"?>\n<!DOCTYPE workbook SYSTEM \"wb.dtd\">\n<workbook/>"
    assertEquals(XmlSecurity.stripLeadingDoctype(in), "<?xml version=\"1.0\"?>\n\n<workbook/>")
  }

  test("GH-350: stripLeadingDoctype tracks quotes and bracket nesting in an internal subset") {
    val in = """<?xml version="1.0"?><!DOCTYPE x [ <!ENTITY e "tricky > ] value"> ]><x/>"""
    assertEquals(XmlSecurity.stripLeadingDoctype(in), """<?xml version="1.0"?><x/>""")
  }

  test("GH-350: stripLeadingDoctype skips prolog comments before the doctype") {
    val in = "<?xml version=\"1.0\"?><!-- produced by tool --><!DOCTYPE x SYSTEM \"x.dtd\"><x/>"
    assertEquals(
      XmlSecurity.stripLeadingDoctype(in),
      "<?xml version=\"1.0\"?><!-- produced by tool --><x/>"
    )
  }

  test("GH-350: stripLeadingDoctype skips comments inside the internal subset (apostrophe)") {
    // The apostrophe inside the comment must not open quote state and swallow the doctype end.
    val in = """<?xml version="1.0"?><!DOCTYPE x [ <!-- don't --> ]><x/>"""
    assertEquals(XmlSecurity.stripLeadingDoctype(in), """<?xml version="1.0"?><x/>""")
  }

  test("GH-350: stripLeadingDoctype skips comments inside the internal subset (bracket)") {
    // The ']' inside the comment must not close the subset early and truncate the strip.
    val in = """<?xml version="1.0"?><!DOCTYPE x [ <!-- ] --> ]><x/>"""
    assertEquals(XmlSecurity.stripLeadingDoctype(in), """<?xml version="1.0"?><x/>""")
  }

  test("GH-350: doctype with a comment-laden internal subset parses via parseSafe") {
    val xml = "<?xml version=\"1.0\"?>\n<!DOCTYPE x [ <!-- don't ] --> <!ELEMENT x ANY> ]>\n<x/>"
    XmlSecurity.parseSafe(xml, "test.xml") match
      case Right(elem) => assertEquals(elem.label, "x")
      case Left(err) => fail(s"comment-laden internal subset must parse: ${err.message}")
  }

  test("GH-350: stripLeadingDoctype strips a doctype behind a UTF-8 BOM (and drops the BOM)") {
    // Xerces rejects U+FEFF at the start of a character stream, so once bytes are decoded the
    // BOM char is junk: drop it along with the doctype.
    val in = "\uFEFF<?xml version=\"1.0\"?>\n<!DOCTYPE x SYSTEM \"x.dtd\">\n<x/>"
    assertEquals(XmlSecurity.stripLeadingDoctype(in), "<?xml version=\"1.0\"?>\n\n<x/>")
  }

  test("GH-350: BOM + DOCTYPE part parses via parseSafe") {
    val xml =
      "\uFEFF<?xml version=\"1.0\"?>\n<!DOCTYPE x SYSTEM \"http://example.invalid/x.dtd\">\n<x/>"
    XmlSecurity.parseSafe(xml, "test.xml") match
      case Right(elem) => assertEquals(elem.label, "x")
      case Left(err) => fail(s"BOM + DOCTYPE part must parse: ${err.message}")
  }

  test("GH-350: stripLeadingDoctypeStream strips a doctype (with BOM) at the byte level") {
    val xml = "\uFEFF<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
      "<!DOCTYPE x [ <!-- don't ] --> <!ENTITY e \"v\"> ]>\n<x>héllo</x>"
    val in = new java.io.ByteArrayInputStream(xml.getBytes(java.nio.charset.StandardCharsets.UTF_8))
    val out = new String(
      XmlSecurity.stripLeadingDoctypeStream(in).readAllBytes(),
      java.nio.charset.StandardCharsets.UTF_8
    )
    assertEquals(out, "\uFEFF<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n\n<x>héllo</x>")
  }

  test("GH-350: stripLeadingDoctypeStream passes doctype-free input through byte-identical") {
    val bytes = "<?xml version=\"1.0\"?>\n<x>日本語 🚀</x>"
      .getBytes(java.nio.charset.StandardCharsets.UTF_8)
    val out = XmlSecurity
      .stripLeadingDoctypeStream(new java.io.ByteArrayInputStream(bytes))
      .readAllBytes()
    assert(java.util.Arrays.equals(out, bytes), "doctype-free stream must be untouched")
  }

  test("GH-350: stripLeadingDoctype leaves content after the first element untouched") {
    val in = "<x><![CDATA[<!DOCTYPE y [ ]>]]></x>"
    assertEquals(XmlSecurity.stripLeadingDoctype(in), in)
  }

  test("GH-350: stripLeadingDoctype leaves an unterminated doctype for the parser to report") {
    val in = "<?xml version=\"1.0\"?><!DOCTYPE x [ <!ENTITY e \"v\">"
    assertEquals(XmlSecurity.stripLeadingDoctype(in), in)
  }

  test("GH-350: stripLeadingDoctype preserves newlines for accurate error line numbers") {
    val in = "<?xml version=\"1.0\"?>\n<!DOCTYPE x [\n<!ELEMENT x ANY>\n]>\n<x/>"
    assertEquals(XmlSecurity.stripLeadingDoctype(in), "<?xml version=\"1.0\"?>\n\n\n\n<x/>")
  }

  test("XlsxReader rejects comment relationship path traversal") {
    val workbookXml = """<?xml version="1.0" encoding="UTF-8"?>
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
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
</Types>"""

    val rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    val workbookRels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""

    val sheetRels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../../workbook.xml"/>
</Relationships>"""

    val sheet1 = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"""

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
    addEntry("xl/worksheets/sheet1.xml", sheet1)
    addEntry("xl/worksheets/_rels/sheet1.xml.rels", sheetRels)

    zos.close()
    val maliciousBytes = baos.toByteArray

    val tempFile = tempDir.resolve("comment-path-traversal.xlsx")
    Files.write(tempFile, maliciousBytes)

    val result = XlsxReader.read(tempFile)

    assert(
      result.isLeft,
      s"Reader should reject comment relationship path traversal, got: $result"
    )
  }
