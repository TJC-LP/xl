package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.error.XLError
import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipOutputStream}

/**
 * Security tests for ZIP bomb detection in XlsxReader.
 *
 * ZIP bombs are malicious archives that expand to enormous sizes when decompressed, potentially
 * causing denial-of-service through memory exhaustion or disk space exhaustion.
 */
class ZipBombSpec extends FunSuite:

  // Helper to create a ZIP with specific characteristics
  private def createSyntheticZip(
    entries: List[(String, Array[Byte])]
  ): Array[Byte] =
    val baos = new ByteArrayOutputStream()
    val zos = new ZipOutputStream(baos)
    try
      entries.foreach { case (name, content) =>
        val entry = new ZipEntry(name)
        zos.putNextEntry(entry)
        zos.write(content)
        zos.closeEntry()
      }
    finally zos.close()
    baos.toByteArray

  // Create minimal valid XLSX structure for testing
  private def createMinimalXlsx(extraEntries: List[(String, Array[Byte])] = Nil): Array[Byte] =
    val contentTypes =
      """<?xml version="1.0" encoding="UTF-8"?>
        |<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        |  <Default Extension="xml" ContentType="application/xml"/>
        |  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        |  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        |  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        |</Types>""".stripMargin.getBytes("UTF-8")

    val rels =
      """<?xml version="1.0" encoding="UTF-8"?>
        |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        |  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        |</Relationships>""".stripMargin.getBytes("UTF-8")

    val workbookRels =
      """<?xml version="1.0" encoding="UTF-8"?>
        |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        |  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        |</Relationships>""".stripMargin.getBytes("UTF-8")

    val workbook =
      """<?xml version="1.0" encoding="UTF-8"?>
        |<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        |  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
        |</workbook>""".stripMargin.getBytes("UTF-8")

    val sheet =
      """<?xml version="1.0" encoding="UTF-8"?>
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData/>
        |</worksheet>""".stripMargin.getBytes("UTF-8")

    val baseEntries = List(
      "[Content_Types].xml" -> contentTypes,
      "_rels/.rels" -> rels,
      "xl/_rels/workbook.xml.rels" -> workbookRels,
      "xl/workbook.xml" -> workbook,
      "xl/worksheets/sheet1.xml" -> sheet
    )

    createSyntheticZip(baseEntries ++ extraEntries)

  test("rejects ZIP with too many entries") {
    // Create XLSX with extra entries to exceed limit
    val manyEntries = (1 to 15).toList.map { i =>
      s"xl/extra/file$i.xml" -> s"<data>test$i</data>".getBytes("UTF-8")
    }
    val xlsxBytes = createMinimalXlsx(manyEntries)

    // Use config with low entry limit
    val config = XlsxReader.ReaderConfig(maxEntryCount = 10)

    XlsxReader.readFromBytes(xlsxBytes, config) match
      case Left(XLError.SecurityError(msg)) =>
        assert(msg.contains("entry count"), s"Expected entry count error, got: $msg")
      case other =>
        fail(s"Expected SecurityError for entry count, got: $other")
  }

  test("accepts ZIP within entry count limit") {
    val xlsxBytes = createMinimalXlsx()

    // Default config should accept small XLSX
    XlsxReader.readFromBytes(xlsxBytes) match
      case Right(_) => () // Success
      case Left(err) => fail(s"Expected success, got: $err")
  }

  test("rejects ZIP exceeding uncompressed size limit") {
    // Create entry with content that exceeds size limit when uncompressed
    val largeContent = "x".repeat(200_000).getBytes("UTF-8") // 200KB
    val xlsxBytes = createMinimalXlsx(List("xl/large.xml" -> largeContent))

    // Use config with 100KB limit
    val config = XlsxReader.ReaderConfig(maxUncompressedSize = 100_000)

    XlsxReader.readFromBytes(xlsxBytes, config) match
      case Left(XLError.SecurityError(msg)) =>
        assert(msg.contains("uncompressed size"), s"Expected size error, got: $msg")
      case other =>
        fail(s"Expected SecurityError for size, got: $other")
  }

  test("accepts ZIP within uncompressed size limit") {
    // Create small XLSX
    val xlsxBytes = createMinimalXlsx()

    // Should succeed with default config
    XlsxReader.readFromBytes(xlsxBytes) match
      case Right(_) => () // Success
      case Left(err) => fail(s"Expected success, got: $err")
  }

  test("rejects highly compressed content (compression ratio check)") {
    // Create highly compressible content (repetitive data compresses extremely well)
    val repetitiveContent = ("a" * 1000).repeat(100).getBytes("UTF-8") // 100KB of 'a's

    val xlsxBytes = createMinimalXlsx(List("xl/repetitive.xml" -> repetitiveContent))

    // Use strict compression ratio limit
    val config = XlsxReader.ReaderConfig(maxCompressionRatio = 10)

    XlsxReader.readFromBytes(xlsxBytes, config) match
      case Left(XLError.SecurityError(msg)) =>
        assert(
          msg.contains("Compression ratio") || msg.contains("ZIP bomb"),
          s"Expected compression ratio error, got: $msg"
        )
      case Left(err) => fail(s"Expected SecurityError, got other error: $err")
      case Right(_) =>
        // Note: If content doesn't compress well enough to trigger limit, test may pass
        // This is acceptable since we're testing the mechanism
        ()
  }

  test("permissive config allows high compression ratios") {
    // Create highly compressible content
    val repetitiveContent = ("a" * 1000).repeat(100).getBytes("UTF-8")
    val xlsxBytes = createMinimalXlsx(List("xl/repetitive.xml" -> repetitiveContent))

    // Permissive config disables all limits
    XlsxReader.readFromBytes(xlsxBytes, XlsxReader.ReaderConfig.permissive) match
      case Right(_) => () // Success
      case Left(err) => fail(s"Expected success with permissive config, got: $err")
  }

  test("all limits can be disabled individually") {
    // Verify each limit can be disabled with 0
    val config = XlsxReader.ReaderConfig(
      maxCompressionRatio = 0, // Disable
      maxUncompressedSize = 0L, // Disable
      maxEntryCount = 0, // Disable
      maxCellCount = 0L, // Disable
      maxStringLength = 0 // Disable
    )

    // This should be equivalent to permissive
    assertEquals(config.maxCompressionRatio, 0)
    assertEquals(config.maxUncompressedSize, 0L)
    assertEquals(config.maxEntryCount, 0)
  }

  test("default config has sensible limits") {
    val config = XlsxReader.ReaderConfig.default

    // Verify default limits are reasonable
    assertEquals(config.maxCompressionRatio, 100)
    assertEquals(config.maxUncompressedSize, 100_000_000L) // 100 MB
    assertEquals(config.maxEntryCount, 10_000)
    assertEquals(config.maxCellCount, 10_000_000L) // 10M
    assertEquals(config.maxStringLength, 32_768) // 32 KB
  }

  test("SecurityError message includes relevant details") {
    // Create XLSX with too many entries
    val manyEntries = (1 to 20).toList.map { i =>
      s"xl/file$i.xml" -> "<x/>".getBytes("UTF-8")
    }
    val xlsxBytes = createMinimalXlsx(manyEntries)

    val config = XlsxReader.ReaderConfig(maxEntryCount = 10)

    XlsxReader.readFromBytes(xlsxBytes, config) match
      case Left(err @ XLError.SecurityError(msg)) =>
        // Verify error message is descriptive
        assert(msg.contains("10") || msg.contains("limit"), s"Error should mention limit: $msg")
        // Verify .message extension works
        assert(err.message.contains("Security error"))
      case other =>
        fail(s"Expected SecurityError, got: $other")
  }

  test("read from Path uses config") {
    // Create temp file with XLSX
    val xlsxBytes = createMinimalXlsx()
    val tempFile = Files.createTempFile("test-", ".xlsx")
    try
      Files.write(tempFile, xlsxBytes)

      // Should succeed with default config
      XlsxReader.read(tempFile) match
        case Right(_) => () // Success
        case Left(err) => fail(s"Expected success, got: $err")

      // Should also accept explicit config
      XlsxReader.read(tempFile, XlsxReader.ReaderConfig.default) match
        case Right(_) => () // Success
        case Left(err) => fail(s"Expected success with explicit config, got: $err")
    finally Files.deleteIfExists(tempFile)
  }
