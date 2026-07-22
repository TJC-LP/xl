package com.tjclp.xl.ooxml

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipOutputStream}

import scala.collection.immutable.ArraySeq

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.context.SourceContent
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * Tests for XlsxReader surgical modification support (Phase 3).
 *
 * Verifies that:
 *   1. Reading from file creates SourceContext with indexed unknown parts
 *   2. Reading from bytes creates an in-memory SourceContext (GH-412) with the same preservation
 *      semantics — write(readFromBytes(bytes)) is byte-identical to write(read(path))
 *   3. Unknown parts are indexed but NOT loaded into memory (file reads)
 *   4. PartManifest has accurate metadata
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class XlsxReaderPassthroughSpec extends FunSuite:

  test("read from file creates SourceContext") {
    // Create minimal XLSX with one known part and one unknown part
    val path = createMinimalWorkbookWithChart()

    val wb = XlsxReader
      .read(path)
      .fold(err => fail(s"Failed to read workbook: $err"), identity)

    // Should have SourceContext since read from file
    assert(wb.sourceContext.isDefined, "Reading from file should create SourceContext")

    val ctx = wb.sourceContext.get
    assertEquals(ctx.content, SourceContent.OnDisk(path))
    assert(ctx.isClean, "Freshly read workbook should have clean tracker")

    // PartManifest should contain both known and unknown parts
    val manifest = ctx.partManifest
    assert(manifest.parsedParts.contains("xl/workbook.xml"))
    assert(manifest.unparsedParts.contains("xl/charts/chart1.xml"))

    // Clean up
    Files.deleteIfExists(path)
  }

  test("read from bytes creates an in-memory SourceContext (GH-412)") {
    val bytes = createMinimalWorkbookBytes()

    val wb = XlsxReader
      .readFromBytes(bytes)
      .fold(err => fail(s"Failed to read workbook: $err"), identity)

    // The archive is already resident, so bytes-based reads preserve like path-based reads
    assert(wb.sourceContext.isDefined, "Reading from bytes should create SourceContext")

    val ctx = wb.sourceContext.get
    assertEquals(ctx.content, SourceContent.InMemory(ArraySeq.unsafeWrapArray(bytes)))
    assert(ctx.isClean, "Freshly read workbook should have clean tracker")
    assertEquals(ctx.fingerprint.size, bytes.length.toLong)
    assert(ctx.partManifest.parsedParts.contains("xl/workbook.xml"))
  }

  test("unknown parts are indexed but not loaded") {
    // Create workbook with a large unknown part (chart XML)
    val largeChart = "<chart>" + ("x" * 1024 * 1024) + "</chart>" // 1MB chart
    val path = createWorkbookWithLargeChart(largeChart)

    // Use permissive config since this test creates highly compressible synthetic data
    // that would trigger ZIP bomb detection (not a security test)
    val wb = XlsxReader
      .read(path, XlsxReader.ReaderConfig.permissive)
      .fold(err => fail(s"Failed to read workbook: $err"), identity)

    // Verify SourceContext exists
    assert(wb.sourceContext.isDefined)

    val manifest = wb.sourceContext.get.partManifest

    // Chart should be in unparsed parts (indexed but not loaded)
    assert(manifest.unparsedParts.contains("xl/charts/chart1.xml"))

    // Verify metadata captured
    val chartEntry = manifest.entries.get("xl/charts/chart1.xml")
    assert(chartEntry.isDefined, "Chart entry should be in manifest")
    assert(!chartEntry.get.parsed, "Chart should NOT be marked as parsed")

    // Note: size may not be available from ZipInputStream (returns -1 for DEFLATED entries)
    // This is expected behavior - size is optional metadata

    // Clean up
    Files.deleteIfExists(path)
  }

  test("PartManifest distinguishes parsed vs unparsed parts") {
    val path = createMinimalWorkbookWithChart()

    val wb = XlsxReader
      .read(path)
      .fold(err => fail(s"Failed to read workbook: $err"), identity)

    val manifest = wb.sourceContext.get.partManifest

    // Parsed parts (known to XL)
    val expectedParsed = Set(
      "xl/workbook.xml",
      "xl/worksheets/sheet1.xml",
      "xl/_rels/workbook.xml.rels",
      "_rels/.rels",
      "[Content_Types].xml"
    )
    assert(expectedParsed.subsetOf(manifest.parsedParts))

    // Unparsed parts (unknown to XL - should be preserved)
    assert(manifest.unparsedParts.contains("xl/charts/chart1.xml"))

    // Clean up
    Files.deleteIfExists(path)
  }

  test("backwards compatibility: existing tests using readFromBytes work unchanged") {
    // This test verifies that existing code that uses readFromBytes continues to work
    val bytes = createMinimalWorkbookBytes()

    val result = XlsxReader.readFromBytes(bytes)

    assert(result.isRight, "readFromBytes should succeed")
    val wb = result.getOrElse(fail("Expected Right"))
    assert(wb.sheets.nonEmpty, "Workbook should have sheets")
    // GH-412: bytes-based reads now carry an in-memory SourceContext (preservation parity)
    assert(wb.sourceContext.isDefined, "Bytes-based read should create in-memory SourceContext")
  }

  // Helper: Create minimal XLSX with one worksheet and one chart (unknown part)
  private def createMinimalWorkbookWithChart(): Path =
    val path = Files.createTempFile("test-workbook-chart", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))

    try
      // Content Types
      writeEntry(
        out,
        "[Content_Types].xml",
        """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>"""
      )

      // Root rels
      writeEntry(
        out,
        "_rels/.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
      )

      // Workbook
      writeEntry(
        out,
        "xl/workbook.xml",
        """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
      )

      // Workbook rels
      writeEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
      )

      // Worksheet
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Test</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

      // Chart (unknown part - should be preserved)
      writeEntry(
        out,
        "xl/charts/chart1.xml",
        """<?xml version="1.0"?>
<chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <chart>
    <plotArea/>
  </chart>
</chartSpace>"""
      )

    finally out.close()

    path

  // Helper: Create minimal XLSX bytes without unknown parts
  private def createMinimalWorkbookBytes(): Array[Byte] =
    val baos = new ByteArrayOutputStream()
    val out = new ZipOutputStream(baos)

    try
      // Content Types
      writeEntry(
        out,
        "[Content_Types].xml",
        """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
      )

      // Root rels
      writeEntry(
        out,
        "_rels/.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
      )

      // Workbook
      writeEntry(
        out,
        "xl/workbook.xml",
        """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
      )

      // Workbook rels
      writeEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
      )

      // Worksheet
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Test</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

    finally out.close()

    baos.toByteArray

  // Helper: Create workbook with large chart
  private def createWorkbookWithLargeChart(chartContent: String): Path =
    val path = Files.createTempFile("test-large-chart", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))

    try
      // Same structure as createMinimalWorkbookWithChart but with custom chart content
      writeEntry(
        out,
        "[Content_Types].xml",
        """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>"""
      )

      writeEntry(
        out,
        "_rels/.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
      )

      writeEntry(
        out,
        "xl/workbook.xml",
        """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
      )

      writeEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
      )

      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Test</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

      // Large chart (unknown part)
      writeEntry(out, "xl/charts/chart1.xml", chartContent)

    finally out.close()

    path

  // Helper: Write ZIP entry
  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    val entry = new ZipEntry(name)
    out.putNextEntry(entry)
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()

  // ===== GH-412: bytes-read/path-read write-fidelity law =====
  //
  // A workbook read from a byte array must write EXACTLY like the same workbook read from a
  // path: preserved-but-unmodeled workbook children (externalReferences, workbookProtection,
  // pivotCaches) and unknown parts (externalLink, pivotCacheDefinition) ride through both.

  private val nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  private val nsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  /**
   * Preservation-rich package: every workbook.xml child class from the GH-397 field incident
   * (workbookProtection, externalReferences, pivotCaches) plus the unknown parts behind them.
   */
  private def preservationRichParts: Seq[(String, String)] = Seq(
    "[Content_Types].xml" ->
      """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
</Types>""",
    "_rels/.rels" ->
      """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
    "xl/workbook.xml" ->
      s"""<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="$nsMain" xmlns:r="$nsRel">
  <fileVersion appName="xl"/>
  <workbookPr/>
  <workbookProtection lockStructure="1"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <externalReferences><externalReference r:id="rId2"/></externalReferences>
  <definedNames><definedName name="MyName">Sheet1!$$A$$1</definedName></definedNames>
  <calcPr calcId="191029"/>
  <pivotCaches><pivotCache cacheId="0" r:id="rId5"/></pivotCaches>
</workbook>""",
    "xl/_rels/workbook.xml.rels" ->
      """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>""",
    "xl/worksheets/sheet1.xml" ->
      s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>""",
    "xl/styles.xml" ->
      s"""<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="$nsMain">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>""",
    "xl/sharedStrings.xml" ->
      s"""<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="$nsMain" count="1" uniqueCount="1"><si><t>Hello</t></si></sst>""",
    "xl/externalLinks/externalLink1.xml" ->
      s"""<?xml version="1.0" encoding="UTF-8"?>
<externalLink xmlns="$nsMain" xmlns:r="$nsRel">
  <externalBook r:id="rId1">
    <sheetNames><sheetName val="Extern"/></sheetNames>
  </externalBook>
</externalLink>""",
    "xl/externalLinks/_rels/externalLink1.xml.rels" ->
      """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="other.xlsx" TargetMode="External"/>
</Relationships>""",
    "xl/pivotCache/pivotCacheDefinition1.xml" ->
      s"""<?xml version="1.0" encoding="UTF-8"?>
<pivotCacheDefinition xmlns="$nsMain" invalid="0" refreshOnLoad="1"/>"""
  )

  private def preservationRichBytes(): Array[Byte] =
    val baos = new ByteArrayOutputStream()
    val out = new ZipOutputStream(baos)
    try
      preservationRichParts.foreach { case (name, content) => writeEntry(out, name, content) }
      // A STORED binary part: exercises the preserved-copy branch that mirrors STORED
      // size/crc metadata, for both on-disk and in-memory sources.
      val media = Array.tabulate[Byte](64)(i => (i * 7).toByte)
      val stored = new ZipEntry("xl/media/image1.png")
      stored.setMethod(ZipEntry.STORED)
      stored.setSize(media.length.toLong)
      stored.setCompressedSize(media.length.toLong)
      val crc = new java.util.zip.CRC32
      crc.update(media)
      stored.setCrc(crc.getValue)
      out.putNextEntry(stored)
      out.write(media)
      out.closeEntry()
    finally out.close()
    baos.toByteArray

  private def zipEntryNames(bytes: Array[Byte]): Vector[String] =
    val zip = new java.util.zip.ZipInputStream(new ByteArrayInputStream(bytes))
    try
      Iterator
        .continually(Option(zip.getNextEntry))
        .takeWhile(_.isDefined)
        .flatten
        .map(_.getName)
        .toVector
    finally zip.close()

  private def zipEntryString(bytes: Array[Byte], name: String): String =
    val zip = new java.util.zip.ZipInputStream(new ByteArrayInputStream(bytes))
    try
      Iterator
        .continually(Option(zip.getNextEntry))
        .takeWhile(_.isDefined)
        .flatten
        .find(_.getName == name)
        .map(_ => new String(zip.readAllBytes(), "UTF-8"))
        .getOrElse(fail(s"entry $name missing from zip: ${zipEntryNames(bytes)}"))
    finally zip.close()

  private def assertPreservedSurvives(outputBytes: Array[Byte], clue: String): Unit =
    val wbXml = zipEntryString(outputBytes, "xl/workbook.xml")
    assert(wbXml.contains("<externalReferences>"), s"$clue: externalReferences dropped:\n$wbXml")
    assert(wbXml.contains("<workbookProtection"), s"$clue: workbookProtection dropped:\n$wbXml")
    assert(wbXml.contains("<pivotCaches>"), s"$clue: pivotCaches dropped:\n$wbXml")
    val entries = zipEntryNames(outputBytes)
    assert(
      entries.contains("xl/externalLinks/externalLink1.xml"),
      s"$clue: externalLink part dropped: $entries"
    )
    assert(
      entries.contains("xl/pivotCache/pivotCacheDefinition1.xml"),
      s"$clue: pivotCacheDefinition part dropped: $entries"
    )
    assert(
      entries.contains("xl/media/image1.png"),
      s"$clue: STORED media part dropped: $entries"
    )

  test("GH-412 law: clean bytes-read write is byte-identical to clean path-read write") {
    val fixtureBytes = preservationRichBytes()
    val fixturePath = Files.createTempFile("gh412-fixture", ".xlsx")
    Files.write(fixturePath, fixtureBytes)
    val outViaPath = Files.createTempFile("gh412-via-path", ".xlsx")
    val outViaBytes = Files.createTempFile("gh412-via-bytes", ".xlsx")
    try
      val viaPath = XlsxReader
        .read(fixturePath)
        .fold(err => fail(s"path read failed: $err"), identity)
      val viaBytes = XlsxReader
        .readFromBytes(fixtureBytes)
        .fold(err => fail(s"bytes read failed: $err"), identity)

      XlsxWriter
        .write(viaPath, outViaPath)
        .fold(err => fail(s"path-side write failed: $err"), identity)
      XlsxWriter
        .write(viaBytes, outViaBytes)
        .fold(err => fail(s"bytes-side write failed: $err"), identity)

      val pathOut = Files.readAllBytes(outViaPath)
      val bytesOut = Files.readAllBytes(outViaBytes)
      assertPreservedSurvives(bytesOut, "bytes-read clean write")
      assert(
        java.util.Arrays.equals(pathOut, bytesOut),
        s"write(read(path)) and write(readFromBytes(bytes)) must be byte-identical " +
          s"(path ${pathOut.length}B vs bytes ${bytesOut.length}B)"
      )
    finally
      Files.deleteIfExists(fixturePath)
      Files.deleteIfExists(outViaPath)
      Files.deleteIfExists(outViaBytes)
  }

  test("GH-412 law: surgical write after identical edit is byte-identical across read forms") {
    val fixtureBytes = preservationRichBytes()
    val fixturePath = Files.createTempFile("gh412-dirty-fixture", ".xlsx")
    Files.write(fixturePath, fixtureBytes)
    val outViaPath = Files.createTempFile("gh412-dirty-via-path", ".xlsx")
    val outViaBytes = Files.createTempFile("gh412-dirty-via-bytes", ".xlsx")
    try
      def edit(wb: Workbook): Workbook =
        (for sheet <- wb("Sheet1")
        yield wb.put(sheet.put(ref"A1" -> "Modified")))
          .fold(e => fail(s"edit failed: $e"), identity)

      val viaPath = edit(
        XlsxReader.read(fixturePath).fold(err => fail(s"path read failed: $err"), identity)
      )
      val viaBytes = edit(
        XlsxReader
          .readFromBytes(fixtureBytes)
          .fold(err => fail(s"bytes read failed: $err"), identity)
      )

      XlsxWriter
        .write(viaPath, outViaPath)
        .fold(err => fail(s"path-side write failed: $err"), identity)
      XlsxWriter
        .write(viaBytes, outViaBytes)
        .fold(err => fail(s"bytes-side write failed: $err"), identity)

      val pathOut = Files.readAllBytes(outViaPath)
      val bytesOut = Files.readAllBytes(outViaBytes)
      assertPreservedSurvives(bytesOut, "bytes-read surgical write")
      assert(
        java.util.Arrays.equals(pathOut, bytesOut),
        s"surgical writes must be byte-identical across read forms " +
          s"(path ${pathOut.length}B vs bytes ${bytesOut.length}B)"
      )
    finally
      Files.deleteIfExists(fixturePath)
      Files.deleteIfExists(outViaPath)
      Files.deleteIfExists(outViaBytes)
  }

  test("GH-412: mutating the caller's array after read cannot corrupt the write") {
    val fixtureBytes = preservationRichBytes()
    val pristine = fixtureBytes.clone()
    val out = Files.createTempFile("gh412-mutation", ".xlsx")
    try
      val wb = XlsxReader
        .readFromBytes(fixtureBytes)
        .fold(err => fail(s"bytes read failed: $err"), identity)

      // Corrupt the caller's array AFTER the read: the reader must have taken a private
      // snapshot, so the clean write still reproduces the ORIGINAL archive.
      java.util.Arrays.fill(fixtureBytes, 0.toByte)

      XlsxWriter.write(wb, out).fold(err => fail(s"write failed: $err"), identity)
      val outBytes = Files.readAllBytes(out)
      assertPreservedSurvives(outBytes, "post-mutation write")
      assert(
        java.util.Arrays.equals(outBytes, pristine),
        "clean write must reproduce the archive as read, immune to caller mutation"
      )
    finally Files.deleteIfExists(out)
  }
