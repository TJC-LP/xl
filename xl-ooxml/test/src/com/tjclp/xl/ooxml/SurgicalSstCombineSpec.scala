package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * GH-277: three defects in XlsxWriter's surgical SST combination, all found by the wave-1 GH-17
 * investigation and pinned here:
 *
 *   1. ORPHANED SST: a no-new-strings surgical write set `sstForSheets=None`, so the regenerated
 *      modified sheet re-emitted ALL its text cells as t="inlineStr" while the preserved
 *      sharedStrings.xml still shipped — wasted size and a mixed representation nothing should
 *      produce. Regenerated sheets must keep referencing the (preserved or combined) SST.
 *   1. NFC STORAGE: the combine path appended `SharedStrings.normalize(str)` (NFC) instead of the
 *      original string — non-NFC text silently changed bytes on surgical write. Normalization is
 *      for COMPARISON only, never STORAGE (and since GH-289 even comparison is exact).
 *   1. REFERENCE COUNT: `combinedTotalCount` added `newEntries.size` (each new string counted
 *      once); the SST `count` attribute counts REFERENCES, so a new string used from N cells must
 *      add N.
 */
class SurgicalSstCombineSpec extends FunSuite:

  private val existingStrings = Vector("Alpha", "Beta", "Gamma")
  private val nfd = "Cafe\u0301" // NFD: e + combining acute (5 codepoints)
  private val nfdUtf8 = nfd.getBytes("UTF-8")

  private def codepoints(s: String): List[Int] = s.codePoints().toArray.toList

  // ========== 1. No-new-strings surgical write must NOT orphan the SST ==========

  test("GH-277.1: numeric surgical edit keeps t=\"s\" SST references in the regenerated sheet") {
    val source = createRawSstFixture()
    val wb = XlsxReader.read(source).fold(err => fail(s"Read failed: $err"), identity)

    val modified = wb("Sheet1")
      .map(sheet => wb.put(sheet.put(ref"D1" -> 42)))
      .fold(err => fail(s"Modify failed: $err"), identity)

    val output = Files.createTempFile("gh277-orphan", ".xlsx")
    XlsxWriter.write(modified, output).fold(err => fail(s"Write failed: $err"), identity)

    // SST is still copied verbatim (no new strings)
    assertEquals(
      readEntry(output, "xl/sharedStrings.xml").toSeq,
      readEntry(source, "xl/sharedStrings.xml").toSeq,
      "no-new-strings SST should remain byte-identical"
    )

    // The regenerated sheet must reference the preserved SST, not inline every string
    val sheetXml = new String(readEntry(output, "xl/worksheets/sheet1.xml"), "UTF-8")
    assert(
      !sheetXml.contains("inlineStr"),
      s"regenerated sheet re-emitted text as inlineStr, orphaning the SST:\n$sheetXml"
    )
    assert(
      sheetXml.contains("t=\"s\""),
      s"regenerated sheet should reference SST entries via t=\"s\":\n$sheetXml"
    )

    // And the values still read back
    val reloaded = XlsxReader.read(output).fold(err => fail(s"Reload failed: $err"), identity)
    List(ref"A1", ref"A2", ref"A3").zip(existingStrings).foreach { case (r, s) =>
      assertEquals(textAt(reloaded, r), Some(s))
    }

    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== 2. Combine path must store ORIGINAL strings, not NFC-normalized ==========

  test("GH-277.2: combine path appends the ORIGINAL codepoints (NFD survives surgical write)") {
    val source = createRawSstFixture()
    val wb = XlsxReader.read(source).fold(err => fail(s"Read failed: $err"), identity)

    val modified = wb("Sheet1")
      .map(sheet => wb.put(sheet.put(ref"D1" -> nfd)))
      .fold(err => fail(s"Modify failed: $err"), identity)

    val output = Files.createTempFile("gh277-nfc", ".xlsx")
    XlsxWriter.write(modified, output).fold(err => fail(s"Write failed: $err"), identity)

    // Raw SST bytes must contain the decomposed UTF-8 sequence, not the NFC rewrite
    val sstBytes = readEntry(output, "xl/sharedStrings.xml")
    assert(
      containsSlice(sstBytes, nfdUtf8),
      s"combined SST lost the NFD codepoints (stored NFC instead):\n${new String(sstBytes, "UTF-8")}"
    )

    val reloaded = XlsxReader.read(output).fold(err => fail(s"Reload failed: $err"), identity)
    val got = textAt(reloaded, ref"D1").getOrElse(fail("D1 missing"))
    assertEquals(codepoints(got), List(67, 97, 102, 101, 0x301), "NFD changed bytes on reload")

    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== 3. SST count attribute counts REFERENCES of new strings ==========

  test("GH-277.3: a new string referenced from 3 cells adds 3 to the SST count attribute") {
    val source = createRawSstFixture()
    val wb = XlsxReader.read(source).fold(err => fail(s"Read failed: $err"), identity)

    val modified = wb("Sheet1")
      .map(sheet =>
        wb.put(
          sheet.put(ref"D1" -> "Repeated").put(ref"D2" -> "Repeated").put(ref"D3" -> "Repeated")
        )
      )
      .fold(err => fail(s"Modify failed: $err"), identity)

    val output = Files.createTempFile("gh277-count", ".xlsx")
    XlsxWriter.write(modified, output).fold(err => fail(s"Write failed: $err"), identity)

    val sstXml = new String(readEntry(output, "xl/sharedStrings.xml"), "UTF-8")
    // Source: count=3 (three referenced strings). New: "Repeated" x3 references → count=6.
    assert(
      sstXml.contains("count=\"6\""),
      s"SST count should be 3 preserved + 3 references of the new string = 6. SST:\n$sstXml"
    )
    assert(
      sstXml.contains("uniqueCount=\"4\""),
      s"SST uniqueCount should be 3 preserved + 1 new unique = 4. SST:\n$sstXml"
    )

    val reloaded = XlsxReader.read(output).fold(err => fail(s"Reload failed: $err"), identity)
    List(ref"D1", ref"D2", ref"D3").foreach { r =>
      assertEquals(textAt(reloaded, r), Some("Repeated"))
    }

    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== helpers ==========

  private def textAt(wb: Workbook, r: ARef): Option[String] =
    wb("Sheet1").toOption.flatMap(_.cells.get(r)).map(_.value).collect {
      case CellValue.Text(s) => s
    }

  private def containsSlice(haystack: Array[Byte], needle: Array[Byte]): Boolean =
    haystack.indexOfSlice(needle) >= 0

  private def readEntry(path: Path, entryName: String): Array[Byte] =
    val zip = new ZipFile(path.toFile)
    try
      val entry = zip.getEntry(entryName)
      assert(entry != null, s"$entryName missing from $path")
      val is = zip.getInputStream(entry)
      try is.readAllBytes()
      finally is.close()
    finally zip.close()

  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    out.putNextEntry(new ZipEntry(name))
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()

  /**
   * Excel-shaped raw fixture: SST with three referenced strings (count=3, uniqueCount=3), Sheet1
   * referencing all three via t="s".
   */
  private def createRawSstFixture(): Path =
    val path = Files.createTempFile("gh277-fixture", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))
    out.setLevel(1)
    try
      writeEntry(
        out,
        "[Content_Types].xml",
        """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""
      )
      writeEntry(
        out,
        "xl/sharedStrings.xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
          "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"3\">" +
          existingStrings.map(s => s"<si><t>$s</t></si>").mkString +
          "</sst>"
      )
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="A2" t="s"><v>1</v></c></row>
    <row r="3"><c r="A3" t="s"><v>2</v></c></row>
  </sheetData>
</worksheet>"""
      )
    finally out.close()
    path
