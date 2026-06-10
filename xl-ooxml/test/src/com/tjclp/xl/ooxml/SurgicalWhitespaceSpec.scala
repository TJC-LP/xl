package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.{SstPolicy, WriterConfig}
import munit.FunSuite

/**
 * Regression guard for GH-17: SST entries must keep exact whitespace (trailing/double/leading
 * spaces, embedded newlines) through SURGICAL writes — read with source handle, modify one
 * unrelated cell, write.
 *
 * Covers every SST sub-path of the surgical writer:
 *   1. Modified sheet introduces a NEW string → preserved SST combined with new entries and
 *      re-serialized (XlsxWriter combine logic + SharedStrings.toXml/writeSax)
 *   1. Modified sheet introduces NO new strings → SST copied verbatim, sheet regenerated
 *   1. Modification on a DIFFERENT sheet → referencing sheet copied verbatim, SST combined
 *      (preserved entries must keep their indices)
 *   1. Fixture authored by the xl writer itself (not raw XML), then surgically modified
 *
 * Each test asserts BOTH the reloaded domain values and the raw xl/sharedStrings.xml bytes
 * (exact text content and xml:space="preserve" attributes).
 */
class SurgicalWhitespaceSpec extends FunSuite:

  // Whitespace corpus from GH-17 (real-world Syndigo file)
  private val trailingTwo = "Capital.  " // two trailing spaces
  private val doubleInterior = "billion.  Resolute" // double interior space
  private val leadingTwo = "  lead" // two leading spaces
  private val embeddedNewline = "or \nOther" // embedded newline
  private val plain = "Plain"

  private val whitespaceCells: Map[ARef, String] = Map(
    ref"A1" -> trailingTwo,
    ref"A2" -> doubleInterior,
    ref"A3" -> leadingTwo,
    ref"A4" -> embeddedNewline
  )

  // ========== Scenario 1: new string on the SST-referencing sheet (SST combine path) ==========

  test("GH-17: surgical write with new string preserves SST whitespace (combine path)") {
    val source = createRawSstFixture(twoSheets = false)
    val wb = XlsxReader.read(source).fold(err => fail(s"Read failed: $err"), identity)

    // Precondition: the read path preserves whitespace (issue confirmed this works)
    assertWhitespaceValues(wb, "Sheet1", "precondition (read)")

    val modified = wb("Sheet1")
      .map(sheet => wb.put(sheet.put(ref"D1" -> "Surgical Update")))
      .fold(err => fail(s"Modify failed: $err"), identity)

    val output = Files.createTempFile("gh17-combine", ".xlsx")
    XlsxWriter.write(modified, output).fold(err => fail(s"Write failed: $err"), identity)

    // Domain values survive the full surgical round-trip
    val reloaded = XlsxReader.read(output).fold(err => fail(s"Reload failed: $err"), identity)
    assertWhitespaceValues(reloaded, "Sheet1", "after surgical write (combine path)")
    assertEquals(textAt(reloaded, "Sheet1", ref"D1"), Some("Surgical Update"))

    // Prove the combine path actually ran: SST was re-serialized (new string forces regeneration)
    val outSst = new String(readSstBytes(output), "UTF-8")
    assert(
      !java.util.Arrays.equals(readSstBytes(output), readSstBytes(source)),
      "SST should be regenerated (not verbatim-copied) when a new string is introduced"
    )
    assert(outSst.contains("Surgical Update"), s"new string missing from combined SST:\n$outSst")

    // Raw SST bytes: exact text content + xml:space="preserve" attributes survive re-emission
    assertRawSstWhitespace(output, "combine path")

    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== Scenario 2: numeric edit, no new strings (SST verbatim-copy path) ==========

  test("GH-17: surgical numeric edit copies SST verbatim and keeps whitespace") {
    val source = createRawSstFixture(twoSheets = false)
    val wb = XlsxReader.read(source).fold(err => fail(s"Read failed: $err"), identity)

    val modified = wb("Sheet1")
      .map(sheet => wb.put(sheet.put(ref"D1" -> 42)))
      .fold(err => fail(s"Modify failed: $err"), identity)

    val output = Files.createTempFile("gh17-verbatim", ".xlsx")
    XlsxWriter.write(modified, output).fold(err => fail(s"Write failed: $err"), identity)

    // No new strings → preserved SST must be byte-identical (strongest guarantee)
    assertEquals(
      readSstBytes(output).toSeq,
      readSstBytes(source).toSeq,
      "SST should be copied byte-for-byte when no new strings are introduced"
    )

    // Domain values survive even though the modified sheet itself was regenerated
    val reloaded = XlsxReader.read(output).fold(err => fail(s"Reload failed: $err"), identity)
    assertWhitespaceValues(reloaded, "Sheet1", "after surgical write (verbatim SST path)")
    assertEquals(
      reloaded("Sheet1").toOption.flatMap(_.cells.get(ref"D1")).map(_.value),
      Some(CellValue.Number(BigDecimal(42)))
    )

    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== Scenario 3: edit a DIFFERENT sheet (preserved sheet + combined SST) ==========

  test("GH-17: surgical edit on another sheet keeps SST indices and whitespace") {
    val source = createRawSstFixture(twoSheets = true)
    val wb = XlsxReader.read(source).fold(err => fail(s"Read failed: $err"), identity)

    val modified = wb("Sheet2")
      .map(sheet => wb.put(sheet.put(ref"D1" -> "Surgical Update")))
      .fold(err => fail(s"Modify failed: $err"), identity)

    val output = Files.createTempFile("gh17-cross-sheet", ".xlsx")
    XlsxWriter.write(modified, output).fold(err => fail(s"Write failed: $err"), identity)

    // Sheet1 (untouched, references SST by index) must be copied byte-for-byte
    assertEquals(
      readEntry(output, "xl/worksheets/sheet1.xml").toSeq,
      readEntry(source, "xl/worksheets/sheet1.xml").toSeq,
      "Unmodified sheet1 should be byte-identical"
    )

    // Combined SST must keep preserved entries at their original indices, with whitespace
    val reloaded = XlsxReader.read(output).fold(err => fail(s"Reload failed: $err"), identity)
    assertWhitespaceValues(reloaded, "Sheet1", "preserved sheet after cross-sheet edit")
    assertEquals(textAt(reloaded, "Sheet2", ref"D1"), Some("Surgical Update"))

    // Combine path proof: SST re-serialized with the new entry appended after preserved ones
    assert(
      !java.util.Arrays.equals(readSstBytes(output), readSstBytes(source)),
      "SST should be regenerated (not verbatim-copied) when a new string is introduced"
    )
    assertRawSstWhitespace(output, "cross-sheet combine path")

    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== Scenario 3b: same combine path through the StAX backend ==========

  test("GH-17: SaxStax backend surgical write preserves SST whitespace (combine path)") {
    val source = createRawSstFixture(twoSheets = false)
    val wb = XlsxReader.read(source).fold(err => fail(s"Read failed: $err"), identity)

    val modified = wb("Sheet1")
      .map(sheet => wb.put(sheet.put(ref"D1" -> "Surgical Update")))
      .fold(err => fail(s"Modify failed: $err"), identity)

    val output = Files.createTempFile("gh17-saxstax", ".xlsx")
    XlsxWriter
      .writeWith(modified, output, WriterConfig.saxStax)
      .fold(err => fail(s"Write failed: $err"), identity)

    val reloaded = XlsxReader.read(output).fold(err => fail(s"Reload failed: $err"), identity)
    assertWhitespaceValues(reloaded, "Sheet1", "after surgical write (SaxStax backend)")
    assertEquals(textAt(reloaded, "Sheet1", ref"D1"), Some("Surgical Update"))

    // writeSax is a separate emission path from toXml: assert its raw bytes independently
    assertRawSstWhitespace(output, "SaxStax combine path")

    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== Scenario 4: fixture authored by the xl writer itself ==========

  test("GH-17: xl-authored SST round-trips whitespace through surgical modification") {
    val sheet = Sheet("Data").put(
      ref"A1" -> trailingTwo,
      ref"A2" -> doubleInterior,
      ref"A3" -> leadingTwo,
      ref"A4" -> embeddedNewline,
      ref"B1" -> plain
    )
    val source = Files.createTempFile("gh17-xl-authored", ".xlsx")
    XlsxWriter
      .writeWith(Workbook(Vector(sheet)), source, WriterConfig(sstPolicy = SstPolicy.Always))
      .fold(err => fail(s"Initial write failed: $err"), identity)

    // Precondition: the writer actually produced an SST (not inline strings) with exact bytes
    assertRawSstWhitespace(source, "xl-authored fixture")

    val wb = XlsxReader.read(source).fold(err => fail(s"Read failed: $err"), identity)
    assertWhitespaceValues(wb, "Data", "precondition (read of xl-authored file)")

    val modified = wb("Data")
      .map(s => wb.put(s.put(ref"D1" -> "Surgical Update")))
      .fold(err => fail(s"Modify failed: $err"), identity)

    val output = Files.createTempFile("gh17-xl-authored-out", ".xlsx")
    XlsxWriter.write(modified, output).fold(err => fail(s"Write failed: $err"), identity)

    val reloaded = XlsxReader.read(output).fold(err => fail(s"Reload failed: $err"), identity)
    assertWhitespaceValues(reloaded, "Data", "after surgical write (xl-authored fixture)")
    assertEquals(textAt(reloaded, "Data", ref"D1"), Some("Surgical Update"))

    assertRawSstWhitespace(output, "xl-authored surgical output")

    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== Assertion helpers ==========

  private def textAt(wb: Workbook, sheetName: String, r: ARef): Option[String] =
    wb(sheetName).toOption.flatMap(_.cells.get(r)).map(_.value).collect {
      case CellValue.Text(s) => s
    }

  private def assertWhitespaceValues(wb: Workbook, sheetName: String, where: String): Unit =
    whitespaceCells.foreach { case (r, expected) =>
      val got = textAt(wb, sheetName, r)
      assertEquals(
        got,
        Some(expected),
        s"$where: $sheetName!${r.toA1} expected '${escapeWs(expected)}' " +
          s"but got ${got.map(s => s"'${escapeWs(s)}'").getOrElse("<missing>")}"
      )
    }

  /** Assert raw sharedStrings.xml bytes keep exact text and xml:space="preserve" attributes. */
  private def assertRawSstWhitespace(path: Path, where: String): Unit =
    val sst = new String(readEntry(path, "xl/sharedStrings.xml"), "UTF-8")
    List(trailingTwo, doubleInterior, leadingTwo, embeddedNewline).foreach { s =>
      assert(
        sst.contains(s"""<t xml:space="preserve">$s</t>"""),
        s"$where: raw SST lost whitespace or xml:space for '${escapeWs(s)}'. SST:\n$sst"
      )
    }

  /** Make whitespace visible in failure messages (spaces → '·', control chars escaped). */
  private def escapeWs(s: String): String =
    s.flatMap {
      case ' ' => "·"
      case '\n' => "\\n"
      case '\r' => "\\r"
      case '\t' => "\\t"
      case c => c.toString
    }

  // ========== ZIP helpers ==========

  private def readEntry(path: Path, entryName: String): Array[Byte] =
    val zip = new ZipFile(path.toFile)
    try
      val entry = zip.getEntry(entryName)
      assert(entry != null, s"$entryName missing from $path")
      val is = zip.getInputStream(entry)
      try is.readAllBytes()
      finally is.close()
    finally zip.close()

  private def readSstBytes(path: Path): Array[Byte] = readEntry(path, "xl/sharedStrings.xml")

  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    out.putNextEntry(new ZipEntry(name))
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()

  // ========== Raw XML fixture (Excel-shaped, bypasses the xl writer on the way in) ==========

  /**
   * Build an xlsx whose SST carries the GH-17 whitespace corpus with xml:space="preserve",
   * exactly as Excel emits it. Sheet1 references all entries via t="s"; the optional Sheet2 holds
   * only a number (so cross-sheet edits leave Sheet1 untouched).
   */
  private def createRawSstFixture(twoSheets: Boolean): Path =
    val path = Files.createTempFile(s"gh17-fixture-${if twoSheets then "2s" else "1s"}", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))
    out.setLevel(1)

    val sheet2ContentType =
      if twoSheets then
        """
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"""
      else ""
    val sheet2Decl = if twoSheets then """
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>""" else ""
    val sheet2Rel =
      if twoSheets then """
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>"""
      else ""
    val sstRelId = if twoSheets then "rId3" else "rId2"

    try
      writeEntry(
        out,
        "[Content_Types].xml",
        s"""<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>$sheet2ContentType
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
        s"""<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>$sheet2Decl
  </sheets>
</workbook>"""
      )

      writeEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        s"""<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>$sheet2Rel
  <Relationship Id="$sstRelId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""
      )

      // SST entries on single lines: <t> content is EXACTLY the corpus strings
      writeEntry(
        out,
        "xl/sharedStrings.xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
          "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"5\" uniqueCount=\"5\">" +
          s"<si><t xml:space=\"preserve\">$trailingTwo</t></si>" +
          s"<si><t xml:space=\"preserve\">$doubleInterior</t></si>" +
          s"<si><t xml:space=\"preserve\">$leadingTwo</t></si>" +
          s"<si><t xml:space=\"preserve\">$embeddedNewline</t></si>" +
          s"<si><t>$plain</t></si>" +
          "</sst>"
      )

      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>4</v></c><c r="C1"><v>7</v></c></row>
    <row r="2"><c r="A2" t="s"><v>1</v></c></row>
    <row r="3"><c r="A3" t="s"><v>2</v></c></row>
    <row r="4"><c r="A4" t="s"><v>3</v></c></row>
  </sheetData>
</worksheet>"""
      )

      if twoSheets then
        writeEntry(
          out,
          "xl/worksheets/sheet2.xml",
          """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"""
        )
    finally out.close()

    path
