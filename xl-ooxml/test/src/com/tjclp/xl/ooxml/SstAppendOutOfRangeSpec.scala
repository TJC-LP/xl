package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.jdk.CollectionConverters.*

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * GH-781: a write that copies a sheet part byte for byte and adds shared strings must not change
 * what that part's cells read as. A `t="s"` value that names no entry (an index at or past the
 * table's size, ECMA-376 Part 1 §18.4.9, or a value SharedStringIndex does not read as an index)
 * reads as #REF! or #VALUE! in XlsxReader; an appended string could give it a value. When a copied
 * part holds such a value, or cannot be read, the write adds no string and writes the new strings
 * inline.
 */
class SstAppendOutOfRangeSpec extends FunSuite:

  private val table = Vector("Alpha", "Beta", "Gamma")

  private val mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  // ========== Writer ==========

  test(
    "GH-781: an index one past the table in a copied sheet still reads as #REF! after an append"
  ) {
    val source = fixture(sheet2Index = "3")
    val output = editAndWrite(source, "NewText")

    val reloaded = read(output)
    assertEquals(valueAt(reloaded, "Sheet2", ref"A1"), Some(CellValue.Error(CellError.Ref)))
    assertEquals(valueAt(reloaded, "Sheet1", ref"A2"), Some(CellValue.Text("NewText")))
    assertEquals(valueAt(reloaded, "Sheet1", ref"A1"), Some(CellValue.Text("Alpha")))

    assertEquals(
      entry(output, "xl/worksheets/sheet2.xml").toSeq,
      entry(source, "xl/worksheets/sheet2.xml").toSeq
    )
    val sst = text(output, "xl/sharedStrings.xml")
    assert(!sst.contains("NewText"), s"the table gained an entry a copied index now names:\n$sst")
    assert(text(output, "xl/worksheets/sheet1.xml").contains("inlineStr"))
    cleanup(source, output)
  }

  test("GH-781: with the append blocked the count attribute counts only table references") {
    val source = fixture(sheet2Index = "3")
    val output = editAndWrite(source, "NewText")
    val sst = text(output, "xl/sharedStrings.xml")
    // Sheet1!A1 references Alpha; Sheet2!A1's t="s" cell is still counted; NewText is inline
    assert(sst.contains("count=\"2\""), sst)
    assert(sst.contains("uniqueCount=\"3\""), sst)
    cleanup(source, output)
  }

  test("GH-781: Int.MaxValue and a 30-digit index in a copied sheet block the append") {
    List("2147483647", "9" * 30, " 3 ").foreach { idx =>
      val source = fixture(sheet2Index = idx)
      val before = valueAt(read(source), "Sheet2", ref"A1")
      val output = editAndWrite(source, "NewText")
      assertEquals(valueAt(read(output), "Sheet2", ref"A1"), before, s"index '$idx'")
      assert(!text(output, "xl/sharedStrings.xml").contains("NewText"), s"index '$idx'")
      cleanup(source, output)
    }
  }

  test("GH-781: in-range indices keep the append plan and the copied part's bytes") {
    val source = fixture(sheet2Index = "2")
    val output = editAndWrite(source, "NewText")

    val reloaded = read(output)
    assertEquals(valueAt(reloaded, "Sheet2", ref"A1"), Some(CellValue.Text("Gamma")))
    assertEquals(valueAt(reloaded, "Sheet1", ref"A2"), Some(CellValue.Text("NewText")))
    assertEquals(
      entry(output, "xl/worksheets/sheet2.xml").toSeq,
      entry(source, "xl/worksheets/sheet2.xml").toSeq
    )

    val sst = text(output, "xl/sharedStrings.xml")
    assert(sst.contains("<si><t>NewText</t></si></sst>"), sst)
    assert(sst.contains("uniqueCount=\"4\""), sst)
    val sheet1 = text(output, "xl/worksheets/sheet1.xml")
    assert(!sheet1.contains("inlineStr"), sheet1)
    assert(sheet1.contains("<v>3</v>"), sheet1)
    cleanup(source, output)
  }

  test("GH-781: an out-of-range index in the edited sheet itself does not block the append") {
    // The edited sheet is regenerated, so its #REF! cell is written as an error value
    val source = fixture(sheet2Index = "2", sheet1Index = "7")
    val output = editAndWrite(source, "NewText")
    val reloaded = read(output)
    assertEquals(valueAt(reloaded, "Sheet1", ref"A1"), Some(CellValue.Error(CellError.Ref)))
    assertEquals(valueAt(reloaded, "Sheet1", ref"A2"), Some(CellValue.Text("NewText")))
    assert(text(output, "xl/sharedStrings.xml").contains("NewText"))
    cleanup(source, output)
  }

  test("GH-781: a copied revision log's out-of-range <nc> blocks the append") {
    val revisionLog =
      """<?xml version="1.0"?>
<revisions xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rcc rId="1" sId="2"><nc r="A1" t="s"><v>5</v></nc></rcc>
</revisions>"""
    val source =
      fixture(sheet2Index = "2", extra = Map("xl/revisions/revisionLog1.xml" -> revisionLog))
    val output = editAndWrite(source, "NewText")
    assert(
      entryNames(output).contains("xl/revisions/revisionLog1.xml"),
      "fixture premise: the revision log is copied"
    )
    assert(!text(output, "xl/sharedStrings.xml").contains("NewText"))
    assertEquals(valueAt(read(output), "Sheet1", ref"A2"), Some(CellValue.Text("NewText")))
    cleanup(source, output)
  }

  test("GH-781: index spellings read the same before and after an append") {
    // in the table's terms: +3, +03, " 3 " and the zero-padded 3 name position 3 (past a
    // three-entry table); the non-ASCII digits name no position; the zero-padded 2 names Gamma
    val padded = "0" * 25
    val blocking = List("+3", "+03", "\u0663", "\uFF13", " 3 ", padded + "3")
    (blocking :+ (padded + "2")).foreach { idx =>
      val source = fixture(sheet2Index = idx)
      val before = valueAt(read(source), "Sheet2", ref"A1")
      val output = editAndWrite(source, "NewText")
      val clue = s"index [$idx]"
      assertEquals(valueAt(read(output), "Sheet2", ref"A1"), before, clue)
      assertEquals(valueAt(read(output), "Sheet1", ref"A2"), Some(CellValue.Text("NewText")), clue)
      assertEquals(
        entry(output, "xl/worksheets/sheet2.xml").toSeq,
        entry(source, "xl/worksheets/sheet2.xml").toSeq,
        clue
      )
      val appended = text(output, "xl/sharedStrings.xml").contains("NewText")
      assertEquals(appended, !blocking.contains(idx), clue)
      cleanup(source, output)
    }
    val gamma = fixture(sheet2Index = padded + "2")
    assertEquals(valueAt(read(gamma), "Sheet2", ref"A1"), Some(CellValue.Text("Gamma")))
    cleanup(gamma)
  }

  test("GH-781: a copied part that cannot be read blocks the append") {
    val source = fixture(
      sheet2Index = "2",
      extra = Map("xl/revisions/revisionLog1.xml" -> "<revisions><rcc><nc")
    )
    val output = editAndWrite(source, "NewText")
    assert(!text(output, "xl/sharedStrings.xml").contains("NewText"))
    assertEquals(valueAt(read(output), "Sheet1", ref"A2"), Some(CellValue.Text("NewText")))
    assertEquals(valueAt(read(output), "Sheet2", ref"A1"), Some(CellValue.Text("Gamma")))
    cleanup(source, output)
  }

  test("GH-781: formula escaping adds no string while a copied part blocks the append") {
    val revisionLog =
      """<?xml version="1.0"?>
<revisions xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rcc rId="1" sId="2"><nc r="A1" t="s"><v>1</v></nc><oc r="A1" t="s"><v>3</v></oc></rcc>
</revisions>"""
    val source =
      fixture(sheet2Index = "2", extra = Map("xl/revisions/revisionLog1.xml" -> revisionLog))
    val output = editAndWrite(source, "NewText", WriterConfig.secure)
    assertEquals(
      entry(output, "xl/revisions/revisionLog1.xml").toSeq,
      entry(source, "xl/revisions/revisionLog1.xml").toSeq
    )
    val sst = text(output, "xl/sharedStrings.xml")
    assert(!sst.contains("NewText"), sst)
    assert(sst.contains("<si><t>Alpha</t></si><si><t>Beta</t></si><si><t>Gamma</t></si>"), sst)
    // Sheet1!A1 (Alpha) and Sheet2!A1 (Gamma) are table references; NewText is inline
    assert(sst.contains("count=\"2\""), sst)
    val reloaded = read(output)
    assertEquals(valueAt(reloaded, "Sheet1", ref"A2"), Some(CellValue.Text("NewText")))
    assertEquals(valueAt(reloaded, "Sheet2", ref"A1"), Some(CellValue.Text("Gamma")))
    cleanup(source, output)
  }

  test("GH-781: escaping with a blocked append counts a copied macro sheet's t=\"s\" cells") {
    // the macro sheet is copied byte for byte; its index 5 names no entry, so the append is
    // blocked, and its three t="s" cells stay table references in the count
    val macroXml = s"""<?xml version="1.0" encoding="UTF-8"?>
<xm:macrosheet xmlns="$mainNs" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>5</v></c></row></sheetData>
</xm:macrosheet>"""
    val source = fixture(sheet2Index = "2", macroSheet = Some(macroXml))
    val output = editAndWrite(source, "NewText", WriterConfig.secure)
    assertEquals(
      entry(output, "xl/macrosheets/sheet1.xml").toSeq,
      entry(source, "xl/macrosheets/sheet1.xml").toSeq
    )
    val sst = text(output, "xl/sharedStrings.xml")
    assert(!sst.contains("NewText"), sst)
    assert(sst.contains("<si><t>Alpha</t></si><si><t>Beta</t></si><si><t>Gamma</t></si>"), sst)
    // Sheet1!A1 (Alpha) and Sheet2!A1 (Gamma) plus the macro sheet's three cells
    assert(sst.contains("count=\"5\""), sst)
    val reloaded = read(output)
    assertEquals(valueAt(reloaded, "Sheet1", ref"A2"), Some(CellValue.Text("NewText")))
    assertEquals(valueAt(reloaded, "Sheet2", ref"A1"), Some(CellValue.Text("Gamma")))
    cleanup(source, output)
  }

  test("GH-781: a write creates no table while a copied part holds a t=\"s\" value") {
    // the source has no table, so Sheet2!A1's index names nothing; a new table would name it
    val source = fixture(sheet2Index = "0", withTable = false)
    val before = valueAt(read(source), "Sheet2", ref"A1")
    val wb = read(source)
    val edited = wb("Sheet1")
      .map(s => wb.put(s.put(ref"A2" -> "NewText")))
      .fold(err => fail(s"Edit failed: $err"), identity)
    val output = Files.createTempFile("gh781-out", ".xlsx")
    XlsxWriter
      .writeWith(edited, output, WriterConfig(sstPolicy = SstPolicy.Always))
      .fold(err => fail(s"Write failed: $err"), identity)
    assert(!entryNames(output).contains("xl/sharedStrings.xml"))
    assertEquals(valueAt(read(output), "Sheet2", ref"A1"), before)
    assertEquals(valueAt(read(output), "Sheet1", ref"A2"), Some(CellValue.Text("NewText")))
    cleanup(source, output)
  }

  test("GH-784: a sheet regenerated without being modified is recounted") {
    // Adding a sheet regenerates every sheet. Sheet2 is not modified, but its out-of-range t="s"
    // cell is written as #REF! (t="e"), so it no longer counts as a reference. The macro sheet is
    // copied byte for byte and keeps the append/recount plan (no orphan to remove).
    val macroXml = s"""<?xml version="1.0" encoding="UTF-8"?>
<xm:macrosheet xmlns="$mainNs" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>2</v></c></row></sheetData>
</xm:macrosheet>"""
    val source = fixture(sheet2Index = "3", macroSheet = Some(macroXml), sstCount = 4)
    val edited =
      read(source).addSheet(Sheet(SheetName.unsafe("New"))).fold(e => fail(e.message), identity)
    val output = Files.createTempFile("gh784-out", ".xlsx")
    XlsxWriter.write(edited, output).fold(err => fail(s"Write failed: $err"), identity)

    val sheet2 = text(output, "xl/worksheets/sheet2.xml")
    assert(!sheet2.contains("t=\"s\""), s"fixture premise: Sheet2 is regenerated\n$sheet2")
    val sst = text(output, "xl/sharedStrings.xml")
    // Sheet1!A1 (Alpha) plus the macro sheet's two cells
    assert(sst.contains("count=\"3\""), sst)
    assert(sst.contains("uniqueCount=\"3\""), sst)
    assertEquals(valueAt(read(output), "Sheet2", ref"A1"), Some(CellValue.Error(CellError.Ref)))
    cleanup(source, output)
  }

  // ========== Fixture helpers ==========

  private def editAndWrite(
    source: Path,
    newText: String,
    config: WriterConfig = WriterConfig()
  ): Path =
    val wb = read(source)
    val edited = wb("Sheet1")
      .map(s => wb.put(s.put(ref"A2" -> newText)))
      .fold(err => fail(s"Edit failed: $err"), identity)
    val output = Files.createTempFile("gh781-out", ".xlsx")
    XlsxWriter.writeWith(edited, output, config).fold(err => fail(s"Write failed: $err"), identity)
    output

  private def read(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"Read failed: $err"), identity)

  private def valueAt(wb: Workbook, sheet: String, r: ARef): Option[CellValue] =
    wb(SheetName.unsafe(sheet)).toOption.flatMap(_.cells.get(r)).map(_.value)

  private def entry(path: Path, name: String): Array[Byte] =
    val zip = new ZipFile(path.toFile)
    try
      val e = zip.getEntry(name)
      assert(e != null, s"$name missing from $path")
      val in = zip.getInputStream(e)
      try in.readAllBytes()
      finally in.close()
    finally zip.close()

  private def entryNames(path: Path): Set[String] =
    val zip = new ZipFile(path.toFile)
    try zip.entries().asScala.map(_.getName).toSet
    finally zip.close()

  private def text(path: Path, name: String): String =
    new String(entry(path, name), StandardCharsets.UTF_8)

  private def cleanup(paths: Path*): Unit = paths.foreach(Files.deleteIfExists)

  /**
   * Two sheets over a three-entry table (Alpha, Beta, Gamma). Sheet1!A1 and Sheet2!A1 are `t="s"`
   * cells holding the given index text.
   */
  private def fixture(
    sheet2Index: String,
    sheet1Index: String = "0",
    extra: Map[String, String] = Map.empty,
    withTable: Boolean = true,
    macroSheet: Option[String] = None,
    sstCount: Int = 2
  ): Path =
    val path = Files.createTempFile("gh781-fixture", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))
    def put(name: String, content: String): Unit =
      out.putNextEntry(new ZipEntry(name))
      out.write(content.getBytes(StandardCharsets.UTF_8))
      out.closeEntry()
    def sheetXml(idx: String): String =
      s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>$idx</v></c></row></sheetData></worksheet>"""
    val ws = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    try
      put(
        "[Content_Types].xml",
        s"""<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="$ws"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="$ws"/>
""" + macroSheet.fold("")(_ =>
          """<Override PartName="/xl/macrosheets/sheet1.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/>"""
        ) + """
""" + (if withTable then
         """<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"""
       else "") + """
</Types>"""
      )
      put(
        "_rels/.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
      )
      put(
        "xl/workbook.xml",
        """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"""
      )
      put(
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
""" + macroSheet.fold("")(_ =>
          """<Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet1.xml"/>"""
        ) + """
""" + (if withTable then
         """<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>"""
       else "") + """
</Relationships>"""
      )
      if withTable then
        put(
          "xl/sharedStrings.xml",
          s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="$sstCount" uniqueCount="3">""" +
            table.map(s => s"<si><t>$s</t></si>").mkString + "</sst>"
        )
      put("xl/worksheets/sheet1.xml", sheetXml(sheet1Index))
      put("xl/worksheets/sheet2.xml", sheetXml(sheet2Index))
      macroSheet.foreach(put("xl/macrosheets/sheet1.xml", _))
      extra.foreach(put)
    finally out.close()
    path
