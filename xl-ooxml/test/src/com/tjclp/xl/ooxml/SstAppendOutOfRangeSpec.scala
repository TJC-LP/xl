package com.tjclp.xl.ooxml

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.jdk.CollectionConverters.*

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * GH-781: a write that copies a sheet part byte for byte and adds shared strings must not change
 * what that part's cells read as. A `t="s"` index at or past the table's size names no entry
 * (ECMA-376 Part 1, §18.4.9) and reads as #REF!; an appended string would give it a value. When a
 * copied part holds such an index, the write appends nothing and writes the new strings inline.
 */
class SstAppendOutOfRangeSpec extends ScalaCheckSuite:

  private val table = Vector("Alpha", "Beta", "Gamma")

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

  // ========== Scanner ==========

  private def scan(xml: String, tableSize: Int): Boolean =
    SharedStringBound.referencesPast(
      new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)),
      tableSize
    )

  private def sheet(cells: String): String =
    s"""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">$cells</row></sheetData></worksheet>"""

  test("GH-781: scanner flags the first index at or past the table and nothing below it") {
    assert(!scan(sheet("""<c r="A1" t="s"><v>2</v></c>"""), 3))
    assert(scan(sheet("""<c r="A1" t="s"><v>3</v></c>"""), 3))
    assert(scan(sheet("""<c r="A1" t="s"><v>2147483647</v></c>"""), 3))
    assert(scan(sheet(s"""<c r="A1" t="s"><v>${"1" * 40}</v></c>"""), 3))
    assert(scan(sheet("""<c r="A1" t="s"><v> 3 </v></c>"""), 3))
    assert(scan(sheet("""<c r="A1" t="s"><v>0</v></c>"""), 0))
  }

  test("GH-781: scanner ignores non-shared cells and values that name no position") {
    assert(!scan(sheet("""<c r="A1"><v>99</v></c><c r="B1" t="n"><v>99</v></c>"""), 3))
    assert(!scan(sheet("""<c r="A1" t="str"><v>99</v></c>"""), 3))
    assert(!scan(sheet("""<c r="A1" t="s"><v>-4</v></c><c r="B1" t="s"><v>x</v></c>"""), 3))
    assert(!scan(sheet("""<c r="A1" t="s"><f>1</f></c>"""), 3))
  }

  test("GH-781: scanner reads revision <nc>/<oc> values and treats unreadable XML as past") {
    val rev =
      """<revisions xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><rcc rId="1" sId="1"><oc r="A1" t="s"><v>9</v></oc></rcc></revisions>"""
    assert(scan(rev, 3))
    assert(!scan(rev, 10))
    assert(scan("<worksheet><sheetData>", 3))
  }

  test("GH-781: cell-bearing parts are worksheets, macro sheets and revision logs") {
    assert(SharedStringBound.isCellBearingPart("xl/worksheets/sheet1.xml"))
    assert(SharedStringBound.isCellBearingPart("xl/macrosheets/sheet1.xml"))
    assert(SharedStringBound.isCellBearingPart("xl/revisions/revisionLog1.xml"))
    assert(!SharedStringBound.isCellBearingPart("xl/worksheets/_rels/sheet1.xml.rels"))
    assert(!SharedStringBound.isCellBearingPart("xl/revisions/revisionHeaders.xml"))
    assert(!SharedStringBound.isCellBearingPart("xl/sharedStrings.xml"))
  }

  property("GH-781: scanner answers whether any index is at or past the table size") {
    val gen = for
      size <- Gen.choose(0, 50)
      indices <- Gen.listOf(Gen.oneOf(Gen.choose(0, 60), Gen.const(Int.MaxValue)))
    yield (size, indices)
    Prop.forAll(gen) { case (size, indices) =>
      val cells = indices.zipWithIndex.map { case (i, n) =>
        s"""<c r="A${n + 1}" t="s"><v>$i</v></c>"""
      }.mkString
      scan(sheet(cells), size) == indices.exists(_ >= size)
    }
  }

  // ========== Fixture helpers ==========

  private def editAndWrite(source: Path, newText: String): Path =
    val wb = read(source)
    val edited = wb("Sheet1")
      .map(s => wb.put(s.put(ref"A2" -> newText)))
      .fold(err => fail(s"Edit failed: $err"), identity)
    val output = Files.createTempFile("gh781-out", ".xlsx")
    XlsxWriter.write(edited, output).fold(err => fail(s"Write failed: $err"), identity)
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
    extra: Map[String, String] = Map.empty
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
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
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
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""
      )
      put(
        "xl/sharedStrings.xml",
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="3">""" +
          table.map(s => s"<si><t>$s</t></si>").mkString + "</sst>"
      )
      put("xl/worksheets/sheet1.xml", sheetXml(sheet1Index))
      put("xl/worksheets/sheet2.xml", sheetXml(sheet2Index))
      extra.foreach(put)
    finally out.close()
    path
