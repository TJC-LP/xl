package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.xml.{Elem, XML}

import com.tjclp.xl.addressing.Row
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.{col, ref}
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}
import munit.FunSuite

/**
 * GH-558: `insert-rows` / `delete-rows` duplicated property-only row records. The model shift was
 * right (`Sheet.shiftAxis` remaps `rowProperties`); the worksheet writer re-emitted every cell-free
 * SOURCE `<row>` with its ORIGINAL `ht`/`hidden`/`outlineLevel` at the ORIGINAL index and then
 * emitted the shifted domain property at the NEW index, so spacer heights landed on content rows
 * and hidden/outline residue on visible rows.
 *
 * Rule under test: the domain `rowProperties` is authoritative for every attribute the reader
 * models (`s`/`customFormat`, `ht`/`customHeight`, `hidden`, `outlineLevel`, `collapsed`); a
 * preserved source row contributes only its UNMODELLED attributes (`spans`, `thickBot`, `thickTop`,
 * `x14ac:dyDescent`); every row index is emitted at most once; a source row left with neither cells
 * nor attributes is dropped.
 */
class PropertyOnlyRowShiftSpec extends FunSuite:

  private val sheetPart = "xl/worksheets/sheet1.xml"

  /** Row attributes the reader lifts into [[RowProperties]] — the ones the domain owns. */
  private val modelledAttrs =
    Set("s", "customFormat", "ht", "customHeight", "hidden", "outlineLevel", "collapsed")

  private val backends = List(
    "ScalaXml" -> WriterConfig(backend = XmlBackend.ScalaXml),
    "SaxStax" -> WriterConfig(backend = XmlBackend.SaxStax)
  )

  // ===== fixture / io helpers =====

  /** The #558 repro: text at A1, A2, A4, A6, `=A2&A4` at A7; property-only rows 3, 5 and 8. */
  private def issueSheet: Sheet =
    Sheet("Sheet")
      .put(ref"A1", CellValue.Text("a"))
      .put(ref"A2", CellValue.Text("b"))
      .put(ref"A4", CellValue.Text("d"))
      .put(ref"A6", CellValue.Text("f"))
      .put(ref"A7", CellValue.Formula("A2&A4", None))
      .setRowProperties(Row.from1(3), RowProperties(height = Some(3.0)))
      .setRowProperties(Row.from1(5), RowProperties(hidden = true, outlineLevel = Some(1)))
      .setRowProperties(Row.from1(8), RowProperties(height = Some(5.15)))

  private def writeTo(wb: Workbook, label: String, config: WriterConfig = WriterConfig()): Path =
    val out = Files.createTempFile(s"xl-558-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter
      .writeWith(wb, out, config)
      .fold(err => fail(s"$label write failed: ${err.message}"), _ => ())
    out

  private def reread(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"re-read failed: ${err.message}"), identity)

  private def firstSheet(wb: Workbook): Sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  /**
   * A pure core edit on a read workbook must be flagged like StructuralEditor does, or the writer
   * verbatim-copies the clean source.
   */
  private def markAllModified(wb: Workbook): Workbook =
    wb.copy(sourceContext =
      wb.sourceContext.map(ctx => wb.sheets.indices.foldLeft(ctx)((c, i) => c.markSheetModified(i)))
    )

  private def editSheet(wb: Workbook)(f: Sheet => Sheet): Workbook =
    markAllModified(wb.copy(sheets = wb.sheets.map(f)))

  // ===== raw worksheet XML inspection =====

  private def rowElems(sheetXml: String): Seq[Elem] =
    (XML.loadString(sheetXml) \ "sheetData" \ "row").collect { case e: Elem => e }

  /** `r -> (attribute -> value)` without `r`; fails when any row index is emitted twice. */
  private def rowAttrs(sheetXml: String): Map[Int, Map[String, String]] =
    val rows = rowElems(sheetXml)
    val indices = rows.map(_ \@ "r")
    assertEquals(indices, indices.distinct, s"every row index must be emitted once:\n$sheetXml")
    rows.map(e => (e \@ "r").toInt -> (e.attributes.asAttrMap - "r")).toMap

  /** Rows carrying at least one modelled (domain-owned) attribute. */
  private def propertyRows(rows: Map[Int, Map[String, String]]): Map[Int, Map[String, String]] =
    rows.collect {
      case (r, attrs) if attrs.keySet.exists(modelledAttrs.contains) =>
        r -> attrs.filter((k, _) => modelledAttrs.contains(k))
    }

  private def cellRefs(sheetXml: String): Set[String] =
    (XML.loadString(sheetXml) \ "sheetData" \ "row" \ "c").map(_ \@ "r").toSet

  /** `<row ...>` start tags in document order, self-closing slash normalized away. */
  private def rowStartTags(sheetXml: String): List[String] =
    """<row\b[^>]*>""".r
      .findAllIn(sheetXml)
      .map(tag => tag.stripSuffix(">").stripSuffix("/").trim + ">")
      .toList

  private val expectedAfterInsert = Map(
    4 -> Map("ht" -> "3.0", "customHeight" -> "1"),
    6 -> Map("hidden" -> "1", "outlineLevel" -> "1"),
    9 -> Map("ht" -> "5.15", "customHeight" -> "1")
  )

  private val expectedAfterDelete = Map(
    3 -> Map("ht" -> "3.0", "customHeight" -> "1"),
    4 -> Map("hidden" -> "1", "outlineLevel" -> "1"),
    7 -> Map("ht" -> "5.15", "customHeight" -> "1")
  )

  // ===== the issue repro, on both serializers of the preserved-metadata writer =====

  backends.foreach { case (backendName, config) =>
    test(
      s"GH-558: insert-rows emits each property-only row once, at the shifted index ($backendName)"
    ) {
      val read = reread(writeTo(Workbook(issueSheet), "src"))
      // insert one row before 1-based row 2 (0-based at = 1)
      val out = writeTo(editSheet(read)(_.insertRows(1, 1)), s"insert-$backendName", config)
      val xml = entryText(out, sheetPart)
      val rows = rowAttrs(xml)

      assertEquals(propertyRows(rows), expectedAfterInsert, xml)
      // the content rows that now sit where the properties used to be carry no residue
      Seq(3, 5, 8).foreach { r =>
        val stale = rows.getOrElse(r, Map.empty).keySet.intersect(modelledAttrs)
        assert(stale.isEmpty, s"row $r must not keep the source row's modelled attrs $stale:\n$xml")
      }
      // the inserted blank row has neither cells nor properties: no <row r="2"> record at all
      assert(!rows.contains(2), s"a row with neither cells nor attributes must be dropped:\n$xml")
      assertEquals(cellRefs(xml), Set("A1", "A3", "A5", "A7", "A8"))

      // the model reads back exactly what the file says
      val back = firstSheet(reread(out))
      assertEquals(back.rowProperties.keySet.map(_.index1), Set(4, 6, 9))
      assertEquals(back.getRowProperties(Row.from1(4)).height, Some(3.0))
      assert(back.getRowProperties(Row.from1(6)).hidden)
      assertEquals(back.getRowProperties(Row.from1(6)).outlineLevel, Some(1))
      assertEquals(back.getRowProperties(Row.from1(9)).height, Some(5.15))
    }
  }

  test("GH-558: delete-rows emits each property-only row once, at the shifted index") {
    val read = reread(writeTo(Workbook(issueSheet), "src"))
    // delete 1-based row 4 (0-based at = 3)
    val out = writeTo(editSheet(read)(_.deleteRows(3, 1)), "delete")
    val xml = entryText(out, sheetPart)
    val rows = rowAttrs(xml)

    assertEquals(propertyRows(rows), expectedAfterDelete, xml)
    Seq(5, 8).foreach { r =>
      val stale = rows.getOrElse(r, Map.empty).keySet.intersect(modelledAttrs)
      assert(stale.isEmpty, s"row $r must not keep the source row's modelled attrs $stale:\n$xml")
    }
    assertEquals(cellRefs(xml), Set("A1", "A2", "A5", "A6"))
    assertEquals(firstSheet(reread(out)).rowProperties.keySet.map(_.index1), Set(3, 4, 7))
  }

  test(
    "GH-558 law: write -> read -> insert -> write -> read -> delete -> write restores the rows"
  ) {
    val original = writeTo(Workbook(issueSheet), "law-src")
    val inserted = writeTo(editSheet(reread(original))(_.insertRows(1, 1)), "law-inserted")
    val restored = writeTo(editSheet(reread(inserted))(_.deleteRows(1, 1)), "law-restored")

    val before = entryText(original, sheetPart)
    val after = entryText(restored, sheetPart)
    assertEquals(rowAttrs(after), rowAttrs(before), s"row records must round-trip:\n$after")
    assertEquals(cellRefs(after), cellRefs(before))
    assertEquals(
      firstSheet(reread(restored)).rowProperties,
      firstSheet(reread(original)).rowProperties
    )
  }

  // ===== unmodelled attributes and the untouched-book guarantee =====

  private def zipEntry(out: ZipOutputStream, name: String, content: String): Unit =
    out.putNextEntry(new ZipEntry(name))
    out.write(content.getBytes(StandardCharsets.UTF_8))
    out.closeEntry()

  /** Minimal single-sheet package around a hand-written worksheet part (no styles part needed). */
  private def minimalXlsx(worksheetXml: String): Path =
    val path = Files.createTempFile("xl-558-foreign-", ".xlsx")
    path.toFile.deleteOnExit()
    val out = new ZipOutputStream(Files.newOutputStream(path))
    try
      zipEntry(
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
      zipEntry(
        out,
        "_rels/.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
      )
      zipEntry(
        out,
        "xl/workbook.xml",
        """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Alpha" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
      )
      zipEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
      )
      zipEntry(out, "xl/worksheets/sheet1.xml", worksheetXml)
    finally out.close()
    path

  /**
   * Excel-shaped rows: `spans` and `x14ac:dyDescent` on every record, a thick-bordered cell-free
   * spacer (row 3), a cell-free row carrying ONLY unmodelled attributes (row 5), and a hidden
   * outline row (row 6). `ht` is written the way our writer prints a Double so the identity check
   * below is exact.
   */
  private val excelShapedRows: List[String] = List(
    """<row r="1" spans="1:2" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c></row>""",
    """<row r="3" spans="1:2" ht="3.0" customHeight="1" thickBot="1" x14ac:dyDescent="0.25"/>""",
    """<row r="5" spans="1:2" thickTop="1" x14ac:dyDescent="0.3"/>""",
    """<row r="6" spans="1:2" hidden="1" outlineLevel="1" x14ac:dyDescent="0.25"/>""",
    """<row r="8" spans="1:2" ht="5.15" customHeight="1" x14ac:dyDescent="0.25"><c r="A8"><v>8</v></c></row>"""
  )

  private def excelShapedFixture(): Path =
    minimalXlsx(
      s"""<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
  <sheetData>
    ${excelShapedRows.mkString("\n    ")}
  </sheetData>
</worksheet>"""
    )

  test("GH-558: unmodelled attributes on cell-free source rows ride through an unrelated edit") {
    val read = reread(excelShapedFixture())
    assertEquals(firstSheet(read).rowProperties.keySet.map(_.index1), Set(3, 6, 8))

    val out = writeTo(editSheet(read)(_.put(ref"B1", CellValue.Text("edited"))), "unmodelled")
    val xml = entryText(out, sheetPart)
    val rows = rowAttrs(xml)

    // thickBot / thickTop / spans / dyDescent are not modelled: the preserved record carries them
    assertEquals(
      rows.get(3),
      Some(
        Map(
          "spans" -> "1:2",
          "ht" -> "3.0",
          "customHeight" -> "1",
          "thickBot" -> "1",
          "x14ac:dyDescent" -> "0.25"
        )
      ),
      xml
    )
    // a cell-free row with ONLY unmodelled attributes still has something to say: it survives
    assertEquals(
      rows.get(5),
      Some(Map("spans" -> "1:2", "thickTop" -> "1", "x14ac:dyDescent" -> "0.3")),
      xml
    )
    assertEquals(cellRefs(xml), Set("A1", "B1", "A8"))
  }

  test("GH-558: regenerating an untouched sheet keeps every <row> start tag identical") {
    val src = excelShapedFixture()
    // no domain edit; force the regeneration path (a clean sheet would be copied verbatim). The
    // DOM serializer keeps Excel's attribute order, so the start tags compare byte-for-byte.
    val out =
      writeTo(markAllModified(reread(src)), "identity", WriterConfig(backend = XmlBackend.ScalaXml))
    assertEquals(
      rowStartTags(entryText(out, sheetPart)),
      rowStartTags(entryText(src, sheetPart)),
      "domain props applied over the stripped source attributes must reproduce the source record"
    )
    // and the SaxStax serializer emits the same attribute sets (it orders attributes itself)
    val sax = writeTo(
      markAllModified(reread(src)),
      "identity-sax",
      WriterConfig(backend = XmlBackend.SaxStax)
    )
    assertEquals(rowAttrs(entryText(sax, sheetPart)), rowAttrs(entryText(src, sheetPart)))
  }

  test(
    "GH-558: a shift on an Excel-shaped sheet moves the modelled attrs and nothing lands twice"
  ) {
    val read = reread(excelShapedFixture())
    val out = writeTo(editSheet(read)(_.insertRows(1, 1)), "excel-shift")
    val xml = entryText(out, sheetPart)
    val rows = rowAttrs(xml)
    assertEquals(
      propertyRows(rows),
      Map(
        4 -> Map("ht" -> "3.0", "customHeight" -> "1"),
        7 -> Map("hidden" -> "1", "outlineLevel" -> "1"),
        9 -> Map("ht" -> "5.15", "customHeight" -> "1")
      ),
      xml
    )
    // the thick-border hint is unmodelled and cannot travel with the model: it stays on the source
    // record, which keeps ONLY its unmodelled attributes at the old index
    assertEquals(
      rows.get(3),
      Some(Map("spans" -> "1:2", "thickBot" -> "1", "x14ac:dyDescent" -> "0.25")),
      xml
    )
    assertEquals(cellRefs(xml), Set("A1", "A9"))
  }

  // ===== the symmetric <cols> path =====

  test("GH-558: delete-cols on the only property-bearing column leaves no stale <col> behind") {
    val sheet = Sheet("Cols")
      .put(ref"A1", CellValue.Text("a"))
      .put(ref"C1", CellValue.Text("c"))
      .setColumnProperties(col"C", ColumnProperties(width = Some(20.0)))
    val read = reread(writeTo(Workbook(sheet), "cols-src"))
    assertEquals(firstSheet(read).columnProperties.keySet.map(_.index0), Set(2))

    // delete column C (0-based at = 2): its cell and its width both go
    val out = writeTo(editSheet(read)(_.deleteColumns(2, 1)), "cols-deleted")
    val xml = entryText(out, sheetPart)
    assert(!xml.contains("<col "), s"the deleted column's width must not resurrect:\n$xml")
    assertEquals(firstSheet(reread(out)).columnProperties, Map.empty)
  }
