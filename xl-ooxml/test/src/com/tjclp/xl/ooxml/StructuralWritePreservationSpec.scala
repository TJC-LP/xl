package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import munit.FunSuite

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.sheets.PageSetup
import com.tjclp.xl.tables.TableSpec

/**
 * GH-429 writer integration: after a pure structural row insert, every range-bearing feature the
 * sheet carries — Excel-shaped data validations, `_xlnm.Print_Area`/`Print_Titles`, tables, and the
 * sheet-level autoFilter — moves together through write → re-read, and the dv/autoFilter CLEAN
 * gates correctly go dirty. The autoFilter overlay preserves filterColumn children verbatim,
 * `Remove` strips the element, and the identity fast-path keeps an untouched filter byte-stable
 * through an unrelated cell edit.
 */
class StructuralWritePreservationSpec extends FunSuite:

  private def writeTo(wb: Workbook, label: String): Path =
    val out = Files.createTempFile(s"xl-429-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter
      .writeWith(wb, out, WriterConfig())
      .fold(err => fail(s"$label write failed: ${err.message}"), _ => ())
    out

  private def reread(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"re-read failed: ${err.message}"), identity)

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) =>
          new String(zip.getInputStream(e).readAllBytes(), java.nio.charset.StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  private def markAllModified(wb: Workbook): Workbook =
    wb.copy(sourceContext =
      wb.sourceContext.map(ctx => wb.sheets.indices.foldLeft(ctx)((c, i) => c.markSheetModified(i)))
    )

  test("GH-429: DV + print area + table + autoFilter all move through insert -> write -> re-read") {
    val table = TableSpec
      .fromColumnNames("T1", "T1", ref"A1:C5", Vector("Region", "Product", "Units"))
      .fold(e => fail(s"table: $e"), identity)
    val sheet = Sheet("Alpha")
      .put(ref"A1", CellValue.Text("Region"))
      .put(ref"B1", CellValue.Text("Product"))
      .put(ref"C1", CellValue.Text("Units"))
      .put(ref"A2", CellValue.Text("North"))
      .put(ref"B2", CellValue.Text("Anvil"))
      .put(ref"C2", CellValue.Number(12))
      .withDataValidation(
        ref"H5:H10",
        DataValidation.listOf("yes", "no").withPrompt("Pick", "yes or no")
      )
      .withTable(table)
      .copy(
        pageSetup = Some(PageSetup(printArea = Some(ref"A1:D10"), repeatRows = Some((1, 2)))),
        autoFilter = Some(AutoFilterState.Ranged(ref"A1:C5"))
      )
    val src = writeTo(Workbook(sheet), "src")

    // read (source context present), pure core structural edit, mark modified like the editor
    val read = reread(src)
    val edited = markAllModified(read.copy(sheets = read.sheets.map(_.insertRows(1, 2))))
    val out = writeTo(edited, "out")
    val result = reread(out)
    val s = result.sheets.headOption.getOrElse(fail("no sheet"))

    // 1. the Excel-shaped DV moved (issue repro: H5:H10 + insert 2 at row 2 -> H7:H12)
    assertEquals(s.typedDataValidations.flatMap(_.ranges.map(_.toA1)), Vector("H7:H12"))
    assertEquals(
      s.typedDataValidations.map(_.messages.prompt),
      Vector(Some("yes or no")),
      "prompt text must survive the dirty write"
    )
    assert(entryText(out, "xl/worksheets/sheet1.xml").contains("sqref=\"H7:H12\""))

    // 2. print names re-derived from the shifted model
    assertEquals(s.pageSetup.flatMap(_.printArea).map(_.toA1), Some("A1:D12"))
    assertEquals(s.pageSetup.flatMap(_.repeatRows), Some((1, 4)))
    val workbookXml = entryText(out, "xl/workbook.xml")
    assert(workbookXml.contains("Alpha!$A$1:$D$12"), workbookXml)
    assert(workbookXml.contains("Alpha!$1:$4"), workbookXml)

    // 3. the table moved (always regenerated from the model)
    assertEquals(s.tables.get("T1").map(_.range.toA1), Some("A1:C7"))
    assert(entryText(out, "xl/tables/table1.xml").contains("ref=\"A1:C7\""))

    // 4. the sheet autoFilter moved
    assertEquals(s.autoFilter, Some(AutoFilterState.Ranged(ref"A1:C7": CellRange)))
    assert(entryText(out, "xl/worksheets/sheet1.xml").contains("<autoFilter ref=\"A1:C7\""))
  }

  // ===== GH-429 rework: PRESERVED payloads (the reader's declaration-prefixed shape) must =====
  // ===== shift too — the guard that only tolerated bare `<dataValidation…` made every real =====
  // ===== Preserved entry structurally inert while hand-built test payloads passed.        =====

  private def zipEntry(out: ZipOutputStream, name: String, content: String): Unit =
    out.putNextEntry(new ZipEntry(name))
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()

  /**
   * Minimal foreign fixture whose worksheet carries entries the typed models refuse: a Preserved CF
   * block (xr:uid — an envelope attr outside {sqref, pivot}), an imeMode DV, and an unknown-type
   * DV. The reader lifts all three into Preserved payloads via CfCodec.preservedXml, i.e.
   * XML-declaration-prefixed canonical XML — the exact shape SqrefShift must tolerate.
   */
  private def preservedPayloadFixture(): Path =
    val path = Files.createTempFile("xl-429-preserved-fixture", ".xlsx")
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
      zipEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="B5:B6" xr:uid="{00000000-0001-0000-0000-000000000000}"><cfRule type="expression" priority="1"><formula>$B5&gt;0</formula></cfRule></conditionalFormatting>
  <dataValidations count="2">
    <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" imeMode="hiragana" sqref="D5:D6"><formula1>"a,b"</formula1></dataValidation>
    <dataValidation type="futureType" sqref="E5:E6"><formula1>1</formula1></dataValidation>
  </dataValidations>
</worksheet>"""
      )
    finally out.close()
    path

  test(
    "GH-429 rework: reader-produced Preserved DV/CF payloads shift through insert -> write -> re-read"
  ) {
    val read = reread(preservedPayloadFixture())
    val s0 = read.sheets.headOption.getOrElse(fail("no sheet"))
    val dvBefore = s0.dataValidations.collect { case DataValidation.Preserved(x) => x }
    val cfBefore = s0.conditionalFormats.collect { case ConditionalFormat.Preserved(x) => x }
    assertEquals(dvBefore.size, 2, s"imeMode + unknown-type must ride Preserved: $dvBefore")
    assertEquals(cfBefore.size, 1, s"xr:uid envelope must ride Preserved: $cfBefore")
    assert(
      (dvBefore ++ cfBefore).forall(_.startsWith("<?xml ")),
      s"reader payloads open with the XML declaration: ${(dvBefore ++ cfBefore).map(_.take(60))}"
    )

    // the pure core structural edit must move the Preserved sqrefs, everything else byte-intact
    val edited = markAllModified(read.copy(sheets = read.sheets.map(_.insertRows(1, 2))))
    val s1 = edited.sheets.headOption.getOrElse(fail("no sheet"))
    assertEquals(
      s1.dataValidations.collect { case DataValidation.Preserved(x) => x },
      dvBefore.map(
        _.replace("sqref=\"D5:D6\"", "sqref=\"D7:D8\"")
          .replace("sqref=\"E5:E6\"", "sqref=\"E7:E8\"")
      ),
      "Preserved DV payloads must shift their sqref and change nothing else"
    )
    assertEquals(
      s1.conditionalFormats.collect { case ConditionalFormat.Preserved(x) => x },
      cfBefore.map(_.replace("sqref=\"B5:B6\"", "sqref=\"B7:B8\"")),
      "the Preserved CF block must shift its sqref and change nothing else"
    )

    // write: the shifted payloads land in the part with their foreign attrs intact
    val out = writeTo(edited, "preserved-shift")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(sheetXml.contains("sqref=\"D7:D8\""), sheetXml)
    assert(sheetXml.contains("sqref=\"E7:E8\""), sheetXml)
    assert(sheetXml.contains("sqref=\"B7:B8\""), sheetXml)
    assert(sheetXml.contains("imeMode=\"hiragana\""), sheetXml)
    assert(sheetXml.contains("type=\"futureType\""), sheetXml)
    assert(sheetXml.contains("xr:uid"), sheetXml)

    // re-read clean: the shifted payloads parse right back to the same Preserved entries
    val back = reread(out).sheets.headOption.getOrElse(fail("no sheet"))
    assertEquals(
      back.dataValidations.collect { case DataValidation.Preserved(x) => x },
      s1.dataValidations.collect { case DataValidation.Preserved(x) => x }
    )
    assertEquals(
      back.conditionalFormats.collect { case ConditionalFormat.Preserved(x) => x },
      s1.conditionalFormats.collect { case ConditionalFormat.Preserved(x) => x }
    )
  }

  test("GH-429: a table collapsed by deleteRows leaves NO tableParts, rels, or part behind") {
    val table = TableSpec
      .fromColumnNames("T1", "T1", ref"A1:C5", Vector("R", "P", "U"))
      .fold(e => fail(s"table: $e"), identity)
    val sheet = Sheet("Alpha")
      .put(ref"A1", CellValue.Text("R"))
      .put(ref"D9", CellValue.Text("keep"))
      .withTable(table)
    val src = writeTo(Workbook(sheet), "tbl-src")
    val read = reread(src)
    val collapsed = markAllModified(read.copy(sheets = read.sheets.map(_.deleteRows(0, 5))))
    assertEquals(collapsed.sheets.headOption.map(_.tables), Some(Map.empty[String, TableSpec]))
    val out = writeTo(collapsed, "tbl-drop")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(!sheetXml.contains("<tablePart"), "a dropped table must not resurrect tableParts")
    val zip = new ZipFile(out.toFile)
    val entries =
      try
        val it = zip.entries()
        Iterator.continually(it).takeWhile(_.hasMoreElements).map(_.nextElement.getName).toList
      finally zip.close()
    assert(!entries.exists(_.startsWith("xl/tables/")), s"no table part may ship: $entries")
    // the round-trip must stay readable (the pre-fix output failed with 'Missing table file')
    assertEquals(reread(out).sheets.headOption.map(_.tables.size), Some(0))
  }

  test("GH-429: reader lifts a foreign autoFilter @ref; identity fast-path is churn-free") {
    val (path, wb) =
      val p = TestFixtures.copyToTemp("autofilter.xlsx")
      (p, reread(p))
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    assertEquals(sheet.autoFilter, Some(AutoFilterState.Ranged(ref"A1:C6": CellRange)))

    // unrelated cell edit regenerates the sheet; the untouched filter must ride ref-identical
    val edited = wb.put(sheet.put(ref"E1", CellValue.Text("edited")))
    val out = writeTo(edited, "identity")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(sheetXml.contains("<autoFilter ref=\"A1:C6\""), sheetXml)
    assertEquals(
      reread(out).sheets.headOption.flatMap(_.autoFilter),
      Some(AutoFilterState.Ranged(ref"A1:C6": CellRange))
    )
  }

  test("GH-429: AutoFilterState.Remove strips the element on write (no resurrection)") {
    val path = TestFixtures.copyToTemp("autofilter.xlsx")
    val wb = reread(path)
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val edited = markAllModified(wb.put(sheet.copy(autoFilter = Some(AutoFilterState.Remove))))
    val out = writeTo(edited, "remove")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(!sheetXml.contains("<autoFilter"), sheetXml)
    assertEquals(reread(out).sheets.headOption.flatMap(_.autoFilter), None)
  }

  test("GH-429: a structural edit moves the foreign autoFilter with its data") {
    val path = TestFixtures.copyToTemp("autofilter.xlsx")
    val wb = reread(path)
    val edited = markAllModified(wb.copy(sheets = wb.sheets.map(_.insertRows(0, 3))))
    val out = writeTo(edited, "shifted")
    assert(entryText(out, "xl/worksheets/sheet1.xml").contains("<autoFilter ref=\"A4:C9\""))
    assertEquals(
      reread(out).sheets.headOption.flatMap(_.autoFilter),
      Some(AutoFilterState.Ranged(ref"A4:C9": CellRange))
    )
  }
