package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

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
