package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.{ColumnProperties, DataValidation, RowProperties}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.font.Font
import munit.FunSuite

/**
 * GH-445: `ColumnProperties.styleId` / `RowProperties.styleId` emit as `<col style=>` and `<row s=
 * customFormat="1">`, remapped through StyleIndex exactly like cell styleIds — on BOTH writer
 * backends — and parse back into the model so read→modify→write keeps them.
 *
 * Column-default styles are how professional workbooks make a body font the sheet default without
 * touching the Normal style; row-default styles carry full-width spacer/rule formatting.
 */
class ColRowStyleAuthoringSpec extends FunSuite:

  private val bodyFont = Font("Times New Roman", 10.0)
  private val ruleFont = Font("Times New Roman", 8.0)

  private def zipEntryString(path: Path, entry: String): String =
    val zf = new ZipFile(path.toFile)
    try
      val is = zf.getInputStream(zf.getEntry(entry))
      try new String(is.readAllBytes(), StandardCharsets.UTF_8)
      finally is.close()
    finally zf.close()

  /** A sheet with a styled+sized column B and a styled, cell-free spacer row 3. */
  private def styledSheet: Sheet =
    Sheet("Sheet1")
      .put(ref"A1" -> 1, ref"B2" -> 2)
      .setColumnProperties(ref"B1".col, ColumnProperties(width = Some(20.0)))
      .withColumnStyle(ref"B1".col, CellStyle.default.withFont(bodyFont))
      .setRowProperties(ref"A3".row, RowProperties(height = Some(5.1)))
      .withRowStyle(ref"A3".row, CellStyle.default.withFont(ruleFont))

  private def writeRead(wb: Workbook, prefix: String): (Workbook, Path) =
    val out = Files.createTempFile(prefix, ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    (reread, out)

  private val colStylePattern =
    """<col min="2" max="2" width="20.0" style="(\d+)" customWidth="1"/>""".r
  private val rowStylePattern =
    """<row r="3" s="(\d+)" customFormat="1" ht="5.1" customHeight="1"/>""".r

  private def assertStyledXml(xml: String): Unit =
    val colMatch = colStylePattern.findFirstMatchIn(xml)
    assert(colMatch.isDefined, s"<col> must carry style= between width and customWidth: $xml")
    assert(colMatch.exists(_.group(1).toInt > 0), "col style index must be a real xf, not 0")
    val rowMatch = rowStylePattern.findFirstMatchIn(xml)
    assert(rowMatch.isDefined, s"<row r=3> must carry s= + customFormat before ht: $xml")
    assert(rowMatch.exists(_.group(1).toInt > 0), "row style index must be a real xf, not 0")

  private def assertModelStyles(reread: Workbook): Unit =
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    val colFont = sheet
      .getColumnProperties(ref"B1".col)
      .styleId
      .flatMap(sheet.styleRegistry.get)
      .map(_.font)
    assertEquals(colFont, Some(bodyFont), "column style must resolve through the registry")
    val rowProps = sheet.getRowProperties(ref"A3".row)
    assertEquals(rowProps.height, Some(5.1))
    val rowFont = rowProps.styleId.flatMap(sheet.styleRegistry.get).map(_.font)
    assertEquals(rowFont, Some(ruleFont), "row style must resolve through the registry")

  test("GH-445: col style= and row s= emit and round-trip (streaming writer path)") {
    val (reread, out) = writeRead(Workbook(styledSheet), "colrow-sax")
    assertStyledXml(zipEntryString(out, "xl/worksheets/sheet1.xml"))
    assertModelStyles(reread)
    Files.deleteIfExists(out)
  }

  test("GH-445: col style= and row s= emit identically on the DOM writer path") {
    // A data validation forces XlsxWriter off the streaming path onto the DOM writer
    val sheet = styledSheet.withDataValidation(ref"D1:D3", DataValidation.listOf("Yes", "No"))
    val (reread, out) = writeRead(Workbook(sheet), "colrow-dom")
    assertStyledXml(zipEntryString(out, "xl/worksheets/sheet1.xml"))
    assertModelStyles(reread)
    Files.deleteIfExists(out)
  }

  test("GH-445: read → edit cell → write keeps col/row styles (no silent regeneration drop)") {
    val src = Files.createTempFile("colrow-src", ".xlsx")
    XlsxWriter.write(Workbook(styledSheet), src).fold(e => fail(s"seed write failed: $e"), identity)

    val edited = for
      wb <- XlsxReader.read(src)
      sheet <- wb("Sheet1")
    yield wb.put(sheet.put(ref"C1" -> 3))
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("colrow-out", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)
    assertStyledXml(zipEntryString(out, "xl/worksheets/sheet1.xml"))

    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    assertModelStyles(reread)
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-445: withColumnStyle preserves existing column properties (width survives)") {
    val sheet = styledSheet
    val props = sheet.getColumnProperties(ref"B1".col)
    assertEquals(props.width, Some(20.0))
    assert(props.styleId.isDefined, "styleId must be set alongside the width")
  }

  test("GH-445: two columns sharing one style deduplicate to the same xf index") {
    val style = CellStyle.default.withFont(bodyFont)
    val sheet = Sheet("Sheet1")
      .put(ref"A1" -> 1)
      .withColumnStyle(ref"B1".col, style)
      .withColumnStyle(ref"C1".col, style)
    val out = Files.createTempFile("colrow-dedupe", ".xlsx")
    XlsxWriter.write(Workbook(sheet), out).fold(e => fail(s"write failed: $e"), identity)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    // Identical props on adjacent columns coalesce into ONE span with one style index
    val span = """<col min="2" max="3" style="(\d+)"/>""".r.findFirstMatchIn(xml)
    assert(span.isDefined, s"adjacent same-style columns must coalesce into one span: $xml")
    Files.deleteIfExists(out)
  }
