package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import com.tjclp.xl.addressing.Column
import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.ColumnProperties
import com.tjclp.xl.unsafe.*
import munit.FunSuite

/**
 * GH-426: `Sheet.defaultRowHeight` / `Sheet.defaultColumnWidth` serialize as `<sheetFormatPr
 * defaultRowHeight=... defaultColWidth=...>` (the CT_Worksheet slot before `<cols>`) on BOTH writer
 * backends, and the reader populates the model fields back from an Excel-authored element.
 *
 * A preserved `<sheetFormatPr>` is the dirty-write base: unmodeled attributes (baseColWidth,
 * outlineLevelRow, ...) ride through, and unchanged modeled values keep the source element verbatim
 * (identity fast-path — no churn). Sheets that never had the element and set no defaults must not
 * grow one.
 */
class SheetFormatPrRoundTripSpec extends FunSuite:

  private def zipEntryString(path: Path, entry: String): String =
    val zf = new ZipFile(path.toFile)
    try
      val is = zf.getInputStream(zf.getEntry(entry))
      try new String(is.readAllBytes(), StandardCharsets.UTF_8)
      finally is.close()
    finally zf.close()

  private def writeRead(wb: Workbook, config: WriterConfig = WriterConfig()): (Workbook, Path) =
    val out = Files.createTempFile("sheetformatpr", ".xlsx")
    XlsxWriter.writeWith(wb, out, config).fold(e => fail(s"write failed: $e"), identity)
    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    (reread, out)

  test("GH-426: defaultRowHeight round-trips and emits sheetFormatPr (DOM backend)") {
    val wb = Workbook(
      Sheet("Sheet1").put(ref"A1" -> 1).copy(defaultRowHeight = Some(12.75))
    )
    val (reread, out) = writeRead(wb)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("<sheetFormatPr "), s"sheetFormatPr missing: $xml")
    assert(xml.contains("defaultRowHeight=\"12.75\""), s"defaultRowHeight missing: $xml")
    // A model-set default row height is a manually-set one (ECMA-376 18.3.1.81)
    assert(xml.contains("customHeight=\"1\""), s"customHeight companion missing: $xml")
    assert(
      xml.indexOf("<dimension ") < xml.indexOf("<sheetFormatPr "),
      s"sheetFormatPr must follow dimension: $xml"
    )
    assert(
      xml.indexOf("<sheetFormatPr ") < xml.indexOf("<sheetData>"),
      s"sheetFormatPr must precede sheetData: $xml"
    )

    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.defaultRowHeight, Some(12.75))
    Files.deleteIfExists(out)
  }

  test(
    "GH-426: defaultColumnWidth round-trips; width-only sheets carry required defaultRowHeight"
  ) {
    val wb = Workbook(
      Sheet("Sheet1").put(ref"A1" -> 1).copy(defaultColumnWidth = Some(8.43))
    )
    val (reread, out) = writeRead(wb)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("defaultColWidth=\"8.43\""), s"defaultColWidth missing: $xml")
    // CT_SheetFormatPr REQUIRES defaultRowHeight — a fresh width-only element backfills
    // Excel's Calibri-11 default without the customHeight flag (it was not manually set)
    assert(xml.contains("defaultRowHeight=\"15\""), s"required defaultRowHeight missing: $xml")
    assert(!xml.contains("customHeight="), s"backfilled height must not claim customHeight: $xml")

    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.defaultColumnWidth, Some(8.43))
    assertEquals(sheet.defaultRowHeight, Some(15.0))
    Files.deleteIfExists(out)
  }

  test("GH-426: both defaults round-trip on the streaming backend (SaxStax direct emission)") {
    val wb = Workbook(
      Sheet("Sheet1")
        .put(ref"A1" -> 1)
        .copy(
          defaultRowHeight = Some(12.75),
          defaultColumnWidth = Some(8.43),
          columnProperties = Map(Column.from0(0) -> ColumnProperties(width = Some(20.0)))
        )
    )
    val (reread, out) = writeRead(wb, WriterConfig.saxStax)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("<sheetFormatPr "), s"sheetFormatPr missing on streaming path: $xml")
    assert(xml.contains("defaultRowHeight=\"12.75\""), s"defaultRowHeight missing: $xml")
    assert(xml.contains("defaultColWidth=\"8.43\""), s"defaultColWidth missing: $xml")
    assert(xml.contains("customHeight=\"1\""), s"customHeight companion missing: $xml")
    assert(
      xml.indexOf("<sheetFormatPr ") < xml.indexOf("<cols>"),
      s"sheetFormatPr must precede cols (CT_Worksheet order): $xml"
    )

    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.defaultRowHeight, Some(12.75))
    assertEquals(sheet.defaultColumnWidth, Some(8.43))
    Files.deleteIfExists(out)
  }

  test("GH-426: Excel-authored sheetFormatPr populates the model fields on read") {
    val src = rawWorksheetFixture(
      """<sheetFormatPr baseColWidth="8" defaultColWidth="9.140625" defaultRowHeight="12.75" outlineLevelRow="2" outlineLevelCol="1"/>"""
    )
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val sheet = wb("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.defaultColumnWidth, Some(9.140625))
    assertEquals(sheet.defaultRowHeight, Some(12.75))
    Files.deleteIfExists(src)
  }

  test("GH-426: unmodeled sheetFormatPr attrs ride a dirty write; unchanged values do not churn") {
    val src = rawWorksheetFixture(
      """<sheetFormatPr baseColWidth="8" defaultColWidth="9.140625" defaultRowHeight="12.75" outlineLevelRow="2" outlineLevelCol="1"/>"""
    )
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val sheet0 = wb("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)

    val out = Files.createTempFile("sheetformatpr-ride", ".xlsx")
    XlsxWriter
      .write(wb.put(sheet0.put(ref"B1" -> 2)), out)
      .fold(e => fail(s"write failed: $e"), identity)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("baseColWidth=\"8\""), s"unmodeled attr lost on rewrite: $xml")
    // GH-448: outlineLevelRow/Col graduated from preserved to MODELED — they now derive from
    // row/col outline properties (the zoomScale graduation precedent). This fixture carries the
    // summary attrs with NO outlined rows/cols behind them, so the recomputed (consistent) form
    // drops them; books with real outlined rows keep them (CanonicalXmlSpec, GroupingCommandSpec).
    assert(!xml.contains("outlineLevelRow="), s"inconsistent summary attr must not survive: $xml")
    assert(!xml.contains("outlineLevelCol="), s"inconsistent summary attr must not survive: $xml")
    assert(xml.contains("defaultColWidth=\"9.140625\""), s"source width lost on rewrite: $xml")
    assert(xml.contains("defaultRowHeight=\"12.75\""), s"source height lost on rewrite: $xml")
    // Identity fast-path: values unchanged since read → no gratuitous customHeight flag
    assert(!xml.contains("customHeight="), s"unchanged element must ride verbatim: $xml")
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-426: changing defaultRowHeight on a read sheet overlays the preserved element") {
    val src = rawWorksheetFixture("""<sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>""")
    val edited = for
      wb <- XlsxReader.read(src)
      sheet <- wb("Sheet1")
    yield wb.put(sheet.copy(defaultRowHeight = Some(12.75)))
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("sheetformatpr-edit", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("defaultRowHeight=\"12.75\""), s"model height must win: $xml")
    assert(xml.contains("customHeight=\"1\""), s"manually-set height needs customHeight: $xml")
    assert(xml.contains("baseColWidth=\"8\""), s"unmodeled attr lost on overlay: $xml")
    assert(!xml.contains("defaultRowHeight=\"15\""), s"stale source height must be replaced: $xml")

    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.defaultRowHeight, Some(12.75))
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-426: no gratuitous sheetFormatPr when no defaults are set (both backends)") {
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1))
    for config <- Seq(WriterConfig(), WriterConfig.saxStax) do
      val out = Files.createTempFile("sheetformatpr-none", ".xlsx")
      XlsxWriter.writeWith(wb, out, config).fold(e => fail(s"write failed: $e"), identity)
      val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
      assert(!xml.contains("sheetFormatPr"), s"gratuitous sheetFormatPr ($config): $xml")
      Files.deleteIfExists(out)
  }

  test("GH-426: a source sheet without sheetFormatPr stays without one across a dirty write") {
    val src = rawWorksheetFixture("")
    val edited = for
      wb <- XlsxReader.read(src)
      sheet <- wb("Sheet1")
    yield wb.put(sheet.put(ref"B1" -> 2))
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("sheetformatpr-absent", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(!xml.contains("sheetFormatPr"), s"element must not appear uninvited: $xml")
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  // ========== helpers ==========

  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    out.putNextEntry(new ZipEntry(name))
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()

  /**
   * Minimal Excel-shaped single-sheet fixture with the given raw worksheet XML injected before
   * `<sheetData>` (the SheetViewRoundTripSpec pattern) — for feeding foreign `<sheetFormatPr>`
   * shapes the XL writer never produces.
   */
  private def rawWorksheetFixture(preSheetDataXml: String): Path =
    val path = Files.createTempFile("sheetformatpr-fixture", ".xlsx")
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
        s"""<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  $preSheetDataXml
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"""
      )
    finally out.close()
    path
