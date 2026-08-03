package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.{FreezePane, SheetView}
import com.tjclp.xl.unsafe.*
import munit.FunSuite

/**
 * GH-258: sheet view settings (showGridLines, zoomScale) serialize into
 * `<sheetViews><sheetView .../>` and parse back. Freeze panes and view settings must share ONE
 * sheetView element.
 *
 * GH-372: the reader populates `Sheet.freezePane` from frozen `<pane>` elements and
 * `SheetView.tabSelected` from the sheetView attribute, so read→modify→write keeps view state
 * through model regeneration.
 */
class SheetViewRoundTripSpec extends FunSuite:

  private def zipEntryString(path: Path, entry: String): String =
    val zf = new ZipFile(path.toFile)
    try
      val is = zf.getInputStream(zf.getEntry(entry))
      try new String(is.readAllBytes(), StandardCharsets.UTF_8)
      finally is.close()
    finally zf.close()

  private def writeRead(wb: Workbook): (Workbook, Path) =
    val out = Files.createTempFile("sheetview", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    (reread, out)

  test("GH-258: gridlines-off + zoom round-trips (write → read → equality)") {
    val view = SheetView(showGridLines = false, zoomScale = Some(85))
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withViewSettings(view))
    val (reread, out) = writeRead(wb)

    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.viewSettings, Some(view))

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("showGridLines=\"0\""), s"showGridLines attr missing: $xml")
    assert(xml.contains("zoomScale=\"85\""), s"zoomScale attr missing: $xml")
    Files.deleteIfExists(out)
  }

  test("GH-258: default view settings (gridlines on, no zoom) round-trip when explicitly set") {
    val view = SheetView(showGridLines = true, zoomScale = None)
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withViewSettings(view))
    val (reread, out) = writeRead(wb)

    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.viewSettings, Some(view))
    Files.deleteIfExists(out)
  }

  test("GH-258: workbook without view settings reads back as None (passive default)") {
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1))
    val (reread, out) = writeRead(wb)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.viewSettings, None)
    Files.deleteIfExists(out)
  }

  test("GH-258: freeze panes + view settings share ONE sheetView element and round-trip") {
    val view = SheetView(showGridLines = false, zoomScale = Some(120))
    val wb = Workbook(
      Sheet("Sheet1")
        .put(ref"A1" -> "Header", ref"A2" -> 1)
        .freezeAt(ref"B2")
        .withViewSettings(view)
    )
    val (reread, out) = writeRead(wb)

    // Domain: view settings round-trip
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.viewSettings, Some(view))

    // XML: exactly one <sheetView ...> element carrying BOTH the pane and the view attributes,
    // in spec order (sheetViews before sheetData)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    val sheetViewCount = xml.sliding("<sheetView ".length).count(_ == "<sheetView ")
    assertEquals(sheetViewCount, 1, s"expected one sheetView element: $xml")
    assert(xml.contains("showGridLines=\"0\""), s"view attrs missing: $xml")
    assert(xml.contains("zoomScale=\"120\""), s"zoom attr missing: $xml")
    assert(xml.contains("<pane "), s"freeze pane missing: $xml")
    assert(xml.contains("topLeftCell=\"B2\""), s"freeze anchor missing: $xml")
    assert(
      xml.indexOf("<sheetViews>") < xml.indexOf("<sheetData>"),
      s"sheetViews must precede sheetData: $xml"
    )
    Files.deleteIfExists(out)
  }

  test("GH-382: scrolled frozen pane writes anchor-derived splits and the scroll topLeftCell") {
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).freezeAt(ref"C11", ref"F40"))
    val out = Files.createTempFile("scrolled-pane", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("xSplit=\"2\""), s"xSplit derived from the anchor missing: $xml")
    assert(xml.contains("ySplit=\"10\""), s"ySplit derived from the anchor missing: $xml")
    assert(xml.contains("topLeftCell=\"F40\""), s"scroll target must win topLeftCell: $xml")
    assert(xml.contains("state=\"frozen\""), s"frozen state missing: $xml")
    Files.deleteIfExists(out)
  }

  // ===== GH-372: freeze panes read back into the model =====

  test("GH-372: freeze pane reads back into the model (write → read)") {
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).freezeAt(ref"B2"))
    val (reread, out) = writeRead(wb)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.freezePane, Some(FreezePane.At(ref"B2", None)))
    Files.deleteIfExists(out)
  }

  test("GH-372/GH-382: scrolled frozen pane round-trips through the model") {
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).freezeAt(ref"C11", ref"F40"))
    val (reread, out) = writeRead(wb)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.freezePane, Some(FreezePane.At(ref"C11", Some(ref"F40"))))
    Files.deleteIfExists(out)
  }

  test("GH-372: workbook without freeze panes reads back None (never Some(Remove))") {
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1))
    val (reread, out) = writeRead(wb)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.freezePane, None)
    Files.deleteIfExists(out)
  }

  test("GH-372: read → edit cell → write retains a foreign scrolled pane (model regeneration)") {
    val src = rawWorksheetFixture(
      """<sheetViews><sheetView rightToLeft="1" workbookViewId="0"><pane xSplit="2" ySplit="10" topLeftCell="C40" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>"""
    )
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val sheet0 = wb("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(
      sheet0.freezePane,
      Some(FreezePane.At(ref"C11", Some(ref"C40"))),
      "anchor derives from xSplit/ySplit; pane topLeftCell is the scroll target"
    )

    val out = Files.createTempFile("pane-retain", ".xlsx")
    XlsxWriter
      .write(wb.put(sheet0.put(ref"B1" -> 2)), out)
      .fold(e => fail(s"write failed: $e"), identity)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("xSplit=\"2\""), s"xSplit lost on rewrite: $xml")
    assert(xml.contains("ySplit=\"10\""), s"ySplit lost on rewrite: $xml")
    assert(xml.contains("topLeftCell=\"C40\""), s"scroll target lost on rewrite: $xml")
    assert(xml.contains("state=\"frozen\""), s"frozen state lost on rewrite: $xml")
    assert(xml.contains("rightToLeft=\"1\""), s"foreign sheetView attr lost on rewrite: $xml")
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-372: unscrolled foreign pane reads back with scrolledTo = None") {
    val src = rawWorksheetFixture(
      """<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="10" topLeftCell="C11" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>"""
    )
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val sheet = wb("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.freezePane, Some(FreezePane.At(ref"C11", None)))
    Files.deleteIfExists(src)
  }

  test("GH-372: plain split panes stay unmodeled and ride preservation on edit") {
    // No state attribute → state defaults to "split": xSplit/ySplit are 1/20 pt, not row/col
    // counts, so the pane must NOT enter the model — and must survive a cell edit verbatim.
    val src = rawWorksheetFixture(
      """<sheetViews><sheetView workbookViewId="0"><pane xSplit="2310" ySplit="990" topLeftCell="C11" activePane="bottomRight"/></sheetView></sheetViews>"""
    )
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val sheet = wb("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.freezePane, None, "split panes are out of the modeled subset")

    val out = Files.createTempFile("split-pane", ".xlsx")
    XlsxWriter
      .write(wb.put(sheet.put(ref"B1" -> 2)), out)
      .fold(e => fail(s"write failed: $e"), identity)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("xSplit=\"2310\""), s"split pane lost on rewrite: $xml")
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  // ===== GH-358: foreign sheetPr / tabColor shapes =====

  test("GH-358: unparseable tabColor (auto) stays out of the model and rides preservation") {
    val src = rawWorksheetFixture("""<sheetPr><tabColor auto="1"/></sheetPr>""")
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val sheet = wb("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.tabColor, None, "auto tabColor is outside the typed color subset")

    val out = Files.createTempFile("tabcolor-auto", ".xlsx")
    XlsxWriter
      .write(wb.put(sheet.put(ref"B1" -> 2)), out)
      .fold(e => fail(s"write failed: $e"), identity)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("<tabColor auto=\"1\"/>"), s"foreign tabColor lost on rewrite: $xml")
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-358: setting tabColor on a read sheet keeps foreign sheetPr attrs and children") {
    val src = rawWorksheetFixture(
      """<sheetPr codeName="Hoja1" filterMode="1"><outlinePr summaryBelow="0"/></sheetPr>"""
    )
    val recolored = for
      wb <- XlsxReader.read(src)
      updated <- wb.updateAt(0, _.withTabColor(com.tjclp.xl.styles.color.Color.Rgb(0xff1f4e79)))
    yield updated
    val wb1 = recolored.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("tabcolor-foreign", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("<tabColor rgb=\"FF1F4E79\"/>"), s"tabColor missing: $xml")
    assert(xml.contains("codeName=\"Hoja1\""), s"foreign sheetPr attr lost: $xml")
    assert(xml.contains("filterMode=\"1\""), s"foreign sheetPr attr lost: $xml")
    assert(xml.contains("summaryBelow=\"0\""), s"foreign sheetPr child lost: $xml")
    // CT_SheetPr sequence: tabColor precedes outlinePr
    assert(
      xml.indexOf("<tabColor ") < xml.indexOf("<outlinePr "),
      s"tabColor must be prepended before other sheetPr children: $xml"
    )
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  // ===== GH-372: tabSelected =====

  test("GH-372: tabSelected round-trips (write → read)") {
    val view = SheetView(showGridLines = false, tabSelected = Some(true))
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withViewSettings(view))
    val (reread, out) = writeRead(wb)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.viewSettings, Some(view))

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("tabSelected=\"1\""), s"tabSelected attr missing: $xml")
    Files.deleteIfExists(out)
  }

  test("GH-372: parseSheetView fires when ONLY tabSelected is present") {
    val src = rawWorksheetFixture(
      """<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>"""
    )
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val sheet = wb("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(
      sheet.viewSettings,
      Some(SheetView(showGridLines = true, zoomScale = None, tabSelected = Some(true)))
    )
    Files.deleteIfExists(src)
  }

  test("GH-372: tabSelected = None leaves a preserved tabSelected attribute untouched") {
    val src = rawWorksheetFixture(
      """<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>"""
    )
    val edited = for
      wb <- XlsxReader.read(src)
      updated <- wb.updateAt(0, _.withViewSettings(SheetView(showGridLines = false)))
    yield updated
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("tabsel-preserve", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("showGridLines=\"0\""), s"modeled attr not overlaid: $xml")
    assert(
      xml.contains("tabSelected=\"1\""),
      s"tabSelected=None must leave the preserved attribute: $xml"
    )
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-258: surgical write (cell edit) preserves view settings from the source file") {
    val view = SheetView(showGridLines = false, zoomScale = Some(75))
    val wb0 = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withViewSettings(view))
    val src = Files.createTempFile("sheetview-src", ".xlsx")
    XlsxWriter.write(wb0, src).fold(e => fail(s"seed write failed: $e"), identity)

    val edited = for
      wb <- XlsxReader.read(src)
      sheet <- wb("Sheet1")
    yield wb.put(sheet.put(ref"B1" -> 2))
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("sheetview-out", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)

    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.viewSettings, Some(view))
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-258: view settings can be changed on a read workbook (sheet-level edit)") {
    val wb0 = Workbook(
      Sheet("Sheet1").put(ref"A1" -> 1).withViewSettings(SheetView(showGridLines = true))
    )
    val src = Files.createTempFile("sheetview-edit-src", ".xlsx")
    XlsxWriter.write(wb0, src).fold(e => fail(s"seed write failed: $e"), identity)

    val newView = SheetView(showGridLines = false, zoomScale = Some(60))
    val edited = for
      wb <- XlsxReader.read(src)
      updated <- wb.updateAt(0, _.withViewSettings(newView))
    yield updated
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("sheetview-edit-out", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)

    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.viewSettings, Some(newView))
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  // ===== GH-446: view mode, per-mode zooms, scroll position =====

  test("GH-446: view mode + zoom variants + topLeftCell round-trip (write → read → equality)") {
    val view = SheetView(
      showGridLines = false,
      zoomScale = Some(60),
      view = Some("pageBreakPreview"),
      zoomScaleNormal = Some(70),
      zoomScaleSheetLayoutView = Some(85),
      topLeftCell = Some(ref"A14")
    )
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withViewSettings(view))
    val (reread, out) = writeRead(wb)

    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.viewSettings, Some(view))

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("view=\"pageBreakPreview\""), s"view attr missing: $xml")
    assert(xml.contains("zoomScale=\"60\""), s"zoomScale attr missing: $xml")
    assert(xml.contains("zoomScaleNormal=\"70\""), s"zoomScaleNormal attr missing: $xml")
    assert(
      xml.contains("zoomScaleSheetLayoutView=\"85\""),
      s"zoomScaleSheetLayoutView attr missing: $xml"
    )
    assert(xml.contains("topLeftCell=\"A14\""), s"topLeftCell attr missing: $xml")
    Files.deleteIfExists(out)
  }

  test("GH-446: parseSheetView fires when ONLY view= is present") {
    val src = rawWorksheetFixture(
      """<sheetViews><sheetView view="pageBreakPreview" workbookViewId="0"/></sheetViews>"""
    )
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val sheet = wb("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.viewSettings, Some(SheetView(view = Some("pageBreakPreview"))))
    Files.deleteIfExists(src)
  }

  test("GH-446: an unknown foreign view value stays unmodeled (rides preservation)") {
    val src = rawWorksheetFixture(
      """<sheetViews><sheetView view="weirdMode" workbookViewId="0"/></sheetViews>"""
    )
    val wb = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    val sheet = wb("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.viewSettings, None, "unknown view modes are outside the modeled subset")
    Files.deleteIfExists(src)
  }

  test("GH-446: illegal view mode is rejected at construction") {
    intercept[IllegalArgumentException](SheetView(view = Some("bogus")))
    intercept[IllegalArgumentException](SheetView(zoomScaleNormal = Some(5)))
    intercept[IllegalArgumentException](SheetView(zoomScaleSheetLayoutView = Some(500)))
  }

  test("GH-446: view mode + freeze pane share ONE sheetView (pane keeps its own topLeftCell)") {
    val view = SheetView(
      showGridLines = false,
      view = Some("pageBreakPreview"),
      zoomScaleSheetLayoutView = Some(60)
    )
    val wb = Workbook(
      Sheet("Sheet1").put(ref"A1" -> "H", ref"A2" -> 1).freezeAt(ref"G13").withViewSettings(view)
    )
    val (reread, out) = writeRead(wb)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.viewSettings, Some(view))

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    val sheetViewCount = xml.sliding("<sheetView ".length).count(_ == "<sheetView ")
    assertEquals(sheetViewCount, 1, s"expected one sheetView element: $xml")
    assert(xml.contains("view=\"pageBreakPreview\""), s"view attr missing: $xml")
    assert(xml.contains("xSplit=\"6\""), s"freeze xSplit missing: $xml")
    assert(xml.contains("ySplit=\"12\""), s"freeze ySplit missing: $xml")
    assert(xml.contains("topLeftCell=\"G13\""), s"pane topLeftCell missing: $xml")
    Files.deleteIfExists(out)
  }

  // ========== helpers ==========

  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    out.putNextEntry(new ZipEntry(name))
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()

  /**
   * Minimal Excel-shaped single-sheet fixture with the given raw worksheet XML injected before
   * `<sheetData>` (the ActiveTabRoundTripSpec pattern) — for feeding foreign `<sheetViews>` /
   * `<sheetPr>` shapes the XL writer never produces.
   */
  private def rawWorksheetFixture(preSheetDataXml: String): Path =
    val path = Files.createTempFile("sheetview-fixture", ".xlsx")
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
