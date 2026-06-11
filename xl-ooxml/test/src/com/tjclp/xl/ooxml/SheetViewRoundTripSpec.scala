package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.SheetView
import com.tjclp.xl.unsafe.*
import munit.FunSuite

/**
 * GH-258: sheet view settings (showGridLines, zoomScale) serialize into
 * `<sheetViews><sheetView .../>` and parse back. Freeze panes and view settings must share ONE
 * sheetView element.
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
