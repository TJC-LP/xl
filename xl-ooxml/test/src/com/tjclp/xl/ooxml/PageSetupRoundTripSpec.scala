package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.{HeaderFooter, PageMargins, PageSetup, SheetView}
import com.tjclp.xl.workbooks.DefinedName
import munit.FunSuite

/**
 * GH-259: PageSetup print extensions round-trip — scale/orientation/fit via `<pageSetup>`, margins
 * via `<pageMargins>`, header/footer via `<headerFooter>`, and print area / repeat title rows via
 * sheet-scoped workbook defined names (`_xlnm.Print_Area` / `_xlnm.Print_Titles`).
 */
class PageSetupRoundTripSpec extends FunSuite:

  private def zipEntryString(path: Path, entry: String): String =
    val zf = new ZipFile(path.toFile)
    try
      val is = zf.getInputStream(zf.getEntry(entry))
      try new String(is.readAllBytes(), StandardCharsets.UTF_8)
      finally is.close()
    finally zf.close()

  private def writeRead(wb: Workbook): (Workbook, Path) =
    val out = Files.createTempFile("pagesetup", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)
    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    (reread, out)

  private def sheetSetup(wb: Workbook): PageSetup =
    wb("Sheet1")
      .fold(e => fail(s"sheet missing: $e"), identity)
      .pageSetup
      .getOrElse(fail("pageSetup missing after round-trip"))

  test("GH-259: scale/orientation/fitTo round-trip (previously never serialized)") {
    val setup =
      PageSetup(scale = 85, orientation = Some("landscape"), fitToWidth = Some(1), fitToHeight = Some(2))
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val (reread, out) = writeRead(wb)
    assertEquals(sheetSetup(reread), setup)
    Files.deleteIfExists(out)
  }

  test("GH-259: page margins round-trip") {
    val setup = PageSetup(margins =
      Some(PageMargins(left = 1.0, right = 0.5, top = 0.75, bottom = 0.75, header = 0.25, footer = 0.4))
    )
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val (reread, out) = writeRead(wb)
    assertEquals(sheetSetup(reread), setup)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("<pageMargins"), s"pageMargins element missing: $xml")
    assert(xml.contains("left=\"1.0\""), s"left margin missing: $xml")
    Files.deleteIfExists(out)
  }

  test("GH-259: header/footer with Excel codes round-trip") {
    val setup = PageSetup(headerFooter =
      Some(
        HeaderFooter(
          oddHeader = Some("&LTHE JORDAN COMPANY&RConfidential"),
          oddFooter = Some("&CPage &P of &N — &D")
        )
      )
    )
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val (reread, out) = writeRead(wb)
    assertEquals(sheetSetup(reread), setup)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("<oddHeader>"), s"oddHeader missing: $xml")
    assert(xml.contains("<oddFooter>"), s"oddFooter missing: $xml")
    assert(xml.contains("&amp;P"), s"footer codes must be XML-escaped: $xml")
    Files.deleteIfExists(out)
  }

  test("GH-259: print area round-trips via _xlnm.Print_Area defined name") {
    val setup = PageSetup(printArea = Some(ref"A1:D20"))
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val (reread, out) = writeRead(wb)

    assertEquals(sheetSetup(reread).printArea, Some(ref"A1:D20"))
    // Lifted into PageSetup, not duplicated in metadata
    assertEquals(reread.metadata.definedNames, Vector.empty)

    val workbookXml = zipEntryString(out, "xl/workbook.xml")
    assert(workbookXml.contains("_xlnm.Print_Area"), s"Print_Area name missing: $workbookXml")
    assert(workbookXml.contains("localSheetId=\"0\""), s"sheet scope missing: $workbookXml")
    assert(workbookXml.contains("Sheet1!$A$1:$D$20"), s"area formula wrong: $workbookXml")
    Files.deleteIfExists(out)
  }

  test("GH-259: repeat rows round-trip via _xlnm.Print_Titles defined name") {
    val setup = PageSetup(repeatRows = Some((1, 3)))
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val (reread, out) = writeRead(wb)

    assertEquals(sheetSetup(reread).repeatRows, Some((1, 3)))
    assertEquals(reread.metadata.definedNames, Vector.empty)

    val workbookXml = zipEntryString(out, "xl/workbook.xml")
    assert(workbookXml.contains("_xlnm.Print_Titles"), s"Print_Titles missing: $workbookXml")
    assert(workbookXml.contains("Sheet1!$1:$3"), s"titles formula wrong: $workbookXml")
    Files.deleteIfExists(out)
  }

  test("GH-259: sheet names needing quotes produce quoted print formulas") {
    val setup = PageSetup(printArea = Some(ref"A1:B2"), repeatRows = Some((1, 1)))
    val sheet = Sheet(SheetName.unsafe("Q1 Report")).put(ref"A1", 1).withPageSetup(setup)
    val wb = Workbook(sheet)
    val out = Files.createTempFile("pagesetup-quoted", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val workbookXml = zipEntryString(out, "xl/workbook.xml")
    assert(
      workbookXml.contains("'Q1 Report'!$A$1:$B$2"),
      s"quoted area formula missing: $workbookXml"
    )
    assert(workbookXml.contains("'Q1 Report'!$1:$1"), s"quoted titles formula: $workbookXml")

    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val rereadSheet = reread(SheetName.unsafe("Q1 Report")).fold(e => fail(s"$e"), identity)
    assertEquals(rereadSheet.pageSetup, Some(setup))
    Files.deleteIfExists(out)
  }

  test("GH-259: every print field round-trips together (full PageSetup equality)") {
    val setup = PageSetup(
      scale = 90,
      orientation = Some("landscape"),
      fitToWidth = Some(1),
      fitToHeight = Some(0),
      headerFooter = Some(HeaderFooter(oddHeader = Some("&A"), oddFooter = Some("Page &P of &N"))),
      margins = Some(PageMargins(left = 0.25, right = 0.25, top = 0.5, bottom = 0.5, header = 0.2, footer = 0.2)),
      printArea = Some(ref"A1:H44"),
      repeatRows = Some((1, 2))
    )
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val (reread, out) = writeRead(wb)
    assertEquals(sheetSetup(reread), setup)
    Files.deleteIfExists(out)
  }

  test("GH-259: worksheet elements appear in schema order (pageMargins, pageSetup, headerFooter)") {
    val setup = PageSetup(
      scale = 80,
      margins = Some(PageMargins.default),
      headerFooter = Some(HeaderFooter(oddFooter = Some("&P")))
    )
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val out = Files.createTempFile("pagesetup-order", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    val margins = xml.indexOf("<pageMargins")
    val pageSetup = xml.indexOf("<pageSetup")
    val headerFooter = xml.indexOf("<headerFooter")
    assert(margins >= 0 && pageSetup >= 0 && headerFooter >= 0, s"elements missing: $xml")
    assert(margins < pageSetup, "pageMargins must precede pageSetup")
    assert(pageSetup < headerFooter, "headerFooter must follow pageSetup")
    assert(xml.indexOf("<sheetData>") < margins, "print elements follow sheetData")
    Files.deleteIfExists(out)
  }

  test("GH-259: surgical write (cell edit) preserves print setup from the source file") {
    val setup = PageSetup(
      orientation = Some("portrait"),
      margins = Some(PageMargins(left = 1.0)),
      headerFooter = Some(HeaderFooter(oddFooter = Some("Page &P"))),
      printArea = Some(ref"A1:C10"),
      repeatRows = Some((1, 1))
    )
    val wb0 = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val src = Files.createTempFile("pagesetup-src", ".xlsx")
    XlsxWriter.write(wb0, src).fold(e => fail(s"seed write failed: $e"), identity)

    val edited = for
      wb <- XlsxReader.read(src)
      sheet <- wb("Sheet1")
    yield wb.put(sheet.put(ref"B1" -> 2))
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("pagesetup-out", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)

    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    assertEquals(sheetSetup(reread), setup)
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-259: changing the print area on a read workbook regenerates the defined name") {
    val wb0 = Workbook(
      Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(PageSetup(printArea = Some(ref"A1:B2")))
    )
    val src = Files.createTempFile("pagesetup-edit-src", ".xlsx")
    XlsxWriter.write(wb0, src).fold(e => fail(s"seed write failed: $e"), identity)

    val edited = for
      wb <- XlsxReader.read(src)
      updated <- wb.updateAt(0, _.withPageSetup(PageSetup(printArea = Some(ref"C1:D2"))))
    yield updated
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("pagesetup-edit-out", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)

    val workbookXml = zipEntryString(out, "xl/workbook.xml")
    assert(workbookXml.contains("Sheet1!$C$1:$D$2"), s"updated area missing: $workbookXml")
    assert(!workbookXml.contains("Sheet1!$A$1:$B$2"), s"stale area lingering: $workbookXml")

    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    assertEquals(sheetSetup(reread).printArea, Some(ref"C1:D2"))
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-259: unmodelable Print_Titles (column span) stays in metadata.definedNames verbatim") {
    val colTitles =
      DefinedName("_xlnm.Print_Titles", "Sheet1!$A:$B", localSheetId = Some(0))
    val base = Workbook(Sheet("Sheet1").put(ref"A1" -> 1))
    val wb = base.copy(metadata = base.metadata.copy(definedNames = Vector(colTitles)))
    val (reread, out) = writeRead(wb)

    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.pageSetup.flatMap(_.repeatRows), None)
    assertEquals(reread.metadata.definedNames, Vector(colTitles))
    Files.deleteIfExists(out)
  }

  test("GH-258/GH-259: freeze panes + view settings + print setup round-trip together") {
    val view = SheetView(showGridLines = false, zoomScale = Some(90))
    val setup = PageSetup(
      orientation = Some("landscape"),
      margins = Some(PageMargins.default),
      headerFooter = Some(HeaderFooter(oddFooter = Some("&CPage &P of &N"))),
      printArea = Some(ref"A1:F30"),
      repeatRows = Some((1, 2))
    )
    val wb = Workbook(
      Sheet("Sheet1")
        .put(ref"A1" -> "Title", ref"A2" -> 1)
        .freezeAt(ref"A3")
        .withViewSettings(view)
        .withPageSetup(setup)
    )
    val (reread, out) = writeRead(wb)

    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.viewSettings, Some(view))
    assertEquals(sheet.pageSetup, Some(setup))

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("<pane "), s"freeze pane missing: $xml")
    assert(xml.contains("showGridLines=\"0\""), s"view settings missing: $xml")
    Files.deleteIfExists(out)
  }
