package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.{HeaderFooter, PageMargins, PageSetup, SheetView}
import com.tjclp.xl.styles.color.{Color, ThemeSlot}
import com.tjclp.xl.workbooks.DefinedName
import munit.FunSuite

/**
 * GH-259: PageSetup print extensions round-trip — scale/orientation/fit via `<pageSetup>`, margins
 * via `<pageMargins>`, header/footer via `<headerFooter>`, and print area / repeat title rows via
 * sheet-scoped workbook defined names (`_xlnm.Print_Area` / `_xlnm.Print_Titles`).
 *
 * GH-358: `Sheet.tabColor` ⇄ `<sheetPr><tabColor .../>` with preserve-if-None merge semantics (the
 * sheetPr overlay lives beside the fitToPage reconciliation, hence this spec).
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
      PageSetup(
        scale = 85,
        orientation = Some("landscape"),
        fitToWidth = Some(1),
        fitToHeight = Some(2)
      )
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val (reread, out) = writeRead(wb)
    assertEquals(sheetSetup(reread), setup)
    Files.deleteIfExists(out)
  }

  test("GH-259: page margins round-trip") {
    val setup = PageSetup(margins =
      Some(
        PageMargins(left = 1.0, right = 0.5, top = 0.75, bottom = 0.75, header = 0.25, footer = 0.4)
      )
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

  test("GH-263: sheet named like a cell ref (Q1) produces quoted print formulas") {
    val setup = PageSetup(printArea = Some(ref"A1:B2"), repeatRows = Some((1, 1)))
    val sheet = Sheet(SheetName.unsafe("Q1")).put(ref"A1", 1).withPageSetup(setup)
    val wb = Workbook(sheet)
    val out = Files.createTempFile("pagesetup-cellref-name", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val workbookXml = zipEntryString(out, "xl/workbook.xml")
    assert(workbookXml.contains("'Q1'!$A$1:$B$2"), s"quoted area formula missing: $workbookXml")
    assert(workbookXml.contains("'Q1'!$1:$1"), s"quoted titles formula missing: $workbookXml")

    val reread = XlsxReader.read(out).fold(e => fail(s"read failed: $e"), identity)
    val rereadSheet = reread(SheetName.unsafe("Q1")).fold(e => fail(s"$e"), identity)
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
      margins = Some(
        PageMargins(left = 0.25, right = 0.25, top = 0.5, bottom = 0.5, header = 0.2, footer = 0.2)
      ),
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

  test("GH-266: even/first headers+footers round-trip with differentOddEven/differentFirst") {
    val hf = HeaderFooter(
      oddHeader = Some("&LOdd Head"),
      oddFooter = Some("&CPage &P of &N"),
      evenHeader = Some("&REven Head"),
      evenFooter = Some("&CEven Foot"),
      firstHeader = Some("&CFirst Head"),
      firstFooter = Some("&CFirst Foot"),
      differentOddEven = true,
      differentFirst = true
    )
    val setup = PageSetup(headerFooter = Some(hf))
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val (reread, out) = writeRead(wb)
    assertEquals(sheetSetup(reread), setup)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("differentOddEven=\"1\""), s"differentOddEven attr missing: $xml")
    assert(xml.contains("differentFirst=\"1\""), s"differentFirst attr missing: $xml")
    assert(xml.contains("<evenHeader>"), s"evenHeader missing: $xml")
    assert(xml.contains("<evenFooter>"), s"evenFooter missing: $xml")
    assert(xml.contains("<firstHeader>"), s"firstHeader missing: $xml")
    assert(xml.contains("<firstFooter>"), s"firstFooter missing: $xml")
    // CT_HeaderFooter child sequence: odd, even, first
    val odd = xml.indexOf("<oddHeader>")
    val even = xml.indexOf("<evenHeader>")
    val first = xml.indexOf("<firstHeader>")
    assert(odd < even && even < first, "headerFooter children must follow schema order")
    Files.deleteIfExists(out)
  }

  test("GH-266: even header round-trips without first-page parts (flags independent)") {
    val hf = HeaderFooter(
      oddHeader = Some("&COdd"),
      evenHeader = Some("&CEven"),
      differentOddEven = true
    )
    val setup = PageSetup(headerFooter = Some(hf))
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val (reread, out) = writeRead(wb)
    assertEquals(sheetSetup(reread), setup)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(!xml.contains("differentFirst"), s"spurious differentFirst attr: $xml")
    assert(!xml.contains("<firstHeader>"), s"spurious firstHeader: $xml")
    Files.deleteIfExists(out)
  }

  test("GH-266: surgical write (cell edit) preserves even/first header+footer parts") {
    val hf = HeaderFooter(
      oddFooter = Some("&CPage &P"),
      evenFooter = Some("&CEven &P"),
      firstHeader = Some("&CCover"),
      differentOddEven = true,
      differentFirst = true
    )
    val wb0 = Workbook(
      Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(PageSetup(headerFooter = Some(hf)))
    )
    val src = Files.createTempFile("evenfirst-src", ".xlsx")
    XlsxWriter.write(wb0, src).fold(e => fail(s"seed write failed: $e"), identity)

    val edited = for
      wb <- XlsxReader.read(src)
      sheet <- wb("Sheet1")
    yield wb.put(sheet.put(ref"B1" -> 2))
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("evenfirst-out", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)
    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    assertEquals(sheetSetup(reread).headerFooter, Some(hf))
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-266: fitToWidth/fitToHeight emit sheetPr/pageSetUpPr fitToPage flag") {
    val setup = PageSetup(fitToWidth = Some(1), fitToHeight = Some(0))
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val (reread, out) = writeRead(wb)
    assertEquals(sheetSetup(reread), setup)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("<pageSetUpPr"), s"pageSetUpPr missing: $xml")
    assert(xml.contains("fitToPage=\"1\""), s"fitToPage flag missing: $xml")
    // sheetPr must come FIRST in the worksheet element order (ECMA-376 18.3.1.99)
    val sheetPr = xml.indexOf("<sheetPr")
    assert(sheetPr >= 0, s"sheetPr missing: $xml")
    assert(sheetPr < xml.indexOf("<dimension"), "sheetPr must precede dimension")
    assert(sheetPr < xml.indexOf("<sheetData"), "sheetPr must precede sheetData")
    Files.deleteIfExists(out)
  }

  test("GH-266: no fitTo* means no fitToPage flag (sheetPr not invented)") {
    val setup = PageSetup(scale = 80, orientation = Some("landscape"))
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val (reread, out) = writeRead(wb)
    assertEquals(sheetSetup(reread), setup)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(!xml.contains("fitToPage"), s"spurious fitToPage flag: $xml")
    assert(!xml.contains("<sheetPr"), s"spurious sheetPr element: $xml")
    Files.deleteIfExists(out)
  }

  test("GH-266: surgical edit keeps fitToPage flag from the source file") {
    val setup = PageSetup(fitToWidth = Some(2), fitToHeight = Some(1))
    val wb0 = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(setup))
    val src = Files.createTempFile("fittopage-src", ".xlsx")
    XlsxWriter.write(wb0, src).fold(e => fail(s"seed write failed: $e"), identity)

    val edited = for
      wb <- XlsxReader.read(src)
      sheet <- wb("Sheet1")
    yield wb.put(sheet.put(ref"B1" -> 2))
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("fittopage-out", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("fitToPage=\"1\""), s"fitToPage lost on surgical edit: $xml")
    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    assertEquals(sheetSetup(reread), setup)
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-284: fitToPage=Some(false) actively clears a preserved fitToPage flag") {
    // Seed a file whose sheetPr carries pageSetUpPr fitToPage="1" (derived from fitTo*)
    val wb0 = Workbook(
      Sheet("Sheet1")
        .put(ref"A1" -> 1)
        .withPageSetup(PageSetup(fitToWidth = Some(1), fitToHeight = Some(0)))
    )
    val src = Files.createTempFile("fittopage-clear-src", ".xlsx")
    XlsxWriter.write(wb0, src).fold(e => fail(s"seed write failed: $e"), identity)
    assert(
      zipEntryString(src, "xl/worksheets/sheet1.xml").contains("fitToPage=\"1\""),
      "seed file must carry the flag"
    )

    // Clear the fit settings AND actively strip the flag
    val cleared = for
      wb <- XlsxReader.read(src)
      sheet <- wb("Sheet1")
      setup = sheet.pageSetup.getOrElse(fail("seed pageSetup missing"))
    yield wb.put(
      sheet.withPageSetup(
        setup.copy(fitToWidth = None, fitToHeight = None, fitToPage = Some(false))
      )
    )
    val wb1 = cleared.fold(e => fail(s"edit failed: $e"), identity)

    val out = Files.createTempFile("fittopage-clear-out", ".xlsx")
    XlsxWriter.write(wb1, out).fold(e => fail(s"write failed: $e"), identity)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(!xml.contains("fitToPage"), s"preserved fitToPage flag not cleared: $xml")
    assert(!xml.contains("<sheetPr"), s"empty sheetPr should be dropped, not kept: $xml")

    val reread = XlsxReader.read(out).fold(e => fail(s"reread failed: $e"), identity)
    val rereadSetup = sheetSetup(reread)
    assertEquals(rereadSetup.fitToWidth, None)
    assertEquals(rereadSetup.fitToHeight, None)
    Files.deleteIfExists(src)
    Files.deleteIfExists(out)
  }

  test("GH-284: Some(false) strips only pageSetUpPr — codeName/tabColor sheetPr content stays") {
    import scala.xml.*
    val preserved =
      XmlSecurity
        .parseSafe(
          """<sheetPr codeName="Sheet1"><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1"/></sheetPr>""",
          "test sheetPr"
        )
        .fold(e => fail(s"parse failed: $e"), identity)
    val merged = com.tjclp.xl.ooxml.worksheet.mergeSheetPrElem(
      Some(preserved),
      Some(PageSetup(fitToPage = Some(false))),
      None
    )
    val elem = merged.getOrElse(fail("sheetPr with codeName/tabColor must survive the strip"))
    assertEquals(elem \@ "codeName", "Sheet1", "codeName attribute lost")
    assert((elem \ "tabColor").nonEmpty, "tabColor child lost")
    assert((elem \ "pageSetUpPr").isEmpty, s"pageSetUpPr should be gone: $elem")
  }

  test("GH-284: Some(false) with nothing preserved emits no sheetPr at all") {
    val merged = com.tjclp.xl.ooxml.worksheet.mergeSheetPrElem(
      None,
      Some(PageSetup(fitToPage = Some(false))),
      None
    )
    assertEquals(merged, None)
  }

  // ===== Tab color (GH-358, model half) =====

  test("GH-358: rgb tabColor round-trips (write → read → equality + XML shape)") {
    val navy = Color.Rgb(0xff1f4e79)
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withTabColor(navy))
    val (reread, out) = writeRead(wb)

    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.tabColor, Some(navy))

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("<sheetPr>"), s"sheetPr missing: $xml")
    assert(xml.contains("<tabColor rgb=\"FF1F4E79\"/>"), s"tabColor element missing: $xml")
    Files.deleteIfExists(out)
  }

  test("GH-358: theme tabColor (slot + tint) round-trips") {
    val theme = Color.Theme(ThemeSlot.Accent2, 0.25)
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withTabColor(theme))
    val (reread, out) = writeRead(wb)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.tabColor, Some(theme))
    Files.deleteIfExists(out)
  }

  test("GH-358: workbook without tabColor reads back None") {
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1))
    val (reread, out) = writeRead(wb)
    val sheet = reread("Sheet1").fold(e => fail(s"sheet missing: $e"), identity)
    assertEquals(sheet.tabColor, None)
    Files.deleteIfExists(out)
  }

  test("GH-358: tabColor + fitToPage keep CT_SheetPr child order (tabColor FIRST)") {
    val wb = Workbook(
      Sheet("Sheet1")
        .put(ref"A1" -> 1)
        .withTabColor(Color.Rgb(0xffff0000))
        .withPageSetup(PageSetup(fitToWidth = Some(1)))
    )
    val (_, out) = writeRead(wb)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    val tabColorIdx = xml.indexOf("<tabColor ")
    val pageSetUpPrIdx = xml.indexOf("<pageSetUpPr ")
    assert(tabColorIdx >= 0, s"tabColor missing: $xml")
    assert(pageSetUpPrIdx >= 0, s"pageSetUpPr missing: $xml")
    assert(
      tabColorIdx < pageSetUpPrIdx,
      s"CT_SheetPr sequence: tabColor must precede pageSetUpPr or Excel repairs the file: $xml"
    )
    Files.deleteIfExists(out)
  }

  test("GH-358: tabColor overlay REPLACES an existing tabColor child in place") {
    import scala.xml.*
    val preserved =
      XmlSecurity
        .parseSafe(
          """<sheetPr codeName="Sheet1"><tabColor rgb="FF00FF00"/><pageSetUpPr fitToPage="1"/></sheetPr>""",
          "test sheetPr"
        )
        .fold(e => fail(s"parse failed: $e"), identity)
    val merged = com.tjclp.xl.ooxml.worksheet.mergeSheetPrElem(
      Some(preserved),
      None,
      Some(Color.Rgb(0xff1f4e79))
    )
    val elem = merged.getOrElse(fail("merged sheetPr missing"))
    assertEquals(elem \@ "codeName", "Sheet1", "codeName attribute lost")
    val tabColors = (elem \ "tabColor").collect { case e: Elem => e }
    assertEquals(tabColors.length, 1, s"exactly one tabColor expected: $elem")
    assertEquals(tabColors.headOption.map(_ \@ "rgb"), Some("FF1F4E79"))
    assert((elem \ "pageSetUpPr").nonEmpty, "unrelated pageSetUpPr child lost")
  }

  test("GH-358: tabColor overlay PREPENDS when the preserved sheetPr has no tabColor") {
    import scala.xml.*
    val preserved =
      XmlSecurity
        .parseSafe(
          """<sheetPr><outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr>""",
          "test sheetPr"
        )
        .fold(e => fail(s"parse failed: $e"), identity)
    val merged = com.tjclp.xl.ooxml.worksheet.mergeSheetPrElem(
      Some(preserved),
      None,
      Some(Color.Rgb(0xffff0000))
    )
    val elem = merged.getOrElse(fail("merged sheetPr missing"))
    val childLabels = elem.child.collect { case e: Elem => e.label }
    assertEquals(
      childLabels.headOption,
      Some("tabColor"),
      s"tabColor must be the FIRST CT_SheetPr child (schema order): $childLabels"
    )
    assertEquals(childLabels, Seq("tabColor", "outlinePr", "pageSetUpPr"))
  }

  test("GH-358: tabColor = None leaves a preserved sheetPr untouched (preserve-if-None)") {
    import scala.xml.*
    val preserved =
      XmlSecurity
        .parseSafe(
          """<sheetPr codeName="S1"><tabColor theme="5" tint="0.25"/></sheetPr>""",
          "test sheetPr"
        )
        .fold(e => fail(s"parse failed: $e"), identity)
    val merged =
      com.tjclp.xl.ooxml.worksheet.mergeSheetPrElem(Some(preserved), None, None)
    assertEquals(merged, Some(preserved), "None must not touch the preserved sheetPr")
  }

  test("GH-358: theme tabColor emits theme/tint attributes (not a resolved rgb)") {
    val theme = Color.Theme(ThemeSlot.Accent1, 0.5)
    val wb = Workbook(Sheet("Sheet1").put(ref"A1" -> 1).withTabColor(theme))
    val (_, out) = writeRead(wb)
    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("<tabColor "), s"tabColor missing: $xml")
    assert(xml.contains("theme=\"4\""), s"theme index missing (Accent1 = 4): $xml")
    assert(xml.contains("tint=\"0.5\""), s"tint missing: $xml")
    Files.deleteIfExists(out)
  }

  test("GH-284: fitToPage=Some(true) forces the flag without fitToWidth/fitToHeight") {
    val wb = Workbook(
      Sheet("Sheet1").put(ref"A1" -> 1).withPageSetup(PageSetup(fitToPage = Some(true)))
    )
    val out = Files.createTempFile("fittopage-force", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("fitToPage=\"1\""), s"forced fitToPage flag missing: $xml")

    // Write-only tri-state (freezePane precedent): the reader keeps fitToPage=None and the flag
    // rides preservation — a subsequent cell edit must keep it.
    val edited = for
      rb <- XlsxReader.read(out)
      sheet <- rb("Sheet1")
    yield rb.put(sheet.put(ref"B1" -> 2))
    val wb1 = edited.fold(e => fail(s"edit failed: $e"), identity)
    val out2 = Files.createTempFile("fittopage-force-2", ".xlsx")
    XlsxWriter.write(wb1, out2).fold(e => fail(s"second write failed: $e"), identity)
    assert(
      zipEntryString(out2, "xl/worksheets/sheet1.xml").contains("fitToPage=\"1\""),
      "forced flag lost on subsequent surgical edit"
    )
    Files.deleteIfExists(out)
    Files.deleteIfExists(out2)
  }

  test("GH-284: explicit Some(false) wins over fitToWidth/fitToHeight derivation") {
    val wb = Workbook(
      Sheet("Sheet1")
        .put(ref"A1" -> 1)
        .withPageSetup(PageSetup(fitToWidth = Some(2), fitToPage = Some(false)))
    )
    val out = Files.createTempFile("fittopage-explicit", ".xlsx")
    XlsxWriter.write(wb, out).fold(e => fail(s"write failed: $e"), identity)

    val xml = zipEntryString(out, "xl/worksheets/sheet1.xml")
    assert(xml.contains("fitToWidth=\"2\""), s"fitToWidth must still serialize: $xml")
    assert(!xml.contains("fitToPage"), s"explicit Some(false) must suppress the flag: $xml")
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
