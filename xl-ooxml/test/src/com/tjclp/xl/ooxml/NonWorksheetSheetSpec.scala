package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.xml.Elem

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.metadata.WorkbookMetadataReader
import com.tjclp.xl.workbooks.DefinedName
import munit.FunSuite

/**
 * GH-785: `<sheets>` entries that are not worksheets — Excel 4.0 macro sheets (plain and
 * international), chartsheets and dialog sheets — are read past, kept byte for byte on write, and
 * every `<sheets>` index (`localSheetId`, `activeTab`) is renumbered between the file's positions
 * and the model's worksheet-only positions.
 *
 * Fixture `<sheets>`: Data (sheetId 1), the non-worksheet `Other` (sheetId 2), Data2 (sheetId 3).
 */
class NonWorksheetSheetSpec extends FunSuite:

  private val mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  private val relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  private final case class Kind(
    label: String,
    relType: String,
    partPath: String,
    contentType: String,
    xml: String
  )

  private val macroSheet = Kind(
    "macro sheet",
    "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
    "xl/macrosheets/sheet1.xml",
    "application/vnd.ms-excel.macrosheet+xml",
    s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xm:macrosheet xmlns="$mainNs" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row></sheetData></xm:macrosheet>"""
  )

  private val intlMacro = Kind(
    "international macro sheet",
    "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet",
    "xl/macrosheets/intlsheet1.xml",
    "application/vnd.ms-excel.intlmacrosheet+xml",
    s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xm:macrosheet xmlns="$mainNs" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData/></xm:macrosheet>"""
  )

  private val chartsheet = Kind(
    "chartsheet",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
    "xl/chartsheets/sheet1.xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
    s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<chartsheet xmlns="$mainNs"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>"""
  )

  private val dialogsheet = Kind(
    "dialog sheet",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
    "xl/dialogsheets/sheet1.xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
    s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dialogsheet xmlns="$mainNs"><sheetViews><sheetView workbookViewId="0"/></sheetViews></dialogsheet>"""
  )

  private val strictChartsheet = chartsheet.copy(
    label = "Strict-namespace chartsheet",
    relType = "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet"
  )

  private val kinds = List(macroSheet, intlMacro, chartsheet, dialogsheet, strictChartsheet)

  private val dataName = SheetName.unsafe("Data")
  private val data2Name = SheetName.unsafe("Data2")

  // ========== Read ==========

  kinds.foreach { kind =>
    test(s"GH-785: a ${kind.label} in <sheets> reads; the model holds the worksheets only") {
      val source = fixture(kind)
      val wb = read(source)
      assertEquals(wb.sheetNames.map(_.value), Seq("Data", "Data2"))
      assertEquals(valueAt(wb, "Data", ref"A1"), Some(CellValue.Text("Alpha")))
      assertEquals(valueAt(wb, "Data2", ref"A1"), Some(CellValue.Number(BigDecimal(7))))
      cleanup(source)
    }
  }

  test("GH-785: a name scoped to a worksheet after the macro sheet reads with that worksheet") {
    val source = fixture(macroSheet)
    val wb = read(source)
    val names = wb.metadata.definedNames
    assertEquals(names.find(_.name == "Local").flatMap(_.localSheetId), Some(1))
    assertEquals(names.find(_.name == "Local").map(_.formula), Some("Data2!$A$1"))
    // a name scoped to the macro sheet has no model scope: it is not in the model
    assert(!names.exists(_.name == "MacroLocal"), names.toString)
    // workbook-scoped names that point into the macro sheet stay as text
    assert(names.exists(n => n.name == "_xlnm.Auto_Open" && n.localSheetId.isEmpty))
    // Print_Titles on Data2 (localSheetId 2 in the file) lifts onto Data2
    assert(wb(data2Name).toOption.flatMap(_.pageSetup).exists(_.repeatRows.nonEmpty))
    cleanup(source)
  }

  test("GH-785: activeTab on a worksheet after the macro sheet reads as that worksheet") {
    val source = fixture(macroSheet, activeTab = Some(2))
    assertEquals(read(source).activeSheetIndex, 1)
    cleanup(source)
  }

  test("GH-785: activeTab on the macro sheet reads as the next worksheet") {
    val source = fixture(macroSheet, activeTab = Some(1))
    assertEquals(read(source).activeSheetIndex, 1)
    cleanup(source)
  }

  test("GH-785: the light metadata reader lists worksheets only and renumbers scopes") {
    val source = fixture(macroSheet)
    val meta = WorkbookMetadataReader.read(source).fold(e => fail(e.message), identity)
    assertEquals(meta.sheets.map(_.name.value), Vector("Data", "Data2"))
    assertEquals(meta.definedNames.find(_.name == "Local").flatMap(_.localSheetId), Some(1))
    assert(!meta.definedNames.exists(_.name == "MacroLocal"))
    val list = WorkbookMetadataReader.readSheetList(source).fold(e => fail(e.message), identity)
    assertEquals(list.map(_.name.value), Vector("Data", "Data2"))
    cleanup(source)
  }

  // ========== Write ==========

  kinds.foreach { kind =>
    test(s"GH-785: a cell edit keeps the ${kind.label}'s entry, part, rel and content type") {
      val source = fixture(kind)
      val wb = read(source)
      val output = write(putA2(wb, "NewText"))

      assertEquals(entry(output, kind.partPath).toSeq, entry(source, kind.partPath).toSeq)
      assertEquals(sheetEntries(output), sheetEntries(source))
      val rels = text(output, "xl/_rels/workbook.xml.rels")
      assert(
        rels.contains(s"""Type="${kind.relType}" Target="${kind.partPath.stripPrefix("xl/")}""""),
        rels
      )
      assert(text(output, "[Content_Types].xml").contains(kind.contentType))
      assertEquals(names(output), names(source))

      val reloaded = read(output)
      assertEquals(reloaded.sheetNames.map(_.value), Seq("Data", "Data2"))
      assertEquals(valueAt(reloaded, "Data", ref"A2"), Some(CellValue.Text("NewText")))
      cleanup(source, output)
    }
  }

  test("GH-785: a write that appends a string keeps the macro sheet's t=\"s\" index valid") {
    val source = fixture(macroSheet)
    val output = write(putA2(read(source), "NewText"))
    val sst = XmlSecurity.parseSafe(text(output, "xl/sharedStrings.xml"), "sst").toOption.get
    val strings = (sst \ "si").map(_.text).toVector
    // the macro sheet's cell holds index 1: it still names MacroText
    assertEquals(strings.lift(1), Some("MacroText"))
    assertEquals(entry(output, macroSheet.partPath).toSeq, entry(source, macroSheet.partPath).toSeq)
    cleanup(source, output)
  }

  test("GH-785: a name scoped to Data2 keeps its scope through a defined-name edit") {
    val source = fixture(macroSheet)
    val wb = read(source).withDefinedName("Added", "Data!$B$1")
    val output = write(wb)
    val written = names(output)
    // Data2 is still <sheets> position 2, the macro sheet still position 1
    assert(written.contains(("Local", Some(2), "Data2!$A$1")), written.toString)
    assert(written.contains(("MacroLocal", Some(1), "Other!$A$1")), written.toString)
    assert(written.contains(("_xlnm.Auto_Open", None, "Other!$A$1")), written.toString)
    assert(written.contains(("_xlnm.Print_Titles", Some(2), "Data2!$1:$1")), written.toString)
    assert(written.contains(("Added", None, "Data!$B$1")), written.toString)
    assertEquals(entry(output, macroSheet.partPath).toSeq, entry(source, macroSheet.partPath).toSeq)
    cleanup(source, output)
  }

  test("GH-785: adding a worksheet keeps the macro sheet's entry and allocates a fresh sheetId") {
    val source = fixture(macroSheet)
    val wb =
      read(source).addSheet(Sheet(SheetName.unsafe("New"))).fold(e => fail(e.message), identity)
    val output = write(wb)
    assertEquals(
      sheetEntries(output),
      Vector(
        ("Data", "1", "rId1"),
        ("Other", "2", "rId2"),
        ("Data2", "3", "rId3"),
        ("New", "4", sheetEntries(output).last._3)
      )
    )
    assert(names(output).contains(("Local", Some(2), "Data2!$A$1")))
    assert(names(output).contains(("MacroLocal", Some(1), "Other!$A$1")))
    assertEquals(read(output).sheetNames.map(_.value), Seq("Data", "Data2", "New"))
    cleanup(source, output)
  }

  test("GH-785: deleting the worksheet before the macro sheet moves the macro sheet first") {
    val source = fixture(macroSheet)
    val wb = read(source).delete(dataName).fold(e => fail(e.message), identity)
    val output = write(wb)
    assertEquals(sheetEntries(output), Vector(("Other", "2", "rId2"), ("Data2", "3", "rId3")))
    val written = names(output)
    assert(written.contains(("Local", Some(1), "Data2!$A$1")), written.toString)
    assert(written.contains(("MacroLocal", Some(0), "Other!$A$1")), written.toString)
    assert(written.contains(("_xlnm.Print_Titles", Some(1), "Data2!$1:$1")), written.toString)
    assertEquals(entry(output, macroSheet.partPath).toSeq, entry(source, macroSheet.partPath).toSeq)
    assertEquals(read(output).sheetNames.map(_.value), Seq("Data2"))
    cleanup(source, output)
  }

  test("GH-785: deleting the worksheet after the macro sheet keeps the macro sheet in place") {
    val source = fixture(macroSheet)
    val wb = read(source).delete(data2Name).fold(e => fail(e.message), identity)
    val output = write(wb)
    assertEquals(sheetEntries(output), Vector(("Data", "1", "rId1"), ("Other", "2", "rId2")))
    val written = names(output)
    assert(!written.exists(_._1 == "Local"), written.toString)
    assert(written.contains(("MacroLocal", Some(1), "Other!$A$1")), written.toString)
    cleanup(source, output)
  }

  test("GH-785: renaming the worksheet before the macro sheet keeps the macro sheet after it") {
    val source = fixture(macroSheet)
    val wb =
      read(source)
        .rename(dataName, SheetName.unsafe("Renamed"))
        .fold(e => fail(e.message), identity)
    val output = write(wb)
    assertEquals(
      sheetEntries(output).map((name, _, rId) => (name, rId)),
      Vector(("Renamed", "rId1"), ("Other", "rId2"), ("Data2", "rId3"))
    )
    assert(sheetEntries(output).contains(("Other", "2", "rId2")))
    assert(names(output).contains(("Local", Some(2), "Data2!$A$1")))
    cleanup(source, output)
  }

  test("GH-785: reordering worksheets carries the macro sheet with the worksheet before it") {
    val source = fixture(macroSheet)
    val wb = read(source).reorder(Vector(data2Name, dataName)).fold(e => fail(e.message), identity)
    val output = write(wb)
    assertEquals(
      sheetEntries(output),
      Vector(("Data2", "3", "rId3"), ("Data", "1", "rId1"), ("Other", "2", "rId2"))
    )
    val written = names(output)
    assert(written.contains(("Local", Some(0), "Data2!$A$1")), written.toString)
    assert(written.contains(("MacroLocal", Some(2), "Other!$A$1")), written.toString)
    cleanup(source, output)
  }

  test("GH-785: activeTab maps through the macro sheet on write") {
    val onData2 = fixture(macroSheet, activeTab = Some(2))
    val out1 = write(putA2(read(onData2), "x"))
    assertEquals(activeTab(out1), Some(2))

    // the source's tab on the macro sheet stays there while the model's active sheet is unchanged
    val onMacro = fixture(macroSheet, activeTab = Some(1))
    val out2 = write(putA2(read(onMacro), "x"))
    assertEquals(activeTab(out2), Some(1))

    // choosing a worksheet moves it
    val chosen = read(onMacro).setActiveSheet(0).fold(e => fail(e.message), identity)
    val out3 = write(chosen)
    assertEquals(activeTab(out3), Some(0))
    cleanup(onData2, out1, onMacro, out2, out3)
  }

  test("GH-785: a worksheet may not take the macro sheet's name") {
    val source = fixture(macroSheet)
    val wb =
      read(source).addSheet(Sheet(SheetName.unsafe("OTHER"))).fold(e => fail(e.message), identity)
    val output = Files.createTempFile("gh785-out", ".xlsm")
    val result = XlsxWriter.write(wb, output)
    assert(result.isLeft, s"expected the write to fail, got $result")
    assertEquals(result, Left(com.tjclp.xl.error.XLError.DuplicateSheet("Other")))
    cleanup(source, output)
  }

  test("GH-785: a macro sheet first in <sheets> stays first through edits") {
    val source = fixture(macroSheet, otherFirst = true)
    val wb = read(source)
    assertEquals(wb.sheetNames.map(_.value), Seq("Data", "Data2"))
    assertEquals(wb.metadata.definedNames.find(_.name == "Local").flatMap(_.localSheetId), Some(1))
    val edited = wb.addSheet(Sheet(SheetName.unsafe("New"))).fold(e => fail(e.message), identity)
    val output = write(edited)
    assertEquals(sheetEntries(output).map(_._1), Vector("Other", "Data", "Data2", "New"))
    val written = names(output)
    assert(written.contains(("Local", Some(2), "Data2!$A$1")), written.toString)
    assert(written.contains(("MacroLocal", Some(0), "Other!$A$1")), written.toString)
    cleanup(source, output)
  }

  test(
    "GH-785: with no activeTab, a macro sheet first in <sheets> keeps the tab through a cell edit"
  ) {
    // workbookView/@activeTab defaults to 0, so the macro sheet is the active tab
    val source = fixture(macroSheet, otherFirst = true)
    val output = write(putA2(read(source), "NewText"))
    assert(activeTab(output).forall(_ == 0), s"tab moved off the macro sheet: ${activeTab(output)}")
    cleanup(source, output)
  }

  test("GH-785: a tab on the macro sheet follows it through a reorder") {
    val source = fixture(macroSheet, activeTab = Some(1))
    // the reader makes Data2 (model 1) active; the reorder moves it, and the active sheet, to 0
    val wb = read(source).reorder(Vector(data2Name, dataName)).fold(e => fail(e.message), identity)
    assertEquals(wb.activeSheetIndex, 0)
    val output = write(wb)
    // <sheets> is now Data2, Data, Other: the active sheet is still Data2, so the tab stays on
    // the macro sheet at its new position
    assertEquals(sheetEntries(output).map(_._1), Vector("Data2", "Data", "Other"))
    assertEquals(activeTab(output), Some(2))
    // selecting another worksheet moves the tab to it
    val moved = write(wb.setActiveSheet(1).fold(e => fail(e.message), identity))
    assertEquals(activeTab(moved), Some(1))
    cleanup(source, output, moved)
  }

  // ========== Layout ==========

  test("GH-785: SourceSheetIndex keeps out-of-range indices out of range both ways") {
    val idx = SourceSheetIndex(Vector(true, false, true))
    assertEquals(idx.toModel(0), Some(0))
    assertEquals(idx.toModel(1), None)
    assertEquals(idx.toModel(2), Some(1))
    assertEquals(idx.toModel(5), Some(4))
    val refs = Vector(
      SheetRef(dataName, 1, "rId1"),
      SheetRef(SheetName.unsafe("Other"), 2, "rId2"),
      SheetRef(data2Name, 3, "rId3")
    )
    val rels = Relationships(
      Seq(
        Relationship("rId1", XmlUtil.relTypeWorksheet, "worksheets/sheet1.xml"),
        Relationship("rId2", XmlUtil.relTypeMacrosheet, "macrosheets/sheet1.xml"),
        Relationship("rId3", XmlUtil.relTypeWorksheet, "worksheets/sheet2.xml")
      )
    )
    val layout = SheetLayout.plan(
      refs,
      rels,
      Vector(Some("xl/worksheets/sheet1.xml"), Some("xl/worksheets/sheet2.xml"))
    )
    assertEquals(layout.toOutput(1), 2)
    assertEquals(layout.toOutput(4), 5)
    val dn = DefinedName("Bad", "Data!$A$1", localSheetId = Some(4))
    assertEquals(
      layout.namesToOutput(Vector(dn), Vector.empty).map(_.localSheetId),
      Vector(Some(5))
    )
  }

  // ========== Fixture ==========

  /** `otherFirst`: `<sheets>` is Other, Data, Data2 (MacroLocal then has localSheetId 0). */
  private def fixture(
    kind: Kind,
    activeTab: Option[Int] = None,
    otherFirst: Boolean = false
  ): Path =
    val path = Files.createTempFile("gh785-src", ".xlsm")
    val zip = new ZipOutputStream(Files.newOutputStream(path))
    def put(name: String, content: String): Unit =
      zip.putNextEntry(new ZipEntry(name))
      zip.write(content.getBytes(StandardCharsets.UTF_8))
      zip.closeEntry()
    val data = """<sheet name="Data" sheetId="1" r:id="rId1"/>"""
    val other = """<sheet name="Other" sheetId="2" r:id="rId2"/>"""
    val sheetEntries =
      (if otherFirst then other + data else data + other) +
        """<sheet name="Data2" sheetId="3" r:id="rId3"/>"""
    val otherIdx = if otherFirst then 0 else 1
    val bookViews =
      activeTab.fold("")(t => s"""<bookViews><workbookView activeTab="$t"/></bookViews>""")
    try
      put(
        "[Content_Types].xml",
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/${kind.partPath}" ContentType="${kind.contentType}"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>"""
      )
      put(
        "_rels/.rels",
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="$relNs/officeDocument" Target="xl/workbook.xml"/></Relationships>"""
      )
      put(
        "xl/workbook.xml",
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="$mainNs" xmlns:r="$relNs">$bookViews<sheets>$sheetEntries</sheets><definedNames><definedName name="_xlnm.Auto_Open">Other!$$A$$1</definedName><definedName name="_xlnm.Print_Titles" localSheetId="2">Data2!$$1:$$1</definedName><definedName name="Local" localSheetId="2">Data2!$$A$$1</definedName><definedName name="MacroLocal" localSheetId="$otherIdx">Other!$$A$$1</definedName></definedNames></workbook>"""
      )
      put(
        "xl/_rels/workbook.xml.rels",
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="$relNs/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${kind.relType}" Target="${kind.partPath
            .stripPrefix(
              "xl/"
            )}"/><Relationship Id="rId3" Type="$relNs/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId4" Type="$relNs/sharedStrings" Target="sharedStrings.xml"/></Relationships>"""
      )
      put(
        "xl/worksheets/sheet1.xml",
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="$mainNs"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>5</v></c></row></sheetData></worksheet>"""
      )
      put(
        "xl/worksheets/sheet2.xml",
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="$mainNs"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"""
      )
      put(kind.partPath, kind.xml)
      put(
        "xl/sharedStrings.xml",
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="$mainNs" count="2" uniqueCount="2"><si><t>Alpha</t></si><si><t>MacroText</t></si></sst>"""
      )
    finally zip.close()
    path

  private def read(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"Read failed: $err"), identity)

  private def putA2(wb: Workbook, value: String): Workbook =
    wb.update(dataName, _.put(ref"A2" -> value)).fold(e => fail(e.message), identity)

  private def write(wb: Workbook): Path =
    val output = Files.createTempFile("gh785-out", ".xlsm")
    XlsxWriter.write(wb, output).fold(err => fail(s"Write failed: $err"), identity)
    output

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

  private def text(path: Path, name: String): String =
    new String(entry(path, name), StandardCharsets.UTF_8)

  private def workbookElem(path: Path): Elem =
    XmlSecurity.parseSafe(text(path, "xl/workbook.xml"), "xl/workbook.xml").toOption.get

  /** `<sheets>` entries as (name, sheetId, r:id), in order. */
  private def sheetEntries(path: Path): Vector[(String, String, String)] =
    (workbookElem(path) \ "sheets" \ "sheet").toVector.map { s =>
      (s \@ "name", s \@ "sheetId", s.attribute(relNs, "id").map(_.text).getOrElse(""))
    }

  /** `<definedNames>` as (name, localSheetId, text), in any order. */
  private def names(path: Path): Set[(String, Option[Int], String)] =
    (workbookElem(path) \ "definedNames" \ "definedName").map { d =>
      (d \@ "name", Option(d \@ "localSheetId").filter(_.nonEmpty).map(_.toInt), d.text)
    }.toSet

  private def activeTab(path: Path): Option[Int] =
    (workbookElem(path) \ "bookViews" \ "workbookView").headOption
      .flatMap(v => Option(v \@ "activeTab").filter(_.nonEmpty))
      .map(_.toInt)

  private def cleanup(paths: Path*): Unit = paths.foreach(Files.deleteIfExists)
