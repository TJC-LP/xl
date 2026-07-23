package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import munit.FunSuite

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * GH-375 preservation laws (the CfPreservationSpec contract applied to `<dataValidations>`):
 *
 *   - (a) read → author a validation → write → re-read: every preserved entry exactly once,
 *     authored entry appended, container count correct
 *   - (b) dv-CLEAN cell edit: the emitted dataValidations element equals the source element (the
 *     dirty-gate proof) — the read→edit→write ask of the GH-372 family
 *   - (c) cleared model emits no dataValidations (no resurrection)
 *   - (d) fresh authoring: scratch list dropdowns emit and round-trip typed (incl. the showDropDown
 *     INVERSION and multi-range sqref)
 *   - (e) both writer backends emit identical dataValidations
 */
class DataValidationPreservationSpec extends FunSuite:

  private def writeTo(wb: Workbook, label: String, config: WriterConfig = WriterConfig()): Path =
    val out = Files.createTempFile(s"xl-dv-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter
      .writeWith(wb, out, config)
      .fold(err => fail(s"$label write failed: ${err.message}"), _ => ())
    out

  private def reread(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"re-read failed: ${err.message}"), identity)

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) =>
          new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  private def occurrences(haystack: String, needle: String): Int =
    java.util.regex.Pattern.quote(needle).r.findAllMatchIn(haystack).size

  // ===== fixtures: an Excel-typical validation the typed model does NOT cover (Preserved) =====

  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    out.putNextEntry(new ZipEntry(name))
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()

  /**
   * Minimal single-sheet fixture whose worksheet carries the given raw `<dataValidations>` block
   * (may be empty). The decimal entry carries `imeMode` — a foreign attribute outside the GH-429
   * widened whitelist, so it must ride [[com.tjclp.xl.sheets.DataValidation.Preserved]].
   */
  private def rawDvFixture(dataValidationsXml: String): Path =
    val path = Files.createTempFile("dv-fixture", ".xlsx")
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
    <sheet name="Inputs" sheetId="1" r:id="rId1"/>
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
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
  $dataValidationsXml
</worksheet>"""
      )
    finally out.close()
    path

  private val foreignDecimalDv =
    """<dataValidations count="1">
    <dataValidation type="decimal" operator="between" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Range" error="0 to 1 only" imeMode="hiragana" sqref="D2:D9"><formula1>0</formula1><formula2>1</formula2></dataValidation>
  </dataValidations>"""

  /** The same entry WITHOUT the foreign attr: inside the GH-429 widened typed subset. */
  private val excelTypicalDv =
    """<dataValidations count="1">
    <dataValidation type="decimal" operator="between" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Range" error="0 to 1 only" sqref="D2:D9"><formula1>0</formula1><formula2>1</formula2></dataValidation>
  </dataValidations>"""

  private def readDvFixture(dvXml: String): (Path, Workbook) =
    val path = rawDvFixture(dvXml)
    val wb = XlsxReader.read(path).fold(e => fail(s"fixture read failed: ${e.message}"), identity)
    (path, wb)

  // ===== (d) fresh authoring: the from-scratch replica case =====

  private def freshDvWorkbook: Workbook =
    val sheet = Sheet(SheetName.unsafe("Case"))
      .put(ref"B2", CellValue.Number(BigDecimal(1)))
      .withDataValidation(ref"B2:B10", DataValidation.listOf("1", "2", "3"))
      .withDataValidation(ref"C2:C10", DataValidation.list("$Z$1:$Z$3", allowBlank = false))
    Workbook(Vector(sheet))

  test("(d) fresh workbook: authored list dropdowns emit on a no-source write and round-trip") {
    val wb = freshDvWorkbook
    val out = writeTo(wb, "fresh")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(sheetXml.contains("<dataValidations count=\"2\">"), sheetXml.take(800))
    assert(sheetXml.contains("type=\"list\""), sheetXml.take(800))
    // inline values keep their literal quotes (scala.xml text-escapes them as &quot;);
    // the range ref rides verbatim
    assert(sheetXml.contains("<formula1>&quot;1,2,3&quot;</formula1>"), sheetXml.take(800))
    assert(sheetXml.contains("<formula1>$Z$1:$Z$3</formula1>"), sheetXml.take(800))
    // default flags: allowBlank=true emits, showDropdown=true emits NO attribute (inversion!)
    assert(sheetXml.contains("allowBlank=\"1\""), sheetXml.take(800))
    assert(!sheetXml.contains("showDropDown"), s"dropdown must not be suppressed: $sheetXml")
    assertEquals(reread(out).sheets(0).dataValidations, wb.sheets(0).dataValidations)
  }

  test("(d) the OOXML showDropDown INVERSION: model showDropdown=false emits showDropDown=\"1\"") {
    val sheet = Sheet(SheetName.unsafe("Case"))
      .withDataValidation(ref"B2:B4", DataValidation.list("\"a,b\"", showDropdown = false))
    val out = writeTo(Workbook(Vector(sheet)), "inversion")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(sheetXml.contains("showDropDown=\"1\""), s"suppression attr missing: $sheetXml")
    val back = reread(out).sheets(0).typedDataValidations
    assertEquals(back.map(_.showDropdown), Vector(false))
  }

  test("(d) multi-range sqref emits space-separated and round-trips") {
    val sheet = Sheet(SheetName.unsafe("Case"))
      .withDataValidation(
        Vector(ref"B2:B4": CellRange, ref"D2:D4": CellRange),
        DataValidation.listOf("x", "y")
      )
    val out = writeTo(Workbook(Vector(sheet)), "multirange")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(sheetXml.contains("sqref=\"B2:B4 D2:D4\""), sheetXml.take(800))
    assertEquals(
      reread(out).sheets(0).typedDataValidations.map(_.ranges.map(_.toA1)),
      Vector(Vector("B2:B4", "D2:D4"))
    )
  }

  test("(d) schema position: dataValidations sits after sheetData, before pageMargins/hyperlinks") {
    val sheet = Sheet(SheetName.unsafe("Case"))
      .put(ref"A1", CellValue.Text("v"))
      .withDataValidation(ref"A2:A5", DataValidation.listOf("a"))
      .withPageSetup(
        com.tjclp.xl.sheets.PageSetup(margins = Some(com.tjclp.xl.sheets.PageMargins.default))
      )
    val out = writeTo(Workbook(Vector(sheet)), "order")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    val sheetData = sheetXml.indexOf("<sheetData>")
    val dv = sheetXml.indexOf("<dataValidations")
    val margins = sheetXml.indexOf("<pageMargins")
    assert(sheetData >= 0 && dv >= 0 && margins >= 0, sheetXml)
    assert(sheetData < dv && dv < margins, s"schema order violated: $sheetXml")
  }

  // ===== (a) authored + preserved coexistence =====

  test("(a) author on top of a foreign (Preserved) validation: no duplication, no loss") {
    val (_, wb) = readDvFixture(foreignDecimalDv)
    val before = wb.sheets(0).dataValidations
    assertEquals(before.size, 1)
    assert(
      before(0).isInstanceOf[com.tjclp.xl.sheets.DataValidation.Preserved],
      s"the imeMode foreign attr must force Preserved: $before"
    )

    val updated = wb
      .update(
        wb.sheets(0).name,
        _.withDataValidation(ref"B2:B9", DataValidation.listOf("Low", "Med", "High"))
      )
      .fold(err => fail(s"update failed: $err"), identity)
    val out = writeTo(updated, "author")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assertEquals(occurrences(sheetXml, "<dataValidations"), 1, "exactly ONE container")
    assert(sheetXml.contains("count=\"2\""), s"container count must be 2: $sheetXml")
    assertEquals(occurrences(sheetXml, "type=\"decimal\""), 1, "preserved entry exactly once")
    assertEquals(occurrences(sheetXml, "errorTitle=\"Range\""), 1, "foreign attrs survive")
    assertEquals(occurrences(sheetXml, "<formula1>&quot;Low,Med,High&quot;</formula1>"), 1)

    val back = reread(out).sheets(0).dataValidations
    assertEquals(back.size, 2)
    assertEquals(back.take(1), before, "preserved entry must ride through unchanged")
    assertEquals(back.lastOption, updated.sheets(0).dataValidations.lastOption)
  }

  test("(a) GH-429: the Excel-typical decimal entry now parses TYPED and stays CLEAN-stable") {
    val (in, wb) = readDvFixture(excelTypicalDv)
    wb.sheets(0).dataValidations match
      case Vector(r: com.tjclp.xl.sheets.DataValidation.Rules) =>
        assertEquals(r.ranges.map(_.toA1), Vector("D2:D9"))
        assertEquals(r.messages.errorTitle, Some("Range"))
        assertEquals(r.messages.showInputMessage, true)
      case other => fail(s"decimal+operator+prompts must parse typed post-GH-429: $other")
    // an unrelated cell edit leaves the dv slot CLEAN: source element passes through verbatim
    val edited = wb
      .update(wb.sheets(0).name, _.put(ref"A5", CellValue.Text("unrelated edit")))
      .fold(err => fail(s"update failed: $err"), identity)
    val out = writeTo(edited, "typed-clean")
    def dvElem(p: Path): Option[String] =
      val ws = XmlSecurity
        .parseSafe(entryText(p, "xl/worksheets/sheet1.xml"), "ws")
        .fold(e => fail(e.message), identity)
      (ws \ "dataValidations").headOption.map(_.toString)
    assertEquals(dvElem(out), dvElem(in), "typed-parsing source must still pass through CLEAN")
    assertEquals(reread(out).sheets(0).dataValidations, wb.sheets(0).dataValidations)
  }

  // ===== (b) the dirty-gate proof (the GH-372-family read→edit→write ask) =====

  test("(b) dv-CLEAN cell edit keeps validations on the REWRITTEN sheet (slot passthrough)") {
    val (in, wb) = readDvFixture(foreignDecimalDv)
    val edited = wb
      .update(wb.sheets(0).name, _.put(ref"A5", CellValue.Text("unrelated edit")))
      .fold(err => fail(s"update failed: $err"), identity)
    val out = writeTo(edited, "clean-edit")
    def dvElem(p: Path): Option[String] =
      val ws = XmlSecurity
        .parseSafe(entryText(p, "xl/worksheets/sheet1.xml"), "ws")
        .fold(e => fail(e.message), identity)
      (ws \ "dataValidations").headOption.map(_.toString)
    assert(dvElem(out).isDefined, "validations dropped by the rewrite")
    assertEquals(dvElem(out), dvElem(in), "clean slot must pass the source element through")
    assertEquals(reread(out).sheets(0).dataValidations, wb.sheets(0).dataValidations)
  }

  // ===== (c) deletion honored =====

  test("(c) cleared model emits NO dataValidations (no resurrection)") {
    val (_, wb) = readDvFixture(foreignDecimalDv)
    val cleared = wb
      .update(wb.sheets(0).name, _.copy(dataValidations = Vector.empty))
      .fold(err => fail(s"update failed: $err"), identity)
    val out = writeTo(cleared, "cleared")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(!sheetXml.contains("<dataValidations"), "dataValidations must not resurrect")
    assertEquals(
      reread(out).sheets(0).dataValidations,
      Vector.empty[com.tjclp.xl.sheets.DataValidation]
    )
  }

  test("(c) removing one entry keeps the other (removeDataValidation)") {
    val (_, wb) = readDvFixture(foreignDecimalDv)
    val updated = wb
      .update(
        wb.sheets(0).name,
        _.withDataValidation(ref"B2:B9", DataValidation.listOf("a", "b")).removeDataValidation(0)
      )
      .fold(err => fail(s"update failed: $err"), identity)
    val out = writeTo(updated, "remove-one")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(!sheetXml.contains("type=\"decimal\""), "removed entry must not resurrect")
    assert(sheetXml.contains("count=\"1\""), sheetXml)
    assertEquals(
      reread(out).sheets(0).typedDataValidations.map(_.ranges.map(_.toA1)),
      Vector(Vector("B2:B9"))
    )
  }

  // ===== (e) backend parity =====

  test("(e) backend parity: ScalaXml and SaxStax emit structurally identical dataValidations") {
    val wb = freshDvWorkbook
    val domOut = writeTo(wb, "dom", WriterConfig.scalaXml)
    val saxOut = writeTo(wb, "sax", WriterConfig.saxStax)
    def normalize(n: scala.xml.Node): String =
      def attrsOf(e: scala.xml.Elem): String =
        e.attributes.asAttrMap.toSeq.sorted.map((k, v) => s"$k=$v").mkString(" ")
      n match
        case e: scala.xml.Elem =>
          val kids = e.child.collect { case c: scala.xml.Elem => normalize(c) }.mkString
          val text = e.child.collect { case t: scala.xml.Text => t.data }.mkString.trim
          s"<${e.label} ${attrsOf(e)}>$text$kids</${e.label}>"
        case other => other.text.trim
    def dvSection(p: Path): Seq[String] =
      val ws = XmlSecurity
        .parseSafe(entryText(p, "xl/worksheets/sheet1.xml"), "ws")
        .fold(e => fail(e.message), identity)
      (ws \ "dataValidations").map(normalize)
    assert(dvSection(domOut).nonEmpty, "DOM backend emitted no dataValidations")
    assertEquals(dvSection(saxOut), dvSection(domOut), "backends must emit identical dv")
    assertEquals(reread(saxOut).sheets(0).dataValidations, wb.sheets(0).dataValidations)
    assertEquals(reread(domOut).sheets(0).dataValidations, wb.sheets(0).dataValidations)
  }

  // ===== structural edits keep dropdowns attached through a full write cycle =====

  test("insert-above shifts the emitted sqref (dropdowns stay attached end to end)") {
    val sheet = Sheet(SheetName.unsafe("Case"))
      .withDataValidation(ref"B2:B4", DataValidation.listOf("1", "2"))
      .insertRows(at = 0, count = 2)
    val out = writeTo(Workbook(Vector(sheet)), "shifted")
    assert(entryText(out, "xl/worksheets/sheet1.xml").contains("sqref=\"B4:B6\""))
    assertEquals(
      reread(out).sheets(0).typedDataValidations.map(_.ranges.map(_.toA1)),
      Vector(Vector("B4:B6"))
    )
  }

  // ===== openpyxl smoke: the authored dropdown is visible to a foreign reader =====

  test("openpyxl reads the authored list validation (subprocess smoke)") {
    assume(pythonWithOpenpyxl, "python3 with openpyxl not available")
    val out = writeTo(freshDvWorkbook, "openpyxl")
    val script =
      s"""import openpyxl
         |wb = openpyxl.load_workbook(${pyString(out.toString)})
         |ws = wb["Case"]
         |dvs = ws.data_validations.dataValidation
         |print(len(dvs))
         |print(sorted((dv.formula1 or "") for dv in dvs))
         |""".stripMargin
    val output = runPython(script)
    assert(output.contains("2"), s"openpyxl lost the validations: $output")
    assert(output.contains("\"1,2,3\""), s"openpyxl lost the inline list: $output")
    assert(output.contains("$Z$1:$Z$3"), s"openpyxl lost the range ref: $output")
  }

  private lazy val pythonWithOpenpyxl: Boolean =
    scala.util
      .Try {
        val p = new ProcessBuilder("python3", "-c", "import openpyxl")
          .redirectErrorStream(true)
          .start()
        p.waitFor() == 0
      }
      .getOrElse(false)

  private def pyString(s: String): String =
    "\"" + s.replace("\\", "\\\\").replace("\"", "\\\"") + "\""

  private def runPython(script: String): String =
    val pb = new ProcessBuilder("python3", "-c", script).redirectErrorStream(true)
    pb.environment().put("PYTHONIOENCODING", "utf-8")
    val p = pb.start()
    val output = new String(p.getInputStream.readAllBytes(), StandardCharsets.UTF_8)
    val exit = p.waitFor()
    assertEquals(exit, 0, s"python exited $exit:\n$output")
    output
