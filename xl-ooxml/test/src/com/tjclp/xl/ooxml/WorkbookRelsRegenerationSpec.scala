package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import scala.jdk.CollectionConverters.*

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref

/**
 * GH-320: regenerating workbook.xml must never desync from workbook.xml.rels.
 *
 * LibreOffice numbers its workbook rels theme-first (rId1 = theme, rId3 = first worksheet), so a
 * writer that renumbers sheet rIds to rId1..N while copying the preserved rels verbatim makes the
 * first sheet resolve to xl/theme/theme1.xml — xl's own reader fails ("Missing required child
 * element: sheetData") and openpyxl reads an empty sheet. Excel/openpyxl files survived only
 * because their rels happen to number sheets first.
 *
 * Contract pinned here:
 *   - sheet rIds are REUSED from the source rels (matched by worksheet target), never renumbered
 *   - non-sheet rels (theme, styles, sharedStrings, calcChain, ...) ride through with their ids
 *   - new sheets allocate ids ABOVE every preserved numeric id (no collision)
 *   - every r:id in the regenerated workbook.xml resolves in the regenerated rels, and every
 *     internal rel target resolves to a zip entry
 *   - the metadata-modified (fromDomain) path carries preserved workbook-level elements through:
 *     fileVersion, workbookPr, workbookProtection, bookViews siblings, calcPr, extLst
 */
class WorkbookRelsRegenerationSpec extends FunSuite:

  private val relTypeWorksheetUri =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  private val relTypeThemeUri =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

  private def readFixture(name: String): (Path, Workbook) =
    val path = TestFixtures.copyToTemp(name)
    val wb = XlsxReader
      .read(path)
      .fold(err => fail(s"$name failed to read: ${err.message}"), identity)
    (path, wb)

  private def writeTo(wb: Workbook, label: String): Path =
    val out = Files.createTempFile(s"xl-rels-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter
      .write(wb, out)
      .fold(err => fail(s"$label write failed: ${err.message}"), _ => ())
    out

  private def reread(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"re-read failed: ${err.message}"), identity)

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  private def zipEntryNames(path: Path): Set[String] =
    val zip = new ZipFile(path.toFile)
    try zip.entries().asScala.map(_.getName).toSet
    finally zip.close()

  private def parseEntry(path: Path, name: String): scala.xml.Elem =
    XmlSecurity.parseSafe(entryText(path, name), name).fold(e => fail(e.message), identity)

  /** name -> r:id for every `<sheet>` in workbook.xml. */
  private def sheetRelIds(path: Path): Vector[(String, String)] =
    val wb = parseEntry(path, "xl/workbook.xml")
    (wb \ "sheets" \ "sheet").collect { case e: scala.xml.Elem =>
      val rId = e
        .attribute(XmlUtil.nsRelationships, "id")
        .map(_.text)
        .getOrElse(fail(s"sheet missing r:id: $e"))
      (e \@ "name") -> rId
    }.toVector

  /** Id -> (Type, Target, TargetMode) for every relationship in workbook.xml.rels. */
  private def workbookRels(path: Path): Map[String, (String, String, Option[String])] =
    val ids = relIdList(path)
    assertEquals(ids.distinct, ids, s"duplicate relationship ids in $path: $ids")
    val rels = parseEntry(path, "xl/_rels/workbook.xml.rels")
    (rels \ "Relationship").collect { case e: scala.xml.Elem =>
      (e \@ "Id") -> ((e \@ "Type", e \@ "Target", Option(e \@ "TargetMode").filter(_.nonEmpty)))
    }.toMap

  /** Every relationship Id in declaration order (duplicates visible, unlike the Map view). */
  private def relIdList(path: Path): Vector[String] =
    val rels = parseEntry(path, "xl/_rels/workbook.xml.rels")
    (rels \ "Relationship").collect { case e: scala.xml.Elem => e \@ "Id" }.toVector

  /** Resolve a workbook-rels target against the xl/ base (leading slash = package-absolute). */
  private def resolveTarget(target: String): String =
    val cleaned = if target.startsWith("/") then target.drop(1) else target
    val resolved =
      if cleaned.startsWith("xl/") then java.nio.file.Paths.get(cleaned)
      else java.nio.file.Paths.get("xl").resolve(cleaned)
    resolved.normalize().toString.replace('\\', '/')

  private def edit(wb: Workbook, value: CellValue): Workbook =
    wb.update(wb.sheets(0).name, _.put(ref"Z99", value))
      .fold(err => fail(s"edit failed: $err"), identity)

  // ===== (i) the wave-7 blocked test: full re-read of an edited LibreOffice file =====

  test("GH-320: one-cell edit on small-values-lo.xlsx re-reads fully with values intact") {
    val (_, wb) = readFixture("small-values-lo.xlsx")
    val edited = edit(wb, CellValue.Text("edited"))
    val out = writeTo(edited, "lo-edit")
    val back = reread(out)
    assertEquals(back.sheets(0).name.value, "Values")
    assertEquals(back.sheets(0)(ref"A2").value, CellValue.Text("héllo wörld"))
    assertEquals(back.sheets(0)(ref"A5").value, CellValue.Text("  leading and trailing  "))
    assertEquals(back.sheets(0)(ref"Z99").value, CellValue.Text("edited"))
  }

  test("GH-320: edited LO workbook keeps the source sheet rId and the theme rel") {
    val (_, wb) = readFixture("small-values-lo.xlsx")
    val out = writeTo(edit(wb, CellValue.Number(BigDecimal(7))), "lo-ids")
    // LibreOffice numbered the sheet rId3 (rId1 = theme); the source ids must ride through
    assertEquals(sheetRelIds(out), Vector("Values" -> "rId3"))
    val rels = workbookRels(out)
    val (sheetType, sheetTarget, _) = rels.getOrElse("rId3", fail(s"rId3 missing: $rels"))
    assertEquals(sheetType, relTypeWorksheetUri)
    assertEquals(resolveTarget(sheetTarget), "xl/worksheets/sheet1.xml")
    val (themeType, themeTarget, _) = rels.getOrElse("rId1", fail(s"rId1 missing: $rels"))
    assertEquals(themeType, relTypeThemeUri)
    assertEquals(resolveTarget(themeTarget), "xl/theme/theme1.xml")
  }

  test("GH-320: openpyxl reads the edited LibreOffice file (subprocess smoke)") {
    assume(pythonWithOpenpyxl, "python3 with openpyxl not available")
    val (_, wb) = readFixture("small-values-lo.xlsx")
    val out = writeTo(edit(wb, CellValue.Text("edited")), "lo-openpyxl")
    val script =
      s"""import openpyxl
         |wb = openpyxl.load_workbook(${pyString(out.toString)})
         |ws = wb["Values"]
         |print(ws["A2"].value)
         |print(ws["Z99"].value)
         |""".stripMargin
    val output = runPython(script)
    assert(output.contains("héllo wörld"), s"openpyxl lost A2 (empty sheet?): $output")
    assert(output.contains("edited"), s"openpyxl lost the edit: $output")
  }

  test("GH-320: openpyxl reads the setActiveSheet output (metadata-modified path smoke)") {
    assume(pythonWithOpenpyxl, "python3 with openpyxl not available")
    val (_, wb) = readFixture("formulas-lo.xlsx")
    val activated = wb.setActiveSheet(1).fold(err => fail(s"setActiveSheet failed: $err"), identity)
    val out = writeTo(activated, "lo-active-openpyxl")
    val script =
      s"""import openpyxl
         |wb = openpyxl.load_workbook(${pyString(out.toString)})
         |print(wb.sheetnames)
         |print(wb["Data"]["A5"].value)
         |""".stripMargin
    val output = runPython(script)
    assert(output.contains("['Data', 'Calc']"), s"openpyxl lost the sheets: $output")
    assert(output.contains("50"), s"openpyxl lost Data!A5: $output")
  }

  // ===== the resolution law, every fixture =====

  test("GH-320 LAW: a cell edit keeps every workbook.xml r:id resolving to its part") {
    TestFixtures.all.foreach { fixture =>
      val (_, wb) = readFixture(fixture)
      val out = writeTo(edit(wb, CellValue.Number(BigDecimal(1))), "law")
      val rels = workbookRels(out)
      val entries = zipEntryNames(out)
      // every <sheet r:id> resolves to a worksheet-type rel whose target part exists
      sheetRelIds(out).foreach { case (name, rId) =>
        val (relType, target, _) =
          rels.getOrElse(rId, fail(s"$fixture: sheet '$name' r:id=$rId missing from rels: $rels"))
        assertEquals(
          relType,
          relTypeWorksheetUri,
          s"$fixture: sheet '$name' r:id=$rId resolves to a non-worksheet part ($target)"
        )
        assert(
          entries.contains(resolveTarget(target)),
          s"$fixture: sheet '$name' target $target missing from zip"
        )
      }
      // no duplicate ids, and every internal rel target resolves to a zip entry
      rels.foreach { case (id, (_, target, targetMode)) =>
        if !targetMode.contains("External") then
          assert(
            entries.contains(resolveTarget(target)),
            s"$fixture: rel $id target $target missing from zip"
          )
      }
    }
  }

  // ===== (iii) metadata-modified path: preserved workbook-level elements survive =====

  test("GH-320: setActiveSheet on formulas-lo.xlsx keeps calcPr/fileVersion/protection/extLst") {
    val (_, wb) = readFixture("formulas-lo.xlsx")
    val activated = wb.setActiveSheet(1).fold(err => fail(s"setActiveSheet failed: $err"), identity)
    val out = writeTo(activated, "lo-active")

    val workbookElem = parseEntry(out, "xl/workbook.xml")
    val calcPr = (workbookElem \ "calcPr").headOption.getOrElse(fail("calcPr dropped"))
    assertEquals(calcPr \@ "iterateCount", "100", "calcPr attributes dropped")
    val fileVersion = (workbookElem \ "fileVersion").headOption.getOrElse(fail("fileVersion lost"))
    assertEquals(fileVersion \@ "appName", "Calc")
    assert((workbookElem \ "workbookProtection").nonEmpty, "workbookProtection dropped")
    assert(
      (workbookElem \ "extLst").exists(_.toString.contains("extCalcPr")),
      "loext extLst dropped"
    )
    val workbookPr = (workbookElem \ "workbookPr").headOption.getOrElse(fail("workbookPr lost"))
    assertEquals(workbookPr \@ "backupFile", "false", "unmodeled workbookPr attrs dropped")
    // bookViews: the model's activeTab overlays while sibling attributes ride through (GH-294)
    val view = (workbookElem \ "bookViews" \ "workbookView").headOption
      .getOrElse(fail("workbookView lost"))
    assertEquals(view \@ "activeTab", "1")
    assertEquals(view \@ "showHorizontalScroll", "true", "unmodeled bookViews attrs dropped")
    // schema order: fileVersion < workbookPr < bookViews < sheets < calcPr
    val xml = entryText(out, "xl/workbook.xml")
    val order = List("<fileVersion", "<workbookPr", "<bookViews", "<sheets", "<calcPr")
      .map(tag => tag -> xml.indexOf(tag))
    order.foreach((tag, idx) => assert(idx >= 0, s"$tag missing: $xml"))
    assertEquals(order.map(_._2).sorted, order.map(_._2), s"schema order violated: $order")

    // sheet rIds preserved, theme rel intact, everything resolves
    assertEquals(sheetRelIds(out), Vector("Data" -> "rId3", "Calc" -> "rId4"))
    val rels = workbookRels(out)
    assertEquals(rels.get("rId1").map(_._1), Some(relTypeThemeUri))

    val back = reread(out)
    assertEquals(back.activeSheetIndex, 1)
    assertEquals(
      back.sheets(1)(ref"A1").value,
      CellValue.Formula("SUM(Data!A1:A5)", Some(CellValue.Number(BigDecimal(150))))
    )
  }

  test("GH-320: authored defined name on small-values-lo.xlsx keeps preserved elements") {
    val (_, wb) = readFixture("small-values-lo.xlsx")
    val named = wb.withDefinedName("MyRange", "Values!$A$1:$A$3")
    val out = writeTo(named, "lo-name")
    val workbookElem = parseEntry(out, "xl/workbook.xml")
    assert((workbookElem \ "calcPr").nonEmpty, "calcPr dropped on metadata write")
    assert((workbookElem \ "fileVersion").nonEmpty, "fileVersion dropped on metadata write")
    val dn = (workbookElem \ "definedNames" \ "definedName").headOption
      .getOrElse(fail("authored definedName missing"))
    assertEquals(dn \@ "name", "MyRange")
    assertEquals(sheetRelIds(out), Vector("Values" -> "rId3"))
    val back = reread(out)
    assertEquals(back.metadata.definedNames.map(_.name), Vector("MyRange"))
    assertEquals(back.sheets(0)(ref"A2").value, CellValue.Text("héllo wörld"))
  }

  // ===== new sheets allocate fresh non-colliding ids =====

  test("GH-320: adding a sheet to small-values-lo.xlsx allocates a fresh rId above the max") {
    val (_, wb) = readFixture("small-values-lo.xlsx")
    val extra = Sheet(SheetName.unsafe("Extra")).put(ref"A1", CellValue.Number(BigDecimal(42)))
    val added = wb.insertAt(1, extra).fold(err => fail(s"insertAt failed: $err"), identity)
    val out = writeTo(added, "lo-add")
    val ids = sheetRelIds(out)
    assertEquals(ids.map(_._1), Vector("Values", "Extra"))
    assertEquals(ids.headOption.map(_._2), Some("rId3"), "surviving sheet must keep its source id")
    // LO's rels claim rId1..rId4 — the fresh sheet must allocate above them
    assertEquals(ids.lift(1).map(_._2), Some("rId5"), "fresh sheet id must not collide")
    val rels = workbookRels(out)
    assertEquals(rels.get("rId1").map(_._1), Some(relTypeThemeUri), "theme rel dropped")
    val back = reread(out)
    assertEquals(back.sheetNames.map(_.value), Vector("Values", "Extra"))
    assertEquals(back.sheets(1)(ref"A1").value, CellValue.Number(BigDecimal(42)))
    assertEquals(back.sheets(0)(ref"A2").value, CellValue.Text("héllo wörld"))
  }

  test("GH-320: deleting a sheet from formulas-lo.xlsx drops its rel and keeps the rest") {
    val (_, wb) = readFixture("formulas-lo.xlsx")
    val without = wb.removeAt(1).fold(err => fail(s"removeAt failed: $err"), identity)
    val out = writeTo(without, "lo-delete")
    // the surviving sheet keeps its source id; the deleted sheet's rel does not dangle
    assertEquals(sheetRelIds(out), Vector("Data" -> "rId3"))
    val rels = workbookRels(out)
    val worksheetTargets = rels.values.collect {
      case (relType, target, _) if relType == relTypeWorksheetUri => resolveTarget(target)
    }
    assertEquals(worksheetTargets.toSet, Set("xl/worksheets/sheet1.xml"))
    assertEquals(rels.get("rId1").map(_._1), Some(relTypeThemeUri), "theme rel dropped")
    val entries = zipEntryNames(out)
    rels.foreach { case (id, (_, target, targetMode)) =>
      if !targetMode.contains("External") then
        assert(entries.contains(resolveTarget(target)), s"rel $id dangling: $target")
    }
    val back = reread(out)
    assertEquals(back.sheetNames.map(_.value), Vector("Data"))
    assertEquals(back.sheets(0)(ref"A5").value, CellValue.Number(BigDecimal(50)))
  }

  // ===== stability: the regenerated output is itself a valid surgical source =====

  test("GH-320: a second edit→write cycle on the LO file is stable (fixpoint)") {
    val (_, wb) = readFixture("small-values-lo.xlsx")
    val out1 = writeTo(edit(wb, CellValue.Number(BigDecimal(1))), "lo-cycle1")
    val firstIds = sheetRelIds(out1)
    val wb2 = reread(out1)
    val edited2 = wb2
      .update(wb2.sheets(0).name, _.put(ref"Y98", CellValue.Number(BigDecimal(2))))
      .fold(err => fail(s"second edit failed: $err"), identity)
    val out2 = writeTo(edited2, "lo-cycle2")
    assertEquals(sheetRelIds(out2), firstIds, "sheet rIds must be stable across cycles")
    val back = reread(out2)
    assertEquals(back.sheets(0)(ref"Z99").value, CellValue.Number(BigDecimal(1)))
    assertEquals(back.sheets(0)(ref"Y98").value, CellValue.Number(BigDecimal(2)))
    assertEquals(back.sheets(0)(ref"A2").value, CellValue.Text("héllo wörld"))
  }

  // ===== (iv) Excel-dialect guard: well-ordered rels ride through unchanged =====

  test("GH-320: numeric edit on small-values.xlsx keeps workbook rels semantically identical") {
    val (in, wb) = readFixture("small-values.xlsx")
    val out = writeTo(edit(wb, CellValue.Number(BigDecimal(7))), "openpyxl-edit")
    assertEquals(sheetRelIds(out), Vector("Values" -> "rId1"))
    assertEquals(workbookRels(out), workbookRels(in), "openpyxl rels must ride through unchanged")
    val back = reread(out)
    assertEquals(back.sheets(0)(ref"Z99").value, CellValue.Number(BigDecimal(7)))
  }

  // ===== helpers: openpyxl subprocess (the wave-6b smoke precedent) =====

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
