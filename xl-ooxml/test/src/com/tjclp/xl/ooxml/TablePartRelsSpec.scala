package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.jdk.CollectionConverters.*
import scala.xml.XML

import munit.FunSuite

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.lint.{Finding, LintCategory, WorkbookLint}
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import com.tjclp.xl.tables.TableSpec

/**
 * GH-557: a regenerated worksheet's `<tablePart r:id>` (and a generated `<legacyDrawing r:id>`)
 * must name a Relationship the EMITTED sheet `.rels` part actually carries. The writer used to
 * number table parts positionally (rId1.., or rId3.. behind a comment pair) while copying foreign
 * (openpyxl / Excel) rels verbatim, so the reference resolved to nothing or to the comments rel and
 * Excel repaired the file. The rels are now planned before the worksheet is emitted: a preserved
 * table rel whose target is the emitted part keeps its id (verbatim rels, byte-identical), an
 * unmatched one is dropped and a fresh id allocated past the highest numeric id.
 */
class TablePartRelsSpec extends FunSuite:

  private val relTypePrinterSettings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings"
  private val printerSettingsPart = "xl/printerSettings/printerSettings1.bin"
  private val printerSettingsDefault =
    """<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>"""

  private def tempXlsx(label: String): Path =
    val p = Files.createTempFile(s"xl-557-$label-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  private def write(wb: Workbook, label: String, config: WriterConfig = WriterConfig()): Path =
    val out = tempXlsx(label)
    XlsxWriter
      .writeWith(wb, out, config)
      .fold(err => fail(s"$label write failed: ${err.message}"), _ => ())
    out

  private def read(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"read failed: ${err.message}"), identity)

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  private def entryNames(path: Path): List[String] =
    val zip = new ZipFile(path.toFile)
    try zip.entries().asScala.map(_.getName).toList
    finally zip.close()

  /**
   * Copy `source` to a new zip, rewriting each text entry through `patch(name, text)` and adding
   * `extra` entries (a rels part the fresh write never emitted, a printerSettings blob).
   */
  private def surgery(source: Path, label: String, extra: Map[String, Array[Byte]] = Map.empty)(
    patch: (String, String) => String
  ): Path =
    val out = tempXlsx(label)
    val zip = new ZipFile(source.toFile)
    val zos = new ZipOutputStream(Files.newOutputStream(out))
    try
      zip.entries().asScala.foreach { entry =>
        val bytes = zip.getInputStream(entry).readAllBytes()
        val name = entry.getName
        val content =
          if name.endsWith(".xml") || name.endsWith(".rels") then
            patch(name, new String(bytes, StandardCharsets.UTF_8)).getBytes(StandardCharsets.UTF_8)
          else bytes
        zos.putNextEntry(new ZipEntry(name))
        zos.write(content)
        zos.closeEntry()
      }
      extra.foreach { case (name, bytes) =>
        zos.putNextEntry(new ZipEntry(name))
        zos.write(bytes)
        zos.closeEntry()
      }
    finally
      zos.close()
      zip.close()
    out

  private def lintOf(path: Path): Vector[Finding] =
    WorkbookLint.lint(path).fold(err => fail(s"lint must not error: ${err.message}"), identity)

  private val relFindings =
    Set(LintCategory.UnresolvedRelId, LintCategory.WrongRelType, LintCategory.MissingPart)

  private def assertRelsResolve(path: Path): Unit =
    val bad = lintOf(path).filter(f => relFindings.contains(f.category))
    assert(bad.isEmpty, bad.mkString("\n"))

  private def rels(path: Path, relsEntry: String): Relationships =
    Relationships
      .fromXml(XML.loadString(entryText(path, relsEntry)))
      .fold(err => fail(s"rels parse failed: $err"), identity)

  private def tablePartIds(sheetXml: String): Vector[String] =
    (XML.loadString(sheetXml) \\ "tablePart")
      .flatMap(_.attribute(XmlUtil.nsRelationships, "id"))
      .map(_.text)
      .toVector

  private def legacyDrawingId(sheetXml: String): Option[String] =
    (XML.loadString(sheetXml) \\ "legacyDrawing")
      .flatMap(_.attribute(XmlUtil.nsRelationships, "id"))
      .map(_.text)
      .headOption

  private def tableSheet(base: Sheet, tableName: String): Sheet =
    val table = TableSpec
      .fromColumnNames(tableName, tableName, ref"A1:B6", Vector("Name", "Val"))
      .fold(e => fail(s"table: $e"), identity)
    (1 to 5)
      .foldLeft(base.put(ref"A1" -> "Name").put(ref"B1" -> "Val")) { (s, i) =>
        s.put(ARef.from0(0, i), CellValue.Text(s"n$i")).put(ARef.from0(1, i), CellValue.Number(i))
      }
      .withTable(table)

  /** The issue's openpyxl fixture: a table, a legacy comment and a formula on one sheet. */
  private def openpyxlNumbered(label: String): Path =
    val sheet = tableSheet(Sheet("Data"), "T1")
      .put(ref"D1" -> "note")
      .comment(ref"D1", Comment.plainText("c", Some("tester")))
      .put(ref"E1", CellValue.Formula("SUM(B2:B6)", None))
    val fresh = write(Workbook(Vector(sheet)), s"$label-fresh")
    // xl's own fresh numbering is rId1 comments / rId2 vml / rId3 table; openpyxl writes
    // table=rId1, comments="comments", vmlDrawing="anysvml"
    assertEquals(tablePartIds(entryText(fresh, "xl/worksheets/sheet1.xml")), Vector("rId3"))
    assertEquals(legacyDrawingId(entryText(fresh, "xl/worksheets/sheet1.xml")), Some("rId2"))
    surgery(fresh, label) {
      case ("xl/worksheets/sheet1.xml", xml) =>
        xml
          .replace("""r:id="rId2"""", """r:id="anysvml"""")
          .replace("""r:id="rId3"""", """r:id="rId1"""")
      case ("xl/worksheets/_rels/sheet1.xml.rels", xml) =>
        xml
          .replace("""Id="rId1"""", """Id="comments"""")
          .replace("""Id="rId2"""", """Id="anysvml"""")
          .replace("""Id="rId3"""", """Id="rId1"""")
      case (_, xml) => xml
    }

  test("GH-557: openpyxl-numbered rels + a cell put keeps <tablePart r:id> resolving") {
    val src = openpyxlNumbered("openpyxl")
    val srcRels = entryText(src, "xl/worksheets/_rels/sheet1.xml.rels")
    assert(srcRels.contains("""Id="rId1""""), srcRels)
    assertRelsResolve(src)
    val wb = read(src)
    val out = write(wb.put(wb.sheets(0).put(ref"B2" -> 99)), "openpyxl-out")
    assertRelsResolve(out)
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assertEquals(tablePartIds(sheetXml), Vector("rId1"))
    val outRels = rels(out, "xl/worksheets/_rels/sheet1.xml.rels")
    assertEquals(
      outRels.findById("rId1").map(r => (r.`type`, r.target)),
      Some((XmlUtil.relTypeTable, "../tables/table1.xml"))
    )
    // the comment pair keeps its foreign ids and the sheet keeps pointing at them
    assertEquals(legacyDrawingId(sheetXml), Some("anysvml"))
    // nothing needed re-numbering, so the preserved rels part rides verbatim
    assertEquals(entryText(out, "xl/worksheets/_rels/sheet1.xml.rels"), srcRels)
    val reread = read(out)
    assertEquals(reread.sheets(0).tables.keySet, Set("T1"))
    assertEquals(reread.sheets(0)(ref"B2").value, CellValue.Number(99))
    assertEquals(reread.sheets(0).comments.size, 1)
  }

  test("GH-557: recalc-style rewrite of every sheet keeps the foreign ids too") {
    val src = openpyxlNumbered("openpyxl-all")
    val wb = read(src)
    val all = wb.copy(sourceContext =
      wb.sourceContext.map(ctx => wb.sheets.indices.foldLeft(ctx)((c, i) => c.markSheetModified(i)))
    )
    val out = write(all, "openpyxl-all-out")
    assertRelsResolve(out)
    assertEquals(tablePartIds(entryText(out, "xl/worksheets/sheet1.xml")), Vector("rId1"))
  }

  test("GH-557: preserved rels with printerSettings at rId1 and the table at rId2 (field case)") {
    val fresh = write(Workbook(Vector(tableSheet(Sheet("Budget"), "Budget1"))), "printer-fresh")
    assertEquals(tablePartIds(entryText(fresh, "xl/worksheets/sheet1.xml")), Vector("rId1"))
    val printerRel =
      s"""<Relationship Id="rId1" Type="$relTypePrinterSettings" Target="../printerSettings/printerSettings1.bin"/>"""
    val src =
      surgery(fresh, "printer", Map(printerSettingsPart -> Array[Byte](0, 1, 2, 3))) {
        case ("xl/worksheets/sheet1.xml", xml) =>
          xml.replace("""r:id="rId1"""", """r:id="rId2"""")
        case ("xl/worksheets/_rels/sheet1.xml.rels", xml) =>
          xml
            .replace("""Id="rId1"""", """Id="rId2"""")
            .replace("<Relationship ", printerRel + "<Relationship ")
        case ("[Content_Types].xml", xml) =>
          xml.replace("<Override ", printerSettingsDefault + "<Override ")
        case (_, xml) => xml
      }
    val srcRels = entryText(src, "xl/worksheets/_rels/sheet1.xml.rels")
    assertRelsResolve(src)
    val wb = read(src)
    val out = write(wb.put(wb.sheets(0).put(ref"B2" -> 99)), "printer-out")
    assertRelsResolve(out)
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    // the table keeps rId2: rId1 is the printerSettings rel, which the old positional scheme hit
    assertEquals(tablePartIds(sheetXml), Vector("rId2"))
    assertEquals(entryText(out, "xl/worksheets/_rels/sheet1.xml.rels"), srcRels)
    assert(entryNames(out).contains(printerSettingsPart), entryNames(out).toString)
    assertEquals(read(out).sheets(0).tables.keySet, Set("Budget1"))
  }

  /**
   * Two sheets whose table parts are numbered OPPOSITE to sheet order — Excel numbers table parts
   * in creation order, so a Sheet2 table created first is `table1.xml`: Alpha (sheet1) holds T_A in
   * `xl/tables/table2.xml`, Beta (sheet2) holds T_B in `xl/tables/table1.xml`.
   */
  private def swappedTwoSheet(label: String): Path =
    val wb0 = Workbook(Vector(tableSheet(Sheet("Alpha"), "T_A"), tableSheet(Sheet("Beta"), "T_B")))
    val fresh = write(wb0, s"$label-fresh")
    // xl numbers a fresh book's table parts by sheet order: Alpha → table1.xml, Beta → table2.xml
    val t1 = entryText(fresh, "xl/tables/table1.xml")
    val t2 = entryText(fresh, "xl/tables/table2.xml")
    assert(t1.contains("""name="T_A""""), t1)
    assert(t2.contains("""name="T_B""""), t2)
    // swap the parts; the table's own id follows its file name (the tableColumn ids stay)
    def withTableId(xml: String, id: Int): String =
      """(<table\s[^>]*?\sid=")\d+(")""".r.replaceAllIn(xml, m => s"${m.group(1)}$id${m.group(2)}")
    val src = surgery(fresh, label) {
      case ("xl/tables/table1.xml", _) => withTableId(t2, 1)
      case ("xl/tables/table2.xml", _) => withTableId(t1, 2)
      case ("xl/worksheets/_rels/sheet1.xml.rels", xml) =>
        xml.replace("../tables/table1.xml", "../tables/table2.xml")
      case ("xl/worksheets/_rels/sheet2.xml.rels", xml) =>
        xml.replace("../tables/table2.xml", "../tables/table1.xml")
      case (_, xml) => xml
    }
    assertRelsResolve(src)
    assert(entryText(src, "xl/tables/table2.xml").contains("""name="T_A""""))
    assert(entryText(src, "xl/tables/table1.xml").contains("""name="T_B""""))
    src

  /** The package-level findings a mis-numbered table part produces. */
  private val partFindings =
    relFindings + LintCategory.UnreferencedPart

  private def assertPartsClean(path: Path): Unit =
    val bad = lintOf(path).filter(f => partFindings.contains(f.category))
    assert(bad.isEmpty, bad.mkString("\n"))

  /**
   * Each sheet's `<tablePart r:id>` resolves, through the EMITTED rels, to a part naming ITS table.
   */
  private def assertEachSheetOwnsItsTable(out: Path, expected: Seq[(Int, String)]): Unit =
    expected.foreach { case (n, tableName) =>
      val sheetXml = entryText(out, s"xl/worksheets/sheet$n.xml")
      val sheetRels = rels(out, s"xl/worksheets/_rels/sheet$n.xml.rels")
      val tableRels = sheetRels.findAllByType(XmlUtil.relTypeTable)
      assertEquals(tableRels.size, 1, s"sheet$n must keep exactly one table rel: $sheetRels")
      val ids = tablePartIds(sheetXml)
      assertEquals(ids.size, 1, sheetXml)
      val rel = sheetRels.findById(ids(0)).getOrElse(fail(s"sheet$n: ${ids(0)} unresolved"))
      assertEquals(rel.`type`, XmlUtil.relTypeTable)
      val target = "xl/" + rel.target.stripPrefix("../")
      val tableXml = entryText(out, target)
      assert(
        tableXml.contains(s"""name="$tableName""""),
        s"sheet$n's tablePart must name ITS table ($tableName): $target holds $tableXml"
      )
    }

  /**
   * The swapped book with ONE sheet edited: the sibling rides verbatim — its rels still name the
   * SOURCE part number — so the write must keep every source table in its source part (byte
   * stability) and number only new tables past the max. Renumbering by sheet order pointed both
   * sheets at one part and orphaned the other (`unreferenced-part`, Excel repair).
   */
  private def singleSheetEditKeepsBothTables(editIdx: Int, backend: XmlBackend): Unit =
    val label = s"swap-one-$editIdx-$backend"
    val src = swappedTwoSheet(label)
    val srcSiblingRels =
      entryText(src, s"xl/worksheets/_rels/sheet${2 - editIdx}.xml.rels")
    val wb = read(src)
    val edited = wb.put(wb.sheets(editIdx).put(ref"B2" -> 99))
    val out = write(edited, s"$label-out", WriterConfig(backend = backend))
    assertPartsClean(out)
    assertEachSheetOwnsItsTable(out, Vector((1, "T_A"), (2, "T_B")))
    // the parts keep their source numbers: T_A stays in table2.xml, T_B in table1.xml
    assert(entryText(out, "xl/tables/table2.xml").contains("""name="T_A""""))
    assert(entryText(out, "xl/tables/table1.xml").contains("""name="T_B""""))
    // the untouched sibling's rels ride verbatim
    assertEquals(
      entryText(out, s"xl/worksheets/_rels/sheet${2 - editIdx}.xml.rels"),
      srcSiblingRels
    )
    val reread = read(out)
    assertEquals(reread.sheets(0).tables.keySet, Set("T_A"))
    assertEquals(reread.sheets(1).tables.keySet, Set("T_B"))
    assertEquals(reread.sheets(editIdx)(ref"B2").value, CellValue.Number(99))

  test("GH-557: swapped numbering, ONE sheet edited — each sheet keeps its own table (ScalaXml)") {
    singleSheetEditKeepsBothTables(editIdx = 0, XmlBackend.ScalaXml)
    singleSheetEditKeepsBothTables(editIdx = 1, XmlBackend.ScalaXml)
  }

  test("GH-557: swapped numbering, ONE sheet edited — each sheet keeps its own table (SaxStax)") {
    singleSheetEditKeepsBothTables(editIdx = 0, XmlBackend.SaxStax)
    singleSheetEditKeepsBothTables(editIdx = 1, XmlBackend.SaxStax)
  }

  test(
    "GH-557: a NEW table on a sheet whose kept rels hold printerSettings gets a fresh id past the max"
  ) {
    val fresh = write(Workbook(Vector(tableSheet(Sheet("Budget"), "Budget1"))), "printer2-fresh")
    val printerRel =
      s"""<Relationship Id="rId1" Type="$relTypePrinterSettings" Target="../printerSettings/printerSettings1.bin"/>"""
    val src =
      surgery(fresh, "printer2", Map(printerSettingsPart -> Array[Byte](0, 1, 2, 3))) {
        case ("xl/worksheets/sheet1.xml", xml) =>
          xml.replace("""r:id="rId1"""", """r:id="rId2"""")
        case ("xl/worksheets/_rels/sheet1.xml.rels", xml) =>
          xml
            .replace("""Id="rId1"""", """Id="rId2"""")
            .replace("<Relationship ", printerRel + "<Relationship ")
        case ("[Content_Types].xml", xml) =>
          xml.replace("<Override ", printerSettingsDefault + "<Override ")
        case (_, xml) => xml
      }
    assertRelsResolve(src)
    val wb = read(src)
    val second = TableSpec
      .fromColumnNames("T2", "T2", ref"D1:E6", Vector("K", "V"))
      .fold(e => fail(s"table: $e"), identity)
    val withSecond = (1 to 5)
      .foldLeft(wb.sheets(0).put(ref"D1" -> "K").put(ref"E1" -> "V")) { (s, i) =>
        s.put(ARef.from0(3, i), CellValue.Text(s"k$i")).put(ARef.from0(4, i), CellValue.Number(i))
      }
      .withTable(second)
    val out = write(wb.put(withSecond), "printer2-out")
    assertPartsClean(out)
    val sheetRels = rels(out, "xl/worksheets/_rels/sheet1.xml.rels")
    val ids = sheetRels.relationships.map(_.id)
    assertEquals(ids.distinct.size, ids.size, s"duplicate rel ids: $sheetRels")
    // the printerSettings rel keeps rId1; the source table keeps rId2; the new table is allocated
    // PAST the highest numeric id (rId3), never a positional rId1 that would shadow the printer rel
    assertEquals(
      sheetRels.findById("rId1").map(_.`type`),
      Some(relTypePrinterSettings),
      sheetRels.toString
    )
    assertEquals(tablePartIds(entryText(out, "xl/worksheets/sheet1.xml")), Vector("rId2", "rId3"))
    assertEquals(
      sheetRels.findById("rId2").map(_.target),
      Some("../tables/table1.xml"),
      sheetRels.toString
    )
    assertEquals(
      sheetRels.findById("rId3").map(_.target),
      Some("../tables/table2.xml"),
      sheetRels.toString
    )
    assert(entryText(out, "xl/tables/table1.xml").contains("""name="Budget1""""))
    assert(entryText(out, "xl/tables/table2.xml").contains("""name="T2""""))
    assertEquals(read(out).sheets(0).tables.keySet, Set("Budget1", "T2"))
  }

  test("GH-557: two sheets whose source table parts are numbered opposite to sheet order") {
    val src = swappedTwoSheet("swap")
    val wb = read(src)
    assertEquals(wb.sheets(0).tables.keySet, Set("T_A"))
    assertEquals(wb.sheets(1).tables.keySet, Set("T_B"))
    val edited = wb
      .put(wb.sheets(0).put(ref"B2" -> 99))
      .put(wb.sheets(1).put(ref"B2" -> 99))
    val out = write(edited, "swap-out")
    assertPartsClean(out)
    assertEachSheetOwnsItsTable(out, Vector((1, "T_A"), (2, "T_B")))
    val reread = read(out)
    assertEquals(reread.sheets(0).tables.keySet, Set("T_A"))
    assertEquals(reread.sheets(1).tables.keySet, Set("T_B"))
  }

  test("GH-557 adjacent: a sheet with preserved rels (printerSettings) gaining its FIRST comment") {
    val fresh = write(Workbook(Vector(Sheet("Plain").put(ref"A1" -> 1))), "first-comment-fresh")
    assert(!entryNames(fresh).contains("xl/worksheets/_rels/sheet1.xml.rels"))
    val relsXml =
      s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="$relTypePrinterSettings" Target="../printerSettings/printerSettings1.bin"/></Relationships>"""
    val src = surgery(
      fresh,
      "first-comment",
      Map(
        "xl/worksheets/_rels/sheet1.xml.rels" -> relsXml.getBytes(StandardCharsets.UTF_8),
        printerSettingsPart -> Array[Byte](0, 1, 2, 3)
      )
    ) {
      case ("[Content_Types].xml", xml) =>
        xml.replace("<Override ", printerSettingsDefault + "<Override ")
      case (_, xml) => xml
    }
    assertRelsResolve(src)
    val wb = read(src)
    val commented =
      wb.put(wb.sheets(0).put(ref"B1" -> "note").comment(ref"B1", Comment.plainText("c", None)))
    val out = write(commented, "first-comment-out")
    assertRelsResolve(out)
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    val sheetRels = rels(out, "xl/worksheets/_rels/sheet1.xml.rels")
    val vmlId = legacyDrawingId(sheetXml).getOrElse(fail(s"no legacyDrawing: $sheetXml"))
    assertEquals(
      sheetRels.findById(vmlId).map(_.`type`),
      Some(XmlUtil.relTypeVmlDrawing),
      s"legacyDrawing r:id=$vmlId must name the vmlDrawing rel: $sheetRels"
    )
    // the printerSettings rel rides through untouched
    assertEquals(
      sheetRels.findById("rId1").map(_.`type`),
      Some(relTypePrinterSettings),
      sheetRels.toString
    )
    assertEquals(read(out).sheets(0).comments.size, 1)
  }
