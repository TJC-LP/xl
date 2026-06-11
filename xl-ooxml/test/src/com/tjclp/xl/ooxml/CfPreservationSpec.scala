package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cf.{CfOperator, CfRule, ConditionalFormat}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.Dxf
import com.tjclp.xl.styles.color.Color

/**
 * GH-136 C1-interaction suite (the critical preservation laws): authored rules + preserved unparsed
 * rules coexist without duplication or loss.
 *
 *   - (a) read → author a rule → write → re-read: every preserved rule exactly once, preserved dxf
 *     entries byte-stable, authored priority = max(all)+1, no duplicate priorities
 *   - (b) cf-CLEAN cell edit: emitted cf elements equal the source slot (the dirty-gate proof)
 *   - (c) cleared model emits no conditionalFormatting (no resurrection)
 *   - (d) dxfs are append-only and a fixpoint under repeated read→write
 *   - (e) sheet-structure changes (insert/reorder) keep cf intact (self-healing gate)
 *   - both writer backends emit identical cf + dxfs
 */
class CfPreservationSpec extends FunSuite:

  private val green = Dxf.fill(Color.Rgb(0xff00b050))

  private def readFixture(name: String): (Path, Workbook) =
    val path = TestFixtures.copyToTemp(name)
    val wb = XlsxReader
      .read(path)
      .fold(err => fail(s"$name failed to read: ${err.message}"), identity)
    (path, wb)

  private def writeTo(wb: Workbook, label: String, config: WriterConfig = WriterConfig()): Path =
    val out = Files.createTempFile(s"xl-cf-$label-", ".xlsx")
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
          new String(zip.getInputStream(e).readAllBytes(), java.nio.charset.StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  private def occurrences(haystack: String, needle: String): Int =
    java.util.regex.Pattern.quote(needle).r.findAllMatchIn(haystack).size

  private def allPriorities(sheet: Sheet): Vector[Int] =
    sheet.conditionalFormats.flatMap {
      case ConditionalFormat.Rules(_, rules, _) => rules.flatMap(CfRule.priorityOf)
      case ConditionalFormat.Preserved(xml) => ConditionalFormat.scanPriorities(xml)
    }

  /**
   * The `<dxfs>` children in canonical (parsed, scala.xml-printed) form. Canonicalization is needed
   * for SOURCE comparisons only because styles.xml is always regenerated on write — openpyxl's
   * `<b val="1" />` spacing normalizes — while content and ORDER must be exact (append-only
   * invariant). DxfTableSpec separately proves the prefix is the same Elem objects.
   */
  private def dxfChildren(stylesXml: String): Vector[String] =
    val root = XmlSecurity.parseSafe(stylesXml, "styles").fold(e => fail(e.message), identity)
    (root \ "dxfs").headOption match
      case Some(dxfs) => (dxfs \ "dxf").collect { case e: scala.xml.Elem => e.toString }.toVector
      case None => Vector.empty

  // ===== fixture parse: the typed/Preserved split =====

  test("condformat.xlsx parses with the pinned typed/Preserved split") {
    val (_, wb) = readFixture("condformat.xlsx")
    val sheet = wb.sheets(0)
    val blocks = sheet.conditionalFormats
    assertEquals(blocks.size, 7) // 8 rules in 7 blocks (B2:B9 carries two rules)
    val rules = sheet.typedConditionalFormats.flatMap(_.rules)
    // typed: cellIs (x2: >100 and between), expression, colorScale, dataBar, top10, containsText
    assertEquals(rules.count(_.isInstanceOf[CfRule.CellIs]), 2)
    assertEquals(rules.count(_.isInstanceOf[CfRule.Expression]), 1)
    assertEquals(rules.count(_.isInstanceOf[CfRule.ColorScale]), 1)
    assertEquals(rules.count(_.isInstanceOf[CfRule.DataBar]), 1)
    assertEquals(rules.count(_.isInstanceOf[CfRule.Top10]), 1)
    assertEquals(rules.count(_.isInstanceOf[CfRule.Text]), 1)
    // iconSet is the ONLY Preserved rule; every envelope parsed typed
    val preservedRules = rules.collect { case p: CfRule.Preserved => p }
    assertEquals(preservedRules.size, 1)
    assert(preservedRules.forall(_.xml.contains("iconSet")), preservedRules.toString)
    assertEquals(
      sheet.conditionalFormats.collect { case p: ConditionalFormat.Preserved => p },
      Vector()
    )
    // multi-range sqref parsed as two ranges
    assert(
      sheet.typedConditionalFormats.exists(_.ranges.map(_.toA1) == Vector("H2:H5", "J2:J5")),
      sheet.typedConditionalFormats.map(_.ranges.map(_.toA1)).toString
    )
    // openpyxl dxf dialect (solid fgColor+bgColor) resolved into typed Dxf values
    val cellIsDxfs = rules.collect {
      case CfRule.CellIs(CfOperator.GreaterThan, _, _, Some(d), _, _) => d
    }
    assert(
      cellIsDxfs.exists(
        _.fill.contains(com.tjclp.xl.styles.fill.Fill.Solid(Color.Rgb(0xffffc7ce)))
      ),
      cellIsDxfs.toString
    )
  }

  test("condformat-lo.xlsx (LibreOffice noise attrs) parses total — rules ride Preserved") {
    val (_, wb) = readFixture("condformat-lo.xlsx")
    val sheet = wb.sheets(0)
    assertEquals(sheet.conditionalFormats.size, 7)
    // LO adds default-noise attrs outside the whitelists → rule-level Preserved, envelopes typed
    val rules = sheet.typedConditionalFormats.flatMap(_.rules)
    assert(rules.nonEmpty)
    assert(
      rules.collect { case p: CfRule.Preserved => p }.nonEmpty,
      "expected LO noise attrs to ride Preserved"
    )
    // and the whole thing round-trips losslessly at the model level
    val out = writeTo(wb, "lo-roundtrip")
    assertEquals(reread(out).sheets(0).conditionalFormats, sheet.conditionalFormats)
  }

  test("LAW: every Preserved cf payload XlsxReader produces re-parses (self-contained XML)") {
    // The dirty-write emission contract silently drops payloads that fail to re-parse
    // (CfCodec.parsePreserved), so reader-produced Preserved xml MUST be namespace-self-contained
    // — including prefixed descendants (the x14/xm extLst pairing Excel writes inside dataBar
    // cfRules) whose inline bindings the reader's block-level rebind severs from the chain.
    TestFixtures.all.foreach { name =>
      val (_, wb) = readFixture(name)
      for
        sheet <- wb.sheets
        cf <- sheet.conditionalFormats
      do
        cf match
          case ConditionalFormat.Preserved(xml) =>
            assert(
              XmlSecurity.parseSafe(xml, s"$name block Preserved").isRight,
              s"$name: block-level Preserved payload is not self-contained XML:\n$xml"
            )
          case ConditionalFormat.Rules(_, rules, _) =>
            rules.foreach {
              case CfRule.Preserved(xml, _) =>
                assert(
                  XmlSecurity.parseSafe(xml, s"$name rule Preserved").isRight,
                  s"$name: rule-level Preserved payload is not self-contained XML:\n$xml"
                )
              case _ => ()
            }
    }
  }

  // ===== (a) authored + preserved coexistence =====

  test("(a) author on top of preserved: no duplication, no loss, priority = max+1") {
    val (_, wb) = readFixture("condformat.xlsx")
    val sheetName = wb.sheets(0).name
    val before = wb.sheets(0).conditionalFormats
    val beforePriorities = allPriorities(wb.sheets(0))
    val updated = wb
      .update(
        sheetName,
        _.conditionalFormat(ref"K2:K9", CfRule.cellIs(CfOperator.GreaterThan, "100", green))
      )
      .fold(err => fail(s"update failed: $err"), identity)
    val authored = updated.sheets(0).conditionalFormats.lastOption
    // authored priority allocated ABOVE every existing (typed and Preserved-scanned) priority
    authored match
      case Some(ConditionalFormat.Rules(_, Vector(rule), _)) =>
        assertEquals(CfRule.priorityOf(rule), Some(beforePriorities.max + 1))
      case other => fail(s"expected authored block, got $other")

    val out = writeTo(updated, "author")
    val back = reread(out)
    val after = back.sheets(0).conditionalFormats
    assertEquals(after.size, before.size + 1)
    assertEquals(after.take(before.size), before, "preserved blocks must ride through unchanged")
    assertEquals(after.lastOption, authored, "authored rule must round-trip")
    // every preserved rule exactly once (the iconSet payload appears exactly once in the XML)
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assertEquals(raw"<iconSet\b".r.findAllIn(sheetXml).size, 1)
    // no duplicate priorities
    val priorities = allPriorities(back.sheets(0))
    assertEquals(priorities.distinct.size, priorities.size, s"duplicate priorities: $priorities")
  }

  test("(a) preserved dxf entries stay byte-identical when authoring appends a new dxf") {
    val (in, wb) = readFixture("condformat.xlsx")
    val sheetName = wb.sheets(0).name
    val updated = wb
      .update(
        sheetName,
        _.conditionalFormat(ref"K2:K9", CfRule.cellIs(CfOperator.GreaterThan, "100", green))
      )
      .fold(err => fail(s"update failed: $err"), identity)
    val out = writeTo(updated, "dxf-append")
    val beforeDxfs = dxfChildren(entryText(in, "xl/styles.xml"))
    val afterDxfs = dxfChildren(entryText(out, "xl/styles.xml"))
    assertEquals(afterDxfs.size, beforeDxfs.size + 1, "exactly one appended dxf")
    assertEquals(afterDxfs.take(beforeDxfs.size), beforeDxfs, "source dxfs byte-identical prefix")
    assert(entryText(out, "xl/styles.xml").contains(s"""<dxfs count="${beforeDxfs.size + 1}""""))
    // preserved dxfIds still resolve to the same typed values after re-read
    val back = reread(out)
    val origTyped = wb.sheets(0).typedConditionalFormats.flatMap(_.rules)
    val backTyped = back.sheets(0).typedConditionalFormats.flatMap(_.rules)
    assertEquals(backTyped.take(origTyped.size), origTyped)
  }

  test("(a) author on LO preserved rules: the x14 extLst dataBar block survives a dirty write") {
    val (_, wb) = readFixture("condformat-lo.xlsx")
    val sheetName = wb.sheets(0).name
    val before = wb.sheets(0).conditionalFormats
    val updated = wb
      .update(
        sheetName,
        _.conditionalFormat(ref"K2:K9", CfRule.cellIs(CfOperator.GreaterThan, "100", green))
      )
      .fold(err => fail(s"update failed: $err"), identity)
    val out = writeTo(updated, "lo-author")
    // dirty write must emit ALL 8 blocks — the D2:D9 dataBar rule is Preserved with a prefixed
    // x14 extLst child; a namespace-broken capture fails parsePreserved at emission and silently
    // drops the whole single-rule block
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    def count(sub: String): Int = occurrences(sheetXml, sub)
    assertEquals(count("<conditionalFormatting"), 8, "no block may be dropped")
    assertEquals(count("sqref=\"D2:D9\""), 1, "the dataBar block must survive exactly once")
    // the rule-level extLst x14 GUID pairing rides through exactly once...
    assertEquals(count("{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"), 1)
    // ...and stays paired with the worksheet-level x14:conditionalFormattings (same rule id,
    // once in the cfRule's <x14:id>, once in the worksheet extLst pairing)
    assertEquals(count("{26C9BFD7-9A20-468E-9236-D82474454A88}"), 2)
    // Re-read the cf model through the reader's own per-part pipeline (OoxmlWorksheet routes
    // through WorksheetReader's rebind; CfCodec.parseAll with the output's dxfs — exactly what
    // XlsxReader.convertToDomainSheet does). Full-workbook reread of a MODIFIED LibreOffice file
    // is blocked by a pre-existing cf-UNRELATED writer bug: workbook.xml is regenerated with
    // renumbered sheet rIds (rId1) while the preserved workbook.xml.rels keeps LO's layout
    // (rId1 = theme), so sheet resolution lands on theme1.xml. Reproduces at the parent commit
    // on small-values-lo.xlsx with a plain cell edit; tracked separately.
    def cfModelOf(p: Path): Vector[ConditionalFormat] =
      val ws = worksheet.OoxmlWorksheet
        .fromXml(
          XmlSecurity
            .parseSafe(entryText(p, "xl/worksheets/sheet1.xml"), "ws")
            .fold(e => fail(e.message), identity)
        )
        .fold(e => fail(e), identity)
      val styles = style.WorkbookStyles
        .fromXml(
          XmlSecurity
            .parseSafe(entryText(p, "xl/styles.xml"), "styles")
            .fold(e => fail(e.message), identity)
        )
        .fold(e => fail(e), identity)
      worksheet.CfCodec.parseAll(ws.conditionalFormatting, styles.dxfs)
    val after = cfModelOf(out)
    assertEquals(after.size, before.size + 1)
    assertEquals(after.take(before.size), before, "preserved blocks must ride through unchanged")
    after.lastOption match
      case Some(ConditionalFormat.Rules(ranges, Vector(_), _)) =>
        assertEquals(ranges.map(_.toA1), Vector("K2:K9"))
      case other => fail(s"expected the authored block last, got $other")
  }

  // ===== (b) the dirty-gate proof =====

  test("(b) cf-CLEAN cell edit: emitted cf elements equal the source slot") {
    val (in, wb) = readFixture("condformat.xlsx")
    val sheetName = wb.sheets(0).name
    val modified = wb
      .update(sheetName, _.put(ref"A20", CellValue.Text("unrelated edit")))
      .fold(err => fail(s"update failed: $err"), identity)
    val out = writeTo(modified, "clean-edit")
    // the gate proof: the regenerated worksheet's cf elements are the SOURCE elements (slot
    // passthrough), compared as parsed element text on both sides
    def cfElems(p: Path): Seq[String] =
      OoxmlWorksheet
        .fromXml(
          XmlSecurity
            .parseSafe(entryText(p, "xl/worksheets/sheet1.xml"), "ws")
            .fold(e => fail(e.message), identity)
        )
        .fold(e => fail(e), identity)
        .conditionalFormatting
        .map(_.toString)
    assertEquals(cfElems(out), cfElems(in))
    // dxfs byte-identical (nothing appended on a cf-clean write)
    assertEquals(
      dxfChildren(entryText(out, "xl/styles.xml")),
      dxfChildren(entryText(in, "xl/styles.xml"))
    )
    // and the model survives
    assertEquals(reread(out).sheets(0).conditionalFormats, wb.sheets(0).conditionalFormats)
  }

  test("(b) untouched write keeps the whole worksheet byte-identical (verbatim copy loop)") {
    val (in, wb) = readFixture("condformat.xlsx")
    val out = writeTo(wb, "untouched")
    assertEquals(
      entryText(out, "xl/worksheets/sheet1.xml"),
      entryText(in, "xl/worksheets/sheet1.xml")
    )
    assertEquals(entryText(out, "xl/styles.xml"), entryText(in, "xl/styles.xml"))
  }

  // ===== (c) deletion honored =====

  test("(c) cleared model emits NO conditionalFormatting (no resurrection)") {
    val (_, wb) = readFixture("condformat.xlsx")
    val sheetName = wb.sheets(0).name
    val cleared = wb
      .update(sheetName, _.copy(conditionalFormats = Vector.empty))
      .fold(err => fail(s"update failed: $err"), identity)
    val out = writeTo(cleared, "cleared")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(!sheetXml.contains("<conditionalFormatting"), "cf must not resurrect")
    assertEquals(reread(out).sheets(0).conditionalFormats, Vector.empty[ConditionalFormat])
  }

  test("(c) removing ONE block keeps the others (incl. Preserved payloads)") {
    val (_, wb) = readFixture("condformat.xlsx")
    val sheetName = wb.sheets(0).name
    val without = wb
      .update(sheetName, _.removeConditionalFormat(0))
      .fold(err => fail(s"update failed: $err"), identity)
    val out = writeTo(without, "remove-one")
    val back = reread(out)
    assertEquals(back.sheets(0).conditionalFormats, wb.sheets(0).conditionalFormats.drop(1))
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assertEquals(raw"<iconSet\b".r.findAllIn(sheetXml).size, 1, "preserved iconSet kept")
  }

  test("(c) removing one LO block keeps the x14 extLst dataBar block (second dirty-write path)") {
    val (_, wb) = readFixture("condformat-lo.xlsx")
    val without = wb
      .update(wb.sheets(0).name, _.removeConditionalFormat(0))
      .fold(err => fail(s"update failed: $err"), identity)
    val out = writeTo(without, "lo-remove-one")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assertEquals(
      occurrences(sheetXml, "<conditionalFormatting"),
      6,
      "ONLY the removed block may disappear"
    )
    assertEquals(occurrences(sheetXml, "sqref=\"D2:D9\""), 1, "the dataBar block must survive")
  }

  // ===== (d) dxfs fixpoint =====

  test("(d) dxfs are a fixpoint: second read→write of the same authored model appends nothing") {
    val (_, wb) = readFixture("condformat.xlsx")
    val sheetName = wb.sheets(0).name
    val updated = wb
      .update(
        sheetName,
        _.conditionalFormat(ref"K2:K9", CfRule.cellIs(CfOperator.GreaterThan, "100", green))
      )
      .fold(err => fail(s"update failed: $err"), identity)
    val out1 = writeTo(updated, "fixpoint-1")
    val wb2 = reread(out1)
    // force a regeneration of the sheet (cell edit) — cf is clean, dxfs must not grow
    val edited = wb2
      .update(wb2.sheets(0).name, _.put(ref"A21", CellValue.Text("second pass")))
      .fold(err => fail(s"update failed: $err"), identity)
    val out2 = writeTo(edited, "fixpoint-2")
    assertEquals(
      dxfChildren(entryText(out2, "xl/styles.xml")),
      dxfChildren(entryText(out1, "xl/styles.xml")),
      "dxfs table must be a fixpoint under edit→write cycles"
    )
    assertEquals(reread(out2).sheets(0).conditionalFormats, wb2.sheets(0).conditionalFormats)
  }

  // ===== (e) self-healing under sheet-structure changes =====

  test("(e) inserting a sheet before the cf sheet keeps cf intact (self-healing gate)") {
    val (_, wb) = readFixture("condformat.xlsx")
    val cfModel = wb.sheets(0).conditionalFormats
    val withNew = wb
      .insertSheet(0, Sheet(SheetName.unsafe("Fresh")).put(ref"A1", CellValue.Text("hi")))
      .fold(err => fail(s"insertSheet failed: $err"), identity)
    val out = writeTo(withNew, "insert-sheet")
    val back = reread(out)
    assertEquals(back.sheets.size, 2)
    val cfSheet = back.sheets.find(_.name.value == "CondFmt").getOrElse(fail("CondFmt missing"))
    assertEquals(cfSheet.conditionalFormats, cfModel, "no cf loss on sheet-structure change")
    // the shifted path mapping hands the FRESH sheet the cf sheet's preserved part — the
    // dirty-gate must refuse to copy that cf onto a model without it (no duplication)
    val fresh = back.sheets.find(_.name.value == "Fresh").getOrElse(fail("Fresh missing"))
    assertEquals(fresh.conditionalFormats, Vector.empty[ConditionalFormat], "cf must not leak")
  }

  // ===== fresh-workbook authoring + backend parity =====

  private def freshCfWorkbook: Workbook =
    val sheet = Sheet(SheetName.unsafe("New"))
      .put(ref"A1", CellValue.Number(BigDecimal(150)))
      .put(ref"A2", CellValue.Number(BigDecimal(50)))
      .conditionalFormat(
        ref"A1:A9",
        CfRule.cellIs(CfOperator.GreaterThan, "100", green),
        CfRule.dataBar(Color.Rgb(0xff638ec6))
      )
    Workbook(Vector(sheet))

  test("fresh workbook: authored cf emits on a no-source write and round-trips") {
    val wb = freshCfWorkbook
    val out = writeTo(wb, "fresh")
    val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
    assert(sheetXml.contains("<conditionalFormatting sqref=\"A1:A9\">"), sheetXml.take(500))
    assert(entryText(out, "xl/styles.xml").contains("<dxfs count=\"1\">"))
    assertEquals(reread(out).sheets(0).conditionalFormats, wb.sheets(0).conditionalFormats)
  }

  test("backend parity: ScalaXml and SaxStax emit structurally identical cf and dxfs") {
    val wb = freshCfWorkbook
    val domOut = writeTo(wb, "dom", WriterConfig.scalaXml)
    val saxOut = writeTo(wb, "sax", WriterConfig.saxStax)
    // The existing parity-spec pattern: compare PARSED structure with attributes canonicalized —
    // the StAX writer alphabetizes streamed-Elem attributes (SaxSupport.metaDataAttributes), a
    // deterministic, pre-existing cross-backend skew shared by every preserved element.
    def normalize(n: scala.xml.Node): String =
      def attrsOf(e: scala.xml.Elem): String =
        e.attributes.asAttrMap.toSeq.sorted.map((k, v) => s"$k=$v").mkString(" ")
      n match
        case e: scala.xml.Elem =>
          val kids = e.child.collect { case c: scala.xml.Elem => normalize(c) }.mkString
          val text = e.child.collect { case t: scala.xml.Text => t.data }.mkString.trim
          s"<${e.label} ${attrsOf(e)}>$text$kids</${e.label}>"
        case other => other.text.trim
    def cfSection(p: Path): Seq[String] =
      val ws = XmlSecurity
        .parseSafe(entryText(p, "xl/worksheets/sheet1.xml"), "ws")
        .fold(e => fail(e.message), identity)
      (ws \ "conditionalFormatting").map(normalize)
    def dxfsSection(p: Path): Seq[String] =
      val st = XmlSecurity
        .parseSafe(entryText(p, "xl/styles.xml"), "styles")
        .fold(e => fail(e.message), identity)
      (st \ "dxfs").map(normalize)
    assert(cfSection(domOut).nonEmpty, "DOM backend emitted no cf")
    assertEquals(cfSection(saxOut), cfSection(domOut), "backends must emit identical cf")
    assertEquals(dxfsSection(saxOut), dxfsSection(domOut), "backends must emit identical dxfs")
    // strongest check: both backends read back to the SAME typed model
    assertEquals(reread(saxOut).sheets(0).conditionalFormats, wb.sheets(0).conditionalFormats)
    assertEquals(reread(domOut).sheets(0).conditionalFormats, wb.sheets(0).conditionalFormats)
  }
