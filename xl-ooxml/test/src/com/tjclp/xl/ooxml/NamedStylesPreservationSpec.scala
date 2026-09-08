package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
import scala.xml.{Elem, MetaData, Null, PrefixedAttribute, Text, TopScope, UnprefixedAttribute}

import com.tjclp.xl.api.{Sheet, Workbook}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.style.{OoxmlStyles, StyleIndex, WorkbookStyles}
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.units.StyleId
import munit.FunSuite

/**
 * GH-610: an Excel-authored styles part keeps its named styles through xl.
 *
 * `named-styles-excel.xlsx` is an Excel model's styles.xml stripped to 8 named styles (Normal +
 * seven house styles), 10 cellXfs — two of which (6 and 8) differ ONLY in `xfId`: "Comma 2" applied
 * vs the same formatting typed by hand — plus the recent-colours palette, the table-style defaults
 * and the x14/x15 `extLst`. Every write used to collapse that to one `Normal`, `xfId="0"` on every
 * xf, and drop the palette and extLst.
 *
 * Laws pinned here: the reader keeps every cell on the exact cellXf its `s=` named (positional
 * registry, twins included); a regenerating write carries `cellStyleXfs`, `cellStyles`,
 * `tableStyles`, `colors` and `extLst` verbatim (byte-identical on the DOM backend, structurally on
 * StAX) with each cellXf's `xfId` intact; a style xl authors is appended with `xfId="0"`; an
 * xl-authored workbook — and a source without masters — still writes exactly one `Normal` and no
 * `xfId` that could dangle.
 */
@SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
class NamedStylesPreservationSpec extends FunSuite:

  private val fixture = "named-styles-excel.xlsx"
  private val sourceXfIds = Vector(0, 1, 2, 3, 4, 5, 6, 7, 0, 6)
  private val passthrough =
    List("cellStyleXfs", "cellXfs", "cellStyles", "tableStyles", "colors", "extLst")
  private val backends = List(
    "ScalaXml" -> WriterConfig(backend = XmlBackend.ScalaXml),
    "SaxStax" -> WriterConfig(backend = XmlBackend.SaxStax)
  )

  // ---------------------------------------------------------------- helpers

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  private def parse(xml: String): Elem =
    XmlSecurity
      .parseSafe(xml, "styles.xml")
      .fold(e => fail(s"parse failed: ${e.message}"), identity)

  /** The raw text of the first `<tag ...>...</tag>` (or self-closing `<tag .../>`) in `xml`. */
  private def section(xml: String, tag: String): String =
    val start = xml.indexOf(s"<$tag")
    assert(start >= 0, s"<$tag> missing in:\n$xml")
    val close = xml.indexOf('>', start)
    if xml.charAt(close - 1) == '/' then xml.substring(start, close + 1)
    else
      val end = xml.indexOf(s"</$tag>", start)
      assert(end >= 0, s"</$tag> missing in:\n$xml")
      xml.substring(start, end + tag.length + 3)

  private def count(xml: String, tag: String): Int =
    s"""<$tag count="(\\d+)"""".r
      .findFirstMatchIn(xml)
      .map(_.group(1).toInt)
      .getOrElse(fail(s"<$tag count=...> missing in:\n$xml"))

  private def xfIdsOf(cellXfs: String): Vector[Int] =
    """xfId="(\d+)"""".r.findAllMatchIn(cellXfs).map(_.group(1).toInt).toVector

  private def styleAttr(sheetXml: String, cell: String): Option[String] =
    s"""<c r="$cell"[^>]*\\bs="(\\d+)"""".r.findFirstMatchIn(sheetXml).map(_.group(1))

  private def readFixture(): (Path, Workbook) =
    val in = TestFixtures.copyToTemp(fixture)
    (in, XlsxReader.read(in).fold(e => fail(s"read failed: ${e.message}"), identity))

  private def writeTemp(wb: Workbook, config: WriterConfig, tag: String): Path =
    val out = Files.createTempFile(s"xl-gh610-$tag-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter.writeWith(wb, out, config).fold(e => fail(s"write failed: ${e.message}"), identity)
    out

  /** Attribute-order-, namespace- and whitespace-insensitive shape, for DOM/StAX parity. */
  private def normalize(elem: Elem): Elem =
    val attrs = attributePairs(elem.attributes)
      .filterNot { case (key, _) => key == "xmlns" || key.startsWith("xmlns:") }
      .sortBy(_._1)
      .foldLeft(Null: MetaData) { case (acc, (key, value)) =>
        key.split(":", 2) match
          case Array(prefix, local) => new PrefixedAttribute(prefix, local, value, acc)
          case _ => new UnprefixedAttribute(key, value, acc)
      }
    val children = elem.child.flatMap {
      case t: Text if t.text.forall(_.isWhitespace) => None
      case e: Elem => Some(normalize(e))
      case other => Some(other)
    }
    elem.copy(prefix = null, scope = TopScope, attributes = attrs, child = children)

  private def attributePairs(attrs: MetaData): List[(String, String)] = attrs match
    case Null => Nil
    case PrefixedAttribute(pre, key, value, next) =>
      (s"$pre:$key" -> value.text) :: attributePairs(next)
    case UnprefixedAttribute(key, value, next) => (key -> value.text) :: attributePairs(next)

  private def child(root: Elem, tag: String): Elem =
    XmlUtil.getChildren(root, tag).headOption.getOrElse(fail(s"<$tag> missing in:\n$root"))

  // ---------------------------------------------------------------- reader

  test("reader: WorkbookStyles carries every cellXf's xfId and the passthrough sections") {
    val (in, _) = readFixture()
    val styles = WorkbookStyles
      .fromXml(parse(entryText(in, "xl/styles.xml")))
      .fold(e => fail(s"styles parse failed: $e"), identity)

    assertEquals(styles.xfIds, sourceXfIds)
    assertEquals(styles.xfIdAt(8), 0)
    assertEquals(styles.xfIdAt(99), 0, "out of range reads as Normal")
    assert(styles.preserved.hasNamedStyles)
    assertEquals(styles.preserved.cellStyleXfs.map(XmlUtil.getChildren(_, "xf").size), Some(8))
    val names = styles.preserved.cellStyles.toList
      .flatMap(XmlUtil.getChildren(_, "cellStyle"))
      .map(_ \@ "name")
    assertEquals(
      names,
      List(
        "Bad 2",
        "Calculation 2",
        "Comma 2",
        "Heading 1 2",
        "Input 2",
        "Normal",
        "Percent 3 4",
        "Smart Forecast"
      )
    )
    assertEquals(styles.preserved.colors.map(c => (c \ "mruColors" \ "color").size), Some(9))
    assertEquals(
      styles.preserved.tableStyles.map(_ \@ "defaultTableStyle"),
      Some("TableStyleMedium2")
    )
    assertEquals(styles.preserved.extLst.map(XmlUtil.getChildren(_, "ext").size), Some(2))
  }

  test("reader: a styles part without the sections yields the empty passthrough") {
    val styles = WorkbookStyles
      .fromXml(parse(XmlUtil.compact(OoxmlStyles.minimal.toXml)))
      .fold(e => fail(s"styles parse failed: $e"), identity)
    // xl's own output carries the Normal pair: that IS a cellStyleXfs section, so it rides through
    assert(styles.preserved.hasNamedStyles)
    assertEquals(styles.preserved.colors, None)
    assertEquals(styles.preserved.extLst, None)
    assertEquals(styles.preserved.tableStyles, None)
    assertEquals(styles.xfIds, Vector(0))
  }

  test("reader: a cell keeps the exact cellXf its s= named — canonical-key twins included") {
    val (_, wb) = readFixture()
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    // ten source cellXfs, ten registry slots: no read-time collapse of the 6/8 twins
    assertEquals(sheet.styleRegistry.size, 10)
    assertEquals(sheet(ref"A1").styleId, Some(StyleId(0)))
    assertEquals(sheet(ref"A7").styleId, Some(StyleId(6)), "Comma 2 applied")
    assertEquals(sheet(ref"A9").styleId, Some(StyleId(8)), "the same formatting, direct")
    assertEquals(sheet(ref"A10").styleId, Some(StyleId(9)))
    val twin6 = sheet.styleRegistry.get(StyleId(6)).getOrElse(fail("slot 6"))
    val twin8 = sheet.styleRegistry.get(StyleId(8)).getOrElse(fail("slot 8"))
    assertEquals(twin6, twin8, "equal styles in two slots")
    // a NEW use of that formatting shares the DIRECT twin (xfId 0), not the "Comma 2" one
    assertEquals(sheet.styleRegistry.indexOf(twin8), Some(StyleId(8)))
  }

  test("a new use of twin formatting shares the direct twin on every path (review item 1)") {
    val (_, wb) = readFixture()
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val comma2 = sheet.styleRegistry.get(StyleId(6)).getOrElse(fail("slot 6"))
    // same sheet: the registry hands out slot 8
    val sameSheet =
      sheet.put(ref"C1", CellValue.Number(BigDecimal(1))).withCellStyle(ref"C1", comma2)
    assertEquals(sameSheet(ref"C1").styleId, Some(StyleId(8)))
    // a NEW sheet has its own registry (slot 1); the surgical StyleIndex must land it on 8, not 6
    val fresh =
      Sheet("Fresh").put(ref"A1", CellValue.Number(BigDecimal(2))).withCellStyle(ref"A1", comma2)
    val edited = wb.put(sameSheet).put(fresh)
    val (index, remappings) =
      StyleIndex.fromWorkbook(edited, sheetsRequiringRemapping = Set(0, 1))
    assertEquals(index.styleToIndex.get(comma2.canonicalKey), Some(StyleId(8)))
    assertEquals(remappings.getOrElse(1, Map.empty).get(1), Some(8), "fresh sheet slot 1 -> xf 8")
    // (the fresh sheet's slot 0, CellStyle.default, is a genuinely new xf on this source — the
    // fixture's Normal carries a theme colour — so exactly one xf is appended, none for the twin)
    assertEquals(index.cellStyles.size, 11, "only the fresh registry's default was appended")
    val out = writeTemp(edited, WriterConfig.default, "direct-twin")
    assertEquals(styleAttr(entryText(out, "xl/worksheets/sheet1.xml"), "C1"), Some("8"))
    assertEquals(styleAttr(entryText(out, "xl/worksheets/sheet2.xml"), "A1"), Some("8"))
  }

  // ---------------------------------------------------------------- writer: both backends

  backends.foreach { case (name, config) =>
    test(s"dirty write keeps the named-style tables, xfIds, palette and extLst ($name)") {
      val (in, wb) = readFixture()
      val source = entryText(in, "xl/styles.xml")
      val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
      // a new cell forces the sheet through regeneration (the verbatim-copy loop is not a pass)
      val out = writeTemp(wb.put(sheet.put(ref"C1", CellValue.Text("dirty"))), config, name)
      val styles = entryText(out, "xl/styles.xml")

      assertEquals(count(styles, "cellStyleXfs"), 8, styles)
      assertEquals(count(styles, "cellStyles"), 8, styles)
      assertEquals(count(styles, "cellXfs"), 10, styles)
      assertEquals(xfIdsOf(section(styles, "cellXfs")), sourceXfIds, styles)
      // source cellXfs ride through verbatim: apply* flags and quotePrefix-class attributes intact
      assertEquals(
        """applyFont="1"""".r.findAllIn(section(styles, "cellXfs")).size,
        """applyFont="1"""".r.findAllIn(section(source, "cellXfs")).size,
        "applyFont flags dropped"
      )

      // CT_Stylesheet order, every section present exactly once
      val sections = List("numFmts", "fonts", "fills", "borders") ++
        List("cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst")
      val offsets = sections.map(t => styles.indexOf(s"<$t"))
      assert(offsets.forall(_ >= 0), s"missing section: ${sections.zip(offsets)}\n$styles")
      assertEquals(offsets, offsets.sorted, s"schema order violated: ${sections.zip(offsets)}")
      sections.foreach(t => assertEquals(styles.sliding(t.length + 1).count(_ == s"<$t"), 1, t))

      // passthrough sections: byte-identical on the DOM backend, same shape on StAX (which sorts
      // attributes and rebinds namespaces at its own discretion)
      val (outRoot, srcRoot) = (parse(styles), parse(source))
      passthrough.foreach { tag =>
        assertEquals(normalize(child(outRoot, tag)), normalize(child(srcRoot, tag)), s"<$tag>")
        if name == "ScalaXml" then
          assertEquals(section(styles, tag), section(source, tag), s"<$tag> not byte-verbatim")
      }
      // (StAX writes childless elements long-form, so match values, not tag shapes)
      assert(styles.contains("""<color rgb="FF7030A0""""), styles)
      assert(styles.contains("""defaultTimelineStyle="TimeSlicerStyleLight1""""), styles)

      // the regenerated sheet keeps every s= — the 6/8 twins were not collapsed
      val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
      assertEquals(styleAttr(sheetXml, "A7"), Some("6"), sheetXml)
      assertEquals(styleAttr(sheetXml, "A9"), Some("8"), sheetXml)
      assertEquals(styleAttr(sheetXml, "A10"), Some("9"), sheetXml)

      // and the output reads back the same way
      val back = XlsxReader.read(out).fold(e => fail(s"re-read failed: ${e.message}"), identity)
      val backSheet = back.sheets.headOption.getOrElse(fail("no sheet"))
      assertEquals(backSheet(ref"A9").styleId, Some(StyleId(8)))
      assertEquals(backSheet(ref"C1").value, CellValue.Text("dirty"))
    }

    test(
      s"a restyled cell appends one xl-authored cellXf with xfId 0; source tables stay ($name)"
    ) {
      val (_, wb) = readFixture()
      val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
      val arial = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
      val out = writeTemp(wb.put(sheet.withCellStyle(ref"A2", arial)), config, s"restyle-$name")
      val styles = entryText(out, "xl/styles.xml")

      assertEquals(count(styles, "cellXfs"), 11, styles)
      assertEquals(xfIdsOf(section(styles, "cellXfs")), sourceXfIds :+ 0, styles)
      assertEquals(count(styles, "cellStyleXfs"), 8, styles)
      assertEquals(count(styles, "cellStyles"), 8, styles)
      assertEquals(count(styles, "fonts"), 9, "source fonts + Arial")
      if name == "ScalaXml" then
        // the ten source records are a byte-verbatim prefix; only xl's xf is regenerated
        val (_, sourceSrc) = (wb, entryText(TestFixtures.copyToTemp(fixture), "xl/styles.xml"))
        val sourceRecords = section(sourceSrc, "cellXfs").stripSuffix("</cellXfs>")
        assert(
          section(styles, "cellXfs").startsWith(
            sourceRecords.replace("count=\"10\"", "count=\"11\"")
          ),
          section(styles, "cellXfs")
        )
      val sheetXml = entryText(out, "xl/worksheets/sheet1.xml")
      assertEquals(styleAttr(sheetXml, "A2"), Some("10"), sheetXml)
      assertEquals(styleAttr(sheetXml, "A9"), Some("8"), sheetXml)
    }

    test(s"an xl-authored workbook still writes exactly one Normal master and xfId 0 ($name)") {
      val sheet = Sheet("S")
        .put(ref"A1", CellValue.Text("x"))
        .withCellStyle(ref"A1", CellStyle.default.withFont(Font("Arial", 12.0, bold = true)))
      val out = writeTemp(Workbook(Vector(sheet)), config, s"fresh-$name")
      val styles = entryText(out, "xl/styles.xml")
      assertEquals(count(styles, "cellStyleXfs"), 1, styles)
      assertEquals(count(styles, "cellStyles"), 1, styles)
      assert(styles.contains("""name="Normal""""), styles)
      assertEquals(xfIdsOf(section(styles, "cellXfs")), Vector(0, 0), styles)
      List("<colors", "<extLst", "<tableStyles").foreach(t => assert(!styles.contains(t), t))
    }
  }

  test("DOM and StAX backends agree on the whole styles part (normalized) for the model shape") {
    val (_, wb) = readFixture()
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val edited = wb.put(sheet.put(ref"C1", CellValue.Text("dirty")))
    val dom = parse(entryText(writeTemp(edited, WriterConfig.scalaXml, "dom"), "xl/styles.xml"))
    val sax = parse(entryText(writeTemp(edited, WriterConfig.saxStax, "sax"), "xl/styles.xml"))
    assertEquals(normalize(sax), normalize(dom))
  }

  test("dropping the source context (scratch write) falls back to one Normal and xfId 0") {
    val (_, wb) = readFixture()
    val scratch = wb.copy(sourceContext = None)
    val styles = entryText(writeTemp(scratch, WriterConfig.default, "scratch"), "xl/styles.xml")
    assertEquals(count(styles, "cellStyleXfs"), 1, styles)
    assertEquals(count(styles, "cellStyles"), 1, styles)
    assert(xfIdsOf(section(styles, "cellXfs")).forall(_ == 0), styles)
    assert(!styles.contains("<colors"), "no palette to carry without a source")
  }

  // ---------------------------------------------------------------- StyleIndex

  test("StyleIndex (source mode): xfIds follow the preserved cellXfs; appended styles get 0") {
    val (_, wb) = readFixture()
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val arial = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
    val edited = wb.put(sheet.withCellStyle(ref"A2", arial))
    val (index, remappings) = StyleIndex.fromWorkbook(edited, sheetsRequiringRemapping = Set(0))
    assertEquals(index.xfIds, sourceXfIds :+ 0)
    assertEquals(index.xfIdAt(10), 0)
    // positional identity for every preserved slot — including twin 8, which canonical-key
    // dedup used to fold onto slot 6
    val remap = remappings.getOrElse(0, fail("sheet 0 has no remapping"))
    sourceXfIds.indices.foreach(i => assertEquals(remap.get(i), Some(i), s"slot $i"))
    assertEquals(remap.get(10), Some(10))
  }

  test("an appended style reuses the FIRST of two equal source fonts (review item 3)") {
    // fixture fonts 0 and 3 parse equal (Calibri 11, theme 1 — they differ only in `scheme`);
    // cellXf 0 references font 0, cellXf 6 ("Comma 2") references font 3
    val (_, wb) = readFixture()
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val font0 = sheet.styleRegistry.get(StyleId(0)).getOrElse(fail("slot 0")).font
    assertEquals(sheet.styleRegistry.get(StyleId(6)).map(_.font), Some(font0), "fixture premise")
    val navy = CellStyle.default.withFont(font0).withFill(Fill.Solid(Color.Rgb(0xff003366)))
    val out = writeTemp(wb.put(sheet.withCellStyle(ref"A2", navy)), WriterConfig.default, "font0")
    val styles = entryText(out, "xl/styles.xml")
    val appended = XmlUtil.getChildren(child(parse(styles), "cellXfs"), "xf").lift(10)
    assertEquals(appended.map(_ \@ "fontId"), Some("0"), styles)
  }

  test("StyleIndex (fresh): no xfIds, so every cellXf resolves to Normal") {
    val sheet = Sheet("S").put(ref"A1", CellValue.Text("x"))
    val (index, _) = StyleIndex.fromWorkbook(Workbook(Vector(sheet)))
    assertEquals(index.xfIds, Vector.empty)
    assertEquals(index.xfIdAt(0), 0)
  }

  // ---------------------------------------------------------------- fragment remapping

  test("remapCellStyleXfs: identity hands back the same element (verbatim bytes)") {
    val el = parse(
      """<cellStyleXfs count="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">""" +
        """<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>""" +
        """<xf numFmtId="9" fontId="1" fillId="3" borderId="0" applyFont="0"><alignment horizontal="right"/></xf>""" +
        "</cellStyleXfs>"
    )
    val same = OoxmlStyles.remapCellStyleXfs(el, identity, identity, identity)
    assert(same eq el, "identity remap must not rebuild the fragment")
  }

  test("remapCellStyleXfs: a moved fill re-points fillId only, attribute order intact") {
    val el = parse(
      """<cellStyleXfs count="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">""" +
        """<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>""" +
        """<xf numFmtId="9" fontId="1" fillId="3" borderId="0" applyFont="0"><alignment horizontal="right"/></xf>""" +
        "</cellStyleXfs>"
    )
    val moved = OoxmlStyles.remapCellStyleXfs(el, identity, i => if i == 3 then 2 else i, identity)
    val out = XmlUtil.compact(moved)
    assert(
      out.contains(
        """<xf numFmtId="9" fontId="1" fillId="2" borderId="0" applyFont="0"><alignment horizontal="right"/></xf>"""
      ),
      out
    )
    assert(out.contains("""<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>"""), out)
    assert(out.contains("""count="2""""), out)
  }

  test("followTable: positional identity, then equal-value fallback, then unchanged") {
    val source = Vector("none", "gray", "red", "red")
    val out = Vector("none", "gray", "red") // the serializer's dedup dropped the twin
    val f = OoxmlStyles.followTable(source, out, out.zipWithIndex.toMap)
    assertEquals(f(0), 0)
    assertEquals(f(2), 2)
    assertEquals(f(3), 2, "the dropped twin follows its equal")
    assertEquals(f(7), 7, "a dangling source reference is left alone")
  }

  // ---------------------------------------------------------------- degenerate source

  test("a source with xfIds but no masters gets the single Normal pair and xfId 0 everywhere") {
    val stylesXml =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
        |<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
        |<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        |<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="5"/></cellXfs>
        |</styleSheet>""".stripMargin
    val sheetXml =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
        |<row r="1"><c r="A1" s="1"><v>0.5</v></c></row>
        |</sheetData></worksheet>""".stripMargin
    val wb = XlsxReader
      .readFromBytes(rawXlsx(stylesXml, sheetXml))
      .fold(e => fail(s"read failed: ${e.message}"), identity)
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val out = writeTemp(
      wb.put(sheet.put(ref"B1", CellValue.Text("dirty"))),
      WriterConfig.default,
      "nomasters"
    )
    val styles = entryText(out, "xl/styles.xml")
    assertEquals(count(styles, "cellStyleXfs"), 1, styles)
    assertEquals(count(styles, "cellStyles"), 1, styles)
    assertEquals(xfIdsOf(section(styles, "cellXfs")), Vector(0, 0), "xfId 5 would dangle")
    assertEquals(styleAttr(entryText(out, "xl/worksheets/sheet1.xml"), "A1"), Some("1"))
  }

  test("empty masters (<cellStyleXfs count=\"0\"/>) are repaired to one Normal, xfIds to 0") {
    val stylesXml =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
        |<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
        |<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        |<cellStyleXfs count="0"/>
        |<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="1" applyNumberFormat="1"/></cellXfs>
        |<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        |</styleSheet>""".stripMargin
    val sheetXml =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
        |<row r="1"><c r="A1" s="1"><v>0.5</v></c></row>
        |</sheetData></worksheet>""".stripMargin
    val parts = WorkbookStyles
      .fromXml(parse(stylesXml))
      .fold(e => fail(s"styles parse failed: $e"), identity)
      .preserved
    assertEquals(parts.masterCount, 0)
    assert(!parts.hasNamedStyles, "an empty masters table is not a named-style table")
    val wb = XlsxReader
      .readFromBytes(rawXlsx(stylesXml, sheetXml))
      .fold(e => fail(s"read failed: ${e.message}"), identity)
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val out = writeTemp(
      wb.put(sheet.put(ref"B1", CellValue.Text("dirty"))),
      WriterConfig.default,
      "empty-masters"
    )
    val styles = entryText(out, "xl/styles.xml")
    assertEquals(count(styles, "cellStyleXfs"), 1, styles)
    assert(styles.contains("""<cellStyleXfs count="1"><xf"""), styles)
    assertEquals(count(styles, "cellStyles"), 1, styles)
    assertEquals(xfIdsOf(section(styles, "cellXfs")), Vector(0, 0), styles)
    // the source record itself rode through (its flag survives), only its xfId was repaired
    assert(section(styles, "cellXfs").contains("""applyNumberFormat="1""""), styles)
  }

  test("an out-of-range xfId in a source with masters is repaired to 0; in-range ones stay") {
    val stylesXml =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
        |<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
        |<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        |<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        |<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="1"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="7"/></cellXfs>
        |<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Pct" xfId="1"/></cellStyles>
        |</styleSheet>""".stripMargin
    val sheetXml =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
        |<row r="1"><c r="A1" s="1"><v>0.5</v></c><c r="B1" s="2"><v>0.25</v></c></row>
        |</sheetData></worksheet>""".stripMargin
    val wb = XlsxReader
      .readFromBytes(rawXlsx(stylesXml, sheetXml))
      .fold(e => fail(s"read failed: ${e.message}"), identity)
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    val out = writeTemp(
      wb.put(sheet.put(ref"C1", CellValue.Text("dirty"))),
      WriterConfig.default,
      "oor-xfid"
    )
    val styles = entryText(out, "xl/styles.xml")
    assertEquals(count(styles, "cellStyleXfs"), 2, styles)
    assertEquals(count(styles, "cellStyles"), 2, styles)
    assertEquals(xfIdsOf(section(styles, "cellXfs")), Vector(0, 1, 0), "7 dangles past 2 masters")
    val sheetOut = entryText(out, "xl/worksheets/sheet1.xml")
    assertEquals(styleAttr(sheetOut, "A1"), Some("1"))
    assertEquals(styleAttr(sheetOut, "B1"), Some("2"))
  }

  private def rawXlsx(stylesXml: String, sheetXml: String): Array[Byte] =
    val contentTypes =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        |<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        |<Default Extension="xml" ContentType="application/xml"/>
        |<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        |<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        |<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
        |</Types>""".stripMargin
    val rootRels =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        |<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        |</Relationships>""".stripMargin
    val workbook =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        |<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
        |</workbook>""".stripMargin
    val workbookRels =
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        |<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        |<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        |</Relationships>""".stripMargin
    val output = new ByteArrayOutputStream()
    val zip = new ZipOutputStream(output)
    try
      List(
        "[Content_Types].xml" -> contentTypes,
        "_rels/.rels" -> rootRels,
        "xl/workbook.xml" -> workbook,
        "xl/_rels/workbook.xml.rels" -> workbookRels,
        "xl/styles.xml" -> stylesXml,
        "xl/worksheets/sheet1.xml" -> sheetXml
      ).foreach { case (name, content) =>
        zip.putNextEntry(new ZipEntry(name))
        zip.write(content.getBytes(StandardCharsets.UTF_8))
        zip.closeEntry()
      }
    finally zip.close()
    output.toByteArray
