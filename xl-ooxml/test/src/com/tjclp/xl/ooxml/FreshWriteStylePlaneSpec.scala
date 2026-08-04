package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cf.{CfRule, ConditionalFormat}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * GH-471: fresh write after read (`copy(sourceContext = None)`) must re-register the style plane.
 *
 *   - (a) numFmt cross-wire: the source declares custom formats at non-contiguous ids (194, 300,
 *     407); the fresh write renumbers `<numFmts>` to 164+ so every `<xf>` numFmtId must be
 *     re-mapped through the rebuilt registry (keyed by format CODE) — keeping the SOURCE ids
 *     silently changes cell display formats.
 *   - (b) dxfs under-registration: `CfRule.Preserved` payloads carry source dxfId refs (7, 9) but
 *     the fresh write emits only the modeled dxfs — the referenced `<dxf>` payloads must be carried
 *     into the rebuilt table and the refs renumbered, or the shipped file has out-of-range dxfIds.
 */
// Test code uses .get/.head for brevity in assertions
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class FreshWriteStylePlaneSpec extends FunSuite:

  private val nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  private val contentTypesXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""

  private val rootRelationshipsXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

  private val workbookRelationshipsXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

  private val workbookXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="$nsMain"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""

  private def buildXlsx(stylesXml: String, worksheetXml: String): Array[Byte] =
    val parts = Vector(
      "[Content_Types].xml" -> contentTypesXml,
      "_rels/.rels" -> rootRelationshipsXml,
      "xl/workbook.xml" -> workbookXml,
      "xl/_rels/workbook.xml.rels" -> workbookRelationshipsXml,
      "xl/styles.xml" -> stylesXml,
      "xl/worksheets/sheet1.xml" -> worksheetXml
    )
    val baos = new ByteArrayOutputStream()
    val zipOut = new ZipOutputStream(baos)
    parts.foreach { case (name, text) =>
      zipOut.putNextEntry(new ZipEntry(name))
      zipOut.write(text.getBytes(StandardCharsets.UTF_8))
      zipOut.closeEntry()
    }
    zipOut.close()
    baos.toByteArray

  private def readBytes(bytes: Array[Byte]): Workbook =
    XlsxReader
      .readFromBytes(bytes)
      .fold(err => fail(s"fixture failed to read: ${err.message}"), identity)

  private def writeFresh(wb: Workbook, label: String): Path =
    val out = Files.createTempFile(s"xl-gh471-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter
      .write(wb.copy(sourceContext = None), out)
      .fold(err => fail(s"fresh write failed: ${err.message}"), _ => ())
    out

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) =>
          new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  private def parseEntry(path: Path, name: String): scala.xml.Elem =
    XmlSecurity.parseSafe(entryText(path, name), name).fold(e => fail(e.message), identity)

  private def reread(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"re-read failed: ${err.message}"), identity)

  private def styleOf(sheet: Sheet, r: com.tjclp.xl.addressing.ARef): CellStyle =
    sheet.cells
      .get(r)
      .flatMap(_.styleId)
      .flatMap(sheet.styleRegistry.get)
      .getOrElse(fail(s"no style on ${r.toA1}"))

  // ===== (a) numFmt cross-wire =====

  private val fyCode = "\"FY\"0\"A\"_)"
  private val multCode = "0.0\"x\""
  private val bpsCode = "0.000\"y\""

  private val numFmtStylesXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="$nsMain">
  <numFmts count="3">
    <numFmt numFmtId="194" formatCode="&quot;FY&quot;0&quot;A&quot;_)"/>
    <numFmt numFmtId="300" formatCode="0.0&quot;x&quot;"/>
    <numFmt numFmtId="407" formatCode="0.000&quot;y&quot;"/>
  </numFmts>
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="194" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="300" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="407" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="15" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

  private val numFmtWorksheetXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain">
  <sheetData>
    <row r="1">
      <c r="A1" s="1"><v>2026</v></c>
      <c r="B1" s="2"><v>2.5</v></c>
      <c r="C1" s="3"><v>0.125</v></c>
      <c r="D1" s="4"><v>45000</v></c>
    </row>
  </sheetData>
</worksheet>"""

  test("fresh write re-maps cellXf numFmtIds through the rebuilt registry (GH-471a)") {
    val wb = readBytes(buildXlsx(numFmtStylesXml, numFmtWorksheetXml))
    val sheet = wb.sheets(0)
    // Sanity: source ids and codes arrive intact
    assertEquals(styleOf(sheet, ref"A1").numFmt, NumFmt.Custom(fyCode))
    assertEquals(styleOf(sheet, ref"A1").numFmtId, Some(194))
    assertEquals(styleOf(sheet, ref"C1").numFmtId, Some(407))

    val out = writeFresh(wb, "numfmt")

    // Structural invariant: every custom-range numFmtId referenced by cellXfs is DECLARED
    val styles = parseEntry(out, "xl/styles.xml")
    val declared: Map[Int, String] = (styles \ "numFmts" \ "numFmt").collect {
      case e: scala.xml.Elem =>
        e.attribute("numFmtId").get.text.toInt -> e.attribute("formatCode").get.text
    }.toMap
    val referenced: Vector[Int] = (styles \ "cellXfs" \ "xf").toVector
      .flatMap(_.attribute("numFmtId"))
      .map(_.text.toInt)
    referenced.filter(_ >= NumFmt.FirstCustomId).foreach { id =>
      assert(
        declared.contains(id),
        s"cellXf references numFmtId $id but <numFmts> declares only ${declared.keySet}"
      )
    }
    // The three custom codes survive with resolvable ids
    assertEquals(declared.values.toSet, Set(fyCode, multCode, bpsCode))
    // The raw built-in id 15 (a d-mmm-yy Date variant) is NOT remapped away
    assert(referenced.contains(15), s"built-in numFmtId 15 lost: $referenced")

    // Behavioral proof: reopened formats resolve to the source codes
    val sheet2 = reread(out).sheets(0)
    assertEquals(styleOf(sheet2, ref"A1").numFmt, NumFmt.Custom(fyCode))
    assertEquals(styleOf(sheet2, ref"B1").numFmt, NumFmt.Custom(multCode))
    assertEquals(styleOf(sheet2, ref"C1").numFmt, NumFmt.Custom(bpsCode))
    assertEquals(styleOf(sheet2, ref"D1").numFmt, NumFmt.Date)
  }

  // ===== (b) dxfs under-registration =====

  private val dxfStylesXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="$nsMain">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="10">
    <dxf/><dxf/><dxf/><dxf/><dxf/>
    <dxf><font><color rgb="FFFF0000"/></font></dxf>
    <dxf/>
    <dxf><font><b/></font><alignment horizontal="center"/></dxf>
    <dxf/>
    <dxf><fill><patternFill><bgColor rgb="FF123456"/></patternFill></fill></dxf>
  </dxfs>
</styleSheet>"""

  private val dxfWorksheetXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="2"><c r="A2"><v>2</v></c></row>
    <row r="3"><c r="A3"><v>3</v></c></row>
    <row r="4"><c r="A4"><v>4</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A1:A4">
    <cfRule type="expression" dxfId="5" priority="1"><formula>$$A1&gt;0</formula></cfRule>
    <cfRule type="cellIs" dxfId="7" priority="2" operator="greaterThan"><formula>1</formula></cfRule>
    <cfRule type="aboveAverage" dxfId="9" priority="3"/>
  </conditionalFormatting>
</worksheet>"""

  test("fresh write re-registers preserved CF rules' dxfs and renumbers dxfId refs (GH-471b)") {
    val wb = readBytes(buildXlsx(dxfStylesXml, dxfWorksheetXml))
    val rules = wb
      .sheets(0)
      .conditionalFormats
      .collect { case ConditionalFormat.Rules(_, rs, _) =>
        rs
      }
      .flatten
    // Sanity: the pinned typed/Preserved split — expression typed with its dxf,
    // cellIs (untypeable dxf: alignment) and aboveAverage (unmodeled type) ride Preserved
    assertEquals(rules.count(_.isInstanceOf[CfRule.Expression]), 1)
    assertEquals(rules.count(_.isInstanceOf[CfRule.Preserved]), 2)
    assert(rules.collect { case CfRule.Expression(_, dxf, _, _) => dxf }.head.isDefined)

    val out = writeFresh(wb, "dxfs")

    val styles = parseEntry(out, "xl/styles.xml")
    val dxfChildren: Vector[scala.xml.Elem] =
      (styles \ "dxfs" \ "dxf").collect { case e: scala.xml.Elem => e }.toVector
    val declaredCount = (styles \ "dxfs").headOption
      .flatMap(_.attribute("count"))
      .map(_.text.toInt)
      .getOrElse(0)
    assertEquals(declaredCount, dxfChildren.size)

    val sheetXml = parseEntry(out, "xl/worksheets/sheet1.xml")
    val cfRules = (sheetXml \ "conditionalFormatting" \ "cfRule").collect {
      case e: scala.xml.Elem => e
    }.toVector
    assertEquals(cfRules.size, 3)
    val refs: Map[String, Int] = cfRules.flatMap { r =>
      val tpe = r.attribute("type").map(_.text).getOrElse("")
      r.attribute("dxfId").map(a => tpe -> a.text.toInt)
    }.toMap

    // Every shipped dxfId ref resolves into the shipped table
    refs.foreach { case (tpe, id) =>
      assert(
        id >= 0 && id < dxfChildren.size,
        s"cfRule type=$tpe references dxfId $id but <dxfs> has ${dxfChildren.size} entries"
      )
    }
    // ... and points at the SOURCE dxf content it referenced
    val cellIsDxf = dxfChildren(refs("cellIs")).toString
    assert(cellIsDxf.contains("alignment"), s"cellIs dxf content lost: $cellIsDxf")
    val aboveAvgDxf = dxfChildren(refs("aboveAverage")).toString
    assert(aboveAvgDxf.contains("FF123456"), s"aboveAverage dxf content lost: $aboveAvgDxf")

    // Behavioral proof: reopened, the preserved rules' dxf refs still resolve (the cellIs rule
    // still rides Preserved — its dxf stays untypeable — but its ref is now in range) and the
    // typed expression rule keeps its dxf
    val rules2 = reread(out)
      .sheets(0)
      .conditionalFormats
      .collect { case ConditionalFormat.Rules(_, rs, _) =>
        rs
      }
      .flatten
    assert(rules2.collect { case CfRule.Expression(_, dxf, _, _) => dxf }.head.isDefined)
    assertEquals(rules2.count(_.isInstanceOf[CfRule.Preserved]), 2)
  }
