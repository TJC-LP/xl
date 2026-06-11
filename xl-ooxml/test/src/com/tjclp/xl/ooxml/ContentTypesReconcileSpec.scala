package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import scala.jdk.CollectionConverters.*

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * GH-314: a metadata-modified surgical write used to rebuild [Content_Types].xml from
 * `ContentTypes.minimal` plus an allowlist of re-adds, silently dropping preserved overrides for
 * exotic parts (pivots, custom XML, macro payloads) whose bytes still ride the verbatim copy loop.
 * A part shipping without a content-type registration corrupts the package.
 *
 * The fix reconciles from the PRESERVED part: preserved defaults/overrides are merged with the
 * model-required entries (model wins on conflict, except the /xl/workbook.xml dialect — see
 * `ContentTypes.reconcile`).
 */
class ContentTypesReconcileSpec extends FunSuite:

  private val customXmlOverride = "/customXml/item1.xml"
  private val macroWorkbookCt = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
  private val vbaProjectCt = "application/vnd.ms-office.vbaProject"

  // ========== exotic override survives a metadata-only write (GH-314 repro) ==========

  test("GH-314: customXml override survives a metadata-only surgical write") {
    val source = createExoticFixture()
    val wb = XlsxReader.read(source).fold(err => fail(s"Read failed: $err"), identity)

    // Metadata-only change: no sheet content touched, workbook.xml regenerated
    val modified = wb.withDefinedName("Answer", "Sheet1!$A$1")

    val output = Files.createTempFile("gh314-metadata", ".xlsx")
    XlsxWriter.write(modified, output).fold(err => fail(s"Write failed: $err"), identity)

    // The exotic part itself rides the copy loop
    assert(entryExists(output, "customXml/item1.xml"), "preserved customXml part should ship")

    // ...and must stay registered in [Content_Types].xml
    val ct = new String(readEntry(output, "[Content_Types].xml"), "UTF-8")
    assert(
      ct.contains(s"PartName=\"$customXmlOverride\""),
      s"metadata-modified write dropped the preserved customXml override. CT:\n$ct"
    )

    // Model-required entries still present (model wins on its own parts)
    assert(ct.contains("PartName=\"/xl/worksheets/sheet1.xml\""), s"worksheet override lost:\n$ct")
    assert(ct.contains("PartName=\"/xl/styles.xml\""), s"styles override lost:\n$ct")

    // And the defined name actually landed
    val reloaded = XlsxReader.read(output).fold(err => fail(s"Reload failed: $err"), identity)
    assert(
      reloaded.metadata.definedNames.exists(_.name == "Answer"),
      "defined name missing after metadata write"
    )

    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== exotic override survives a cell-edit write (preserved branch pin) ==========

  test("GH-314: customXml override survives a cell-edit surgical write") {
    val source = createExoticFixture()
    val wb = XlsxReader.read(source).fold(err => fail(s"Read failed: $err"), identity)

    val modified = wb("Sheet1")
      .map(sheet => wb.put(sheet.put(ref"D1" -> 42)))
      .fold(err => fail(s"Modify failed: $err"), identity)

    val output = Files.createTempFile("gh314-celledit", ".xlsx")
    XlsxWriter.write(modified, output).fold(err => fail(s"Write failed: $err"), identity)

    assert(entryExists(output, "customXml/item1.xml"), "preserved customXml part should ship")
    val ct = new String(readEntry(output, "[Content_Types].xml"), "UTF-8")
    assert(
      ct.contains(s"PartName=\"$customXmlOverride\""),
      s"cell-edit write dropped the preserved customXml override. CT:\n$ct"
    )

    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== macro dialect: model must not clobber /xl/workbook.xml content type ==========

  test("GH-314: macro-enabled workbook dialect survives a metadata-only surgical write") {
    val source = createExoticFixture(macroEnabled = true)
    val wb = XlsxReader.read(source).fold(err => fail(s"Read failed: $err"), identity)

    val modified = wb.withDefinedName("Answer", "Sheet1!$A$1")

    val output = Files.createTempFile("gh314-macro", ".xlsm")
    XlsxWriter.write(modified, output).fold(err => fail(s"Write failed: $err"), identity)

    assert(entryExists(output, "xl/vbaProject.bin"), "preserved vbaProject.bin should ship")
    val ct = new String(readEntry(output, "[Content_Types].xml"), "UTF-8")
    assert(
      ct.contains(macroWorkbookCt),
      s"metadata write rewrote the macro-enabled workbook content type. CT:\n$ct"
    )
    assert(
      ct.contains(s"ContentType=\"$vbaProjectCt\""),
      s"metadata write dropped the vbaProject bin Default. CT:\n$ct"
    )

    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== GH-321: comment overrides follow EMITTED paths on metadata-modified writes ==========

  test("GH-321: metadata-only write of a foreign-dialect comments file has no dangling overrides") {
    val source = TestFixtures.copyToTemp("comments-hyperlinks.xlsx")
    val wb = XlsxReader.read(source).fold(err => fail(s"Read failed: $err"), identity)

    // Metadata-only edit: no sheet content touched → the regenerate-from-minimal CT branch
    val modified = wb.withDefinedName("Answer", "Notes!$A$1")

    val output = Files.createTempFile("gh321-metadata", ".xlsx")
    XlsxWriter.write(modified, output).fold(err => fail(s"Write failed: $err"), identity)

    val entries = zipEntryNames(output)
    // The openpyxl-dialect comment + VML parts ship at their SOURCE paths (identity-mapped,
    // GH-315/GH-292), never at the legacy index paths
    assert(
      entries.contains("xl/comments/comment1.xml"),
      s"foreign-dialect comment part missing from output: $entries"
    )
    assert(
      entries.contains("xl/drawings/commentsDrawing1.vml"),
      s"foreign-dialect VML part missing from output: $entries"
    )
    assert(!entries.contains("xl/comments1.xml"), "comment part emitted at the legacy index path")

    val ct = parseContentTypes(output)
    // The shipped comment part is registered at its ACTUAL path...
    assertEquals(
      ct.overrides.get("/xl/comments/comment1.xml"),
      Some(XmlUtil.ctComments),
      s"shipped comment part unregistered. Overrides: ${ct.overrides}"
    )
    // ...the shipped VML part is covered by the vml extension Default...
    assert(
      ct.defaults.keysIterator.exists(_.equalsIgnoreCase("vml")),
      s"vml Default missing. Defaults: ${ct.defaults}"
    )
    // ...and NO override points at a part that does not ship (index-derived registration would
    // leave /xl/comments1.xml + /xl/drawings/vmlDrawing1.vml dangling here)
    val dangling = ct.overrides.keys.filterNot(p => entries.contains(p.stripPrefix("/")))
    assert(dangling.isEmpty, s"overrides point at parts that do not ship: ${dangling.mkString(", ")}")

    // The comment itself round-trips
    val reloaded = XlsxReader.read(output).fold(err => fail(s"Reload failed: $err"), identity)
    val sheet = reloaded("Notes").fold(err => fail(s"Sheet missing: $err"), identity)
    assert(sheet.comments.contains(ref"A1"), "comment lost on metadata-only write")

    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ========== pure reconcile law: preserved ∪ model, model wins except workbook dialect ==========

  test("ContentTypes.reconcile: union with model-wins, preserved workbook dialect kept") {
    val preserved = ContentTypes(
      defaults = Map("rels" -> XmlUtil.ctRelationships, "bin" -> vbaProjectCt),
      overrides = Map(
        "/xl/workbook.xml" -> macroWorkbookCt,
        customXmlOverride -> "application/xml",
        "/xl/styles.xml" -> "stale/preserved-styles-ct"
      )
    )
    val model = ContentTypes.minimal(hasStyles = true, hasSharedStrings = false, sheetCount = 2)

    val merged = ContentTypes.reconcile(preserved, model)

    // Preserved-only entries survive
    assertEquals(merged.defaults.get("bin"), Some(vbaProjectCt))
    assertEquals(merged.overrides.get(customXmlOverride), Some("application/xml"))
    // Model wins on conflicts for parts it regenerates
    assertEquals(merged.overrides.get("/xl/styles.xml"), Some(XmlUtil.ctStyles))
    // Model-required entries present
    assertEquals(
      merged.overrides.get("/xl/worksheets/sheet2.xml"),
      Some(XmlUtil.ctWorksheet)
    )
    // ...except the workbook dialect, which the model does not track
    assertEquals(merged.overrides.get("/xl/workbook.xml"), Some(macroWorkbookCt))
  }

  // ========== helpers ==========

  private def entryExists(path: Path, entryName: String): Boolean =
    val zip = new ZipFile(path.toFile)
    try zip.getEntry(entryName) != null
    finally zip.close()

  private def zipEntryNames(path: Path): Set[String] =
    val zip = new ZipFile(path.toFile)
    try zip.entries().asScala.map(_.getName).toSet
    finally zip.close()

  private def parseContentTypes(path: Path): ContentTypes =
    val xml = scala.xml.XML.loadString(new String(readEntry(path, "[Content_Types].xml"), "UTF-8"))
    ContentTypes.fromXml(xml).fold(err => fail(s"Content types parse failed: $err"), identity)

  private def readEntry(path: Path, entryName: String): Array[Byte] =
    val zip = new ZipFile(path.toFile)
    try
      val entry = zip.getEntry(entryName)
      assert(entry != null, s"$entryName missing from $path")
      val is = zip.getInputStream(entry)
      try is.readAllBytes()
      finally is.close()
    finally zip.close()

  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    out.putNextEntry(new ZipEntry(name))
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()

  /**
   * Excel-shaped raw fixture carrying an exotic part: customXml/item1.xml registered via an
   * Override the writer's domain model knows nothing about. With `macroEnabled` the workbook main
   * part uses the macro dialect content type and a vbaProject.bin payload rides along (bin
   * Default), mirroring a real .xlsm.
   */
  private def createExoticFixture(macroEnabled: Boolean = false): Path =
    val path = Files.createTempFile("gh314-fixture", if macroEnabled then ".xlsm" else ".xlsx")
    val workbookCt =
      if macroEnabled then macroWorkbookCt
      else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    val macroDefault =
      if macroEnabled then s"""  <Default Extension="bin" ContentType="$vbaProjectCt"/>\n"""
      else ""
    val out = new ZipOutputStream(Files.newOutputStream(path))
    out.setLevel(1)
    try
      writeEntry(
        out,
        "[Content_Types].xml",
        s"""<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
$macroDefault  <Override PartName="/xl/workbook.xml" ContentType="$workbookCt"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="$customXmlOverride" ContentType="application/xml"/>
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
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
      )
      writeEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
      )
      writeEntry(
        out,
        "xl/styles.xml",
        """<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"""
      )
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"""
      )
      writeEntry(
        out,
        "customXml/item1.xml",
        """<?xml version="1.0"?><root xmlns="urn:example:custom"><value>42</value></root>"""
      )
      if macroEnabled then
        out.putNextEntry(new ZipEntry("xl/vbaProject.bin"))
        out.write(Array[Byte](0x01, 0x02, 0x03, 0x04))
        out.closeEntry()
    finally out.close()
    path
