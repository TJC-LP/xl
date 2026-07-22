package com.tjclp.xl.ooxml.lint

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{XlsxReader, XlsxWriter}
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import munit.FunSuite

/**
 * Structural lint (GH-397): CT_Workbook / CT_Worksheet child-order violations and r:id references
 * that don't resolve in the paired .rels or resolve to a wrong-typed part — the Excel-repair
 * classes lenient readers accept silently.
 *
 * Field incident pinned here: a pipeline zip-patched `<externalReferences>` AFTER `<extLst>`
 * (CT_Workbook wants it between `<sheets>` and `<definedNames>`) with dangling rIds; Excel showed
 * the repair dialog while xl 0.12.6 read the file silently.
 *
 * Also pins the acceptance invariant: xl's OWN output must lint clean, including round-trips of
 * workbooks carrying preserved externalReferences/workbookProtection (previously re-emitted after
 * extLst — the very violation this lint flags).
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class WorkbookLintSpec extends FunSuite:

  private val nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  private val nsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  // ===== Base package parts (clean minimal workbook) =====

  private val contentTypesXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"""

  private val rootRelsXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

  private val workbookXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="$nsMain" xmlns:r="$nsRel">
  <fileVersion appName="xl"/>
  <workbookPr/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames><definedName name="MyName">Sheet1!$$A$$1</definedName></definedNames>
  <calcPr calcId="191029"/>
</workbook>"""

  private val workbookRelsXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""

  private val worksheetXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
    </row>
  </sheetData>
</worksheet>"""

  private val stylesXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="$nsMain">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

  private val sharedStringsXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="$nsMain" count="0" uniqueCount="0"/>"""

  private val baseParts: Map[String, String] = Map(
    "[Content_Types].xml" -> contentTypesXml,
    "_rels/.rels" -> rootRelsXml,
    "xl/workbook.xml" -> workbookXml,
    "xl/_rels/workbook.xml.rels" -> workbookRelsXml,
    "xl/worksheets/sheet1.xml" -> worksheetXml,
    "xl/styles.xml" -> stylesXml,
    "xl/sharedStrings.xml" -> sharedStringsXml
  )

  // ===== External-reference variant (the field-incident structure, in VALID form) =====

  private val externalLinkXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<externalLink xmlns="$nsMain" xmlns:r="$nsRel">
  <externalBook r:id="rId1">
    <sheetNames><sheetName val="Extern"/></sheetNames>
  </externalBook>
</externalLink>"""

  private val externalLinkRelsXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="other.xlsx" TargetMode="External"/>
</Relationships>"""

  private val contentTypesWithExternalXml = contentTypesXml.replace(
    "</Types>",
    """  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>"""
  )

  private val workbookRelsWithExternalXml = workbookRelsXml.replace(
    "</Relationships>",
    """  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>"""
  )

  /** Valid workbook.xml carrying every preserved-element class from the field incident. */
  private val workbookWithExternalXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="$nsMain" xmlns:r="$nsRel">
  <fileVersion appName="xl"/>
  <workbookPr/>
  <workbookProtection lockStructure="1"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <externalReferences><externalReference r:id="rId2"/></externalReferences>
  <definedNames><definedName name="MyName">Sheet1!$$A$$1</definedName></definedNames>
  <calcPr calcId="191029"/>
</workbook>"""

  private val externalParts: Map[String, String] = baseParts ++ Map(
    "[Content_Types].xml" -> contentTypesWithExternalXml,
    "xl/workbook.xml" -> workbookWithExternalXml,
    "xl/_rels/workbook.xml.rels" -> workbookRelsWithExternalXml,
    "xl/externalLinks/externalLink1.xml" -> externalLinkXml,
    "xl/externalLinks/_rels/externalLink1.xml.rels" -> externalLinkRelsXml
  )

  // ===== Helpers =====

  private def zipBytes(parts: Map[String, String]): Array[Byte] =
    val baos = ByteArrayOutputStream()
    val zos = ZipOutputStream(baos)
    parts.foreach { case (name, content) =>
      zos.putNextEntry(ZipEntry(name))
      zos.write(content.getBytes(StandardCharsets.UTF_8))
      zos.closeEntry()
    }
    zos.close()
    baos.toByteArray

  private def lintOf(parts: Map[String, String]): Vector[Finding] =
    WorkbookLint
      .lintBytes(zipBytes(parts))
      .fold(err => fail(s"lint must not error on a parseable package: $err"), identity)

  private def tempFile(bytes: Array[Byte]): Path =
    val path = Files.createTempFile("lint-spec", ".xlsx")
    Files.write(path, bytes)
    path

  private def readEntry(zipPath: Path, entry: String): String =
    val zip = new ZipFile(zipPath.toFile)
    try
      val e = Option(zip.getEntry(entry)).getOrElse(fail(s"missing $entry in ${zipPath}"))
      new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
    finally zip.close()

  // ===== Clean packages =====

  test("clean minimal workbook has no findings") {
    assertEquals(lintOf(baseParts), Vector.empty[Finding])
  }

  test("clean workbook with externalReferences in schema position has no findings") {
    assertEquals(lintOf(externalParts), Vector.empty[Finding])
  }

  test("lint(path) agrees with lintBytes for the same package") {
    val bytes = zipBytes(externalParts)
    val path = tempFile(bytes)
    try
      assertEquals(WorkbookLint.lint(path), WorkbookLint.lintBytes(bytes))
      assertEquals(WorkbookLint.lint(path), Right(Vector.empty[Finding]))
    finally Files.deleteIfExists(path)
  }

  // ===== CT_Workbook child order =====

  test("GH-397 repro: externalReferences after extLst is flagged as ChildOrder") {
    // The field incident: a pipeline zip-patched <externalReferences> after <extLst>.
    val corrupted = workbookWithExternalXml
      .replace("<externalReferences><externalReference r:id=\"rId2\"/></externalReferences>\n", "")
      .replace(
        "</workbook>",
        "<extLst/><externalReferences><externalReference r:id=\"rId2\"/></externalReferences></workbook>"
      )
    val findings = lintOf(externalParts + ("xl/workbook.xml" -> corrupted))
    val orderFindings = findings.filter(_.category == LintCategory.ChildOrder)
    assertEquals(orderFindings.size, 1, s"expected exactly one ChildOrder finding, got: $findings")
    val f = orderFindings.head
    assertEquals(f.part, "xl/workbook.xml")
    assert(f.locator.contains("externalReferences"), s"locator should name the element: $f")
    assert(f.message.contains("extLst"), s"message should name the misordered pair: $f")
  }

  test("workbookProtection after calcPr is flagged as ChildOrder") {
    val corrupted = workbookXml.replace(
      "<calcPr calcId=\"191029\"/>",
      "<calcPr calcId=\"191029\"/><workbookProtection lockStructure=\"1\"/>"
    )
    val findings = lintOf(baseParts + ("xl/workbook.xml" -> corrupted))
    assertEquals(findings.map(_.category), Vector(LintCategory.ChildOrder))
    assert(findings.head.message.contains("workbookProtection"), findings.head.toString)
  }

  test("mc:AlternateContent and xr:revisionPtr are position-transparent (no findings)") {
    val withMc = workbookXml.replace(
      "<workbookPr/>",
      """<workbookPr/>
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"/></mc:AlternateContent>
  <xr:revisionPtr revIDLastSave="0" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"/>"""
    )
    assertEquals(lintOf(baseParts + ("xl/workbook.xml" -> withMc)), Vector.empty[Finding])
  }

  test("unknown main-namespace child is skipped by the order check") {
    val withUnknown = workbookXml.replace(
      "<definedNames>",
      "<futureThing/><definedNames>"
    )
    assertEquals(lintOf(baseParts + ("xl/workbook.xml" -> withUnknown)), Vector.empty[Finding])
  }

  // ===== Workbook-level r:id resolution =====

  test("sheet r:id with no matching relationship is flagged as UnresolvedRelId") {
    val corrupted = workbookXml.replace("r:id=\"rId1\"", "r:id=\"rId99\"")
    val findings = lintOf(baseParts + ("xl/workbook.xml" -> corrupted))
    assertEquals(findings.map(_.category), Vector(LintCategory.UnresolvedRelId))
    val f = findings.head
    assertEquals(f.part, "xl/workbook.xml")
    assert(f.locator.contains("rId99"), s"locator should carry the dangling id: $f")
    assert(f.locator.contains("Sheet1"), s"locator should carry the sheet name: $f")
  }

  test("sheet without r:id attribute is flagged as UnresolvedRelId") {
    val corrupted = workbookXml.replace(" r:id=\"rId1\"", "")
    val findings = lintOf(baseParts + ("xl/workbook.xml" -> corrupted))
    assertEquals(findings.map(_.category), Vector(LintCategory.UnresolvedRelId))
    assert(findings.head.message.contains("r:id"), findings.head.toString)
  }

  test("externalReference r:id resolving to a styles-typed rel is flagged as WrongRelType") {
    // The renumbered-rels incident class: the id exists but points at a different-typed part.
    val corrupted = workbookWithExternalXml.replace(
      "<externalReference r:id=\"rId2\"/>",
      "<externalReference r:id=\"rId3\"/>"
    )
    val findings = lintOf(externalParts + ("xl/workbook.xml" -> corrupted))
    assertEquals(findings.map(_.category), Vector(LintCategory.WrongRelType))
    val f = findings.head
    assert(f.message.contains("styles"), s"message should name the actual type: $f")
    assert(f.message.contains("externalLink"), s"message should name the expected type: $f")
  }

  test("pivotCache r:id with no matching relationship is flagged as UnresolvedRelId") {
    // pivotCaches injected after calcPr — its schema slot, so the only finding is the rel id
    val withPivot = workbookXml.replace(
      "<calcPr calcId=\"191029\"/>",
      "<calcPr calcId=\"191029\"/><pivotCaches><pivotCache cacheId=\"1\" r:id=\"rId77\"/></pivotCaches>"
    )
    val findings = lintOf(baseParts + ("xl/workbook.xml" -> withPivot))
    assertEquals(findings.map(_.category), Vector(LintCategory.UnresolvedRelId))
    assert(findings.head.locator.contains("rId77"), findings.head.toString)
  }

  test("worksheet-typed rel whose target part is missing is flagged as MissingPart") {
    val corrupted = workbookRelsXml.replace("worksheets/sheet1.xml", "worksheets/sheet9.xml")
    val findings = lintOf(baseParts + ("xl/_rels/workbook.xml.rels" -> corrupted))
    assertEquals(findings.map(_.category), Vector(LintCategory.MissingPart))
    assert(findings.head.message.contains("sheet9.xml"), findings.head.toString)
  }

  test("worksheet-typed rel pointing at non-worksheet content is flagged as WrongRelType") {
    // Type says worksheet but the part's root element is <styleSheet> — the zip-patch renumber class.
    val corrupted = workbookRelsXml.replace("worksheets/sheet1.xml", "styles.xml")
    val findings = lintOf(baseParts + ("xl/_rels/workbook.xml.rels" -> corrupted))
    assert(
      findings.exists(f =>
        f.category == LintCategory.WrongRelType && f.message.contains("styleSheet")
      ),
      s"expected WrongRelType naming the actual root, got: $findings"
    )
  }

  // ===== CT_Worksheet child order =====

  test("worksheet mergeCells before sheetData is flagged as ChildOrder on the sheet part") {
    val corrupted =
      s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain">
  <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
  <sheetData/>
</worksheet>"""
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> corrupted))
    assertEquals(findings.map(_.category), Vector(LintCategory.ChildOrder))
    assertEquals(findings.head.part, "xl/worksheets/sheet1.xml")
    assert(findings.head.message.contains("mergeCells"), findings.head.toString)
  }

  // ===== Worksheet-level r:id resolution =====

  test("hyperlink r:id missing from sheet rels is flagged as UnresolvedRelId on the sheet part") {
    val corrupted =
      s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain" xmlns:r="$nsRel">
  <sheetData/>
  <hyperlinks><hyperlink ref="A1" r:id="rId9"/></hyperlinks>
</worksheet>"""
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> corrupted))
    assertEquals(findings.map(_.category), Vector(LintCategory.UnresolvedRelId))
    assertEquals(findings.head.part, "xl/worksheets/sheet1.xml")
    assert(findings.head.locator.contains("rId9"), findings.head.toString)
  }

  test("internal hyperlink without r:id is not flagged") {
    val internal =
      s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain">
  <sheetData/>
  <hyperlinks><hyperlink ref="A1" location="Sheet1!B2"/></hyperlinks>
</worksheet>"""
    assertEquals(
      lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> internal)),
      Vector.empty[Finding]
    )
  }

  test("drawing r:id resolving to a table-typed rel is flagged as WrongRelType") {
    val sheet =
      s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain" xmlns:r="$nsRel">
  <sheetData/>
  <drawing r:id="rId1"/>
</worksheet>"""
    val sheetRels =
      """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>"""
    val findings = lintOf(
      baseParts +
        ("xl/worksheets/sheet1.xml" -> sheet) +
        ("xl/worksheets/_rels/sheet1.xml.rels" -> sheetRels)
    )
    assertEquals(findings.map(_.category), Vector(LintCategory.WrongRelType))
    assert(findings.head.message.contains("table"), findings.head.toString)
  }

  // ===== Totality =====

  test("garbage bytes yield Left, never throw") {
    assert(WorkbookLint.lintBytes("not a zip at all".getBytes(StandardCharsets.UTF_8)).isLeft)
  }

  test("package without workbook.xml yields Left") {
    assert(WorkbookLint.lintBytes(zipBytes(baseParts - "xl/workbook.xml")).isLeft)
  }

  test("malformed workbook.xml yields Left (diagnosable, exit-2 class)") {
    val bad = baseParts + ("xl/workbook.xml" -> "<workbook><sheets></workbook>")
    assert(WorkbookLint.lintBytes(zipBytes(bad)).isLeft)
  }

  // ===== Acceptance invariant: xl's own output lints clean (GH-397 adjacent writer bug) =====

  private def roundTripLintsClean(backend: XmlBackend): Unit =
    val source = tempFile(zipBytes(externalParts))
    val output = Files.createTempFile("lint-roundtrip", ".xlsx")
    try
      val modified = for
        wb <- XlsxReader.read(source)
        sheet <- wb("Sheet1")
      yield wb.put(sheet.put(ref"A1" -> "Modified"))
      val wb = modified.fold(err => fail(s"read/modify failed: $err"), identity)
      XlsxWriter
        .writeWith(wb, output, WriterConfig(backend = backend))
        .fold(err => fail(s"write failed: $err"), identity)

      // Preservation guard: the externalReferences/workbookProtection elements must survive ...
      val wbXml = readEntry(output, "xl/workbook.xml")
      assert(wbXml.contains("<externalReferences>"), s"externalReferences dropped:\n$wbXml")
      assert(wbXml.contains("<workbookProtection"), s"workbookProtection dropped:\n$wbXml")
      // ... in their canonical slots (externalReferences between sheets and definedNames).
      assert(
        wbXml.indexOf("<externalReferences>") < wbXml.indexOf("<definedNames>"),
        s"externalReferences must precede definedNames:\n$wbXml"
      )

      // The lint must agree: zero findings on xl's own output.
      val findings =
        WorkbookLint.lint(output).fold(err => fail(s"lint errored on xl output: $err"), identity)
      assertEquals(findings, Vector.empty[Finding], s"xl's own output must lint clean:\n$wbXml")
    finally
      Files.deleteIfExists(source)
      Files.deleteIfExists(output)

  test("xl round-trip with externalReferences + workbookProtection lints clean (ScalaXml)") {
    roundTripLintsClean(XmlBackend.ScalaXml)
  }

  test("xl round-trip with externalReferences + workbookProtection lints clean (SaxStax)") {
    roundTripLintsClean(XmlBackend.SaxStax)
  }

  test("GH-412: bytes-read round-trip with externalReferences lints clean too") {
    // Same acceptance as the path-based round-trip, through readFromBytes: the in-memory
    // SourceContext must feed the identical preservation machinery.
    val output = Files.createTempFile("lint-roundtrip-bytes", ".xlsx")
    try
      val modified = for
        wb <- XlsxReader.readFromBytes(zipBytes(externalParts))
        sheet <- wb("Sheet1")
      yield wb.put(sheet.put(ref"A1" -> "Modified"))
      val wb = modified.fold(err => fail(s"read/modify failed: $err"), identity)
      XlsxWriter.write(wb, output).fold(err => fail(s"write failed: $err"), identity)

      val wbXml = readEntry(output, "xl/workbook.xml")
      assert(wbXml.contains("<externalReferences>"), s"externalReferences dropped:\n$wbXml")
      assert(wbXml.contains("<workbookProtection"), s"workbookProtection dropped:\n$wbXml")

      val findings =
        WorkbookLint.lint(output).fold(err => fail(s"lint errored on xl output: $err"), identity)
      assertEquals(
        findings,
        Vector.empty[Finding],
        s"xl's bytes-read output must lint clean:\n$wbXml"
      )
    finally Files.deleteIfExists(output)
  }

  test("fresh scratch workbook written by xl lints clean") {
    val output = Files.createTempFile("lint-fresh", ".xlsx")
    try
      val sheet = Sheet("Data").put(ref"A1" -> "hello")
      XlsxWriter
        .write(Workbook(Vector(sheet)), output)
        .fold(err => fail(s"write failed: $err"), identity)
      assertEquals(WorkbookLint.lint(output), Right(Vector.empty[Finding]))
    finally Files.deleteIfExists(output)
  }

  // ===== GH-413 (1): chartsheet / dialogsheet child-order tables =====

  /** Base package plus a second sheet of the given kind wired through workbook rels + CT. */
  private def withSecondSheet(
    relType: String,
    target: String,
    contentType: String,
    sheetXml: String
  ): Map[String, String] =
    baseParts ++ Map(
      "[Content_Types].xml" -> contentTypesXml.replace(
        "</Types>",
        s"""  <Override PartName="/xl/$target" ContentType="$contentType"/>\n</Types>"""
      ),
      "xl/workbook.xml" -> workbookXml.replace(
        "</sheets>",
        "  <sheet name=\"Extra\" sheetId=\"2\" r:id=\"rId5\"/>\n  </sheets>"
      ),
      "xl/_rels/workbook.xml.rels" -> workbookRelsXml.replace(
        "</Relationships>",
        s"""  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/$relType" Target="$target"/>\n</Relationships>"""
      ),
      s"xl/$target" -> sheetXml
    )

  private val chartsheetXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<chartsheet xmlns="$nsMain" xmlns:r="$nsRel">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</chartsheet>"""

  private val chartsheetCt =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"

  private val chartsheetParts: Map[String, String] =
    withSecondSheet("chartsheet", "chartsheets/sheet1.xml", chartsheetCt, chartsheetXml)

  private val misorderedChartsheetXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<chartsheet xmlns="$nsMain">
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
</chartsheet>"""

  private val dialogsheetXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<dialogsheet xmlns="$nsMain">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
</dialogsheet>"""

  private val dialogsheetCt =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml"

  private val dialogsheetParts: Map[String, String] =
    withSecondSheet("dialogsheet", "dialogsheets/sheet1.xml", dialogsheetCt, dialogsheetXml)

  private val misorderedDialogsheetXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<dialogsheet xmlns="$nsMain">
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
</dialogsheet>"""

  test("clean chartsheet part has no findings") {
    assertEquals(lintOf(chartsheetParts), Vector.empty[Finding])
  }

  test("GH-413: chartsheet sheetViews after pageMargins is flagged as ChildOrder") {
    val findings =
      lintOf(chartsheetParts + ("xl/chartsheets/sheet1.xml" -> misorderedChartsheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.ChildOrder))
    assertEquals(findings.head.part, "xl/chartsheets/sheet1.xml")
    assert(findings.head.message.contains("CT_Chartsheet"), findings.head.toString)
    assert(findings.head.message.contains("sheetViews"), findings.head.toString)
  }

  test("GH-413: chartsheet drawing r:id with no sibling rels is flagged as UnresolvedRelId") {
    val withDrawing =
      chartsheetXml.replace("</chartsheet>", "<drawing r:id=\"rId1\"/></chartsheet>")
    val findings = lintOf(chartsheetParts + ("xl/chartsheets/sheet1.xml" -> withDrawing))
    assertEquals(findings.map(_.category), Vector(LintCategory.UnresolvedRelId))
    assertEquals(findings.head.part, "xl/chartsheets/sheet1.xml")
    assert(findings.head.locator.contains("rId1"), findings.head.toString)
  }

  test("clean dialogsheet part has no findings") {
    assertEquals(lintOf(dialogsheetParts), Vector.empty[Finding])
  }

  test("GH-413: dialogsheet sheetViews after sheetFormatPr is flagged as ChildOrder") {
    val findings =
      lintOf(dialogsheetParts + ("xl/dialogsheets/sheet1.xml" -> misorderedDialogsheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.ChildOrder))
    assertEquals(findings.head.part, "xl/dialogsheets/sheet1.xml")
    assert(findings.head.message.contains("CT_Dialogsheet"), findings.head.toString)
  }

  // ===== GH-413 (2): externalLink part's own <externalBook r:id> chain =====

  private val danglingExternalBookXml = externalLinkXml.replace("r:id=\"rId1\"", "r:id=\"rId9\"")

  private val wrongTypeExternalBookRelsXml = externalLinkRelsXml.replace(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  )

  test("GH-413: externalBook r:id not resolving in the externalLink sibling rels is flagged") {
    val findings =
      lintOf(externalParts + ("xl/externalLinks/externalLink1.xml" -> danglingExternalBookXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.UnresolvedRelId))
    assertEquals(findings.head.part, "xl/externalLinks/externalLink1.xml")
    assert(findings.head.locator.contains("externalBook"), findings.head.toString)
    assert(findings.head.locator.contains("rId9"), findings.head.toString)
  }

  test("GH-413: externalBook r:id resolving to a non-externalLinkPath rel is WrongRelType") {
    val findings = lintOf(
      externalParts +
        ("xl/externalLinks/_rels/externalLink1.xml.rels" -> wrongTypeExternalBookRelsXml)
    )
    assertEquals(findings.map(_.category), Vector(LintCategory.WrongRelType))
    assertEquals(findings.head.part, "xl/externalLinks/externalLink1.xml")
    assert(findings.head.message.contains("image"), findings.head.toString)
    assert(findings.head.message.contains("externalLinkPath"), findings.head.toString)
  }

  // ===== GH-413 (3): [Content_Types].xml registration =====

  private val stylesOverrideLine =
    "  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n"

  private val xmlDefaultLine =
    "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"

  private val unregisteredStylesCt: String =
    val ct = contentTypesXml.replace(stylesOverrideLine, "").replace(xmlDefaultLine, "")
    assert(!ct.contains("/xl/styles.xml") && !ct.contains("Extension=\"xml\""), ct)
    ct

  test("GH-413: present-and-referenced part with no Override and no Default is flagged") {
    val findings = lintOf(baseParts + ("[Content_Types].xml" -> unregisteredStylesCt))
    assertEquals(findings.map(_.category), Vector(LintCategory.MissingContentType))
    assertEquals(findings.head.part, "[Content_Types].xml")
    assert(findings.head.locator.contains("/xl/styles.xml"), findings.head.toString)
    assert(findings.head.message.contains("xl/styles.xml"), findings.head.toString)
  }

  test("GH-413: part covered only by an extension Default is treated as registered") {
    // xml Default kept: styles.xml is still registered (content-type CORRECTNESS is out of scope)
    val ct = contentTypesXml.replace(stylesOverrideLine, "")
    assertEquals(lintOf(baseParts + ("[Content_Types].xml" -> ct)), Vector.empty[Finding])
  }

  test("GH-413: package without [Content_Types].xml gets a single MissingContentType finding") {
    val findings = lintOf(baseParts - "[Content_Types].xml")
    assertEquals(findings.map(_.category), Vector(LintCategory.MissingContentType))
    assertEquals(findings.head.part, "[Content_Types].xml")
  }

  // ===== GH-428 class: sqref/ref/dimension tokens past row 1048576 / column XFD =====

  private def worksheetWith(body: String): String =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain">
  $body
</worksheet>"""

  private val overMaxRowSheetXml = worksheetWith(
    """<sheetData/>
  <conditionalFormatting sqref="A1:XFD1048578"><cfRule type="expression" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>"""
  )

  private val overMaxColSheetXml = worksheetWith(
    """<sheetData/>
  <mergeCells count="1"><mergeCell ref="A1:XFE10"/></mergeCells>"""
  )

  private val overMaxDimensionSheetXml = worksheetWith(
    """<dimension ref="A1:B1048577"/>
  <sheetData/>"""
  )

  private val atMaxSheetXml = worksheetWith(
    """<dimension ref="A1:XFD1048576"/>
  <sheetData/>
  <mergeCells count="1"><mergeCell ref="XFD1048576"/></mergeCells>
  <conditionalFormatting sqref="A1:XFD1048576"><cfRule type="expression" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>"""
  )

  test("GH-428 class: conditionalFormatting sqref past row 1048576 is flagged as RefOutOfBounds") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> overMaxRowSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.RefOutOfBounds))
    assertEquals(findings.head.part, "xl/worksheets/sheet1.xml")
    assert(findings.head.locator.contains("A1:XFD1048578"), findings.head.toString)
    assert(findings.head.message.contains("1048578"), findings.head.toString)
  }

  test("GH-428 class: mergeCell ref past column XFD is flagged as RefOutOfBounds") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> overMaxColSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.RefOutOfBounds))
    assert(findings.head.locator.contains("A1:XFE10"), findings.head.toString)
    assert(findings.head.message.contains("XFE"), findings.head.toString)
  }

  test("GH-428 class: dimension ref past the row limit is flagged as RefOutOfBounds") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> overMaxDimensionSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.RefOutOfBounds))
    assert(findings.head.locator.contains("dimension"), findings.head.toString)
  }

  test("GH-428 coordination: ranges ending exactly at XFD1048576 are NOT flagged") {
    // Post-#428 the writer clamps shifted ranges AT the sheet edge — that shape must lint clean.
    assertEquals(
      lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> atMaxSheetXml)),
      Vector.empty[Finding]
    )
  }

  test("GH-428 class: only the offending token of a multi-range sqref is flagged") {
    val sheet = worksheetWith(
      """<sheetData/>
  <conditionalFormatting sqref="A1:B2 C3:XFD1048999"><cfRule type="expression" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>"""
    )
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> sheet))
    assertEquals(findings.map(_.category), Vector(LintCategory.RefOutOfBounds))
    assert(findings.head.locator.contains("C3:XFD1048999"), findings.head.toString)
    assert(!findings.head.locator.contains("A1:B2"), findings.head.toString)
  }

  private val tableSheetXml = worksheetWith(
    """<sheetData/>
  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>"""
  ).replace(s"""<worksheet xmlns="$nsMain">""", s"""<worksheet xmlns="$nsMain" xmlns:r="$nsRel">""")

  private val tableSheetRelsXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>"""

  private val overMaxTableXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="$nsMain" id="1" name="T1" displayName="T1" ref="A1:C1048999"/>"""

  private val overMaxTableParts: Map[String, String] = baseParts +
    ("xl/worksheets/sheet1.xml" -> tableSheetXml) +
    ("xl/worksheets/_rels/sheet1.xml.rels" -> tableSheetRelsXml) +
    ("xl/tables/table1.xml" -> overMaxTableXml)

  test("GH-428 class: table part ref past the row limit is flagged on the table part") {
    val findings = lintOf(overMaxTableParts)
    assertEquals(findings.map(_.category), Vector(LintCategory.RefOutOfBounds))
    assertEquals(findings.head.part, "xl/tables/table1.xml")
    assert(findings.head.locator.contains("A1:C1048999"), findings.head.toString)
  }

  private val sharedTableParts: Map[String, String] = overMaxTableParts ++ Map(
    "xl/workbook.xml" -> workbookXml.replace(
      "</sheets>",
      "  <sheet name=\"Two\" sheetId=\"2\" r:id=\"rId5\"/>\n  </sheets>"
    ),
    "xl/_rels/workbook.xml.rels" -> workbookRelsXml.replace(
      "</Relationships>",
      """  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>"""
    ),
    "xl/worksheets/sheet2.xml" -> tableSheetXml,
    "xl/worksheets/_rels/sheet2.xml.rels" -> tableSheetRelsXml
  )

  test("GH-428 class: a table shared by two sheets' rels is reported once") {
    val findings = lintOf(sharedTableParts)
    assertEquals(findings.map(_.category), Vector(LintCategory.RefOutOfBounds))
    assertEquals(findings.head.part, "xl/tables/table1.xml")
  }

  // ===== GH-413 (4): O(1) SAX scanning mode =====

  private def lintStreamOf(parts: Map[String, String]): Vector[Finding] =
    WorkbookLint
      .lintStreamBytes(zipBytes(parts))
      .fold(err => fail(s"streaming lint must not error on a parseable package: $err"), identity)

  private def parityFixtures: Vector[(String, Map[String, String])] = Vector(
    "clean minimal" -> baseParts,
    "clean external" -> externalParts,
    "clean chartsheet" -> chartsheetParts,
    "clean dialogsheet" -> dialogsheetParts,
    "misordered chartsheet" ->
      (chartsheetParts + ("xl/chartsheets/sheet1.xml" -> misorderedChartsheetXml)),
    "misordered dialogsheet" ->
      (dialogsheetParts + ("xl/dialogsheets/sheet1.xml" -> misorderedDialogsheetXml)),
    "dangling externalBook" ->
      (externalParts + ("xl/externalLinks/externalLink1.xml" -> danglingExternalBookXml)),
    "unregistered styles part" ->
      (baseParts + ("[Content_Types].xml" -> unregisteredStylesCt)),
    "over-max cf sqref" -> (baseParts + ("xl/worksheets/sheet1.xml" -> overMaxRowSheetXml)),
    "over-max mergeCell" -> (baseParts + ("xl/worksheets/sheet1.xml" -> overMaxColSheetXml)),
    "over-max table ref" -> overMaxTableParts,
    "shared over-max table" -> sharedTableParts,
    "at-max boundary" -> (baseParts + ("xl/worksheets/sheet1.xml" -> atMaxSheetXml)),
    "dangling sheet r:id" ->
      (baseParts + ("xl/workbook.xml" -> workbookXml.replace("r:id=\"rId1\"", "r:id=\"rId99\""))),
    "worksheet order violation" -> (baseParts + ("xl/worksheets/sheet1.xml" ->
      s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain">
  <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
  <sheetData/>
</worksheet>""")),
    "dangling hyperlink r:id" -> (baseParts + ("xl/worksheets/sheet1.xml" ->
      s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain" xmlns:r="$nsRel">
  <sheetData/>
  <hyperlinks><hyperlink ref="A1" r:id="rId9"/></hyperlinks>
</worksheet>"""))
  )

  test("GH-413: lintStreamBytes agrees with lintBytes on every fixture (SAX/DOM parity)") {
    parityFixtures.foreach { case (name, parts) =>
      val bytes = zipBytes(parts)
      assertEquals(
        WorkbookLint.lintStreamBytes(bytes),
        WorkbookLint.lintBytes(bytes),
        s"parity broken for fixture: $name"
      )
    }
  }

  test("GH-413: lintStream(path) agrees with lint(path)") {
    val bytes = zipBytes(externalParts)
    val path = tempFile(bytes)
    try
      assertEquals(WorkbookLint.lintStream(path), WorkbookLint.lint(path))
      assertEquals(WorkbookLint.lintStream(path), Right(Vector.empty[Finding]))
    finally Files.deleteIfExists(path)
  }

  test("GH-428 class: the streaming scanner flags over-max sqref (O(1) mode carries the check)") {
    val findings = lintStreamOf(baseParts + ("xl/worksheets/sheet1.xml" -> overMaxRowSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.RefOutOfBounds))
    assert(findings.head.locator.contains("A1:XFD1048578"), findings.head.toString)
  }

  test("GH-413: streaming mode yields Left on malformed worksheet xml (parity with DOM mode)") {
    val bad = zipBytes(
      baseParts + ("xl/worksheets/sheet1.xml" -> "<worksheet><sheetData></worksheet>")
    )
    assert(WorkbookLint.lintStreamBytes(bad).isLeft)
    assert(WorkbookLint.lintBytes(bad).isLeft)
  }

  test("GH-413: streaming mode yields Left on garbage bytes, never throws") {
    assert(
      WorkbookLint.lintStreamBytes("not a zip at all".getBytes(StandardCharsets.UTF_8)).isLeft
    )
  }
