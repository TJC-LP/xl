package com.tjclp.xl.ooxml.lint

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{TestFixtures, XlsxReader, XlsxWriter}
import com.tjclp.xl.ooxml.writer.{WriterConfig, XmlBackend}
import com.tjclp.xl.sheets.dataTableSyntax.*
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
    // GH-460: the re-pointed rel also strands the present sheet1.xml — no relationship reaches it
    assertEquals(
      findings.map(_.category),
      Vector(LintCategory.MissingPart, LintCategory.UnreferencedPart)
    )
    assert(findings.head.message.contains("sheet9.xml"), findings.head.toString)
    assertEquals(findings(1).part, "xl/worksheets/sheet1.xml")
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

  // ===== GH-458: Microsoft xlExternalLinkPath variants are valid externalBook targets =====

  /** Excel writes these for broken/special external-book paths — a real in-the-wild state. */
  private def msVariantRelsXml(variant: String): String = externalLinkRelsXml.replace(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
    s"http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/$variant"
  )

  test("GH-458: externalBook rel typed xlPathMissing (Microsoft variant) is not flagged") {
    val findings = lintOf(
      externalParts +
        ("xl/externalLinks/_rels/externalLink1.xml.rels" -> msVariantRelsXml("xlPathMissing"))
    )
    assertEquals(findings, Vector.empty[Finding])
  }

  test("GH-458: externalBook rel typed xlLibraryPath (Microsoft variant) is not flagged") {
    val findings = lintOf(
      externalParts +
        ("xl/externalLinks/_rels/externalLink1.xml.rels" -> msVariantRelsXml("xlLibraryPath"))
    )
    assertEquals(findings, Vector.empty[Finding])
  }

  test("GH-529: suffix-less Microsoft variants (xlStartup, xlLibrary, xlAltStartup) are valid") {
    for variant <- List("xlStartup", "xlLibrary", "xlAltStartup") do
      val findings = lintOf(
        externalParts +
          ("xl/externalLinks/_rels/externalLink1.xml.rels" -> msVariantRelsXml(variant))
      )
      assertEquals(findings, Vector.empty[Finding], s"variant $variant must be accepted")
  }

  test("GH-458: a genuinely wrong externalBook rel type (image) still flags WrongRelType") {
    val findings = lintOf(
      externalParts +
        ("xl/externalLinks/_rels/externalLink1.xml.rels" -> wrongTypeExternalBookRelsXml)
    )
    assertEquals(findings.map(_.category), Vector(LintCategory.WrongRelType))
    assert(findings.head.message.contains("image"), findings.head.toString)
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

  // ===== GH-442: data-table record integrity (torn / unseeded) =====

  private val dtRecord =
    """<f t="dataTable" ref="D5:F6" dt2D="1" dtr="1" r1="B1" r2="B2" ca="1"/>"""

  /**
   * Excel's own 2-D data table shape over interior D5:F6: corner formula C4, row axis D4:F4, column
   * axis C5:C6, inputs B1/B2, the record on the interior top-left (corner-only) and every other
   * interior cell a plain cached value. Fully seeded, so it must lint clean under any calcMode.
   * Each fixture below varies exactly one thing off this base.
   */
  private val seededDataTableSheetXml = worksheetWith(
    s"""<sheetData>
    <row r="1"><c r="B1"><v>1</v></c></row>
    <row r="2"><c r="B2"><v>10</v></c></row>
    <row r="4"><c r="C4"><f>B1*B2</f><v>10</v></c><c r="D4"><v>1</v></c><c r="E4"><v>2</v></c><c r="F4"><v>3</v></c></row>
    <row r="5"><c r="C5"><v>10</v></c><c r="D5">$dtRecord<v>10</v></c><c r="E5"><v>20</v></c><c r="F5"><v>30</v></c></row>
    <row r="6"><c r="C6"><v>20</v></c><c r="D6"><v>20</v></c><c r="E6"><v>40</v></c><c r="F6"><v>60</v></c></row>
  </sheetData>"""
  )

  /** The put/putf tear: a real formula landed on interior E5. */
  private val tornInteriorSheetXml = seededDataTableSheetXml
    .replace("""<c r="E5"><v>20</v></c>""", """<c r="E5"><f>C4*2</f><v>20</v></c>""")

  /** The GH-435 delete-band shape: the row input was structurally removed, del1 records it. */
  private val delInputSheetXml =
    seededDataTableSheetXml.replace(""" r1="B1" r2="B2" ca="1"""", """ del1="1" r2="B2" ca="1"""")

  /** A 2-D record with no r1 and no del1 — the input reference is simply gone. */
  private val missingInputSheetXml =
    seededDataTableSheetXml.replace(""" r1="B1" r2="B2"""", """ r2="B2"""")

  /** Corner overwritten: the record survives at E5 but D5 (the ref's top-left) is a plain value. */
  private val overwrittenCornerSheetXml = seededDataTableSheetXml
    .replace(s"""<c r="D5">$dtRecord<v>10</v></c>""", """<c r="D5"><v>10</v></c>""")
    .replace("""<c r="E5"><v>20</v></c>""", s"""<c r="E5">$dtRecord<v>20</v></c>""")

  /** Every interior cache stripped — the state a builder ships before seeding. */
  private val unseededDataTableSheetXml = seededDataTableSheetXml
    .replace(s"""<c r="D5">$dtRecord<v>10</v></c>""", s"""<c r="D5">$dtRecord</c>""")
    .replace("""<c r="E5"><v>20</v></c>""", "")
    .replace("""<c r="F5"><v>30</v></c>""", "")
    .replace("""<c r="D6"><v>20</v></c>""", "")
    .replace("""<c r="E6"><v>40</v></c>""", "")
    .replace("""<c r="F6"><v>60</v></c>""", "")

  /** One interior cell absent: pins that the uncached count is arithmetic over the whole ref. */
  private val partlySeededDataTableSheetXml =
    seededDataTableSheetXml.replace("""<c r="F6"><v>60</v></c>""", "")

  /** A record whose ref leaves no room for the corner formula or either axis. */
  private val noRoomDataTableSheetXml = worksheetWith(
    """<sheetData>
    <row r="1"><c r="A1"><f t="dataTable" ref="A1:B2" dt2D="1" dtr="1" r1="D1" r2="D2" ca="1"/><v>1</v></c><c r="B1"><v>2</v></c></row>
    <row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>
  </sheetData>"""
  )

  private val autoNoTableWorkbookXml = workbookXml
    .replace(
      """<calcPr calcId="191029"/>""",
      """<calcPr calcId="191029" calcMode="autoNoTable"/>"""
    )

  private def dtParts(sheetXml: String): Map[String, String] =
    baseParts + ("xl/worksheets/sheet1.xml" -> sheetXml)

  private def autoNoTableParts(sheetXml: String): Map[String, String] =
    dtParts(sheetXml) + ("xl/workbook.xml" -> autoNoTableWorkbookXml)

  test("GH-442: a fully seeded data table is clean, including under calcMode=autoNoTable") {
    assertEquals(lintOf(dtParts(seededDataTableSheetXml)), Vector.empty[Finding])
    assertEquals(lintOf(autoNoTableParts(seededDataTableSheetXml)), Vector.empty[Finding])
  }

  test("GH-442: a formula inside the record interior is flagged as DataTableTorn") {
    val findings = lintOf(dtParts(tornInteriorSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.DataTableTorn))
    assertEquals(findings.head.part, "xl/worksheets/sheet1.xml")
    assert(findings.head.locator.contains("""ref="D5:F6""""), findings.head.toString)
    assert(findings.head.message.contains("E5"), findings.head.toString)
    assert(findings.head.message.contains("formula"), findings.head.toString)
  }

  test("GH-442: del1 on the record is flagged as DataTableTorn (input removed)") {
    val findings = lintOf(dtParts(delInputSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.DataTableTorn))
    assert(findings.head.message.contains("""del1="1""""), findings.head.toString)
    assert(findings.head.message.contains("structural delete"), findings.head.toString)
  }

  test("GH-442: a record missing its required input reference is flagged as DataTableTorn") {
    val findings = lintOf(dtParts(missingInputSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.DataTableTorn))
    assert(findings.head.message.contains("r1"), findings.head.toString)
  }

  test("GH-442: a record with no room for its corner and axes is flagged as DataTableTorn") {
    val findings = lintOf(dtParts(noRoomDataTableSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.DataTableTorn))
    assert(findings.head.locator.contains("""ref="A1:B2""""), findings.head.toString)
    assert(findings.head.message.contains("row 1 or column A"), findings.head.toString)
  }

  test("GH-442: a record ref whose top-left carries no record is flagged as DataTableTorn") {
    val findings = lintOf(dtParts(overwrittenCornerSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.DataTableTorn))
    assert(findings.head.message.contains("D5"), findings.head.toString)
  }

  test("GH-442: an uncached interior in a calcMode=autoNoTable book is DataTableUnseeded") {
    val findings = lintOf(autoNoTableParts(unseededDataTableSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.DataTableUnseeded))
    assertEquals(findings.head.part, "xl/worksheets/sheet1.xml")
    assert(findings.head.locator.contains("""ref="D5:F6""""), findings.head.toString)
    assert(findings.head.message.contains("6 of 6"), findings.head.toString)
    assert(findings.head.message.contains("autoNoTable"), findings.head.toString)
    assert(findings.head.message.contains("xl recalc --tables"), findings.head.toString)
  }

  test("GH-442: the uncached count is arithmetic over the whole ref (absent cells count)") {
    val findings = lintOf(autoNoTableParts(partlySeededDataTableSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.DataTableUnseeded))
    assert(findings.head.message.contains("1 of 6"), findings.head.toString)
  }

  test("GH-442: the same uncached interior is clean when the book is not calcMode=autoNoTable") {
    // GH-419 doctrine: plain calcMode="auto" books recompute their tables on open and self-heal.
    assertEquals(lintOf(dtParts(unseededDataTableSheetXml)), Vector.empty[Finding])
  }

  test("GH-442: an unseedable record is reported torn, never unseeded") {
    // Seeding cannot fix a del1 record, so pointing at `xl recalc --tables` would be wrong advice.
    assertEquals(
      lintOf(autoNoTableParts(delInputSheetXml)).map(_.category),
      Vector(LintCategory.DataTableTorn)
    )
  }

  test("GH-442: a torn-but-seedable record is reported torn AND unseeded") {
    assertEquals(
      lintOf(autoNoTableParts(tornInteriorSheetXml)).map(_.category),
      Vector(LintCategory.DataTableTorn, LintCategory.DataTableUnseeded)
    )
  }

  test("GH-442: Excel-authored data-table fixtures lint clean") {
    // The design-panel oracles (see DataTableAuthoringSpec / PROVENANCE.md): every-interior
    // records (formula-records.xlsx) and corner-only records must both pass untouched.
    List("datatable-excel.xlsx", "datatable-excel-edge.xlsx", "formula-records.xlsx").foreach {
      name =>
        val findings = WorkbookLint
          .lint(TestFixtures.copyToTemp(name))
          .fold(err => fail(s"lint errored on $name: $err"), identity)
        assertEquals(
          findings.filter(f => dataTableCategories.contains(f.category)),
          Vector.empty[Finding],
          s"$name must produce no data-table findings"
        )
    }
  }

  private val dataTableCategories =
    Set(LintCategory.DataTableTorn, LintCategory.DataTableUnseeded)

  test("GH-442: sheet.dataTable output lints clean; a formula put into its interior does not") {
    val output = Files.createTempFile("lint-dt-authored", ".xlsx")
    try
      val authored = Sheet("Data")
        .put(ref"C4", CellValue.Formula("B1*B2"))
        .put(ref"D4" -> 1, ref"E4" -> 2, ref"F4" -> 3)
        .put(ref"C5" -> 10, ref"C6" -> 20)
        .dataTable(
          CellRange.parse("D5:F6").fold(fail(_), identity),
          ref"B1",
          ref"B2",
          Seq.fill(2)(Seq.fill(3)(CellValue.Number(BigDecimal(1))))
        )
        .fold(err => fail(s"authoring failed: $err"), identity)
      XlsxWriter
        .write(Workbook(Vector(authored)), output)
        .fold(err => fail(s"write failed: $err"), identity)
      assertEquals(WorkbookLint.lint(output), Right(Vector.empty[Finding]))

      // The unguarded tear GH-442 exists to catch: put lands a real formula in the interior.
      val torn = authored.put(ref"E5", CellValue.Formula("C4*2"))
      XlsxWriter
        .write(Workbook(Vector(torn)), output)
        .fold(err => fail(s"write failed: $err"), identity)
      val findings =
        WorkbookLint.lint(output).fold(err => fail(s"lint errored: $err"), identity)
      assertEquals(findings.map(_.category), Vector(LintCategory.DataTableTorn))
      assert(findings.head.message.contains("E5"), findings.head.toString)
    finally Files.deleteIfExists(output)
  }

  // ===== GH-456: <f> text carrying the display form's leading '=' =====

  private val leadingEqualsSheetXml = worksheetWith(
    """<sheetData>
    <row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>=A1*2</f><v>4</v></c></row>
  </sheetData>"""
  )

  private val cleanFormulaSheetXml = worksheetWith(
    """<sheetData>
    <row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>4</v></c></row>
  </sheetData>"""
  )

  /** Seven offending formulas in document order A1, B1, C1, A2, B2, C2, A3. */
  private val manyLeadingEqualsSheetXml = worksheetWith(
    """<sheetData>
    <row r="1"><c r="A1"><f>=1</f></c><c r="B1"><f>=2</f></c><c r="C1"><f>=3</f></c></row>
    <row r="2"><c r="A2"><f>=4</f></c><c r="B2"><f>=5</f></c><c r="C2"><f>=6</f></c></row>
    <row r="3"><c r="A3"><f>=7</f></c></row>
  </sheetData>"""
  )

  test("GH-456: <f> text starting with '=' is flagged as FormulaLeadingEquals") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> leadingEqualsSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.FormulaLeadingEquals))
    val f = findings.head
    assertEquals(f.part, "xl/worksheets/sheet1.xml")
    assert(f.locator.contains("B1"), s"locator should carry the first offending cell: $f")
    assert(f.message.contains("1 formula(s)"), s"message should carry the total count: $f")
    assert(f.message.contains("B1"), s"message should name the offending cell: $f")
    assert(f.message.contains("re-writing"), s"message should say a re-write heals it: $f")
  }

  test("GH-456: leading-'=' findings aggregate to ONE finding per part (first 5 + total count)") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> manyLeadingEqualsSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.FormulaLeadingEquals))
    val f = findings.head
    assertEquals(f.part, "xl/worksheets/sheet1.xml")
    // Locator names the FIRST offending cell so the finding is actionable without re-deriving it.
    assert(f.locator.contains("A1"), s"locator should carry the first offending cell: $f")
    assert(f.message.contains("7 formula(s)"), s"message should carry the total count: $f")
    assert(
      f.message.contains("first 5: A1, B1, C1, A2, B2"),
      s"message should sample the first 5 offending cells in document order: $f"
    )
    // The sample is bounded: cells past the first 5 never appear (no unbounded finding growth).
    assert(!f.message.contains("C2"), s"message must not carry cells past the sample: $f")
    assert(!f.message.contains("A3"), s"message must not carry cells past the sample: $f")
    assert(f.message.contains("re-writing"), s"message should say a re-write heals it: $f")
  }

  test("GH-456: clean <f> text produces no findings") {
    assertEquals(
      lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> cleanFormulaSheetXml)),
      Vector.empty[Finding]
    )
  }

  test("GH-456: streaming mode flags the leading '=' identically") {
    assertEquals(
      lintStreamOf(baseParts + ("xl/worksheets/sheet1.xml" -> leadingEqualsSheetXml)),
      lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> leadingEqualsSheetXml))
    )
    assertEquals(
      lintStreamOf(baseParts + ("xl/worksheets/sheet1.xml" -> leadingEqualsSheetXml))
        .map(_.category),
      Vector(LintCategory.FormulaLeadingEquals)
    )
  }

  // ===== GH-525: dangling external-workbook ordinals =====

  /** `[1]` resolves (externalParts declares one entry); `[2]` (twice) and `[5]` dangle. */
  private val danglingExternalRefSheetXml = worksheetWith(
    """<sheetData>
    <row r="1">
      <c r="A1"><f>'[1]Extern'!A1</f><v>1</v></c>
      <c r="B1"><f>'[2]Model Guidance (Big 3)'!AQ5</f><v>2</v></c>
      <c r="C1"><f>[2]Extern!B1*2</f><v>4</v></c>
    </row>
    <row r="2"><c r="A2"><f>SUM('[5]Q1:Q4'!A1)</f><v>5</v></c></row>
  </sheetData>"""
  )

  private val externalOkSheetXml = worksheetWith(
    """<sheetData>
    <row r="1"><c r="A1"><f>'[1]Extern'!A1+[1]Extern!B1</f><v>2</v></c></row>
  </sheetData>"""
  )

  /** Bracket forms that must NEVER count: structured refs and string literals. */
  private val bracketNoiseSheetXml = worksheetWith(
    """<sheetData>
    <row r="1">
      <c r="A1"><f>SUM(Table1[3])</f><v>1</v></c>
      <c r="B1"><f>SUM(Table1[[#This Row],[3]])</f><v>2</v></c>
      <c r="C1" t="str"><f>"see [9] note"</f><v>see [9] note</v></c>
    </row>
  </sheetData>"""
  )

  private val danglingNameWorkbookXml = workbookWithExternalXml.replace(
    "<definedName name=\"MyName\">Sheet1!$A$1</definedName>",
    "<definedName name=\"MyName\">Sheet1!$A$1</definedName>" +
      "<definedName name=\"ExtRate\">'[4]Src'!$A$1</definedName>"
  )

  test("GH-525: externalOrdinals reads every external-ref form and nothing else") {
    assertEquals(WorkbookLint.externalOrdinals("'[3]GP - Customer'!$J$43"), Set(3))
    assertEquals(WorkbookLint.externalOrdinals("[5]Data!B2+[5]Data!C2"), Set(5))
    assertEquals(
      WorkbookLint.externalOrdinals(
        "INDEX('[5]Model Guidance (Big 3)'!AQ$5:AQ$15,MATCH($C78,'[5]Model Guidance (Big 3)'!$G$5:$G$15,0))"
      ),
      Set(5)
    )
    assertEquals(WorkbookLint.externalOrdinals("SUM('[2]Q1:Q4'!A1)"), Set(2))
    assertEquals(WorkbookLint.externalOrdinals("[2]!ExtName"), Set(2))
    assertEquals(WorkbookLint.externalOrdinals("=[1]Sheet1!A1"), Set(1))
    assertEquals(WorkbookLint.externalOrdinals("'[12]Ext'!A1*[3]Ext!B1"), Set(12, 3))
    // ordinal 0 is the self-workbook
    assertEquals(WorkbookLint.externalOrdinals("[0]!ThisBookName"), Set.empty[Int])
    // structured references — including a column literally named "3" — never count
    assertEquals(WorkbookLint.externalOrdinals("SUM(Table1[3])"), Set.empty[Int])
    assertEquals(
      WorkbookLint.externalOrdinals("SUM(Table1[[#This Row],[3]])"),
      Set.empty[Int]
    )
    // string literals are opaque, including escaped quotes and R1C1 text
    assertEquals(WorkbookLint.externalOrdinals("\"see [3] note\"&A1"), Set.empty[Int])
    assertEquals(WorkbookLint.externalOrdinals("INDIRECT(\"R[1]C[2]\",FALSE)"), Set.empty[Int])
    assertEquals(WorkbookLint.externalOrdinals("\"a\"\"[7]\"\"b\""), Set.empty[Int])
    // brackets inside a quoted name past position 0 cannot be an ordinal (sheet names forbid [])
    assertEquals(WorkbookLint.externalOrdinals("'It''s [2]'!A1"), Set.empty[Int])
    // file-name form and degenerate brackets are skipped, not findings
    assertEquals(WorkbookLint.externalOrdinals("[Book1.xlsx]Sheet1!A1"), Set.empty[Int])
    assertEquals(WorkbookLint.externalOrdinals("A1*[unclosed"), Set.empty[Int])
    assertEquals(WorkbookLint.externalOrdinals("A1*[]B2"), Set.empty[Int])
  }

  test("GH-525: formulas referencing external ordinals past the declared count are flagged") {
    val findings =
      lintOf(externalParts + ("xl/worksheets/sheet1.xml" -> danglingExternalRefSheetXml))
    assertEquals(
      findings.map(_.category),
      Vector(LintCategory.ExternalRefDangling, LintCategory.ExternalRefDangling)
    )
    val ord2 = findings(0)
    val ord5 = findings(1)
    assertEquals(ord2.part, "xl/worksheets/sheet1.xml")
    assert(ord2.message.contains("2 formula(s) reference external workbook [2]"), ord2.toString)
    assert(ord2.message.contains("only 1 <externalReference> entry"), ord2.toString)
    assert(ord2.locator.contains("B1"), s"locator carries the first offending cell: $ord2")
    assert(ord5.message.contains("1 formula(s) reference external workbook [5]"), ord5.toString)
    assert(ord5.message.contains("first at A2"), ord5.toString)
  }

  test("GH-525: with no externalReferences block, any external ordinal dangles") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> externalOkSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.ExternalRefDangling))
    assert(
      findings.head.message.contains("no <externalReference> entries"),
      findings.head.toString
    )
    assert(findings.head.message.contains("[1]"), findings.head.toString)
  }

  test("GH-525: ordinals within the declared count are clean") {
    assertEquals(
      lintOf(externalParts + ("xl/worksheets/sheet1.xml" -> externalOkSheetXml)),
      Vector.empty[Finding]
    )
  }

  test("GH-525: structured references and string literals never count") {
    assertEquals(
      lintOf(externalParts + ("xl/worksheets/sheet1.xml" -> bracketNoiseSheetXml)),
      Vector.empty[Finding]
    )
  }

  test("GH-525: a defined name with a dangling external ordinal is flagged on workbook.xml") {
    val findings = lintOf(externalParts + ("xl/workbook.xml" -> danglingNameWorkbookXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.ExternalRefDangling))
    val f = findings.head
    assertEquals(f.part, "xl/workbook.xml")
    assert(f.locator.contains("ExtRate"), f.toString)
    assert(f.message.contains("[4]"), f.toString)
  }

  test("GH-525: streaming mode flags dangling ordinals identically") {
    val parts = externalParts + ("xl/worksheets/sheet1.xml" -> danglingExternalRefSheetXml)
    assertEquals(lintStreamOf(parts), lintOf(parts))
    assertEquals(
      lintStreamOf(parts).map(_.category),
      Vector(LintCategory.ExternalRefDangling, LintCategory.ExternalRefDangling)
    )
  }

  // ===== GH-528: defined-name validity (fold-duplicates + definitely-illegal names) =====

  private def workbookWithNames(namesXml: String): String = workbookXml.replace(
    "<definedNames><definedName name=\"MyName\">Sheet1!$A$1</definedName></definedNames>",
    s"<definedNames>$namesXml</definedNames>"
  )

  private val foldDupWorkbookXml = workbookWithNames(
    """<definedName name="g" hidden="1">Sheet1!$A$1</definedName>""" +
      """<definedName name="ｇ" hidden="1">Sheet1!$A$2</definedName>"""
  )

  test("GH-528: names colliding under case/width folding in the same scope are flagged") {
    val findings = lintOf(baseParts + ("xl/workbook.xml" -> foldDupWorkbookXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.DefinedNameInvalid))
    val f = findings.head
    assertEquals(f.part, "xl/workbook.xml")
    assert(f.message.contains("\"g\""), f.toString)
    assert(f.message.contains("\"ｇ\""), f.toString)
    assert(f.message.contains("workbook scope"), f.toString)
  }

  test("GH-528: small vs base kana names collide (Excel's kana-size-insensitive comparison)") {
    val wb = workbookWithNames(
      """<definedName name="ぁ" hidden="1">Sheet1!$A$1</definedName>""" +
        """<definedName name="あ" hidden="1">Sheet1!$A$2</definedName>"""
    )
    val findings = lintOf(baseParts + ("xl/workbook.xml" -> wb))
    assertEquals(findings.map(_.category), Vector(LintCategory.DefinedNameInvalid))
  }

  test("GH-528: the same folded name on different scopes is legal shadowing, not a collision") {
    val wb = workbookWithNames(
      """<definedName name="Rate">Sheet1!$A$1</definedName>""" +
        """<definedName name="RATE" localSheetId="0">Sheet1!$A$2</definedName>"""
    )
    assertEquals(lintOf(baseParts + ("xl/workbook.xml" -> wb)), Vector.empty[Finding])
  }

  test("GH-528: a name past Excel's 255-character limit is flagged") {
    val long = "N" * 256
    val wb = workbookWithNames(s"""<definedName name="$long">Sheet1!$$A$$1</definedName>""")
    val findings = lintOf(baseParts + ("xl/workbook.xml" -> wb))
    assertEquals(findings.map(_.category), Vector(LintCategory.DefinedNameInvalid))
    assert(findings.head.message.contains("256"), findings.head.toString)
  }

  test("GH-528: a name carrying whitespace is flagged") {
    val wb = workbookWithNames("""<definedName name="My Name">Sheet1!$A$1</definedName>""")
    val findings = lintOf(baseParts + ("xl/workbook.xml" -> wb))
    assertEquals(findings.map(_.category), Vector(LintCategory.DefinedNameInvalid))
    assert(findings.head.message.contains("whitespace"), findings.head.toString)
  }

  test("GH-528: exact duplicates in the same scope are the degenerate fold collision") {
    val wb = workbookWithNames(
      """<definedName name="Twice">Sheet1!$A$1</definedName>""" +
        """<definedName name="Twice">Sheet1!$A$2</definedName>"""
    )
    val findings = lintOf(baseParts + ("xl/workbook.xml" -> wb))
    assertEquals(findings.map(_.category), Vector(LintCategory.DefinedNameInvalid))
    assert(findings.head.message.contains("2 defined names"), findings.head.toString)
  }

  // ===== GH-413 (4): O(1) SAX scanning mode =====

  // ===== GH-555: stale calculation chain =====

  private val calcChainContentTypesXml = contentTypesXml.replace(
    "</Types>",
    """  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>"""
  )

  private val calcChainRelsXml = workbookRelsXml.replace(
    "</Relationships>",
    """  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>"""
  )

  /**
   * B1 and C1 hold formulas; C2 is a shared-formula member (self-closing `<f/>`); A1 is a value.
   */
  private val formulaSheetXml = worksheetWith(
    """<sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><f>A1*2</f><v>2</v></c>
      <c r="C1"><f t="shared" ref="C1:C2" si="0">B1+1</f><v>3</v></c>
    </row>
    <row r="2">
      <c r="C2"><f t="shared" si="0"/><v>3</v></c>
    </row>
  </sheetData>"""
  )

  private def calcChainXml(entries: String): String =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<calcChain xmlns="$nsMain">$entries</calcChain>"""

  private def calcChainParts(entries: String): Map[String, String] = baseParts ++ Map(
    "[Content_Types].xml" -> calcChainContentTypesXml,
    "xl/_rels/workbook.xml.rels" -> calcChainRelsXml,
    "xl/worksheets/sheet1.xml" -> formulaSheetXml,
    "xl/calcChain.xml" -> calcChainXml(entries)
  )

  test("GH-555: a chain whose entries all name formula cells is clean (i carries forward)") {
    assertEquals(
      lintOf(calcChainParts("""<c r="B1" i="1"/><c r="C1"/><c r="C2" l="1"/>""")),
      Vector.empty[Finding]
    )
  }

  test("GH-555: entries naming cells without a formula are flagged once on xl/calcChain.xml") {
    val findings = lintOf(calcChainParts("""<c r="B1" i="1"/><c r="A1"/><c r="Z9"/>"""))
    assertEquals(
      findings.map(f => (f.part, f.category)),
      Vector(("xl/calcChain.xml", LintCategory.CalcChainStale))
    )
    val finding = findings.headOption.getOrElse(fail("expected one finding"))
    assertEquals(finding.locator, """<c r="A1" i="1">""")
    assert(finding.message.contains("2 of 3 entries"), finding.message)
    assert(finding.message.contains("first: A1, Z9"), finding.message)
    assert(finding.message.contains("Override"), finding.message)
  }

  test("GH-555: an entry for a sheet id workbook.xml does not declare is flagged") {
    val findings = lintOf(calcChainParts("""<c r="B1" i="1"/><c r="B1" i="7"/>"""))
    assertEquals(findings.map(_.category), Vector(LintCategory.CalcChainStale))
    val finding = findings.headOption.getOrElse(fail("expected one finding"))
    assertEquals(finding.locator, """<c r="B1" i="7">""")
    assert(finding.message.contains("sheetId 7"), finding.message)
  }

  test("GH-555: entries before any i attribute are skipped, not misattributed") {
    assertEquals(lintOf(calcChainParts("""<c r="Z9"/><c r="B1" i="1"/>""")), Vector.empty[Finding])
  }

  test("GH-555: calcChainEntries reads i carry-forward and upper-cases refs") {
    val chain = scala.xml.XML.loadString(
      calcChainXml("""<c r="b1" i="1"/><c r="C1"/><c r="D4" i="2"/><c r="e5"/>""")
    )
    assertEquals(
      WorkbookLint.calcChainEntries(chain),
      Vector(1 -> "B1", 1 -> "C1", 2 -> "D4", 2 -> "E5")
    )
  }

  test("GH-555: a chain that is not well-formed XML is one finding, and other rules still run") {
    val parts = calcChainParts("""<c r="B1" i="1"/>""") +
      ("xl/calcChain.xml" -> "<calcChain><c r=\"B1\" i=\"1\"></calcChain") +
      ("xl/worksheets/sheet1.xml" -> leadingEqualsSheetXml)
    val findings = lintOf(parts)
    assertEquals(
      findings.map(_.category).sortBy(_.slug),
      Vector(LintCategory.CalcChainStale, LintCategory.FormulaLeadingEquals)
    )
    val chain =
      findings.find(_.category == LintCategory.CalcChainStale).getOrElse(fail("no chain finding"))
    assertEquals(chain.locator, "<calcChain>")
    assert(chain.message.contains("not well-formed XML"), chain.message)
    assertEquals(lintStreamOf(parts), findings)
  }

  test("GH-555: i is the sheetId attribute, not the sheet position") {
    val renumbered = calcChainParts("""<c r="B1" i="7"/><c r="C1"/>""") +
      ("xl/workbook.xml" -> workbookXml.replace("""sheetId="1"""", """sheetId="7""""))
    assertEquals(lintOf(renumbered), Vector.empty[Finding])
    val positional = calcChainParts("""<c r="B1" i="1"/>""") +
      ("xl/workbook.xml" -> workbookXml.replace("""sheetId="1"""", """sheetId="7""""))
    val findings = lintOf(positional)
    assertEquals(findings.map(_.category), Vector(LintCategory.CalcChainStale))
    assert(findings.headOption.exists(_.message.contains("sheetId 1")), findings.toString)
    assertEquals(lintStreamOf(renumbered), Vector.empty[Finding])
  }

  test("GH-555: entries inside an array-formula range are formula cells (anchor carries the <f>)") {
    val arraySheet = worksheetWith(
      """<sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="array" ref="B1:B3">TRANSPOSE(A1:C1)</f><v>1</v></c></row>
    <row r="2"><c r="B2"><v>2</v></c></row>
    <row r="3"><c r="B3"><v>3</v></c><c r="D3"><v>9</v></c></row>
  </sheetData>"""
    )
    val parts = calcChainParts("""<c r="B1" i="1"/><c r="B2"/><c r="B3"/>""") +
      ("xl/worksheets/sheet1.xml" -> arraySheet)
    assertEquals(lintOf(parts), Vector.empty[Finding])
    assertEquals(lintStreamOf(parts), Vector.empty[Finding])
    val outside = parts + ("xl/calcChain.xml" -> calcChainXml("""<c r="B1" i="1"/><c r="D3"/>"""))
    assertEquals(lintOf(outside).map(_.category), Vector(LintCategory.CalcChainStale))
    assert(
      lintOf(outside).headOption.exists(_.message.contains("first: D3")),
      lintOf(outside).toString
    )
  }

  test("GH-555: an entry for a declared sheet whose part is missing is left to missing-part") {
    val parts = calcChainParts("""<c r="B1" i="1"/>""") - "xl/worksheets/sheet1.xml"
    val categories = lintOf(parts).map(_.category)
    assert(categories.contains(LintCategory.MissingPart), categories.toString)
    assert(!categories.contains(LintCategory.CalcChainStale), categories.toString)
  }

  test("GH-555: streaming mode reports identical findings") {
    val parts = calcChainParts("""<c r="B1" i="1"/><c r="A1"/><c r="B1" i="7"/>""")
    assertEquals(lintStreamOf(parts), lintOf(parts))
    assertEquals(lintStreamOf(parts).size, 2)
    val clean = calcChainParts("""<c r="B1" i="1"/><c r="C1"/><c r="C2"/>""")
    assertEquals(lintStreamOf(clean), Vector.empty[Finding])
  }

  // ===== GH-577 / GH-588: post-2007 functions stored bare (xlfn-missing) =====

  /**
   * D1, E1, the cfRule `<formula>` and the dataValidation `<formula1>` are bare; A1 carries the
   * prefix, B1 is an Excel 2007 function, C1 hides `IFS(` inside a string literal.
   */
  private val bareXlfnSheetXml = worksheetWith(
    """<sheetData>
    <row r="1">
      <c r="A1"><f>_xlfn.IFS(B1=1,1,TRUE,0)</f><v>1</v></c>
      <c r="B1"><f>SUMIFS(C1:C3,D1:D3,1)</f><v>0</v></c>
      <c r="C1" t="str"><f>"IFS(" &amp; "x"</f><v>IFS(x</v></c>
      <c r="D1"><f>IFS(B1=1,1,TRUE,0)</f><v>1</v></c>
      <c r="E1"><f>xlookup(A1,A1:A3,B1:B3)</f><v>1</v></c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="A1:A3"><cfRule type="expression" priority="1"><formula>IFS(A1=1,TRUE,TRUE,FALSE)</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="custom" sqref="D1:D3"><formula1>MAXIFS(B1:B3,A1:A3,C1)</formula1></dataValidation></dataValidations>"""
  )

  /** The issue's repro: a hand-built CF rule calling IFS without the prefix, nothing else. */
  private val bareCfIfsSheetXml = worksheetWith(
    """<sheetData/>
  <conditionalFormatting sqref="A1:A3"><cfRule type="expression" priority="1"><formula>IFS(A1=1,TRUE,TRUE,FALSE)</formula></cfRule></conditionalFormatting>"""
  )

  /** Every slot prefixed the way Excel writes it — including x14 rules' `<xm:f>`. */
  private val prefixedXlfnSheetXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><f>_xlfn.IFS(B1=1,1,TRUE,0)</f><v>1</v></c><c r="B1"><f>_xlfn._xlws.FILTER(A1:A3,B1:B3)</f><v>1</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A1:A3"><cfRule type="expression" priority="1"><formula>_xlfn.IFS(A1=1,TRUE,TRUE,FALSE)</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="custom" sqref="D1:D3"><formula1>_xlfn.MAXIFS(B1:B3,A1:A3,C1)</formula1></dataValidation></dataValidations>
  <extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression" priority="2"><xm:f>_xlfn.XLOOKUP(A1,A1:A3,B1:B3)=1</xm:f></x14:cfRule><xm:sqref>A1:A3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>
</worksheet>"""

  /** Seven bare cells in document order A1, B1, C1, A2, B2, C2, A3 (the aggregation shape). */
  private val manyBareXlfnSheetXml = worksheetWith(
    """<sheetData>
    <row r="1"><c r="A1"><f>IFS(1,1)</f></c><c r="B1"><f>CONCAT(1)</f></c><c r="C1"><f>IFS(1,1)</f></c></row>
    <row r="2"><c r="A2"><f>IFNA(1,1)</f></c><c r="B2"><f>IFS(1,1)</f></c><c r="C2"><f>IFS(1,1)</f></c></row>
    <row r="3"><c r="A3"><f>IFS(1,1)</f></c></row>
  </sheetData>"""
  )

  private val bareNameWorkbookXml = workbookWithNames(
    """<definedName name="Best">MAXIFS(Sheet1!$B$1:$B$3,Sheet1!$A$1:$A$3,1)</definedName>""" +
      """<definedName name="Fine">_xlfn.MAXIFS(Sheet1!$B$1:$B$3,Sheet1!$A$1:$A$3,1)</definedName>"""
  )

  /**
   * The openpyxl class the lint exists for: `_xlfn.` present on LET, `_xlpm.` absent from its
   * parameters — Excel reports the book as unreadable content on open.
   */
  private val halfPrefixedLetSheetXml = worksheetWith(
    """<sheetData>
    <row r="1"><c r="A1"><f>_xlfn.LET(x,1,x+1)</f><v>2</v></c></row>
  </sheetData>"""
  )

  /** A part whose ROOT is a formula-text label: never a site (it has no parent to name one). */
  private val rootFormulaSheetXml = "<f>IFS(1,1)</f>"

  test("GH-588: the lint's rule is the writer's — an _xlfn.LET with bare parameters is flagged") {
    // FormulaStorage.bareFutureCalls runs toStored's scanner with recording callbacks, so the
    // half-prefixed LET (which toStored completes to _xlfn.LET(_xlpm.x,1,_xlpm.x+1)) is a finding
    // here exactly because the writer would change it; FormulaStorageSpec pins the invariant.
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> halfPrefixedLetSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.XlfnMissing))
    val f = findings.head
    assertEquals(f.locator, """<c r="A1"><f>""")
    assert(f.message.contains("1 formula(s)"), f.toString)
    assert(f.message.contains("(LET; A1)"), f.toString)
    assert(f.message.contains("_xlpm."), f.toString)
  }

  test("GH-654: a bare @ is reported as the token in the file, not as a SINGLE call") {
    val bareAtSheetXml = worksheetWith(
      """<sheetData>
    <row r="1"><c r="B1"><f>@acq</f><v>1</v></c><c r="C1"><f>_xlfn.SINGLE(acq)</f><v>1</v></c></row>
  </sheetData>"""
    )
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> bareAtSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.XlfnMissing))
    val f = findings.head
    assertEquals(f.locator, """<c r="B1"><f>""")
    assert(f.message.contains("1 formula(s)"), f.toString)
    assert(f.message.contains("bare @"), f.toString)
    assert(f.message.contains("_xlfn.SINGLE"), f.toString)
    assert(f.message.contains("(B1)"), f.toString)
    assert(f.message.contains("repairs"), f.toString)
    assert(!f.message.contains("#NAME?"), s"no call is missing a prefix here: ${f.message}")
    assert(!f.message.contains("(SINGLE"), s"the formula holds no SINGLE call: ${f.message}")
  }

  test("GH-655: a bare spill reference x# is reported as the token in the file") {
    val bareSpillSheetXml = worksheetWith(
      """<sheetData>
    <row r="1"><c r="B1"><f>SUM(A1#)</f><v>6</v></c><c r="C1"><f>SUM(_xlfn.ANCHORARRAY(A1))</f><v>6</v></c></row>
  </sheetData>"""
    )
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> bareSpillSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.XlfnMissing))
    val f = findings.head
    assertEquals(f.locator, """<c r="B1"><f>""")
    assert(f.message.contains("bare x#"), f.toString)
    assert(f.message.contains("_xlfn.ANCHORARRAY"), f.toString)
    assert(f.message.contains("(B1)"), f.toString)
    assert(f.message.contains("repairs"), f.toString)
    assert(!f.message.contains("#NAME?"), f.toString)
    assert(!f.message.contains("(ANCHORARRAY"), f.toString)
  }

  test("GH-588: a part whose root element is formula text records no site in either mode") {
    val parts = baseParts + ("xl/worksheets/sheet1.xml" -> rootFormulaSheetXml)
    val dom = lintOf(parts)
    assert(!dom.exists(_.category == LintCategory.XlfnMissing), dom.toString)
    assertEquals(lintStreamOf(parts), dom)
  }

  test("GH-588: a hand-built bare IFS in a CF rule is flagged as XlfnMissing on the sheet part") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> bareCfIfsSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.XlfnMissing))
    val f = findings.head
    assertEquals(f.part, "xl/worksheets/sheet1.xml")
    assertEquals(f.locator, "<cfRule><formula>")
    assert(f.message.contains("IFS"), f.toString)
    assert(f.message.contains("1 formula(s)"), f.toString)
    assert(f.message.contains("_xlfn."), f.toString)
    assert(f.message.contains("#NAME?"), f.toString)
  }

  test(
    "GH-593: the remediation states the healing condition — regenerate the part, not re-author"
  ) {
    // GH-588 pinned the old condition ("only when xl regenerates it; re-authoring identical text
    // does not"): the CF / DV / name gates compared models, and bare text parses to the same model
    // as prefixed text. GH-593 made the gates storage-form aware (FutureFunctionPrefixSpec pins
    // the writer), so the honest condition is now: any in-memory write that regenerates the
    // worksheet (CF/DV) or workbook.xml (names) heals the slot. The finding states that in one
    // clause and points at the lint docs, which carry the residuals (an unmodeled preserved rule,
    // an x14 <xm:f>, an untouched worksheet, a --stream write) slot by slot (cli.md,
    // LIMITATIONS.md).
    val cf = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> bareCfIfsSheetXml)).head
    val dn = lintOf(baseParts + ("xl/workbook.xml" -> bareNameWorkbookXml)).head
    Vector(cf, dn).foreach { f =>
      assert(!f.message.contains("re-writing the affected part"), f.toString)
      assert(!f.message.contains("re-authoring identical text does not"), f.toString)
      assert(!f.message.contains("only when xl regenerates it"), f.toString)
      assert(f.message.contains("any in-memory write regenerating the worksheet"), f.toString)
      assert(f.message.contains("workbook.xml (names) heals it"), f.toString)
      assert(f.message.contains("docs/reference/cli.md"), f.toString)
      // CLI-level detail (verbs, --stream) lives in the docs, not in a library finding
      assert(!f.message.contains("--stream"), f.toString)
      assert(!f.message.contains("name add"), f.toString)
      assert(f.message.length <= 320, s"${f.message.length} chars: ${f.message}")
    }
  }

  test("GH-588: bare calls in <f>, <formula> and <formula1> aggregate to ONE finding per part") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> bareXlfnSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.XlfnMissing))
    val f = findings.head
    assertEquals(f.part, "xl/worksheets/sheet1.xml")
    // the first offending site in document order is the cell D1 (A1 is prefixed, B1 is 2007,
    // C1 only mentions IFS( inside a string literal)
    assertEquals(f.locator, """<c r="D1"><f>""")
    assert(f.message.contains("4 formula(s)"), f.toString)
    assert(f.message.contains("IFS, MAXIFS, XLOOKUP"), f.toString)
    assert(f.message.contains("D1, E1, <cfRule><formula>, <dataValidation><formula1>"), f.toString)
  }

  test("GH-588: the sample is bounded to the first 5 sites plus a total count") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> manyBareXlfnSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.XlfnMissing))
    val f = findings.head
    assertEquals(f.locator, """<c r="A1"><f>""")
    assert(f.message.contains("7 formula(s)"), f.toString)
    assert(f.message.contains("first 5: A1, B1, C1, A2, B2"), f.toString)
    assert(f.message.contains("CONCAT, IFNA, IFS"), f.toString)
    assert(!f.message.contains("A3"), f.toString)
  }

  test("GH-588: Excel's own storage form (incl. _xlfn._xlws. and x14 <xm:f>) is clean") {
    assertEquals(
      lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> prefixedXlfnSheetXml)),
      Vector.empty[Finding]
    )
  }

  test("GH-588: a bare post-2007 call in a <definedName> is flagged on workbook.xml") {
    val findings = lintOf(baseParts + ("xl/workbook.xml" -> bareNameWorkbookXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.XlfnMissing))
    val f = findings.head
    assertEquals(f.part, "xl/workbook.xml")
    assertEquals(f.locator, """<definedName name="Best">""")
    assert(f.message.contains("MAXIFS"), f.toString)
    assert(f.message.contains("1 defined name(s)"), f.toString)
    assert(!f.message.contains("Fine"), f.toString)
  }

  test("GH-588: streaming mode flags bare calls identically (cells, CF, DV, names)") {
    val parts = baseParts +
      ("xl/worksheets/sheet1.xml" -> bareXlfnSheetXml) +
      ("xl/workbook.xml" -> bareNameWorkbookXml)
    assertEquals(lintStreamOf(parts), lintOf(parts))
    assertEquals(
      lintStreamOf(parts).map(f => (f.part, f.category)),
      Vector(
        "xl/workbook.xml" -> LintCategory.XlfnMissing,
        "xl/worksheets/sheet1.xml" -> LintCategory.XlfnMissing
      )
    )
  }

  private def customDv(formula: String): com.tjclp.xl.sheets.DataValidation.Rules =
    com.tjclp.xl.sheets.DataValidation
      .Rules(Vector.empty, com.tjclp.xl.sheets.DvKind.Custom(formula))

  test("GH-588: an xl-written book with post-2007 functions in every slot lints clean") {
    import com.tjclp.xl.cf.{CfOperator, CfPoint, CfRule, Cfvo}
    import com.tjclp.xl.styles.color.Color
    val sheet = Sheet("Data")
      .put(ref"A1" -> "a")
      .put(ref"B1" -> 1)
      .put(ref"C1" -> CellValue.Formula("IFS(B1=1,1,TRUE,0)", Some(CellValue.Number(1))))
      .put(ref"D1" -> CellValue.Formula("FILTER(A1:A3,B1:B3)", None))
      .put(ref"E1" -> CellValue.Formula("LET(x,B1,x+1)", None))
      .conditionalFormat(
        ref"A1:A3",
        CfRule.Expression("IFS(A1=1,1,TRUE,0)", None, 1),
        CfRule.CellIs(CfOperator.GreaterThan, "XLOOKUP(C1,A1:A3,B1:B3)", None, None, 2),
        CfRule.ColorScale(
          CfPoint(Cfvo.Formula("MINIFS(B1:B3,A1:A3,C1)"), Color.Rgb(0xffff0000)),
          None,
          CfPoint(Cfvo.Max, Color.Rgb(0xff00ff00)),
          3
        )
      )
      .withDataValidation(ref"F1:F3", customDv("MAXIFS(B1:B3,A1:A3,C1)"))
    val wb = Workbook(Vector(sheet))
      .withDefinedName("Best", "MAXIFS(Data!$B$1:$B$3,Data!$A$1:$A$3,Data!$C$1)")
    Vector(XmlBackend.ScalaXml, XmlBackend.SaxStax).foreach { backend =>
      val path = Files.createTempFile("lint-xlfn", ".xlsx")
      try
        XlsxWriter
          .writeWith(wb, path, WriterConfig(backend = backend))
          .fold(e => fail(s"write failed: $e"), identity)
        assertEquals(WorkbookLint.lint(path), Right(Vector.empty[Finding]), backend.toString)
        assertEquals(WorkbookLint.lintStream(path), Right(Vector.empty[Finding]), backend.toString)
      finally Files.deleteIfExists(path)
    }
  }

  // ===== GH-460: empty inline strings (openpyxl's serialization of value="") =====

  private val emptyInlineSheetXml = worksheetWith(
    """<sheetData>
    <row r="1">
      <c r="A1" s="0" t="inlineStr"/>
      <c r="B1" t="inlineStr"><is/></c>
      <c r="C1" t="str"/>
      <c r="D1" t="inlineStr"><is><t>ok</t></is></c>
      <c r="E1" t="str"><v>x</v></c>
      <c r="F1" t="str"><f>D1</f></c>
      <c r="G1" t="inlineStr"><is><r><t>rich</t></r></is></c>
      <c r="H1" t="inlineStr"><v>legacy</v></c>
    </row>
  </sheetData>"""
  )

  private val textfulInlineSheetXml = worksheetWith(
    """<sheetData>
    <row r="1">
      <c r="D1" t="inlineStr"><is><t>ok</t></is></c>
      <c r="E1" t="str"><v>x</v></c>
      <c r="F1" t="str"><f>D1</f></c>
      <c r="G1" t="inlineStr"><is><r><t>rich</t></r></is></c>
      <c r="H1" t="inlineStr"><v>legacy</v></c>
      <c r="I1"/>
      <c r="J1" s="0"/>
    </row>
  </sheetData>"""
  )

  private val manyEmptyInlineSheetXml = worksheetWith(
    "<sheetData><row r=\"1\">" +
      (0 until 7).map(i => s"""<c r="${('A' + i).toChar}1" t="inlineStr"/>""").mkString +
      "</row></sheetData>"
  )

  test("GH-460: <c t=\"inlineStr\"/> with no <is> is flagged as EmptyInlineStr, once per part") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> emptyInlineSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.EmptyInlineStr))
    val f = findings.head
    assertEquals(f.part, "xl/worksheets/sheet1.xml")
    assertEquals(f.locator, """<c r="A1" t="inlineStr"/>""")
    assert(f.message.startsWith("2 "), f.message)
    assert(f.message.contains("A1, C1"), f.message)
    assert(f.message.contains("openpyxl"), f.message)
    // shapes that carry text (D1, E1, G1, H1), a value-less formula (F1) and an <is/> that is
    // present but empty (B1 — the empty STRING to every reader) are not the class
    Vector("B1", "D1", "E1", "F1", "G1", "H1").foreach { r =>
      assert(!f.message.contains(r), f.message)
    }
  }

  test("GH-460: the empty-inline sample is bounded to the first 5 cells plus a total count") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> manyEmptyInlineSheetXml))
    assertEquals(findings.size, 1)
    assert(findings.head.message.startsWith("7 "), findings.head.message)
    assert(findings.head.message.contains("first 5: A1, B1, C1, D1, E1, …"), findings.head.message)
  }

  test("GH-460: inline strings that carry text, value-less formulas and typeless cells are clean") {
    assertEquals(
      lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> textfulInlineSheetXml)),
      Vector.empty[Finding]
    )
  }

  test("GH-460: streaming mode flags empty inline strings identically") {
    val parts = baseParts + ("xl/worksheets/sheet1.xml" -> emptyInlineSheetXml)
    assertEquals(lintStreamOf(parts), lintOf(parts))
    assertEquals(lintStreamOf(parts).map(_.category), Vector(LintCategory.EmptyInlineStr))
  }

  test("GH-460 addendum: the shape lint flags is the shape the in-memory reader now opens") {
    // Before: lint passed the file clean while XlsxReader failed every read verb on it.
    val bytes = zipBytes(baseParts + ("xl/worksheets/sheet1.xml" -> emptyInlineSheetXml))
    assertEquals(
      WorkbookLint.lintBytes(bytes).map(_.map(_.category)),
      Right(Vector(LintCategory.EmptyInlineStr))
    )
    val wb = XlsxReader
      .readFromBytes(bytes)
      .fold(err => fail(s"the reader must tolerate a childless inlineStr: $err"), identity)
    val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
    Vector(ref"A1", ref"C1").foreach(r => assertEquals(sheet(r).value, CellValue.Empty))
    // an <is/> is an inline string that is present but empty — the empty string, as streaming reads it
    assertEquals(sheet(ref"B1").value, CellValue.Text(""))
    assertEquals(sheet(ref"D1").value, CellValue.Text("ok"))
    assertEquals(sheet(ref"E1").value, CellValue.Text("x"))
    assertEquals(sheet(ref"H1").value, CellValue.Text("legacy"))
  }

  // ===== GH-460: mc:Ignorable naming undeclared prefixes (the ElementTree re-prefix class) =====

  private val nsMc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
  private val nsX14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
  private val nsXr = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"

  private def worksheetRoot(rootAttrs: String, body: String = "<sheetData/>"): String =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain" $rootAttrs>
  $body
</worksheet>"""

  /** ElementTree kept the Ignorable VALUE but re-prefixed its declaration: x14ac and xr dangle. */
  private val undeclaredIgnorableSheetXml =
    worksheetRoot(s"""xmlns:mc="$nsMc" xmlns:ns1="$nsX14ac" mc:Ignorable="x14ac xr"""")

  private val declaredIgnorableSheetXml = worksheetRoot(
    s"""xmlns:mc="$nsMc" xmlns:x14ac="$nsX14ac" xmlns:xr="$nsXr" mc:Ignorable="x14ac xr""""
  )

  /** sheetPr: x14ac declared locally, xr on the root (clean); ext: x14ac lives on a SIBLING. */
  private val nestedIgnorableSheetXml = worksheetRoot(
    s"""xmlns:mc="$nsMc" xmlns:xr="$nsXr"""",
    s"""<sheetPr xmlns:x14ac="$nsX14ac" mc:Ignorable="x14ac xr"/>
  <sheetData/>
  <extLst><ext uri="{1}" mc:Ignorable="x14ac"/></extLst>"""
  )

  private val elementTreeRootSheetXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<ns0:worksheet xmlns:ns0="$nsMain"><ns0:sheetData/></ns0:worksheet>"""

  private val unboundPrefixSheetXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<ns0:worksheet><ns0:sheetData/></ns0:worksheet>"""

  private val undeclaredIgnorableWorkbookXml = workbookXml.replace(
    s"""<workbook xmlns="$nsMain" xmlns:r="$nsRel">""",
    s"""<workbook xmlns="$nsMain" xmlns:r="$nsRel" xmlns:mc="$nsMc" mc:Ignorable="x15">"""
  )

  private val undeclaredIgnorableTableXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="$nsMain" xmlns:mc="$nsMc" mc:Ignorable="xr xr3" id="1" name="T1" displayName="T1" ref="A1:C3"/>"""

  private val undeclaredIgnorableTableParts: Map[String, String] =
    overMaxTableParts + ("xl/tables/table1.xml" -> undeclaredIgnorableTableXml)

  test("GH-460: mc:Ignorable naming a prefix declared on no ancestor is flagged") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> undeclaredIgnorableSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.IgnorableUndeclared))
    val f = findings.head
    assertEquals(f.part, "xl/worksheets/sheet1.xml")
    assertEquals(f.locator, """<worksheet mc:Ignorable="x14ac xr">""")
    assert(f.message.contains("x14ac, xr"), f.message)
    assert(f.message.contains("ElementTree"), f.message)
  }

  test("GH-460: mc:Ignorable whose prefixes are all declared is clean") {
    assertEquals(
      lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> declaredIgnorableSheetXml)),
      Vector.empty[Finding]
    )
  }

  test("GH-460: a prefix declared on an ancestor resolves; one declared on a sibling does not") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> nestedIgnorableSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.IgnorableUndeclared))
    assertEquals(findings.head.locator, """<ext mc:Ignorable="x14ac">""")
    assert(findings.head.message.contains("x14ac"), findings.head.message)
    assert(!findings.head.message.contains("xr"), findings.head.message)
  }

  test("GH-460: an ns0-prefixed root (ElementTree re-serialization signature) is flagged") {
    val findings = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> elementTreeRootSheetXml))
    assertEquals(findings.map(_.category), Vector(LintCategory.IgnorableUndeclared))
    assertEquals(findings.head.locator, "<ns0:worksheet>")
    assert(findings.head.message.contains("ElementTree"), findings.head.message)
    // PR #659 review: Excel and LibreOffice open a namespace-correct prefixed root with every cell
    // intact (verified against both) — the message must say so, never that the part opens blank
    assert(!findings.head.message.contains("blank"), findings.head.message)
    assert(findings.head.message.contains("intact"), findings.head.message)
  }

  test("GH-460: an UNBOUND element prefix is a well-formedness error — Left in both modes") {
    val bytes = zipBytes(baseParts + ("xl/worksheets/sheet1.xml" -> unboundPrefixSheetXml))
    assert(WorkbookLint.lintBytes(bytes).isLeft)
    assert(WorkbookLint.lintStreamBytes(bytes).isLeft)
  }

  test("GH-460: undeclared mc:Ignorable prefixes on workbook.xml and table parts are flagged") {
    val wb = lintOf(baseParts + ("xl/workbook.xml" -> undeclaredIgnorableWorkbookXml))
    assertEquals(
      wb.map(f => (f.part, f.category)),
      Vector(("xl/workbook.xml", LintCategory.IgnorableUndeclared))
    )
    assertEquals(wb.head.locator, """<workbook mc:Ignorable="x15">""")
    val table = lintOf(undeclaredIgnorableTableParts)
    assertEquals(
      table.map(f => (f.part, f.category)),
      Vector(("xl/tables/table1.xml", LintCategory.IgnorableUndeclared))
    )
    assert(table.head.message.contains("xr, xr3"), table.head.message)
  }

  test("GH-460: streaming mode flags undeclared mc:Ignorable prefixes identically") {
    Vector(undeclaredIgnorableSheetXml, nestedIgnorableSheetXml, elementTreeRootSheetXml).foreach {
      sheet =>
        val parts = baseParts + ("xl/worksheets/sheet1.xml" -> sheet)
        assertEquals(lintStreamOf(parts), lintOf(parts))
        assertEquals(lintStreamOf(parts).size, 1)
    }
    assertEquals(
      lintStreamOf(undeclaredIgnorableTableParts),
      lintOf(undeclaredIgnorableTableParts)
    )
  }

  // ===== GH-460: dxfId past the <dxfs> table =====

  private def stylesWithDxfs(dxfsXml: String): String =
    stylesXml.replace("</styleSheet>", s"  $dxfsXml\n</styleSheet>")

  private def cfSheetXml(dxfId: Int): String = worksheetWith(
    s"""<sheetData/>
  <conditionalFormatting sqref="A1:A5"><cfRule type="cellIs" dxfId="$dxfId" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>"""
  )

  private val oneDxfStylesXml =
    stylesWithDxfs("""<dxfs count="1"><dxf><font><b/></font></dxf></dxfs>""")

  private val danglingDxfParts: Map[String, String] = baseParts ++ Map(
    "xl/styles.xml" -> oneDxfStylesXml,
    "xl/worksheets/sheet1.xml" -> cfSheetXml(1)
  )

  private val inRangeDxfParts: Map[String, String] = baseParts ++ Map(
    "xl/styles.xml" -> oneDxfStylesXml,
    "xl/worksheets/sheet1.xml" -> cfSheetXml(0)
  )

  private val dxfTableXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="$nsMain" id="1" name="T1" displayName="T1" ref="A1:C3" headerRowDxfId="0" dataDxfId="3">
  <tableColumns count="1"><tableColumn id="1" name="A" dataDxfId="7"/></tableColumns>
</table>"""

  private val dxfTableParts: Map[String, String] = overMaxTableParts ++ Map(
    "xl/styles.xml" -> oneDxfStylesXml,
    "xl/tables/table1.xml" -> dxfTableXml
  )

  test("GH-460: cfRule dxfId >= the <dxfs> child count is flagged as DxfIdOutOfRange") {
    val findings = lintOf(danglingDxfParts)
    assertEquals(findings.map(_.category), Vector(LintCategory.DxfIdOutOfRange))
    val f = findings.head
    assertEquals(f.part, "xl/worksheets/sheet1.xml")
    assertEquals(f.locator, """<cfRule dxfId="1">""")
    assert(f.message.contains("holds 1 <dxf>"), f.message)
  }

  test("GH-460: the <dxfs count> attribute is not trusted — actual <dxf> children are") {
    val liedCount = stylesWithDxfs("""<dxfs count="5"><dxf><font><b/></font></dxf></dxfs>""")
    val findings = lintOf(danglingDxfParts + ("xl/styles.xml" -> liedCount))
    assertEquals(findings.map(_.category), Vector(LintCategory.DxfIdOutOfRange))
    assert(findings.head.message.contains("holds 1 <dxf>"), findings.head.message)
  }

  test("GH-460: a dxfId within the table is clean; no <dxfs> at all with a dxfId is flagged") {
    assertEquals(lintOf(inRangeDxfParts), Vector.empty[Finding])
    val noDxfs = lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> cfSheetXml(0)))
    assertEquals(noDxfs.map(_.category), Vector(LintCategory.DxfIdOutOfRange))
    assert(noDxfs.head.message.contains("has no <dxfs>"), noDxfs.head.message)
    val noStyles =
      lintOf(baseParts - "xl/styles.xml" + ("xl/worksheets/sheet1.xml" -> cfSheetXml(0)))
    assertEquals(noStyles.map(_.category), Vector(LintCategory.DxfIdOutOfRange))
    assert(noStyles.head.message.contains("no styles part"), noStyles.head.message)
  }

  test("GH-460: table-part dataDxfId / headerRowDxfId index the same <dxfs> table") {
    val findings = lintOf(dxfTableParts)
    assertEquals(
      findings.map(f => (f.part, f.category, f.locator)),
      Vector(
        ("xl/tables/table1.xml", LintCategory.DxfIdOutOfRange, """<table dataDxfId="3">"""),
        ("xl/tables/table1.xml", LintCategory.DxfIdOutOfRange, """<tableColumn dataDxfId="7">""")
      )
    )
  }

  test("GH-460: streaming mode flags out-of-range dxfIds identically") {
    Vector(danglingDxfParts, inRangeDxfParts, dxfTableParts).foreach { parts =>
      assertEquals(lintStreamOf(parts), lintOf(parts))
    }
    assertEquals(lintStreamOf(dxfTableParts).size, 2)
  }

  // ===== GH-460: package reachability =====

  private val orphanMediaParts: Map[String, String] =
    baseParts + ("xl/media/image9.png" -> "not really a png")

  private val drawingRelsSheetRelsXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>"""

  private val drawingRelsXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/a.png" TargetMode="External"/>
</Relationships>"""

  private val drawingXml =
    """<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>"""

  /** sheet rels → drawing → drawing rels → media: a two-hop closure, no <drawing r:id> needed. */
  private val multiHopParts: Map[String, String] = baseParts ++ Map(
    "xl/worksheets/_rels/sheet1.xml.rels" -> drawingRelsSheetRelsXml,
    "xl/drawings/drawing1.xml" -> drawingXml,
    "xl/drawings/_rels/drawing1.xml.rels" -> drawingRelsXml,
    "xl/media/image1.png" -> "png bytes"
  )

  private val malformedDrawingRelsParts: Map[String, String] =
    multiHopParts + ("xl/drawings/_rels/drawing1.xml.rels" -> "<Relationships><Relationship>")

  test("GH-460: a zip entry reachable from no .rels is flagged as UnreferencedPart") {
    val findings = lintOf(orphanMediaParts)
    assertEquals(findings.map(_.category), Vector(LintCategory.UnreferencedPart))
    val f = findings.head
    assertEquals(f.part, "xl/media/image9.png")
    assertEquals(f.locator, """<Relationship Target="/xl/media/image9.png">""")
    assert(f.message.contains("no relationship"), f.message)
  }

  test("GH-460: parts reached through drawing → media rels are referenced (multi-hop closure)") {
    assertEquals(lintOf(multiHopParts), Vector.empty[Finding])
  }

  test("GH-460: *.rels, [Content_Types].xml and [trash]/ entries are never findings") {
    val parts = baseParts ++ Map(
      "[trash]/0000.dat" -> "excel leftovers",
      "xl/_rels/gone.xml.rels" -> rootRelsXml
    )
    assertEquals(lintOf(parts), Vector.empty[Finding])
  }

  test("GH-460: a malformed .rels inside the closure is one finding, never a Left or a flood") {
    val findings = lintOf(malformedDrawingRelsParts)
    assertEquals(
      findings.map(f => (f.part, f.category)),
      Vector(("xl/drawings/_rels/drawing1.xml.rels", LintCategory.UnreferencedPart))
    )
    assert(findings.head.message.contains("not well-formed"), findings.head.message)
    assertEquals(lintStreamOf(malformedDrawingRelsParts), findings)
  }

  test("GH-460: streaming mode reports unreferenced parts identically") {
    assertEquals(lintStreamOf(orphanMediaParts), lintOf(orphanMediaParts))
    assertEquals(lintStreamOf(orphanMediaParts).size, 1)
    assertEquals(lintStreamOf(multiHopParts), Vector.empty[Finding])
  }

  // ===== GH-567: shared-string entries referenced by no cell =====

  private def sstXml(entries: String*): String =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="$nsMain" count="${entries.size}" uniqueCount="${entries.size}">${entries
        .map(t => s"<si><t>$t</t></si>")
        .mkString}</sst>"""

  private def sstSheetXml(indices: Int*): String = worksheetWith(
    "<sheetData><row r=\"1\">" +
      indices.zipWithIndex.map { (idx, i) =>
        s"""<c r="${('A' + i).toChar}1" t="s"><v>$idx</v></c>"""
      }.mkString +
      "</row></sheetData>"
  )

  private val orphanSstParts: Map[String, String] = baseParts ++ Map(
    "xl/sharedStrings.xml" -> sstXml("alpha", "secret counterparty", "gamma"),
    "xl/worksheets/sheet1.xml" -> sstSheetXml(0, 2)
  )

  private val unionSstParts: Map[String, String] = withSecondSheet(
    "worksheet",
    "worksheets/sheet2.xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
    sstSheetXml(1, 2)
  ) ++ Map(
    "xl/sharedStrings.xml" -> sstXml("alpha", "beta", "gamma"),
    "xl/worksheets/sheet1.xml" -> sstSheetXml(0)
  )

  private val pastTableSstParts: Map[String, String] = baseParts ++ Map(
    "xl/sharedStrings.xml" -> sstXml("alpha", "beta", "gamma"),
    "xl/worksheets/sheet1.xml" -> sstSheetXml(0, 1, 2, 7)
  )

  private val noSstPartParts: Map[String, String] =
    baseParts - "xl/sharedStrings.xml" + ("xl/worksheets/sheet1.xml" -> sstSheetXml(0))

  private val malformedSstParts: Map[String, String] =
    orphanMediaParts + ("xl/sharedStrings.xml" -> "<sst><si>")

  test("GH-567: entries referenced by no t=\"s\" cell are flagged once on xl/sharedStrings.xml") {
    val findings = lintOf(orphanSstParts)
    assertEquals(findings.map(_.category), Vector(LintCategory.SharedStringOrphan))
    val f = findings.head
    assertEquals(f.part, "xl/sharedStrings.xml")
    assertEquals(f.locator, "<si> #1")
    assert(f.message.startsWith("1 of 3 shared string"), f.message)
    assert(f.message.contains("index: 1"), f.message)
    // scrubbed text must never reappear in a lint log: indices only
    assert(!f.message.contains("secret"), f.message)
    assert(!f.message.contains("counterparty"), f.message)
  }

  test("GH-567: references are unioned across sheets — every entry used somewhere is clean") {
    assertEquals(lintOf(unionSstParts), Vector.empty[Finding])
  }

  test("GH-567: a t=\"s\" index past the <si> count is reported on the sheet part") {
    val findings = lintOf(pastTableSstParts)
    assertEquals(findings.map(_.category), Vector(LintCategory.SharedStringOrphan))
    val f = findings.head
    assertEquals(f.part, "xl/worksheets/sheet1.xml")
    assertEquals(f.locator, """<c r="D1" t="s"><v>7</v>""")
    assert(f.message.contains("D1 → 7"), f.message)
    assert(f.message.contains("holds 3 <si>"), f.message)
    assert(f.message.contains("#REF!"), f.message)
  }

  test("GH-567: a t=\"s\" cell in a book with no shared-string part is the same class") {
    val findings = lintOf(noSstPartParts)
    assertEquals(findings.map(_.category), Vector(LintCategory.SharedStringOrphan))
    assertEquals(findings.head.part, "xl/worksheets/sheet1.xml")
    assert(findings.head.message.contains("no shared-string part"), findings.head.message)
    // ... and a book with neither the part nor any t="s" cell has nothing to report
    assertEquals(lintOf(baseParts - "xl/sharedStrings.xml"), Vector.empty[Finding])
  }

  test("GH-567: the orphan sample is bounded to 5 indices; named-styles-excel.xlsx carries 8") {
    // The committed derivative of a stripped Excel model: 10 <si>, 2 referenced — exactly the
    // scrubbing scenario of #567 (the text of the 8 unreferenced entries rides in the package).
    val path = TestFixtures.copyToTemp("named-styles-excel.xlsx")
    val findings = WorkbookLint.lint(path).fold(err => fail(err.message), identity)
    val orphans = findings.filter(_.category == LintCategory.SharedStringOrphan)
    assertEquals(orphans.size, 1, findings.mkString("\n"))
    assertEquals(orphans.head.part, "xl/sharedStrings.xml")
    assert(orphans.head.message.startsWith("8 of 10 shared string"), orphans.head.message)
    assert(
      orphans.head.message.contains("first 5 indices: 1, 3, 4, 5, 6, …"),
      orphans.head.message
    )
    assertEquals(WorkbookLint.lintStream(path), Right(findings))
  }

  test("GH-567: a shared-string part that is not well-formed is one finding, other rules run") {
    val findings = lintOf(malformedSstParts)
    assertEquals(
      findings.map(f => (f.part, f.category)),
      Vector(
        ("xl/sharedStrings.xml", LintCategory.SharedStringOrphan),
        ("xl/media/image9.png", LintCategory.UnreferencedPart)
      )
    )
    assert(findings(0).message.contains("not well-formed"), findings(0).message)
    assertEquals(lintStreamOf(malformedSstParts), findings)
  }

  test("GH-567: streaming mode reports identical shared-string findings") {
    Vector(orphanSstParts, unionSstParts, pastTableSstParts, noSstPartParts).foreach { parts =>
      assertEquals(lintStreamOf(parts), lintOf(parts))
    }
    assertEquals(lintStreamOf(orphanSstParts).size, 1)
  }

  test("shared-string accumulation handles sparse words, duplicates and bounded invalid samples") {
    val count = 4097
    def sheet(indices: Seq[Int]): String = worksheetWith(
      "<sheetData>" + indices.zipWithIndex.map { (idx, row) =>
        s"""<row r="${row + 1}"><c r="A${row + 1}" t="s"><v>$idx</v></c></row>"""
      }.mkString + "</sheetData>"
    )
    // Visit the highest word first, repeat indices across words/sheets, and leave one orphan.
    // Invalid indices must be sampled before they can allocate storage (especially Int.MaxValue).
    val first = (0 until count by 2).reverse ++ Vector(0, 64, 128, 256, 4096) ++
      Vector(-1, count, Int.MaxValue, 5000, 6000, 7000)
    val second = (1 until count - 2 by 2) ++ Vector(64, 128, 4096)
    val parts = withSecondSheet(
      "worksheet",
      "worksheets/sheet2.xml",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
      sheet(second)
    ) ++ Map(
      "xl/sharedStrings.xml" -> sstXml((0 until count).map(i => s"value-$i")*),
      "xl/worksheets/sheet1.xml" -> sheet(first)
    )
    val findings = lintOf(parts)
    assertEquals(lintStreamOf(parts), findings)
    assertEquals(findings.size, 2)
    val invalid =
      findings.find(_.part == "xl/worksheets/sheet1.xml").getOrElse(fail("no invalid indices"))
    assertEquals(invalid.category, LintCategory.SharedStringOrphan)
    assert(invalid.message.startsWith("6 cell(s)"), invalid.message)
    assert(invalid.message.contains(Int.MaxValue.toString), invalid.message)
    assert(!invalid.message.contains("7000"), invalid.message)
    val orphan = findings.find(_.part == "xl/sharedStrings.xml").getOrElse(fail("no orphan"))
    assertEquals(orphan.category, LintCategory.SharedStringOrphan)
    assert(orphan.message.startsWith("1 of 4097 shared string"), orphan.message)
    assert(orphan.message.contains("index: 4095"), orphan.message)
  }

  test("GH-567: a fresh SST-dialect write by xl lints clean (every entry is referenced)") {
    // >10 text cells with duplicates → SstPolicy.Auto emits a shared-string table
    val cells = (1 to 12).map { i =>
      com.tjclp.xl.addressing.ARef.parse(s"A$i").fold(fail(_), identity) -> s"v${i % 3}"
    }
    val sheet = cells.foldLeft(Sheet("Data")) { case (s, (r, v)) => s.put(r -> v) }
    val bytes =
      XlsxWriter.writeToBytes(Workbook(Vector(sheet))).fold(e => fail(e.message), identity)
    assert(
      new String(bytes, StandardCharsets.ISO_8859_1).contains("sharedStrings.xml"),
      "the write must have chosen the SST dialect for this test to mean anything"
    )
    assertEquals(WorkbookLint.lintBytes(bytes), Right(Vector.empty[Finding]))
    assertEquals(WorkbookLint.lintStreamBytes(bytes), Right(Vector.empty[Finding]))
  }

  test("GH-567 carve-out: replacing text in a foreign SST book leaves an orphan the lint reports") {
    // The documented outcome of the lint-only decision: xl appends the new string and re-points
    // the cell, but never prunes a preserved table — the replaced text stays in the package and
    // `shared-string-orphan` is the ONE finding on xl's own output. (No SST compaction this wave.)
    val source = baseParts ++ Map(
      "xl/sharedStrings.xml" -> sstXml("plain", "other"),
      "xl/worksheets/sheet1.xml" -> sstSheetXml(0, 1)
    )
    assertEquals(lintOf(source), Vector.empty[Finding])
    val edited = for
      wb <- XlsxReader.readFromBytes(zipBytes(source))
      sheet <- wb("Sheet1")
      out <- XlsxWriter.writeToBytes(wb.put(sheet.put(ref"A1" -> "REPLACED")))
    yield out
    val bytes = edited.fold(err => fail(s"read/modify/write failed: $err"), identity)
    val sst = readZipEntry(bytes, "xl/sharedStrings.xml")
    assert(sst.contains("<t>plain</t>"), sst) // the replaced text is still in the package
    assert(sst.contains("""uniqueCount="3""""), sst)
    val findings = WorkbookLint.lintBytes(bytes).fold(err => fail(err.message), identity)
    assertEquals(
      findings.map(_.category),
      Vector(LintCategory.SharedStringOrphan),
      findings.toString
    )
    assertEquals(findings.head.part, "xl/sharedStrings.xml")
    assert(findings.head.message.startsWith("1 of 3 shared string"), findings.head.message)
    assert(findings.head.message.contains("index: 0"), findings.head.message)
  }

  private def readZipEntry(bytes: Array[Byte], entry: String): String =
    val zip = new java.util.zip.ZipInputStream(new java.io.ByteArrayInputStream(bytes))
    try
      Iterator
        .continually(zip.getNextEntry)
        .takeWhile(_ != null)
        .find(_.getName == entry)
        .map(_ => new String(zip.readAllBytes(), StandardCharsets.UTF_8))
        .getOrElse(fail(s"missing $entry"))
    finally zip.close()

  // ===== GH-460 / GH-567: the committed corpus and xl's own edits of it =====

  private val newCategories: Set[LintCategory] = Set(
    LintCategory.EmptyInlineStr,
    LintCategory.IgnorableUndeclared,
    LintCategory.DxfIdOutOfRange,
    LintCategory.UnreferencedPart,
    LintCategory.SharedStringOrphan
  )

  test("GH-460/567: over the committed corpus, only named-styles-excel trips a new rule") {
    // false-positive guard: every well-formed fixture (openpyxl, LibreOffice, Excel, derived)
    TestFixtures.all.foreach { name =>
      val path = TestFixtures.copyToTemp(name)
      val findings = WorkbookLint.lint(path).fold(err => fail(s"$name: ${err.message}"), identity)
      val fresh = findings.filter(f => newCategories.contains(f.category))
      val expected =
        if name == "named-styles-excel.xlsx" then Vector(LintCategory.SharedStringOrphan)
        else Vector.empty[LintCategory]
      assertEquals(fresh.map(_.category), expected, s"$name: ${fresh.mkString("\n")}")
      assertEquals(WorkbookLint.lintStream(path), Right(findings), name)
    }
  }

  test("GH-460: a regenerated Excel-authored sheet keeps its mc:Ignorable prefixes declared") {
    // Excel roots carry mc:Ignorable="x14ac xr xr2 xr3": the writer must re-declare every one on
    // the regenerated root, or the new rule would flag xl's own output. SharedStringOrphan is
    // excluded here because the FOREIGN source already carries 8 orphans (see the fixture test).
    val path = TestFixtures.copyToTemp("named-styles-excel.xlsx")
    val edited = for
      wb <- XlsxReader.read(path)
      sheet <- wb.sheets.headOption.toRight(
        com.tjclp.xl.error.XLError.ParseError(path.toString, "no sheet")
      )
      out <- XlsxWriter.writeToBytes(wb.put(sheet.put(ref"Z1" -> "edited")))
    yield out
    val bytes = edited.fold(err => fail(s"read/modify/write failed: $err"), identity)
    val findings = WorkbookLint.lintBytes(bytes).fold(err => fail(err.message), identity)
    assertEquals(
      findings.filter(f => newCategories.contains(f.category)).map(_.category),
      Vector(LintCategory.SharedStringOrphan),
      findings.mkString("\n")
    )
  }

  // ===== Severity tiers (PR #659 review): repair vs hygiene =====

  test("severity: orphan shared strings and unreferenced parts are hygiene, the rest repair") {
    assertEquals(lintOf(orphanSstParts).map(_.severity), Vector(LintSeverity.Hygiene))
    assertEquals(lintOf(orphanMediaParts).map(_.severity), Vector(LintSeverity.Hygiene))
    // the past-the-table half of the shared-string rule is a repair (Excel repairs, xl reads #REF!)
    assertEquals(lintOf(pastTableSstParts).map(_.severity), Vector(LintSeverity.Repair))
    assertEquals(lintOf(danglingDxfParts).map(_.severity), Vector(LintSeverity.Repair))
    // a generated-prefix root opens intact in Excel and LibreOffice: hygiene, not repair
    assertEquals(
      lintOf(baseParts + ("xl/worksheets/sheet1.xml" -> elementTreeRootSheetXml)).map(_.severity),
      Vector(LintSeverity.Hygiene)
    )
    // the streaming scanner assigns the same tier
    assertEquals(lintStreamOf(orphanSstParts).map(_.severity), Vector(LintSeverity.Hygiene))
    assertEquals(lintStreamOf(pastTableSstParts).map(_.severity), Vector(LintSeverity.Repair))
  }

  test("severity: slugs are the stable spellings the CLI publishes") {
    assertEquals(LintSeverity.Repair.slug, "repair")
    assertEquals(LintSeverity.Hygiene.slug, "hygiene")
    // a Finding built without a tier is a repair — the conservative default for every rule that
    // does not say otherwise
    assertEquals(Finding("p", LintCategory.ChildOrder, "<x>", "m").severity, LintSeverity.Repair)
  }

  // ===== Macro sheets (Excel 4.0 XLM, PR #659 review): a sheet kind of their own =====

  private val nsXm = "http://schemas.microsoft.com/office/excel/2006/main"
  private val relTypeMacrosheet =
    "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet"
  private val macrosheetCt = "application/vnd.ms-excel.macrosheet+xml"

  /** As Excel writes it: an `xm:macrosheet` root whose children are main-namespace elements. */
  private val macrosheetXml =
    s"""<?xml version="1.0" encoding="UTF-8"?>
<xm:macrosheet xmlns="$nsMain" xmlns:r="$nsRel" xmlns:xm="$nsXm">
  <sheetPr codeName="Macro1"/>
  <dimension ref="A1"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row></sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</xm:macrosheet>"""

  /** Sheet1 references SST entry 0; the macrosheet is the ONLY reference to entry 1. */
  private def macrosheetPartsWith(sheetXml: String): Map[String, String] = baseParts ++ Map(
    "[Content_Types].xml" -> contentTypesXml.replace(
      "</Types>",
      s"""  <Override PartName="/xl/macrosheets/sheet1.xml" ContentType="$macrosheetCt"/>\n</Types>"""
    ),
    "xl/workbook.xml" -> workbookXml.replace(
      "</sheets>",
      "  <sheet name=\"Macro1\" sheetId=\"2\" r:id=\"rId5\"/>\n  </sheets>"
    ),
    "xl/_rels/workbook.xml.rels" -> workbookRelsXml.replace(
      "</Relationships>",
      s"""  <Relationship Id="rId5" Type="$relTypeMacrosheet" Target="macrosheets/sheet1.xml"/>\n</Relationships>"""
    ),
    "xl/sharedStrings.xml" -> sstXml("alpha", "macro only"),
    "xl/worksheets/sheet1.xml" -> sstSheetXml(0),
    "xl/macrosheets/sheet1.xml" -> sheetXml
  )

  private val macrosheetParts: Map[String, String] = macrosheetPartsWith(macrosheetXml)

  private val misorderedMacrosheetXml = macrosheetXml
    .replace("  <dimension ref=\"A1\"/>\n", "")
    .replace("</sheetData>\n", "</sheetData>\n  <dimension ref=\"A1\"/>\n")

  test(
    "macrosheet: an xlMacrosheet <sheet> is a sheet kind — its t=\"s\" cells count, no finding"
  ) {
    // Before the kind existed the workbook-level check called the rel wrong-typed ("expected
    // worksheet") and the SST rule reported entry 1 as an orphan — on a book Excel writes itself.
    assertEquals(lintOf(macrosheetParts), Vector.empty[Finding])
    assertEquals(lintStreamOf(macrosheetParts), Vector.empty[Finding])
  }

  test("macrosheet: child order is checked against CT_Macrosheet in both modes") {
    val parts = macrosheetPartsWith(misorderedMacrosheetXml)
    val findings = lintOf(parts)
    assertEquals(
      findings.map(f => (f.part, f.category)),
      Vector(("xl/macrosheets/sheet1.xml", LintCategory.ChildOrder))
    )
    assert(findings.head.message.contains("CT_Macrosheet"), findings.head.message)
    assertEquals(lintStreamOf(parts), findings)
  }

  test("macrosheet: a sheet rel of a genuinely foreign type is still wrong-rel-type") {
    val parts = macrosheetParts + ("xl/_rels/workbook.xml.rels" -> workbookRelsXml.replace(
      "</Relationships>",
      s"""  <Relationship Id="rId5" Type="$nsRel/styles" Target="macrosheets/sheet1.xml"/>\n</Relationships>"""
    ))
    assert(
      lintOf(parts).exists(f =>
        f.part == "xl/workbook.xml" && f.category == LintCategory.WrongRelType
      ),
      lintOf(parts).toString
    )
  }

  // ===== GH-460 / PR #659 review: dxfId is read only where the schema puts one =====

  private val cellDxfSheetXml = worksheetWith(
    """<sheetData><row r="1"><c r="A1" dxfId="99"><v>1</v></c></row></sheetData>"""
  )

  private val cellDxfParts: Map[String, String] = baseParts ++ Map(
    "xl/styles.xml" -> oneDxfStylesXml,
    "xl/worksheets/sheet1.xml" -> cellDxfSheetXml
  )

  private val sortConditionDxfSheetXml = worksheetWith(
    """<sheetData/>
  <autoFilter ref="A1:A5"><sortState ref="A2:A5"><sortCondition ref="A2:A5" dxfId="4"/></sortState></autoFilter>"""
  )

  private val sortConditionDxfParts: Map[String, String] = baseParts ++ Map(
    "xl/styles.xml" -> oneDxfStylesXml,
    "xl/worksheets/sheet1.xml" -> sortConditionDxfSheetXml
  )

  test("GH-460: a dxfId on an element that cannot carry one (a cell) is not a finding") {
    // The scanners probe the dxf attributes only on cfRule / sortCondition / table / tableColumn,
    // so the million <c>/<v>/<f> of a large sheet cost no attribute lookups (PR #659 review)
    assertEquals(lintOf(cellDxfParts), Vector.empty[Finding])
    assertEquals(lintStreamOf(cellDxfParts), Vector.empty[Finding])
  }

  test("GH-460: sortCondition dxfId past the table is flagged in both modes") {
    val findings = lintOf(sortConditionDxfParts)
    assertEquals(findings.map(_.category), Vector(LintCategory.DxfIdOutOfRange))
    assert(findings.head.locator.contains("sortCondition"), findings.head.locator)
    assertEquals(lintStreamOf(sortConditionDxfParts), findings)
  }

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
</worksheet>""")),
    // GH-442: the data-table checks read sheetData cells, so every record shape needs parity
    "clean seeded data table" -> dtParts(seededDataTableSheetXml),
    "torn data-table interior" -> dtParts(tornInteriorSheetXml),
    "del1 data-table record" -> dtParts(delInputSheetXml),
    "input-less data-table record" -> dtParts(missingInputSheetXml),
    "no-room data-table record" -> dtParts(noRoomDataTableSheetXml),
    "overwritten data-table corner" -> dtParts(overwrittenCornerSheetXml),
    "unseeded autoNoTable data table" -> autoNoTableParts(unseededDataTableSheetXml),
    "partly seeded autoNoTable data table" -> autoNoTableParts(partlySeededDataTableSheetXml),
    "torn autoNoTable data table" -> autoNoTableParts(tornInteriorSheetXml),
    // GH-456: <f> text observation must agree between the DOM and SAX scanners, including the
    // aggregated per-part shape (sample order + total count)
    "leading-equals formula" ->
      (baseParts + ("xl/worksheets/sheet1.xml" -> leadingEqualsSheetXml)),
    "many leading-equals formulas" ->
      (baseParts + ("xl/worksheets/sheet1.xml" -> manyLeadingEqualsSheetXml)),
    "clean formula text" -> (baseParts + ("xl/worksheets/sheet1.xml" -> cleanFormulaSheetXml)),
    // GH-525: external-ordinal observation must agree between the DOM and SAX scanners,
    // including the per-ordinal aggregation (counts + first-cell locators)
    "dangling external ordinals" ->
      (externalParts + ("xl/worksheets/sheet1.xml" -> danglingExternalRefSheetXml)),
    "clean external ordinals" ->
      (externalParts + ("xl/worksheets/sheet1.xml" -> externalOkSheetXml)),
    "bracket noise formulas" ->
      (externalParts + ("xl/worksheets/sheet1.xml" -> bracketNoiseSheetXml)),
    "dangling defined-name ordinal" ->
      (externalParts + ("xl/workbook.xml" -> danglingNameWorkbookXml)),
    // GH-528: defined-name validity is workbook-level (always DOM) but must agree in both modes
    "fold-duplicate defined names" ->
      (baseParts + ("xl/workbook.xml" -> foldDupWorkbookXml)),
    // GH-588: formula-text observation (<f>, <formula>, <formula1>/<formula2>, <xm:f>) must agree
    // between the DOM and SAX scanners, including the per-part aggregation
    "bare xlfn everywhere" -> (baseParts + ("xl/worksheets/sheet1.xml" -> bareXlfnSheetXml)),
    "bare xlfn cf rule" -> (baseParts + ("xl/worksheets/sheet1.xml" -> bareCfIfsSheetXml)),
    "many bare xlfn cells" ->
      (baseParts + ("xl/worksheets/sheet1.xml" -> manyBareXlfnSheetXml)),
    "prefixed xlfn everywhere" ->
      (baseParts + ("xl/worksheets/sheet1.xml" -> prefixedXlfnSheetXml)),
    "bare xlfn defined name" -> (baseParts + ("xl/workbook.xml" -> bareNameWorkbookXml)),
    "half-prefixed LET cell" ->
      (baseParts + ("xl/worksheets/sheet1.xml" -> halfPrefixedLetSheetXml)),
    // the root element is never a formula-text site, in either scanner
    "root-level formula element" ->
      (baseParts + ("xl/worksheets/sheet1.xml" -> rootFormulaSheetXml)),
    // GH-529: suffix-less Microsoft externalLinkPath variants resolve clean in both modes
    "ms xlStartup external" ->
      (externalParts +
        ("xl/externalLinks/_rels/externalLink1.xml.rels" -> msVariantRelsXml("xlStartup"))),
    // GH-458: Microsoft xlExternalLinkPath variants resolve clean in both modes
    "ms xlPathMissing external" ->
      (externalParts +
        ("xl/externalLinks/_rels/externalLink1.xml.rels" -> msVariantRelsXml("xlPathMissing"))),
    // GH-460: <c t=...>/<is>/<v> observation (empty inline strings) must agree between scanners
    "empty inline strings" -> (baseParts + ("xl/worksheets/sheet1.xml" -> emptyInlineSheetXml)),
    "many empty inline strings" ->
      (baseParts + ("xl/worksheets/sheet1.xml" -> manyEmptyInlineSheetXml)),
    "textful inline strings" ->
      (baseParts + ("xl/worksheets/sheet1.xml" -> textfulInlineSheetXml)),
    // GH-460: mc:Ignorable prefix resolution (DOM scope chain vs SAX prefix-mapping stack)
    "undeclared mc:Ignorable" ->
      (baseParts + ("xl/worksheets/sheet1.xml" -> undeclaredIgnorableSheetXml)),
    "declared mc:Ignorable" ->
      (baseParts + ("xl/worksheets/sheet1.xml" -> declaredIgnorableSheetXml)),
    "nested mc:Ignorable" -> (baseParts + ("xl/worksheets/sheet1.xml" -> nestedIgnorableSheetXml)),
    "ns0-prefixed root" -> (baseParts + ("xl/worksheets/sheet1.xml" -> elementTreeRootSheetXml)),
    "undeclared mc:Ignorable workbook" ->
      (baseParts + ("xl/workbook.xml" -> undeclaredIgnorableWorkbookXml)),
    "undeclared mc:Ignorable table" -> undeclaredIgnorableTableParts,
    // GH-460: dxfId attribute capture on sheet-class and table parts
    "dangling cfRule dxfId" -> danglingDxfParts,
    "in-range cfRule dxfId" -> inRangeDxfParts,
    "dangling table dxfIds" -> dxfTableParts,
    "dxfId on a cell (not a dxf bearer)" -> cellDxfParts,
    "dangling sortCondition dxfId" -> sortConditionDxfParts,
    // PR #659 review: XLM macro sheets are a sheet kind (t="s" cells feed the SST union)
    "clean macrosheet" -> macrosheetParts,
    "misordered macrosheet" -> macrosheetPartsWith(misorderedMacrosheetXml),
    // GH-460: package reachability is mode-independent but must agree
    "orphan media part" -> orphanMediaParts,
    "multi-hop rels closure" -> multiHopParts,
    "malformed drawing rels" -> malformedDrawingRelsParts,
    // GH-567: t="s" <v> index capture (SAX buffers <v> text only for t="s" cells)
    "orphan shared strings" -> orphanSstParts,
    "unioned shared strings" -> unionSstParts,
    "past-table shared-string index" -> pastTableSstParts,
    "t=s without a shared-string part" -> noSstPartParts,
    "malformed shared strings" -> malformedSstParts
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

  // ===== GH-663: <f> text the formula oracle rejects =====

  /**
   * The stand-in for the CLI's parser: xl-ooxml cannot see xl-evaluator, so the rule takes the
   * oracle as a parameter. Unbalanced parentheses are the class the field incident carried.
   */
  private val parenCheck: WorkbookLint.FormulaCheck = text =>
    if text.count(_ == '(') != text.count(_ == ')') then Some(s"unbalanced parentheses in '$text'")
    else None

  private val rejectAll: WorkbookLint.FormulaCheck = text => Some(s"rejected '$text'")

  private val unparseableSheetXml = worksheetWith(
    """<sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row>
    <row r="3"><c r="A3"><f>SUM(A1:A2</f><v>3</v></c></row>
  </sheetData>"""
  )

  /** Seven rejected formulas in document order A1, B1, C1, A2, B2, C2, A3. */
  private val manyUnparseableSheetXml = worksheetWith(
    """<sheetData>
    <row r="1"><c r="A1"><f>SUM(1</f></c><c r="B1"><f>SUM(2</f></c><c r="C1"><f>SUM(3</f></c></row>
    <row r="2"><c r="A2"><f>SUM(4</f></c><c r="B2"><f>SUM(5</f></c><c r="C2"><f>SUM(6</f></c></row>
    <row r="3"><c r="A3"><f>SUM(7</f></c></row>
  </sheetData>"""
  )

  /**
   * The shapes the rule must never judge: a shared-formula dependent (empty `<f>`), an array
   * formula (ordinary text, judged like any other), a leading-'=' formula (judged WITHOUT the '='
   * so GH-456 stays the only finding), and a data-table record (its text is display, GH-430).
   */
  private val formulaShapesSheetXml = worksheetWith(
    """<sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c></row>
    <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>
    <row r="3"><c r="C3"><f t="array" ref="C3:C4">A1:A2*2</f><v>2</v></c></row>
    <row r="4"><c r="D4"><f>=A1*2</f><v>2</v></c></row>
  </sheetData>"""
  )

  private def unparseableOf(findings: Vector[Finding]): Vector[Finding] =
    findings.filter(_.category == LintCategory.FormulaUnparseable)

  test("GH-663: <f> text the oracle rejects is a repair-tier FormulaUnparseable finding") {
    val parts = baseParts + ("xl/worksheets/sheet1.xml" -> unparseableSheetXml)
    val findings = WorkbookLint
      .lintBytes(zipBytes(parts), parenCheck)
      .fold(err => fail(s"lint must not error: $err"), identity)
    assertEquals(findings.map(_.category), Vector(LintCategory.FormulaUnparseable))
    val f = findings.head
    assertEquals(f.severity, LintSeverity.Repair)
    assertEquals(f.part, "xl/worksheets/sheet1.xml")
    assert(f.locator.contains("A3"), s"locator should carry the offending cell: $f")
    assert(f.message.contains("1 formula(s)"), s"message should carry the total count: $f")
    assert(f.message.contains("A3"), s"message should name the offending cell: $f")
    assert(f.message.contains("SUM(A1:A2"), s"message should quote the stored text: $f")
    assert(
      f.message.contains("unbalanced parentheses in 'SUM(A1:A2'"),
      s"message should carry the oracle's diagnostic: $f"
    )
    assert(f.message.contains("repair"), s"message should say Excel repairs the file: $f")
  }

  test("GH-663: rejected formulas aggregate to ONE finding per part (first 5 + total count)") {
    val parts = baseParts + ("xl/worksheets/sheet1.xml" -> manyUnparseableSheetXml)
    val findings = WorkbookLint
      .lintBytes(zipBytes(parts), parenCheck)
      .fold(err => fail(s"lint must not error: $err"), identity)
    assertEquals(findings.map(_.category), Vector(LintCategory.FormulaUnparseable))
    val f = findings.head
    assert(f.locator.contains("A1"), s"locator should carry the first offending cell: $f")
    assert(f.message.contains("7 formula(s)"), s"message should carry the total count: $f")
    assert(
      f.message.contains("first 5: A1, B1, C1, A2, B2"),
      s"message should sample the first 5 offending cells in document order: $f"
    )
    assert(!f.message.contains("C2"), s"message must not carry cells past the sample: $f")
    assert(!f.message.contains("A3"), s"message must not carry cells past the sample: $f")
  }

  test("GH-663: without an oracle the rule is off — the default entry points stay unchanged") {
    val parts = baseParts + ("xl/worksheets/sheet1.xml" -> unparseableSheetXml)
    assertEquals(lintOf(parts), Vector.empty[Finding])
    assertEquals(lintStreamOf(parts), Vector.empty[Finding])
  }

  test("GH-663: streaming mode flags the rejected formula identically") {
    val parts = baseParts + ("xl/worksheets/sheet1.xml" -> manyUnparseableSheetXml)
    val bytes = zipBytes(parts)
    val dom = WorkbookLint.lintBytes(bytes, parenCheck).fold(e => fail(s"$e"), identity)
    val sax = WorkbookLint.lintStreamBytes(bytes, parenCheck).fold(e => fail(s"$e"), identity)
    assertEquals(sax, dom)
    assertEquals(sax.map(_.category), Vector(LintCategory.FormulaUnparseable))
  }

  test("GH-663: shared dependents and data-table records are never judged; '=' is stripped first") {
    // An oracle that rejects everything it is shown: whatever survives was never shown to it.
    val shapes = baseParts + ("xl/worksheets/sheet1.xml" -> formulaShapesSheetXml)
    val shown = Vector.newBuilder[String]
    val recording: WorkbookLint.FormulaCheck = text =>
      shown += text
      None
    val bytes = zipBytes(shapes)
    val judged = WorkbookLint.lintBytes(bytes, recording).fold(e => fail(s"$e"), identity)
    assertEquals(judged.map(_.category), Vector(LintCategory.FormulaLeadingEquals))
    // the shared master, the array formula and the '='-stripped display form; NOT the empty
    // dependent
    assertEquals(shown.result(), Vector("A1*2", "A1:A2*2", "A1*2"))
    val streamShown = Vector.newBuilder[String]
    val streamRecording: WorkbookLint.FormulaCheck = text =>
      streamShown += text
      None
    WorkbookLint.lintStreamBytes(bytes, streamRecording).fold(e => fail(s"$e"), identity)
    assertEquals(streamShown.result(), Vector("A1*2", "A1:A2*2", "A1*2"))

    // A data-table record's text is display, never formula text (GH-430): the oracle is not asked
    // even when the record carries LibreOffice's `TABLE(…)` spelling; the corner formula still is
    val recordWithText = dtRecord.stripSuffix("/>") + ">TABLE(B1,B2)</f>"
    val dtShown = Vector.newBuilder[String]
    val dtRecording: WorkbookLint.FormulaCheck = text =>
      dtShown += text
      if text.contains("TABLE") then Some("display text") else None
    val dt = zipBytes(dtParts(seededDataTableSheetXml.replace(dtRecord, recordWithText)))
    assertEquals(
      unparseableOf(WorkbookLint.lintBytes(dt, dtRecording).fold(e => fail(s"$e"), identity)),
      Vector.empty[Finding]
    )
    assertEquals(dtShown.result(), Vector("B1*B2"))
    assertEquals(
      unparseableOf(WorkbookLint.lintStreamBytes(dt, dtRecording).fold(e => fail(s"$e"), identity)),
      Vector.empty[Finding]
    )
    assertEquals(dtShown.result(), Vector("B1*B2", "B1*B2"))
    // and an oracle that rejects everything still never sees the record
    assertEquals(
      unparseableOf(WorkbookLint.lintBytes(dt, rejectAll).fold(e => fail(s"$e"), identity))
        .map(_.locator),
      Vector("""<c r="C4"><f>""")
    )
  }

  test("GH-663: the slug is formula-unparseable") {
    assertEquals(LintCategory.FormulaUnparseable.slug, "formula-unparseable")
  }
