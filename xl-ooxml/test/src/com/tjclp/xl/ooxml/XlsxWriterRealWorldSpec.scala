package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import munit.FunSuite

/**
 * Real-world integration test for surgical modification.
 *
 * This test validates the surgical modification system against a complex workbooks
 * that mimics real-world Excel files with:
 *   - Multiple sheets (some hidden)
 *   - Charts and drawings
 *   - Conditional formatting
 *   - Named ranges (defined names)
 *   - Theme colors in styles
 *   - Unknown parts (comments, custom XML)
 *   - Rich row attributes
 *   - Shared strings table
 *   - Differential formats (dxfs)
 *
 * **Success Criteria (All must pass)**:
 *   1. No Excel corruption warning when opening output
 *   2. Unknown parts (charts, drawings) preserved byte-for-byte
 *   3. Defined names preserved
 *   4. Conditional formatting works
 *   5. Hidden sheets remain hidden
 *   6. Theme colors correct (not black)
 *   7. Modified cell has new content
 *   8. Unmodified sheets byte-identical
 *
 * This is the PRIMARY validation test for the surgical modification system.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class XlsxWriterRealWorldSpec extends FunSuite:

  test("complex workbooks with charts, conditionalFormatting, hidden sheets, named ranges") {
    // This test validates ALL 26 fixes from the surgical-mod branch against
    // a single comprehensive real-world-like workbooks

    val source = createComplexWorkbook()

    // Modify one cell in Sheet1, leave everything else untouched
    val modified = for
      wb <- XlsxReader.read(source)
      sheet <- wb("Sheet1")
      updatedSheet <- sheet.put(ref"A1" -> "Modified by XL")
      updated <- wb.put(updatedSheet)
    yield updated

    val wb = modified.fold(err => fail(s"Failed to modify: $err"), identity)

    assert(wb.sourceContext.isDefined, "Workbook should have SourceContext")
    assert(!wb.sourceContext.get.isClean, "Modified workbooks should be dirty")

    // Verify only Sheet1 modified
    val tracker = wb.sourceContext.get.modificationTracker
    assertEquals(tracker.modifiedSheets, Set(0), "Only sheets 0 should be modified")

    // Write back
    val output = Files.createTempFile("complex-real-world", ".xlsx")
    XlsxWriter
      .write(wb, output)
      .fold(err => fail(s"Failed to write: $err"), identity)

    // ===== Validation 1: Unknown Parts Preserved Byte-for-Byte =====
    val sourceZip = new ZipFile(source.toFile)
    val outputZip = new ZipFile(output.toFile)

    // Chart should be preserved
    val sourceChart = readEntryBytes(sourceZip, sourceZip.getEntry("xl/charts/chart1.xml"))
    val outputChart = readEntryBytes(outputZip, outputZip.getEntry("xl/charts/chart1.xml"))
    assertEquals(
      outputChart.toSeq,
      sourceChart.toSeq,
      "Chart should be byte-identical (unknown part preservation)"
    )

    // Drawing should be preserved
    val sourceDrawing = readEntryBytes(sourceZip, sourceZip.getEntry("xl/drawings/drawing1.xml"))
    val outputDrawing = readEntryBytes(outputZip, outputZip.getEntry("xl/drawings/drawing1.xml"))
    assertEquals(
      outputDrawing.toSeq,
      sourceDrawing.toSeq,
      "Drawing should be byte-identical (unknown part preservation)"
    )

    // Comments should be preserved
    val sourceComments = readEntryBytes(sourceZip, sourceZip.getEntry("xl/comments1.xml"))
    val outputComments = readEntryBytes(outputZip, outputZip.getEntry("xl/comments1.xml"))
    assertEquals(
      outputComments.toSeq,
      sourceComments.toSeq,
      "Comments should be byte-identical (unknown part preservation)"
    )

    // ===== Validation 2: Defined Names Preserved =====
    val workbookXml = readEntryString(outputZip, outputZip.getEntry("xl/workbook.xml"))
    assert(
      workbookXml.contains("<definedNames>"),
      "Defined names section missing"
    )
    assert(
      workbookXml.contains("MyNamedRange"),
      "Named range 'MyNamedRange' missing"
    )
    assert(
      workbookXml.contains("SalesTotal"),
      "Named range 'SalesTotal' missing"
    )

    // ===== Validation 3: Conditional Formatting Preserved =====
    val sheet1Xml = readEntryString(outputZip, outputZip.getEntry("xl/worksheets/sheet1.xml"))
    assert(
      sheet1Xml.contains("<conditionalFormatting"),
      "Conditional formatting missing from Sheet1"
    )
    assert(
      sheet1Xml.contains("dataBar") || sheet1Xml.contains("cellIs"),
      "Conditional formatting rules missing"
    )

    // ===== Validation 4: Hidden Sheets Preserved =====
    assert(
      workbookXml.contains("""state="hidden""""),
      "Hidden sheets state missing (Sheet3 should be hidden)"
    )

    // ===== Validation 5: Theme Colors Preserved =====
    val stylesXml = readEntryString(outputZip, outputZip.getEntry("xl/styles.xml"))
    assert(
      stylesXml.contains("theme=\"0\"") || stylesXml.contains("theme=\"1\""),
      "Theme colors missing - would cause black fill corruption"
    )
    assert(
      stylesXml.contains("tint="),
      "Theme color tints missing"
    )

    // ===== Validation 6: dxfs (Differential Formats) Preserved =====
    assert(
      stylesXml.contains("<dxfs"),
      "dxfs section missing (conditional formatting styles lost)"
    )

    // ===== Validation 7: Namespace Preservation (mc:Ignorable) =====
    assert(
      workbookXml.contains("mc:Ignorable"),
      "mc:Ignorable attribute missing - PRIMARY corruption trigger"
    )

    // ===== Validation 8: SST Correctness =====
    val sstXml = readEntryString(outputZip, outputZip.getEntry("xl/sharedStrings.xml"))
    val countMatch = """count="(\d+)"""".r.findFirstMatchIn(sstXml)
    val uniqueCountMatch = """uniqueCount="(\d+)"""".r.findFirstMatchIn(sstXml)

    assert(countMatch.isDefined, "SST count attribute missing")
    assert(uniqueCountMatch.isDefined, "SST uniqueCount attribute missing")

    val count = countMatch.get.group(1).toInt
    val uniqueCount = uniqueCountMatch.get.group(1).toInt

    assert(
      count >= uniqueCount,
      s"SST count=$count should be >= uniqueCount=$uniqueCount"
    )

    // ===== Validation 9: Row Attributes Preserved =====
    assert(
      sheet1Xml.contains("spans="),
      "Row spans attribute missing"
    )
    assert(
      sheet1Xml.contains("dyDescent="),
      "Row dyDescent namespaced attribute missing"
    )

    // ===== Validation 10: Modified Cell Present =====
    // Note: Cell uses SST reference, so check SST instead of sheets XML
    assert(
      sstXml.contains("Modified by XL"),
      "Modified cell content missing from SST"
    )

    // ===== Validation 11: Unmodified Sheets Byte-Identical =====
    val sourceSheet2 = readEntryBytes(sourceZip, sourceZip.getEntry("xl/worksheets/sheet2.xml"))
    val outputSheet2 = readEntryBytes(outputZip, outputZip.getEntry("xl/worksheets/sheet2.xml"))
    assertEquals(
      outputSheet2.toSeq,
      sourceSheet2.toSeq,
      "Unmodified Sheet2 should be byte-identical"
    )

    // ===== Validation 12: No Excel Corruption (via round-trip) =====
    // If Excel would show a corruption warning, the file would fail to re-read
    val reloaded = XlsxReader
      .read(output)
      .fold(err => fail(s"Round-trip reload failed (indicates Excel would reject): $err"), identity)

    assertEquals(reloaded.sheets.size, 3, "Should have 3 sheets after round-trip")

    val cell = reloaded("Sheet1")
      .flatMap(s => Right(s.cells.get(ref"A1")))
      .fold(err => fail(s"Failed to get modified cell: $err"), identity)

    assertEquals(
      cell.map(_.value),
      Some(CellValue.Text("Modified by XL")),
      "Modified cell should survive round-trip"
    )

    // ===== Validation 13: Surgical Stats =====
    val partManifest = wb.sourceContext.get.partManifest
    val unparsedCount = partManifest.unparsedParts.size

    assert(
      unparsedCount >= 3,
      s"Should have at least 3 unparsed parts (chart, drawing, comments), found $unparsedCount"
    )

    // Clean up
    sourceZip.close()
    outputZip.close()
    Files.deleteIfExists(source)
    Files.deleteIfExists(output)
  }

  // ===== Helper Methods =====

  private def readEntryBytes(zip: ZipFile, entry: ZipEntry): Array[Byte] =
    val is = zip.getInputStream(entry)
    try is.readAllBytes()
    finally is.close()

  private def readEntryString(zip: ZipFile, entry: ZipEntry): String =
    new String(readEntryBytes(zip, entry), "UTF-8")

  // ===== Complex Workbook Creator =====

  private def createComplexWorkbook(): Path =
    val path = Files.createTempFile("test-complex", ".xlsx")
    val out = new ZipOutputStream(Files.newOutputStream(path))
    out.setLevel(1) // Match production compression level

    try
      // Content Types (comprehensive)
      writeEntry(
        out,
        "[Content_Types].xml",
        """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheets.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
</Types>"""
      )

      // Root rels
      writeEntry(
        out,
        "_rels/.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
      )

      // Workbook with namespaces, defined names, hidden sheets
      writeEntry(
        out,
        "xl/workbook.xml",
        """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" mc:Ignorable="x15">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
    <sheet name="Sheet3" sheetId="3" state="hidden" r:id="rId3"/>
  </sheets>
  <definedNames>
    <definedName name="MyNamedRange">Sheet1!$A$1:$A$10</definedName>
    <definedName name="SalesTotal">Sheet2!$B$5</definedName>
  </definedNames>
</workbook>"""
      )

      // Workbook rels
      writeEntry(
        out,
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""
      )

      // Styles with theme colors, dxfs, multiple formats
      writeEntry(
        out,
        "xl/styles.xml",
        """<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
    <font><i/><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor theme="0" tint="-0.04999"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left styles="thin"><color theme="1" tint="0.5"/></left><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
  <dxfs count="2">
    <dxf>
      <font><color rgb="FF9C0006"/></font>
      <fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill>
    </dxf>
    <dxf>
      <font><color rgb="FF006100"/></font>
      <fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill>
    </dxf>
  </dxfs>
</styleSheet>"""
      )

      // SharedStrings
      writeEntry(
        out,
        "xl/sharedStrings.xml",
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="8" uniqueCount="5">
  <si><t>Original Text</t></si>
  <si><t>Shared String</t></si>
  <si>
    <r><rPr><b/></rPr><t>Bold</t></r>
    <r><t xml:space="preserve"> </t></r>
    <r><rPr><i/></rPr><t>Italic</t></r>
  </si>
  <si><t>Data Point 1</t></si>
  <si><t>Data Point 2</t></si>
</sst>"""
      )

      // Sheet1: With conditional formatting, rich row attributes, SST refs
      writeEntry(
        out,
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <cols>
    <col min="1" max="1" width="20" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" spans="1:3" s="1" customFormat="1" ht="18" customHeight="1" x14ac:dyDescent="0.25">
      <c r="A1" s="1" t="s"><v>0</v></c>
      <c r="B1" s="1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>3</v></c>
      <c r="B2"><v>100</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>4</v></c>
      <c r="B3"><v>75</v></c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="B2:B10">
    <cfRule type="dataBar" priority="1">
      <dataBar>
        <cfvo type="min"/><cfvo type="max"/>
        <color rgb="FF638EC6"/>
      </dataBar>
    </cfRule>
  </conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>"""
      )

      // Sheet1 rels (for drawing/chart)
      writeEntry(
        out,
        "xl/worksheets/_rels/sheet1.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>"""
      )

      // Sheet2: Clean sheets (will be byte-identical after modification)
      writeEntry(
        out,
        "xl/worksheets/sheet2.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>1</v></c>
    </row>
    <row r="5">
      <c r="B5"><v>12345</v></c>
    </row>
  </sheetData>
</worksheet>"""
      )

      // Sheet3: Hidden sheets
      writeEntry(
        out,
        "xl/worksheets/sheet3.xml",
        """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hidden Sheet Data</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
      )

      // Chart (unknown part)
      writeEntry(
        out,
        "xl/charts/chart1.xml",
        """<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:p><a:r><a:t>Sales Chart</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$3</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"""
      )

      // Drawing (unknown part)
      writeEntry(
        out,
        "xl/drawings/drawing1.xml",
        """<?xml version="1.0"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"""
      )

      // Drawing rels
      writeEntry(
        out,
        "xl/drawings/_rels/drawing1.xml.rels",
        """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>"""
      )

      // Comments (unknown part)
      writeEntry(
        out,
        "xl/comments1.xml",
        """<?xml version="1.0"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>Author Name</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0">
      <text>
        <t>This is a comment on A1</t>
      </text>
    </comment>
  </commentList>
</comments>"""
      )

    finally out.close()

    path

  private def writeEntry(out: ZipOutputStream, name: String, content: String): Unit =
    val entry = new ZipEntry(name)
    out.putNextEntry(entry)
    out.write(content.getBytes("UTF-8"))
    out.closeEntry()
