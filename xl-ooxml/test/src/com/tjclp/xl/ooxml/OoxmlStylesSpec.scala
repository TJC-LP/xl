package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import munit.FunSuite
import com.tjclp.xl.api.{Sheet, Workbook}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.{Fill, Color, PatternType, CellStyle, NumFmt}
import com.tjclp.xl.styles.font.{Font, Underline}
import scala.xml.Elem

/** Tests for OOXML Styles serialization (xl/styles.xml) */
class OoxmlStylesSpec extends FunSuite:

  test("defaultFills has None at index 0") {
    // Use empty StyleIndex to test defaultFills behavior
    val styles = OoxmlStyles(StyleIndex.empty)
    val xml = styles.toXml

    val fills = xml \ "fills" \ "fill"
    assert(fills.nonEmpty, "Should have fills in styles.xml")

    // First fill should be patternFill with patternType="none"
    val firstFill = fills(0)
    val patternFill = firstFill \ "patternFill"
    assert(patternFill.nonEmpty, "First fill should have patternFill element")

    val patternType = (patternFill \ "@patternType").text
    assertEquals(patternType, "none", "First default fill should have patternType='none'")
  }

  test("defaultFills has Gray125 at index 1") {
    val styles = OoxmlStyles(StyleIndex.empty)
    val xml = styles.toXml

    val fills = xml \ "fills" \ "fill"
    assert(fills.size >= 2, "Should have at least 2 default fills")

    // Second fill should be patternFill with patternType="gray125"
    val secondFill = fills(1)
    val patternFill = secondFill \ "patternFill"
    assert(patternFill.nonEmpty, "Second fill should have patternFill element")

    val patternType = (patternFill \ "@patternType").text
    assertEquals(patternType, "gray125", "Second default fill should have patternType='gray125'")
  }

  test("toXml emits gray125 pattern in fills section") {
    val styles = OoxmlStyles(StyleIndex.empty)
    val xml = styles.toXml

    // Verify fills section exists with correct count
    val fillsElem = xml \ "fills"
    assert(fillsElem.nonEmpty, "Should have fills element")

    val count = (fillsElem \ "@count").text.toInt
    assert(count >= 2, "Should have at least 2 fills (default fills)")

    // GH-448: the mandatory gray125 placeholder is written BARE, exactly as Excel writes it
    val fills = xml \ "fills" \ "fill"
    val gray125Fill = fills(1)
    val patternFill = gray125Fill \ "patternFill"
    assertEquals((patternFill \ "@patternType").text, "gray125")
    assert((patternFill \ "fgColor").isEmpty, "bare gray125 must not carry fgColor (GH-448)")
    assert((patternFill \ "bgColor").isEmpty, "bare gray125 must not carry bgColor (GH-448)")

    // Verify pattern type is gray125
    val patternType = (patternFill \ "@patternType").text
    assertEquals(patternType, "gray125", "Pattern type should be gray125")
  }

  // ===== GH-404: custom numFmt declarations with built-in-equal codes =====

  test("reader: <numFmts> declarations parse to Custom(code) verbatim, built-in-equal or not") {
    // A <numFmt> table entry is by definition a custom declaration — the reader must not
    // collapse its code to a built-in enum case (that discards the id->code binding and the
    // writer then has nothing to emit for the custom id: GH-404's silent "General" corruption).
    val stylesXml =
      """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <numFmts count="3">
        |    <numFmt numFmtId="164" formatCode="0.00%"/>
        |    <numFmt numFmtId="165" formatCode="General"/>
        |    <numFmt numFmtId="166" formatCode="yyyy-mm-dd"/>
        |  </numFmts>
        |  <fonts count="1"><font><name val="Calibri"/><sz val="11"/></font></fonts>
        |  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
        |  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
        |  <cellXfs count="2">
        |    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        |    <xf numFmtId="164" fontId="0" fillId="0" borderId="0"/>
        |  </cellXfs>
        |</styleSheet>""".stripMargin
    val parsed = XmlSecurity
      .parseSafe(stylesXml, "styles.xml")
      .fold(e => fail(s"xml parse failed: ${e.message}"), identity)
    val wbStyles = WorkbookStyles
      .fromXml(parsed)
      .fold(e => fail(s"styles parse failed: $e"), identity)

    assertEquals(
      wbStyles.customNumFmts,
      Vector(
        164 -> NumFmt.Custom("0.00%"),
        165 -> NumFmt.Custom("General"),
        166 -> NumFmt.Custom("yyyy-mm-dd")
      )
    )
    // The referencing cellXf resolves the declared code, not the built-in twin
    assertEquals(wbStyles.cellStyles(1).numFmt, NumFmt.Custom("0.00%"))
    assertEquals(wbStyles.cellStyles(1).numFmtId, Some(164))
  }

  test("writer (DOM): numFmts table entries emit their format code, never General") {
    val index = StyleIndex.empty.copy(
      numFmts = Vector(
        164 -> NumFmt.PercentDecimal, // built-in-equal binding (what the reader used to produce)
        165 -> NumFmt.Custom("0.00%_)"),
        166 -> NumFmt.ThousandsDecimal
      )
    )
    val xml = OoxmlStyles(index).toXml
    val emitted = (xml \ "numFmts" \ "numFmt").map { n =>
      (n \ "@numFmtId").text.toInt -> (n \ "@formatCode").text
    }.toMap
    assertEquals(emitted, Map(164 -> "0.00%", 165 -> "0.00%_)", 166 -> "#,##0.00"))
  }

  test("writer (SAX): numFmts table entries emit their format code, never General") {
    val index = StyleIndex.empty.copy(
      numFmts = Vector(164 -> NumFmt.PercentDecimal, 165 -> NumFmt.Custom("0.0"))
    )
    val output = new ByteArrayOutputStream()
    OoxmlStyles(index).writeSax(StaxSaxWriter.create(output))
    val xml = new String(output.toByteArray, StandardCharsets.UTF_8)
    assert(xml.contains("formatCode=\"0.00%\""), s"SAX output lost the 0.00% code: $xml")
    assert(xml.contains("formatCode=\"0.0\""), s"SAX output lost the 0.0 code: $xml")
    assert(!xml.contains("formatCode=\"General\""), s"SAX output degraded to General: $xml")
  }

  test("GH-404 field repro: 0.00% survives write -> read -> modify -> write; renders 7.30%") {
    val tempDir = Files.createTempDirectory("xl-gh404-")
    def entryText(path: Path, name: String): String =
      val zip = new ZipFile(path.toFile)
      try
        Option(zip.getEntry(name)) match
          case Some(e) =>
            new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
          case None => fail(s"zip entry $name not found in $path")
      finally zip.close()

    // Step 1: fresh write with a custom format whose code equals built-in id 10 (0.00%)
    val initial = Workbook("S")
    val styled = initial
      .sheets(0)
      .put(ref"A1", CellValue.Number(BigDecimal("0.073")))
      .withCellStyle(ref"A1", CellStyle.default.withNumFmt(NumFmt.Custom("0.00%")))
    val wb = initial
      .update(initial.sheets(0).name, _ => styled)
      .fold(e => fail(s"update failed: $e"), identity)
    val styledPath = tempDir.resolve("styled.xlsx")
    XlsxWriter.write(wb, styledPath).fold(e => fail(s"write failed: ${e.message}"), identity)
    assert(
      entryText(styledPath, "xl/styles.xml").contains("formatCode=\"0.00%\""),
      "fresh write must declare formatCode 0.00%"
    )

    // Step 2: read -> unrelated edit -> write (the field corruption: formatCode became General)
    val read1 = XlsxReader.read(styledPath).fold(e => fail(s"read failed: ${e.message}"), identity)
    val modified = read1
      .update(read1.sheets(0).name, _.put(ref"B1", CellValue.Number(BigDecimal(1))))
      .fold(e => fail(s"update failed: $e"), identity)
    val rtPath = tempDir.resolve("rt.xlsx")
    XlsxWriter.write(modified, rtPath).fold(e => fail(s"write failed: ${e.message}"), identity)

    val rtStyles = entryText(rtPath, "xl/styles.xml")
    assert(
      rtStyles.contains("formatCode=\"0.00%\""),
      s"round-trip degraded the declared 0.00% code: $rtStyles"
    )
    assert(
      !rtStyles.contains("formatCode=\"General\""),
      s"round-trip emitted the GH-404 General corruption: $rtStyles"
    )

    // Step 3: the styled cell still carries the format and renders Excel-correct 7.30%
    val read2 = XlsxReader.read(rtPath).fold(e => fail(s"re-read failed: ${e.message}"), identity)
    val sheet2 = read2.sheets(0)
    val a1 = sheet2(ref"A1")
    val a1Style = a1.styleId
      .flatMap(sheet2.styleRegistry.get)
      .getOrElse(fail("A1 lost its style on round-trip"))
    assertEquals(a1Style.numFmt, NumFmt.Custom("0.00%"))
    assertEquals(NumFmtFormatter.formatValue(a1.value, a1Style.numFmt), "7.30%")
  }

  // ===== GH-423: underline variants (u@val, ST_UnderlineValues) =====

  private def parseStyles(stylesXml: String): WorkbookStyles =
    val parsed = XmlSecurity
      .parseSafe(stylesXml, "styles.xml")
      .fold(e => fail(s"xml parse failed: ${e.message}"), identity)
    WorkbookStyles.fromXml(parsed).fold(e => fail(s"styles parse failed: $e"), identity)

  private def stylesXmlWithFont(fontInner: String): String =
    s"""<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
       |  <fonts count="1"><font>$fontInner</font></fonts>
       |  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
       |  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
       |  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
       |</styleSheet>""".stripMargin

  test("reader: u@val parses typed — bare <u/> Single, double/accounting variants kept (GH-423)") {
    def underlineOf(fontInner: String): Underline =
      parseStyles(stylesXmlWithFont(fontInner)).fonts(0).underline
    assertEquals(underlineOf("""<name val="Calibri"/><sz val="11"/>"""), Underline.None)
    assertEquals(underlineOf("""<name val="Calibri"/><sz val="11"/><u/>"""), Underline.Single)
    assertEquals(
      underlineOf("""<name val="Calibri"/><sz val="11"/><u val="single"/>"""),
      Underline.Single
    )
    assertEquals(
      underlineOf("""<name val="Calibri"/><sz val="11"/><u val="double"/>"""),
      Underline.Double
    )
    assertEquals(
      underlineOf("""<name val="Calibri"/><sz val="11"/><u val="singleAccounting"/>"""),
      Underline.SingleAccounting
    )
    assertEquals(
      underlineOf("""<name val="Calibri"/><sz val="11"/><u val="doubleAccounting"/>"""),
      Underline.DoubleAccounting
    )
    // <u val="none"/> is explicit no-underline; unknown tokens stay lenient (read as Single,
    // the pre-GH-423 truthy behavior for any present <u>)
    assertEquals(
      underlineOf("""<name val="Calibri"/><sz val="11"/><u val="none"/>"""),
      Underline.None
    )
    assertEquals(
      underlineOf("""<name val="Calibri"/><sz val="11"/><u val="wavy"/>"""),
      Underline.Single
    )
  }

  test("writer (DOM): underline variants emit u@val; Single stays bare <u/> (GH-423)") {
    def fontXml(u: Underline): String =
      val index = StyleIndex.empty.copy(fonts = Vector(Font.default.withUnderline(u)))
      (OoxmlStyles(index).toXml \ "fonts" \ "font").toString
    assert(!fontXml(Underline.None).contains("<u"), fontXml(Underline.None))
    assert(fontXml(Underline.Single).contains("<u/>"), fontXml(Underline.Single))
    assert(
      fontXml(Underline.Double).contains("""<u val="double"/>"""),
      fontXml(Underline.Double)
    )
    assert(
      fontXml(Underline.SingleAccounting).contains("""<u val="singleAccounting"/>"""),
      fontXml(Underline.SingleAccounting)
    )
    assert(
      fontXml(Underline.DoubleAccounting).contains("""<u val="doubleAccounting"/>"""),
      fontXml(Underline.DoubleAccounting)
    )
  }

  test("writer (SAX): underline variants emit u@val; Single stays bare <u/> (GH-423)") {
    def fontXml(u: Underline): String =
      val index = StyleIndex.empty.copy(fonts = Vector(Font.default.withUnderline(u)))
      val output = new ByteArrayOutputStream()
      OoxmlStyles(index).writeSax(StaxSaxWriter.create(output))
      new String(output.toByteArray, StandardCharsets.UTF_8)
    // StAX renders the empty element as <u></u>; the point is: no val attribute for Single
    val single = fontXml(Underline.Single)
    assert(single.contains("<u>") || single.contains("<u/>"), single)
    assert(!single.contains("<u val="), single)
    assert(
      fontXml(Underline.SingleAccounting).contains("""<u val="singleAccounting""""),
      fontXml(Underline.SingleAccounting)
    )
    assert(
      fontXml(Underline.Double).contains("""<u val="double""""),
      fontXml(Underline.Double)
    )
  }

  test("GH-423 field repro: singleAccounting band style survives write -> read") {
    val tempDir = Files.createTempDirectory("xl-gh423-")
    val bandStyle = CellStyle.default.withFont(
      Font("Times New Roman", 10.0, underline = Underline.SingleAccounting)
    )
    val initial = Workbook("Model")
    val styled = initial
      .sheets(0)
      .put(ref"B2", CellValue.Text("Actual"))
      .withCellStyle(ref"B2", bandStyle)
    val wb = initial
      .update(initial.sheets(0).name, _ => styled)
      .fold(e => fail(s"update failed: $e"), identity)
    val path = tempDir.resolve("band.xlsx")
    XlsxWriter.write(wb, path).fold(e => fail(s"write failed: ${e.message}"), identity)

    val zip = new ZipFile(path.toFile)
    val stylesEntry =
      try
        new String(
          zip.getInputStream(zip.getEntry("xl/styles.xml")).readAllBytes(),
          StandardCharsets.UTF_8
        )
      finally zip.close()
    assert(
      stylesEntry.contains("""<u val="singleAccounting"/>"""),
      s"styles.xml lost the accounting underline: $stylesEntry"
    )

    val read = XlsxReader.read(path).fold(e => fail(s"read failed: ${e.message}"), identity)
    val sheet = read.sheets(0)
    val b2Style = sheet(ref"B2").styleId
      .flatMap(sheet.styleRegistry.get)
      .getOrElse(fail("B2 lost its style"))
    assertEquals(b2Style.font.underline, Underline.SingleAccounting)
  }

  // ===== GH-425: workbook default (Normal) font =====

  private val houseFont = Font("Times New Roman", 10.0)

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  test("GH-425: withDefaultFont drives font 0 + Normal cellStyleXf; reader populates it back") {
    val tempDir = Files.createTempDirectory("xl-gh425-")
    val initial = Workbook("Model").withDefaultFont(houseFont)
    val wb = initial
      .update(initial.sheets(0).name, _.put(ref"A1", CellValue.Text("typed into white space")))
      .fold(e => fail(s"update failed: $e"), identity)
    val path = tempDir.resolve("house.xlsx")
    XlsxWriter.write(wb, path).fold(e => fail(s"write failed: ${e.message}"), identity)

    val stylesXml = entryText(path, "xl/styles.xml")
    val parsed = XmlSecurity
      .parseSafe(stylesXml, "styles.xml")
      .fold(e => fail(s"xml parse failed: ${e.message}"), identity)

    // Font slot 0 IS the house font (the xfId-0 cellStyleXf hardcodes fontId=0)
    val font0 = (parsed \ "fonts" \ "font").headOption.getOrElse(fail(s"no fonts: $stylesXml"))
    assertEquals((font0 \ "name" \ "@val").text, "Times New Roman", stylesXml)
    assertEquals((font0 \ "sz" \ "@val").text, "10", stylesXml) // GH-448: integral sz
    // Normal master record references it
    val styleXf0 = (parsed \ "cellStyleXfs" \ "xf").headOption
      .getOrElse(fail(s"no cellStyleXfs: $stylesXml"))
    assertEquals((styleXf0 \ "@fontId").text, "0", stylesXml)
    // The unstyled cell's xf 0 resolves to the Normal font, not a stray Calibri entry
    val xf0 = (parsed \ "cellXfs" \ "xf").headOption.getOrElse(fail(s"no cellXfs: $stylesXml"))
    assertEquals((xf0 \ "@fontId").text, "0", stylesXml)
    assert(!stylesXml.contains("Calibri"), s"stock Calibri leaked into the font table: $stylesXml")

    // Reader populates the metadata back
    val read = XlsxReader.read(path).fold(e => fail(s"read failed: ${e.message}"), identity)
    assertEquals(read.metadata.defaultFont, Some(houseFont))
  }

  test("GH-425: both serializer backends emit the default font at slot 0 (DOM == SAX)") {
    val wb = Workbook("S").withDefaultFont(houseFont)
    val (index, _) = StyleIndex.fromWorkbook(wb)
    val domFont0 = (OoxmlStyles(index).toXml \ "fonts" \ "font").headOption
      .getOrElse(fail("DOM: no fonts"))
    assertEquals((domFont0 \ "name" \ "@val").text, "Times New Roman")

    val output = new ByteArrayOutputStream()
    OoxmlStyles(index).writeSax(StaxSaxWriter.create(output))
    val sax = new String(output.toByteArray, StandardCharsets.UTF_8)
    assert(
      sax.contains("""<name val="Times New Roman">""") ||
        sax.contains("""<name val="Times New Roman"/>"""),
      s"SAX backend lost the default font: $sax"
    )
    assert(!sax.contains("Calibri"), s"SAX backend leaked stock Calibri: $sax")
  }

  test("GH-425: a book whose Normal is already non-default keeps it (read -> scratch write)") {
    val tempDir = Files.createTempDirectory("xl-gh425-keep-")
    // Author a house book, read it back, DROP the source context (forces the from-scratch
    // serializer path), write again: Normal must still be the house font.
    val initial = Workbook("Model").withDefaultFont(houseFont)
    val wb = initial
      .update(initial.sheets(0).name, _.put(ref"A1", CellValue.Text("x")))
      .fold(e => fail(s"update failed: $e"), identity)
    val p1 = tempDir.resolve("v1.xlsx")
    XlsxWriter.write(wb, p1).fold(e => fail(s"write failed: ${e.message}"), identity)

    val read = XlsxReader.read(p1).fold(e => fail(s"read failed: ${e.message}"), identity)
    val scratch = read.copy(sourceContext = None)
    val p2 = tempDir.resolve("v2.xlsx")
    XlsxWriter.write(scratch, p2).fold(e => fail(s"write failed: ${e.message}"), identity)

    val parsed = XmlSecurity
      .parseSafe(entryText(p2, "xl/styles.xml"), "styles.xml")
      .fold(e => fail(s"xml parse failed: ${e.message}"), identity)
    val font0 = (parsed \ "fonts" \ "font").headOption.getOrElse(fail("no fonts"))
    assertEquals((font0 \ "name" \ "@val").text, "Times New Roman")

    // And the surgical path (source retained) keeps it too
    val edited = read
      .update(read.sheets(0).name, _.put(ref"B1", CellValue.Text("y")))
      .fold(e => fail(s"update failed: $e"), identity)
    val p3 = tempDir.resolve("v3.xlsx")
    XlsxWriter.write(edited, p3).fold(e => fail(s"write failed: ${e.message}"), identity)
    val parsed3 = XmlSecurity
      .parseSafe(entryText(p3, "xl/styles.xml"), "styles.xml")
      .fold(e => fail(s"xml parse failed: ${e.message}"), identity)
    val font0v3 = (parsed3 \ "fonts" \ "font").headOption.getOrElse(fail("no fonts"))
    assertEquals((font0v3 \ "name" \ "@val").text, "Times New Roman")
  }

  test("GH-425: stock-default books read back defaultFont=None and stay byte-stable") {
    val tempDir = Files.createTempDirectory("xl-gh425-stock-")
    val initial = Workbook("S")
    val wb = initial
      .update(initial.sheets(0).name, _.put(ref"A1", CellValue.Text("x")))
      .fold(e => fail(s"update failed: $e"), identity)
    val path = tempDir.resolve("stock.xlsx")
    XlsxWriter.write(wb, path).fold(e => fail(s"write failed: ${e.message}"), identity)
    val read = XlsxReader.read(path).fold(e => fail(s"read failed: ${e.message}"), identity)
    // Calibri-11 Normal is the modeled default — no phantom Some(Font.default)
    assertEquals(read.metadata.defaultFont, None)
  }

  test("GH-425: explicitly styled cells keep their own font; sentinel styles follow Normal") {
    val tempDir = Files.createTempDirectory("xl-gh425-styled-")
    val arial = Font("Arial", 9.0, bold = true)
    val initial = Workbook("Model").withDefaultFont(houseFont)
    val wb = initial
      .update(
        initial.sheets(0).name,
        s =>
          s.put(ref"A1", CellValue.Text("headline"))
            .withCellStyle(ref"A1", CellStyle.default.withFont(arial))
            .put(ref"B1", CellValue.Number(BigDecimal(1)))
            .withCellStyle(ref"B1", CellStyle.default.withNumFmt(NumFmt.Percent))
      )
      .fold(e => fail(s"update failed: $e"), identity)
    val path = tempDir.resolve("styled.xlsx")
    XlsxWriter.write(wb, path).fold(e => fail(s"write failed: ${e.message}"), identity)

    val read = XlsxReader.read(path).fold(e => fail(s"read failed: ${e.message}"), identity)
    val sheet = read.sheets(0)
    val a1Font = sheet(ref"A1").styleId
      .flatMap(sheet.styleRegistry.get)
      .map(_.font)
      .getOrElse(fail("A1 lost its style"))
    assertEquals(a1Font, arial)
    // B1's style only set a numFmt — its font was the Font.default sentinel, which resolves
    // to the workbook default on write
    val b1Font = sheet(ref"B1").styleId
      .flatMap(sheet.styleRegistry.get)
      .map(_.font)
      .getOrElse(fail("B1 lost its style"))
    assertEquals(b1Font, houseFont)
  }

  // ===== GH-566: non-solid pattern fills (ST_PatternType textures) survive read and write =====

  /**
   * A styles part in the two texture dialects the reader used to drop: openpyxl 3.1.5's
   * `PatternFill(patternType="mediumGray", fgColor="FF808080")` writes NO bgColor (fill 2), and a
   * two-colour `lightUp` hatch (fill 3). xf 1 / xf 2 point at them.
   */
  private val textureStylesXml: String =
    """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      |  <fonts count="1"><font><name val="Calibri"/><sz val="11"/></font></fonts>
      |  <fills count="4">
      |    <fill><patternFill patternType="none"/></fill>
      |    <fill><patternFill patternType="gray125"/></fill>
      |    <fill><patternFill patternType="mediumGray"><fgColor rgb="FF808080"/></patternFill></fill>
      |    <fill><patternFill patternType="lightUp"><fgColor rgb="FF0000FF"/><bgColor rgb="FFFFFF00"/></patternFill></fill>
      |  </fills>
      |  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      |  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      |  <cellXfs count="3">
      |    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
      |    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
      |    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
      |  </cellXfs>
      |</styleSheet>""".stripMargin

  /** Minimal package around [[textureStylesXml]]: A1 hatched mediumGray, A2 lightUp, B1 = 1. */
  private def textureFixture(): Path =
    val path = Files.createTempFile("xl-gh566-", ".xlsx")
    path.toFile.deleteOnExit()
    val zip = new java.util.zip.ZipOutputStream(Files.newOutputStream(path))
    def entry(name: String, content: String): Unit =
      zip.putNextEntry(new java.util.zip.ZipEntry(name))
      zip.write(content.getBytes(StandardCharsets.UTF_8))
      zip.closeEntry()
    try
      entry(
        "[Content_Types].xml",
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          |<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          |<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          |<Default Extension="xml" ContentType="application/xml"/>
          |<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          |<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          |<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
          |</Types>""".stripMargin
      )
      entry(
        "_rels/.rels",
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          |<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
          |</Relationships>""".stripMargin
      )
      entry(
        "xl/workbook.xml",
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          |<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          |<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>
          |</workbook>""".stripMargin
      )
      entry(
        "xl/_rels/workbook.xml.rels",
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          |<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          |<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
          |</Relationships>""".stripMargin
      )
      entry("xl/styles.xml", textureStylesXml)
      entry(
        "xl/worksheets/sheet1.xml",
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
          |<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>hatched</t></is></c><c r="B1"><v>1</v></c></row>
          |<row r="2"><c r="A2" s="2" t="inlineStr"><is><t>hatched too</t></is></c></row>
          |</sheetData></worksheet>""".stripMargin
      )
    finally zip.close()
    path

  private def zipEntryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  private def fillOf(sheet: Sheet, at: com.tjclp.xl.addressing.ARef): Fill =
    sheet(at).styleId
      .flatMap(sheet.styleRegistry.get)
      .map(_.fill)
      .getOrElse(fail(s"${at.toA1} lost its style"))

  test("GH-566: the reader keeps openpyxl's fgColor-only mediumGray texture as a pattern fill") {
    val parsed = parseStyles(textureStylesXml)
    assertEquals(
      parsed.fills(2),
      Fill.Pattern(Some(Color.Rgb(0xff808080)), None, PatternType.MediumGray),
      "fgColor-only texture: foreground kept, background automatic"
    )
    assertEquals(
      parsed.fills(3),
      Fill.pattern(Color.Rgb(0xff0000ff), Color.Rgb(0xffffff00), PatternType.LightUp)
    )
    // the source's bare gray125 IS the writer's mandatory placeholder
    assertEquals(parsed.fills(1), OoxmlStyles.defaultGray125)
    assertEquals(parsed.fills(1), Fill.Pattern(None, None, PatternType.Gray125))
  }

  test("GH-566: Excel's automatic pattern colours (indexed 64/65, auto) keep the texture") {
    val excelDialect = textureStylesXml.replace(
      """<fill><patternFill patternType="mediumGray"><fgColor rgb="FF808080"/></patternFill></fill>""",
      """<fill><patternFill patternType="darkGray"><fgColor indexed="64"/><bgColor indexed="65"/></patternFill></fill>"""
    )
    assertEquals(
      parseStyles(excelDialect).fills(2),
      Fill.Pattern(None, None, PatternType.DarkGray),
      "system indices 64/65 are the automatic colours"
    )
    val autoDialect = textureStylesXml.replace(
      """<fill><patternFill patternType="mediumGray"><fgColor rgb="FF808080"/></patternFill></fill>""",
      """<fill><patternFill patternType="darkTrellis"><fgColor auto="1"/></patternFill></fill>"""
    )
    assertEquals(
      parseStyles(autoDialect).fills(2),
      Fill.Pattern(None, None, PatternType.DarkTrellis),
      "auto=\"1\" is the automatic colour"
    )
  }

  List(
    "ScalaXml" -> com.tjclp.xl.ooxml.writer.WriterConfig.scalaXml,
    "SaxStax" -> com.tjclp.xl.ooxml.writer.WriterConfig.saxStax
  ).foreach { case (backend, config) =>
    test(
      s"GH-566 field repro ($backend): `put B1 2` keeps mediumGray/lightUp fills and their xfs"
    ) {
      val in = textureFixture()
      val wb = XlsxReader.read(in).fold(e => fail(s"read failed: ${e.message}"), identity)
      val edited = wb
        .update(wb.sheets(0).name, _.put(ref"B1", CellValue.Number(BigDecimal(2))))
        .fold(e => fail(s"update failed: $e"), identity)
      val out = Files.createTempFile(s"xl-gh566-out-$backend-", ".xlsx")
      out.toFile.deleteOnExit()
      XlsxWriter
        .writeWith(edited, out, config)
        .fold(e => fail(s"write failed: ${e.message}"), identity)

      // StAX renders a childless element as <e></e>; compare in the minimized form
      val styles = """<([A-Za-z][\w:.-]*)([^<>]*)></\1>""".r
        .replaceAllIn(zipEntryText(out, "xl/styles.xml"), m => s"<${m.group(1)}${m.group(2)}/>")
      assert(styles.contains("""patternType="mediumGray""""), s"mediumGray dropped:\n$styles")
      assert(styles.contains("""patternType="lightUp""""), s"lightUp dropped:\n$styles")
      assert(styles.contains("""<fgColor rgb="FF808080"/>"""), s"mediumGray fgColor lost:\n$styles")
      assert(!styles.contains("mediumgray") && !styles.contains("lightup"), s"lowercase:\n$styles")
      // the source's four fills (none, gray125, mediumGray, lightUp): the bare gray125 is the
      // writer's own placeholder and the textures are kept — nothing dropped, nothing duplicated
      assert(styles.contains("""<fills count="4">"""), s"fills count drifted:\n$styles")
      assertEquals("patternType=\"gray125\"".r.findAllIn(styles).size, 1, styles)

      val back = XlsxReader.read(out).fold(e => fail(s"re-read failed: ${e.message}"), identity)
      val sheet = back.sheets(0)
      assertEquals(
        fillOf(sheet, ref"A1"),
        Fill.Pattern(Some(Color.Rgb(0xff808080)), None, PatternType.MediumGray),
        "A1's xf must still point at the mediumGray fill"
      )
      assertEquals(
        fillOf(sheet, ref"A2"),
        Fill.pattern(Color.Rgb(0xff0000ff), Color.Rgb(0xffffff00), PatternType.LightUp),
        "A2's xf must still point at the lightUp fill"
      )
      assertEquals(sheet(ref"B1").value, CellValue.Number(BigDecimal(2)))
    }
  }
