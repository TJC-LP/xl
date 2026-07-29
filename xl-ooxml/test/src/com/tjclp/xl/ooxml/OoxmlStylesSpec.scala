package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import munit.FunSuite
import com.tjclp.xl.api.Workbook
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

    // Verify gray125 pattern has foreground and background colors
    val fills = xml \ "fills" \ "fill"
    val gray125Fill = fills(1)
    val patternFill = gray125Fill \ "patternFill"

    // Check foreground color
    val fgColor = patternFill \ "fgColor"
    assert(fgColor.nonEmpty, "gray125 pattern should have fgColor")
    val fgRgb = (fgColor \ "@rgb").text
    assert(fgRgb.nonEmpty, "fgColor should have rgb attribute")

    // Check background color
    val bgColor = patternFill \ "bgColor"
    assert(bgColor.nonEmpty, "gray125 pattern should have bgColor")
    val bgRgb = (bgColor \ "@rgb").text
    assert(bgRgb.nonEmpty, "bgColor should have rgb attribute")

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
    assertEquals((font0 \ "sz" \ "@val").text, "10.0", stylesXml)
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
