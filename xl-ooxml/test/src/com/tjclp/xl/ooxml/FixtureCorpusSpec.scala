package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Real-file corpus checks (GH-240): every reader test input here was produced by a FOREIGN writer
 * (openpyxl or LibreOffice), so these assertions cannot be self-fulfilling round-trips of XL's own
 * XML dialect.
 *
 * openpyxl emits inline strings (<is><t>); LibreOffice emits shared strings with
 * xml:space="preserve", explicit t="n"/s="N" attributes, cached formula values, and converts plain
 * booleans into cached TRUE()/FALSE() formulas. Both dialects must parse.
 */
class FixtureCorpusSpec extends FunSuite:

  private def load(name: String): Workbook =
    XlsxReader
      .read(TestFixtures.copyToTemp(name))
      .fold(err => fail(s"$name failed to read: ${err.message}"), identity)

  private def sheetOf(wb: Workbook, name: String, fixture: String): Sheet =
    wb(name).fold(err => fail(s"$fixture missing sheet $name: $err"), identity)

  TestFixtures.all.foreach { fixture =>
    test(s"corpus: $fixture parses via in-memory XlsxReader") {
      val wb = load(fixture)
      assert(wb.sheets.nonEmpty, s"$fixture produced an empty workbook")
      val nonEmptyCells = wb.sheets.map(_.cells.count(_._2.value != CellValue.Empty)).sum
      assert(nonEmptyCells > 0, s"$fixture produced no non-empty cells")
    }
  }

  test("small-values.xlsx: unicode, whitespace, numbers, booleans, dates parse exactly") {
    val sheet = sheetOf(load("small-values.xlsx"), "Values", "small-values.xlsx")
    // Strings: unicode + whitespace-significant (openpyxl inline-string dialect)
    assertEquals(sheet(ref"A1").value, CellValue.Text("plain"))
    assertEquals(sheet(ref"A2").value, CellValue.Text("héllo wörld"))
    assertEquals(sheet(ref"A3").value, CellValue.Text("日本語テキスト"))
    assertEquals(sheet(ref"A4").value, CellValue.Text("emoji 🚀 rocket"))
    assertEquals(sheet(ref"A5").value, CellValue.Text("  leading and trailing  "))
    assertEquals(sheet(ref"A6").value, CellValue.Text("internal  double  spaces"))
    // Numbers incl. lowercase scientific notation (1.23e-10 in the XML)
    assertEquals(sheet(ref"B1").value, CellValue.Number(BigDecimal(42)))
    assertEquals(sheet(ref"B2").value, CellValue.Number(BigDecimal(-17)))
    assertEquals(sheet(ref"B3").value, CellValue.Number(BigDecimal("3.14159")))
    assertEquals(sheet(ref"B4").value, CellValue.Number(BigDecimal("1234567890123")))
    assertEquals(sheet(ref"B5").value, CellValue.Number(BigDecimal("1.23e-10")))
    assertEquals(sheet(ref"B6").value, CellValue.Number(BigDecimal("9990000000000000")))
    // Booleans
    assertEquals(sheet(ref"C1").value, CellValue.Bool(true))
    assertEquals(sheet(ref"C2").value, CellValue.Bool(false))
    // Dates surface as 1900-system serials (style carries the date numFmt)
    assertEquals(sheet(ref"D1").value, CellValue.Number(BigDecimal(45366)))
    assertEquals(sheet(ref"D2").value, CellValue.Number(BigDecimal("45366.6046875")))
    // Digit string must stay text
    assertEquals(sheet(ref"E1").value, CellValue.Text("0123"))
  }

  test("small-values-lo.xlsx: LibreOffice SST dialect parses; booleans become cached formulas") {
    val sheet = sheetOf(load("small-values-lo.xlsx"), "Values", "small-values-lo.xlsx")
    // Same strings, but via sharedStrings.xml with xml:space="preserve" on every entry
    assertEquals(sheet(ref"A2").value, CellValue.Text("héllo wörld"))
    assertEquals(sheet(ref"A3").value, CellValue.Text("日本語テキスト"))
    assertEquals(sheet(ref"A4").value, CellValue.Text("emoji 🚀 rocket"))
    assertEquals(sheet(ref"A5").value, CellValue.Text("  leading and trailing  "))
    assertEquals(sheet(ref"A6").value, CellValue.Text("internal  double  spaces"))
    assertEquals(sheet(ref"E1").value, CellValue.Text("0123"))
    // LibreOffice re-renders 1.23e-10 with a three-digit exponent (1.23E-010)
    assertEquals(sheet(ref"B5").value, CellValue.Number(BigDecimal("1.23E-10")))
    assertEquals(sheet(ref"B6").value, CellValue.Number(BigDecimal("9990000000000000")))
    // LibreOffice converts plain booleans to cached TRUE()/FALSE() formulas (t="b" + <f>)
    assertEquals(sheet(ref"C1").value, CellValue.Formula("TRUE()", Some(CellValue.Bool(true))))
    assertEquals(sheet(ref"C2").value, CellValue.Formula("FALSE()", Some(CellValue.Bool(false))))
    // Dates: identical serials regardless of producer
    assertEquals(sheet(ref"D1").value, CellValue.Number(BigDecimal(45366)))
    assertEquals(sheet(ref"D2").value, CellValue.Number(BigDecimal("45366.6046875")))
  }

  test("formulas.xlsx: cross-sheet refs, leading plus, absolute anchors (no cached values)") {
    val wb = load("formulas.xlsx")
    val data = sheetOf(wb, "Data", "formulas.xlsx")
    val calc = sheetOf(wb, "Calc", "formulas.xlsx")
    assertEquals(data(ref"A5").value, CellValue.Number(BigDecimal(50)))
    // openpyxl emits <f>expr</f><v /> - empty cached value must read as None
    assertEquals(calc(ref"A1").value, CellValue.Formula("SUM(Data!A1:A5)", None))
    assertEquals(calc(ref"A2").value, CellValue.Formula("Data!A1*2", None))
    assertEquals(calc(ref"A3").value, CellValue.Formula("+Data!A2", None))
    assertEquals(calc(ref"A4").value, CellValue.Formula("$A$1+A2", None))
    assertEquals(calc(ref"A5").value, CellValue.Formula("SUM($A$1:A4)", None))
    assertEquals(calc(ref"A6").value, CellValue.Formula("AVERAGE(Data!A1:A5)+Data!B1", None))
  }

  test("formulas-lo.xlsx: LibreOffice computes and caches formula results") {
    val calc = sheetOf(load("formulas-lo.xlsx"), "Calc", "formulas-lo.xlsx")
    assertEquals(
      calc(ref"A1").value,
      CellValue.Formula("SUM(Data!A1:A5)", Some(CellValue.Number(BigDecimal(150))))
    )
    assertEquals(
      calc(ref"A3").value,
      CellValue.Formula("+Data!A2", Some(CellValue.Number(BigDecimal(20))))
    )
    assertEquals(
      calc(ref"A4").value,
      CellValue.Formula("$A$1+A2", Some(CellValue.Number(BigDecimal(170))))
    )
  }

  test("styled.xlsx: fonts, fills, numFmts, alignment+indent, merges survive reading") {
    val sheet = sheetOf(load("styled.xlsx"), "Styled", "styled.xlsx")
    val a1Style = sheet.getCellStyle(ref"A1").getOrElse(fail("A1 style missing"))
    assertEquals(a1Style.font.bold, true)
    assertEquals(a1Style.font.name, "Arial")
    assertEquals(a1Style.font.sizePt, 14.0)
    val b1Style = sheet.getCellStyle(ref"B1").getOrElse(fail("B1 style missing"))
    b1Style.numFmt match
      case NumFmt.Custom(code) =>
        assert(code.contains("#,##0.00"), s"B1 numFmt should be currency-like, was: $code")
      case other => fail(s"expected custom currency numFmt on B1, got: $other")
    val c1Style = sheet.getCellStyle(ref"C1").getOrElse(fail("C1 style missing"))
    assertEquals(c1Style.align.wrapText, true)
    assertEquals(c1Style.align.indent, 2)
    assert(
      sheet.mergedRanges.exists(_.toA1 == "A4:C4"),
      s"expected merge A4:C4, got: ${sheet.mergedRanges.map(_.toA1)}"
    )
    assertEquals(sheet(ref"A4").value, CellValue.Text("Merged Header"))
  }

  test("comments-hyperlinks.xlsx: hyperlink targets survive reading") {
    val sheet = sheetOf(load("comments-hyperlinks.xlsx"), "Notes", "comments-hyperlinks.xlsx")
    assertEquals(
      sheet(ref"B1").hyperlink,
      Some("https://example.com/xl-fixtures"),
      "external hyperlink target"
    )
    assertEquals(sheet(ref"B1").value, CellValue.Text("example link"))
  }

  test("comments-hyperlinks.xlsx: KNOWN GAP - openpyxl comment dialect is invisible to the domain") {
    // openpyxl stores the comment part at xl/comments/comment1.xml (subdirectory,
    // singular) and points to it from the sheet rels with an absolute target
    // (/xl/comments/comment1.xml). XlsxReader only treats xl/comments\d+.xml as a
    // known comment part, so this comment is NOT parsed into Sheet.comments - it
    // rides along as an opaque preserved part instead. Excel itself reads openpyxl
    // comments fine. This test pins the current (lossy) behavior; if it starts
    // failing because the comment appears, the reader learned rels-based comment
    // resolution and this test should be flipped to assert the comment text.
    val sheet = sheetOf(load("comments-hyperlinks.xlsx"), "Notes", "comments-hyperlinks.xlsx")
    assertEquals(
      sheet.comments.get(ref"A1"),
      None,
      "reader unexpectedly resolved the openpyxl-dialect comment - flip this test to assert it"
    )
  }

  test("autofilter.xlsx: data table reads; autoFilter rides through as preserved XML") {
    val sheet = sheetOf(load("autofilter.xlsx"), "Filtered", "autofilter.xlsx")
    assertEquals(sheet(ref"A1").value, CellValue.Text("Region"))
    assertEquals(sheet(ref"C6").value, CellValue.Number(BigDecimal(19)))
    // Round-trip: the autoFilter element must survive an in-memory write
    val out = XlsxWriter.writeToBytes(load("autofilter.xlsx")).fold(
      err => fail(s"write failed: ${err.message}"),
      identity
    )
    val sheetXml = zipEntry(out, "xl/worksheets/sheet1.xml")
    assert(
      sheetXml.contains("autoFilter") && sheetXml.contains("A1:C6"),
      "autoFilter ref=A1:C6 should be preserved through read->write"
    )
  }

  test("chart-bar.xlsx and image.xlsx: workbooks with drawings parse without error") {
    val chartWb = load("chart-bar.xlsx")
    val chartSheet = sheetOf(chartWb, "ChartData", "chart-bar.xlsx")
    assertEquals(chartSheet(ref"B5").value, CellValue.Number(BigDecimal(23)))
    val imageWb = load("image.xlsx")
    val imageSheet = sheetOf(imageWb, "HasImage", "image.xlsx")
    assertEquals(imageSheet(ref"A1").value, CellValue.Text("tiny synthetic png anchored at B2"))
  }

  test("corpus: read -> write -> read is value-stable for every fixture") {
    TestFixtures.all.foreach { fixture =>
      val original = load(fixture)
      val bytes = XlsxWriter
        .writeToBytes(original)
        .fold(err => fail(s"$fixture write failed: ${err.message}"), identity)
      val reread = XlsxReader
        .readFromBytes(bytes)
        .fold(err => fail(s"$fixture re-read failed: ${err.message}"), identity)
      original.sheets.zip(reread.sheets).foreach { (before, after) =>
        assertEquals(after.name, before.name, s"$fixture sheet name drifted")
        val beforeValues = before.cells.collect {
          case (ref, cell) if cell.value != CellValue.Empty => ref -> cell.value
        }
        val afterValues = after.cells.collect {
          case (ref, cell) if cell.value != CellValue.Empty => ref -> cell.value
        }
        assertEquals(
          afterValues,
          beforeValues,
          s"$fixture!${before.name.value} cell values drifted across write->read"
        )
      }
    }
  }

  private def zipEntry(bytes: Array[Byte], name: String): String =
    val zis = new java.util.zip.ZipInputStream(new java.io.ByteArrayInputStream(bytes))
    try
      Iterator
        .continually(Option(zis.getNextEntry))
        .takeWhile(_.isDefined)
        .flatten
        .collectFirst {
          case e if e.getName == name =>
            new String(zis.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8)
        }
        .getOrElse(fail(s"zip entry $name not found"))
    finally zis.close()
