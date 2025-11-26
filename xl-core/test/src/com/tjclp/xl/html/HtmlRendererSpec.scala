package com.tjclp.xl.html

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.{Column, Row}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.dsl.{*, given}
import com.tjclp.xl.richtext.RichText.{*, given}
import com.tjclp.xl.macros.{ref, col}
import com.tjclp.xl.patch.Patch
// Removed: BatchPutMacro is dead code (shadowed by Sheet.put member)  // For batch put extension
import com.tjclp.xl.codec.syntax.*
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.border.{Border, BorderStyle}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.unsafe.*

/** Tests for HTML export functionality */
class HtmlRendererSpec extends FunSuite:

  // ========== Basic HTML Export ==========

  test("toHtml: empty ref") {
    val sheet = Sheet("Test")
    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("<table"), "Should contain table tag")
    assert(html.contains("border-collapse"), "Should have border-collapse styling")
    assert(html.contains("background-color: #FFFFFF"), "Empty cell should have white background")
  }

  test("toHtml: single cell with text") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Hello"))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("<table"), "Should contain table tag")
    assert(html.contains("Hello"), "Should contain cell text")
  }

  test("toHtml: 2x2 grid") {
    val sheet = Sheet("Test")
      .put(
        ref"A1" -> "A1",
        ref"B1" -> "B1",
        ref"A2" -> "A2",
        ref"B2" -> "B2"
      )
      .unsafe

    val html = sheet.toHtml(ref"A1:B2")
    assert(html.contains("<table"))
    assert(html.contains("A1"))
    assert(html.contains("B2"))
    // Should have 2 data rows (tr elements with height style)
    val rowCount = "<tr style=\"height:".r.findAllIn(html).length
    assertEquals(rowCount, 2, "Should have 2 rows")
  }

  // ========== Rich Text Export ==========

  test("toHtml: rich text with bold") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.RichText("Bold".bold + " normal"))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("<b>Bold</b>"), "Should render bold tag")
    assert(html.contains(" normal"), "Should include normal text")
  }

  test("toHtml: rich text with italic") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.RichText("Italic".italic))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("<i>Italic</i>"), "Should render italic tag")
  }

  test("toHtml: rich text with underline") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.RichText("Underline".underline))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("<u>Underline</u>"), "Should render underline tag")
  }

  test("toHtml: rich text with colors") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.RichText("Red".red + " and " + "Green".green))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("color:"), "Should include color style")
    assert(html.contains("<span"), "Should use span for color")
  }

  test("toHtml: rich text with mixed formatting") {
    val text = "Error: ".red.bold + "File not found".underline
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.RichText(text))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("<b>"), "Should have bold")
    assert(html.contains("<u>"), "Should have underline")
    assert(html.contains("color:"), "Should have color")
  }

  // ========== Cell Style Export ==========

  test("toHtml: cell style as inline CSS") {
    val headerStyle = CellStyle.default
      .withFont(Font("Arial", 14.0, bold = true))
      .withFill(Fill.Solid(Color.fromRgb(200, 200, 200)))
      .withAlign(Align(HAlign.Center, VAlign.Middle))

    val sheet = Sheet("Test")
      .put(ref"A1" -> "Header")
      .unsafe
      .withCellStyle(ref"A1", headerStyle)

    val html = sheet.toHtml(ref"A1:A1", includeStyles = true)
    assert(html.contains("font-weight: bold"), "Should include bold CSS")
    assert(html.contains("background-color:"), "Should include background color")
    assert(html.contains("text-align: center"), "Should include alignment")
  }

  test("toHtml: includeStyles=false omits CSS") {
    val boldStyle = CellStyle.default.withFont(Font.default.withBold(true))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Bold")
      .unsafe
      .withCellStyle(ref"A1", boldStyle)

    val html = sheet.toHtml(ref"A1:A1", includeStyles = false)
    // Table still has wrapper style, but cells should not have style attributes
    assert(!html.contains("<td style="), "Cells should not include inline styles")
    assert(!html.contains("font-weight"), "Should not include font styling")
  }

  // ========== HTML Escaping ==========

  test("toHtml: escapes HTML special characters") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("<script>alert('xss')</script>"))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("&lt;script&gt;"), "Should escape < and >")
    assert(!html.contains("<script>"), "Should not contain unescaped script tag")
  }

  test("toHtml: escapes ampersands") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("A & B"))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("A &amp; B"), "Should escape ampersand")
  }

  test("toHtml: escapes quotes") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("""Say "Hello""""))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("&quot;"), "Should escape quotes")
  }

  test("toHtml: escapes comment tooltip content (HTML + quotes)") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "value").unsafe
      .comment(ref"A1", com.tjclp.xl.cells.Comment.plainText("""<b onload="x" data='y'>&test</b>""", Some("""Auth"or'&""")))

    val html = sheet.toHtml(ref"A1:A1", includeComments = true)
    assert(html.contains("""title=""""), s"Tooltip should be present, got: $html")
    assert(html.contains("""Auth&quot;or&#39;&amp;:"""), "Author should be escaped")
    assert(html.contains("""&lt;b onload=&quot;x&quot; data=&#39;y&#39;&gt;"""), "Comment body should escape tags/quotes")
    assert(html.contains("""&amp;test"""), "Ampersand in body should be escaped")
  }

  // ========== Different Cell Types ==========

  test("toHtml: number cells") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> BigDecimal("123.45"))
      .unsafe

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("123.45"))
  }

  test("toHtml: boolean cells") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> true)
      .unsafe

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("TRUE"), "Excel uses uppercase TRUE/FALSE")
  }

  test("toHtml: date cells") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> java.time.LocalDate.of(2025, 11, 10))
      .unsafe

    val html = sheet.toHtml(ref"A1:A1")
    // Date is formatted with NumFmt.Date style (M/d/yy format)
    assert(html.contains("11/10/25"), s"Expected Excel date format, got: $html")
  }

  test("toHtml: formula cells") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Formula("SUM(B1:B10)"))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("=SUM"), "Should include formula")
  }

  // ========== Integration with putMixed ==========

  test("toHtml: financial model example") {
    val sheet = Sheet("Q1 Report")
      .put(
        ref"A1" -> "Revenue",
        ref"B1" -> BigDecimal("1000000"),
        ref"A2" -> "Expenses",
        ref"B2" -> BigDecimal("750000")
      )
      .unsafe

    val html = sheet.toHtml(ref"A1:B2")
    assert(html.contains("Revenue"))
    assert(html.contains("1000000"))
    assert(html.contains("Expenses"))
    assert(html.contains("750000"))
  }

  test("toHtml: mixed rich text and plain cells") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.RichText("Bold".bold + " text"))
      .put(ref"A2", CellValue.Text("Plain text"))

    val html = sheet.toHtml(ref"A1:A2")
    assert(html.contains("<b>Bold</b>"), "Rich text should be formatted")
    assert(html.contains("Plain text"), "Plain text should be included")
  }

  // ========== CSS Escaping (Regression tests for PR #44 feedback) ==========

  test("toHtml: CSS escapes newlines in font names") {
    // Malicious font name with newline - without escaping, the newline would break out of the string
    val maliciousFont = Font("Arial\nEvil", 11.0)
    val style = CellStyle.default.withFont(maliciousFont)
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Test")
      .unsafe
      .withCellStyle(ref"A1", style)

    val html = sheet.toHtml(ref"A1:A1")
    // Newline should be escaped as CSS escape sequence \A (not literal newline)
    assert(!html.contains("Arial\nEvil"), "Literal newline should not appear in CSS")
    assert(html.contains("\\A "), "Newline should be escaped as \\A")
    assert(html.contains("font-family:"), "Should contain font-family property")
  }

  test("toHtml: CSS escapes carriage returns in font names") {
    val fontWithCR = Font("Arial\rEvil", 11.0)
    val style = CellStyle.default.withFont(fontWithCR)
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Test")
      .unsafe
      .withCellStyle(ref"A1", style)

    val html = sheet.toHtml(ref"A1:A1")
    assert(!html.contains("Arial\rEvil"), "Literal CR should not appear in CSS")
    assert(html.contains("\\D "), "Carriage return should be escaped as \\D")
  }

  test("toHtml: CSS removes null bytes from font names") {
    val fontWithNull = Font("Ari\u0000al", 11.0)
    val style = CellStyle.default.withFont(fontWithNull)
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Test")
      .unsafe
      .withCellStyle(ref"A1", style)

    val html = sheet.toHtml(ref"A1:A1")
    assert(!html.contains("\u0000"), "Null bytes should be removed")
    assert(html.contains("Arial"), "Font name should be preserved minus null")
  }

  test("toHtml: CSS escapes quotes in font names") {
    // Single quote in font name should be escaped to prevent breaking out of quoted string
    val fontWithQuotes = Font("Arial's Best", 11.0)
    val style = CellStyle.default.withFont(fontWithQuotes)
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Test")
      .unsafe
      .withCellStyle(ref"A1", style)

    val html = sheet.toHtml(ref"A1:A1")
    // The quote should be escaped, so we shouldn't see an unescaped single quote
    // that could terminate the CSS string early
    assert(html.contains("\\'"), "Single quote should be escaped")
    assert(html.contains("font-family:"), "Should contain font-family property")
  }

  //========== Real-World Use Case ==========

  test("toHtml: complete financial report with rich text and styles") {
    val headerStyle = CellStyle.default
      .withFont(Font("Arial", 14.0, bold = true))
      .withFill(Fill.Solid(Color.fromRgb(220, 220, 220)))
      .withAlign(Align(HAlign.Center, VAlign.Middle))

    val positive = "+12.5%".green.bold
    val negative = "-5.2%".red.bold

    val sheet = Sheet("Report")
      .put(
        ref"A1" -> "Metric",
        ref"B1" -> "Change"
      )
      .unsafe
      .withRangeStyle(ref"A1:B1", headerStyle)
      .put(ref"A2", CellValue.Text("Revenue"))
      .put(ref"B2", CellValue.RichText(positive))
      .put(ref"A3", CellValue.Text("Costs"))
      .put(ref"B3", CellValue.RichText(negative))

    val html = sheet.toHtml(ref"A1:B3")

    // Verify structure
    assert(html.contains("<table"))
    assert(html.contains("Metric"))
    assert(html.contains("Change"))

    // Verify header styles
    assert(html.contains("background-color:"), "Headers should have background")
    assert(html.contains("text-align: center"), "Headers should be centered")

    // Verify rich text
    assert(html.contains("+12.5%"), "Should include percentage")
    assert(html.contains("-5.2%"))
  }

  // ========== Row/Column Sizing Tests ==========

  test("toHtml: includes colgroup with column widths") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")

    val html = sheet.toHtml(ref"A1:A1")

    // Should have colgroup element with col elements
    assert(html.contains("<colgroup>"), "Should have colgroup element")
    assert(html.contains("<col style=\"width:"), "Should have col elements with width")
  }

  test("toHtml: table-layout fixed is used") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")

    val html = sheet.toHtml(ref"A1:A1")

    assert(html.contains("table-layout: fixed"), "Should use table-layout: fixed")
  }

  test("toHtml: explicit column width is used") {
    // Column width 20 chars → (20 * 7 + 5) = 145 pixels
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(20.0)))

    val html = sheet.toHtml(ref"A1:A1")

    assert(html.contains("width: 145px"), s"Column width should be 145px (20 chars), got: $html")
  }

  test("toHtml: explicit row height is used") {
    // Row height 30 points → (30 * 4/3) = 40 pixels
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setRowProperties(Row.from0(0), RowProperties(height = Some(30.0)))

    val html = sheet.toHtml(ref"A1:A1")

    // tr should have height style
    assert(html.contains("height: 40px"), s"Row height should be 40px (30pt), got: $html")
  }

  test("toHtml: hidden column renders as 0px") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setColumnProperties(Column.from0(0), ColumnProperties(hidden = true))

    val html = sheet.toHtml(ref"A1:A1")

    // Hidden column should have 0 width in colgroup
    assert(html.contains("width: 0px"), s"Hidden column should be 0px, got: $html")
  }

  test("toHtml: hidden row renders as 0px") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setRowProperties(Row.from0(0), RowProperties(hidden = true))

    val html = sheet.toHtml(ref"A1:A1")

    // Hidden row should have 0 height
    assert(html.contains("height: 0px"), s"Hidden row should be 0px, got: $html")
  }

  test("toHtml: overflow hidden is applied to cells") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Data")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(10.0)))

    val html = sheet.toHtml(ref"A1:A1", includeStyles = true)

    // Cells should have overflow: hidden to clip content
    assert(html.contains("overflow: hidden"), s"Cells should have overflow: hidden, got: $html")
  }

  test("toHtml: DSL-based sizing works") {
    // Use the row/column DSL
    val patch = col"A".width(15.0).toPatch ++ row(0).height(24.0).toPatch

    val html = Patch
      .applyPatch(Sheet("Test").put(ref"A1" -> "Data"), patch)
      .toHtml(ref"A1:A1")

    // Column A: 15 chars → 110px, Row 0: 24pt → 32px
    assert(html.contains("width: 110px"), s"DSL column width should apply, got: $html")
    assert(html.contains("height: 32px"), s"DSL row height should apply, got: $html")
  }

  test("toHtml: multiple columns with different widths") {
    val sheet = Sheet("Test")
      .put(ref"A1" -> "A", ref"B1" -> "B", ref"C1" -> "C")
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(10.0))) // 75px
      .setColumnProperties(Column.from0(1), ColumnProperties(width = Some(20.0))) // 145px

    val html = sheet.toHtml(ref"A1:C1")

    // Check colgroup has different widths
    assert(html.contains("width: 75px"), s"Column A should be 75px, got: $html")
    assert(html.contains("width: 145px"), s"Column B should be 145px, got: $html")
    // Column C uses default
    assert(html.contains("width: 64px"), s"Column C should use default 64px, got: $html")
  }

  // ========== Indentation Tests ==========

  test("toHtml: cell with indent renders padding-left") {
    val indentStyle = CellStyle.default.withAlign(Align(indent = 2))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Indented")
      .unsafe
      .withCellStyle(ref"A1", indentStyle)

    val html = sheet.toHtml(ref"A1:A1")
    // 2 levels * 21px = 42px padding-left
    assert(html.contains("padding-left: 42px"), s"Should have padding-left for indent=2, got: $html")
  }

  test("toHtml: indent 0 does not add padding") {
    val noIndent = CellStyle.default.withAlign(Align(indent = 0))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Normal")
      .unsafe
      .withCellStyle(ref"A1", noIndent)

    val html = sheet.toHtml(ref"A1:A1")
    assert(!html.contains("padding-left:"), s"Should not have padding-left for indent=0, got: $html")
  }

  test("toHtml: large indent value works") {
    // Excel max indent is 15 levels
    val maxIndent = CellStyle.default.withAlign(Align(indent = 15))
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Max indent")
      .unsafe
      .withCellStyle(ref"A1", maxIndent)

    val html = sheet.toHtml(ref"A1:A1")
    // 15 levels * 21px = 315px
    assert(html.contains("padding-left: 315px"), s"Should have padding-left for indent=15, got: $html")
  }

  // ========== Merged Cells Tests ==========

  test("toHtml: merged cells render with colspan") {
    val range = CellRange.parse("A1:C1").toOption.get
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Merged")
      .merge(range)

    val html = sheet.toHtml(ref"A1:C1")
    assert(html.contains("""colspan="3""""), s"Should have colspan=3 for 3-column merge, got: $html")
    // Should only have 1 td element (interior cells skipped)
    val tdCount = "<td".r.findAllIn(html).length
    assertEquals(tdCount, 1, s"Merged row should have only 1 <td>, got: $html")
  }

  test("toHtml: merged cells render with rowspan") {
    val range = CellRange.parse("A1:A3").toOption.get
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Merged")
      .merge(range)

    val html = sheet.toHtml(ref"A1:A3")
    assert(html.contains("""rowspan="3""""), s"Should have rowspan=3 for 3-row merge, got: $html")
  }

  test("toHtml: merged cells with both colspan and rowspan") {
    val range = CellRange.parse("A1:B2").toOption.get
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Big Merge")
      .merge(range)

    val html = sheet.toHtml(ref"A1:B2")
    assert(html.contains("""colspan="2""""), s"Should have colspan=2, got: $html")
    assert(html.contains("""rowspan="2""""), s"Should have rowspan=2, got: $html")
  }

  test("toHtml: non-anchor merged cells are skipped") {
    val range = CellRange.parse("A1:B1").toOption.get
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Anchor")
      .put(ref"B1" -> "Should be hidden")
      .merge(range)

    val html = sheet.toHtml(ref"A1:B1")
    assert(html.contains("Anchor"), "Anchor cell content should appear")
    assert(!html.contains("Should be hidden"), s"Interior cell content should not appear, got: $html")
  }

  test("toHtml: unmerged cells not affected by merge logic") {
    val range = CellRange.parse("A1:A1").toOption.get
    val sheet = Sheet("Test")
      .put(ref"A1" -> "Normal")
      .put(ref"B1" -> "Also Normal")

    val html = sheet.toHtml(ref"A1:B1")
    assert(html.contains("Normal"), "Normal cell should appear")
    assert(html.contains("Also Normal"), "Other cell should appear")
    assert(!html.contains("colspan"), "No colspan for unmerged cells")
    assert(!html.contains("rowspan"), "No rowspan for unmerged cells")
  }
