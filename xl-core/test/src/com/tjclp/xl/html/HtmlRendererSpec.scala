package com.tjclp.xl.html

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.richtext.RichText.{*, given}
import com.tjclp.xl.macros.ref
// Removed: BatchPutMacro is dead code (shadowed by Sheet.put member)  // For batch put extension
import com.tjclp.xl.codec.syntax.*
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
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("<table>"), "Should contain table tag")
    assert(html.contains("<td></td>"), "Empty cell should render as empty td")
  }

  test("toHtml: single cell with text") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1", CellValue.Text("Hello"))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("<table>"), "Should contain table tag")
    assert(html.contains("Hello"), "Should contain cell text")
  }

  test("toHtml: 2x2 grid") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(
        ref"A1" -> "A1",
        ref"B1" -> "B1",
        ref"A2" -> "A2",
        ref"B2" -> "B2"
      )
      .unsafe

    val html = sheet.toHtml(ref"A1:B2")
    assert(html.contains("<table>"))
    assert(html.contains("A1"))
    assert(html.contains("B2"))
    // Should have 2 rows
    val rowCount = html.split("<tr>").length - 1
    assertEquals(rowCount, 2, "Should have 2 rows")
  }

  // ========== Rich Text Export ==========

  test("toHtml: rich text with bold") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1", CellValue.RichText("Bold".bold + " normal"))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("<b>Bold</b>"), "Should render bold tag")
    assert(html.contains(" normal"), "Should include normal text")
  }

  test("toHtml: rich text with italic") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1", CellValue.RichText("Italic".italic))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("<i>Italic</i>"), "Should render italic tag")
  }

  test("toHtml: rich text with underline") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1", CellValue.RichText("Underline".underline))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("<u>Underline</u>"), "Should render underline tag")
  }

  test("toHtml: rich text with colors") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1", CellValue.RichText("Red".red + " and " + "Green".green))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("color:"), "Should include color style")
    assert(html.contains("<span"), "Should use span for color")
  }

  test("toHtml: rich text with mixed formatting") {
    val text = "Error: ".red.bold + "File not found".underline
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
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

    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
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
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1" -> "Bold")
      .unsafe
      .withCellStyle(ref"A1", boldStyle)

    val html = sheet.toHtml(ref"A1:A1", includeStyles = false)
    assert(!html.contains("style="), "Should not include inline styles")
  }

  // ========== HTML Escaping ==========

  test("toHtml: escapes HTML special characters") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1", CellValue.Text("<script>alert('xss')</script>"))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("&lt;script&gt;"), "Should escape < and >")
    assert(!html.contains("<script>"), "Should not contain unescaped script tag")
  }

  test("toHtml: escapes ampersands") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1", CellValue.Text("A & B"))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("A &amp; B"), "Should escape ampersand")
  }

  test("toHtml: escapes quotes") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1", CellValue.Text("""Say "Hello""""))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("&quot;"), "Should escape quotes")
  }

  test("toHtml: escapes comment tooltip content (HTML + quotes)") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
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
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1" -> BigDecimal("123.45"))
      .unsafe

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("123.45"))
  }

  test("toHtml: boolean cells") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1" -> true)
      .unsafe

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("true"))
  }

  test("toHtml: date cells") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1" -> java.time.LocalDate.of(2025, 11, 10))
      .unsafe

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("2025"))
  }

  test("toHtml: formula cells") {
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1", CellValue.Formula("SUM(B1:B10)"))

    val html = sheet.toHtml(ref"A1:A1")
    assert(html.contains("=SUM"), "Should include formula")
  }

  // ========== Integration with putMixed ==========

  test("toHtml: financial model example") {
    val sheet = Sheet("Q1 Report").getOrElse(fail("Sheet creation failed"))
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
    val sheet = Sheet("Test").getOrElse(fail("Sheet creation failed"))
      .put(ref"A1", CellValue.RichText("Bold".bold + " text"))
      .put(ref"A2", CellValue.Text("Plain text"))

    val html = sheet.toHtml(ref"A1:A2")
    assert(html.contains("<b>Bold</b>"), "Rich text should be formatted")
    assert(html.contains("Plain text"), "Plain text should be included")
  }

  //========== Real-World Use Case ==========

  test("toHtml: complete financial report with rich text and styles") {
    val headerStyle = CellStyle.default
      .withFont(Font("Arial", 14.0, bold = true))
      .withFill(Fill.Solid(Color.fromRgb(220, 220, 220)))
      .withAlign(Align(HAlign.Center, VAlign.Middle))

    val positive = "+12.5%".green.bold
    val negative = "-5.2%".red.bold

    val sheet = Sheet("Report").getOrElse(fail("Sheet creation failed"))
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
    assert(html.contains("<table>"))
    assert(html.contains("Metric"))
    assert(html.contains("Change"))

    // Verify header styles
    assert(html.contains("background-color:"), "Headers should have background")
    assert(html.contains("text-align: center"), "Headers should be centered")

    // Verify rich text
    assert(html.contains("+12.5%"), "Should include percentage")
    assert(html.contains("-5.2%"))
  }
