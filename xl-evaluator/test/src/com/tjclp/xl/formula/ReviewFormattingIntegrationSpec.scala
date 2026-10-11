package com.tjclp.xl.formula

import java.math.BigInteger

import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.render.RenderUtils
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.syntax.*

import munit.FunSuite

/**
 * Integration probes for PRs #706 and #708; expected numeric coercions verified in Excel 16.113.4.
 */
class ReviewFormattingIntegrationSpec extends FunSuite:
  private val sheet = Sheet("Review")

  test("VALUE accepts a long leading group before valid thousands groups") {
    List("1234,567", "12345,678", "+1234,567", "-1234,567", "$1234,567").foreach { input =>
      val expected = BigDecimal(input.replace(",", "").replace("$", ""))
      assertEquals(
        sheet.evaluateFormula(s"=VALUE(\"$input\")"),
        Right(CellValue.Number(expected)),
        input
      )
    }
  }

  test("TEXT converts a partially grouped numeric string before applying numeric format") {
    assertEquals(
      sheet.evaluateFormula("=TEXT(\"1234,567\",\"0.00\")"),
      Right(CellValue.Text("1234567.00"))
    )
  }

  test("TEXT and VALUE keep rejecting malformed non-leading thousands groups") {
    List("1,23", "12,34,567", "1.2,3", "1e3,0", "1$2", "$$123").foreach { input =>
      assertEquals(
        sheet.evaluateFormula(s"=VALUE(\"$input\")"),
        Right(CellValue.Error(CellError.Value)),
        input
      )
      assertEquals(
        sheet.evaluateFormula(s"=TEXT(\"$input\",\"0.00\")"),
        Right(CellValue.Text(input)),
        input
      )
    }
  }

  test("TEXT applies fraction formats to full-precision negative extreme values") {
    val unscaled = new BigInteger("1234567890" * 4).negate
    val value = BigDecimal(new java.math.BigDecimal(unscaled, Int.MinValue))
    val extreme = sheet.put(ref"A1", CellValue.Number(value))
    List("# ?/?", "# ?/8", "?/8").foreach { code =>
      val expected = NumFmtFormatter.formatValue(
        CellValue.Number(value),
        NumFmt.Custom(code),
        NumFmtFormatter.GeneralRule.Text
      )
      assert(expected.startsWith("-"), code)
      assert(expected.length < 80, code)
      assertEquals(
        extreme.evaluateFormula(s"=TEXT(A1,\"$code\")"),
        Right(CellValue.Text(expected)),
        code
      )
    }
  }

  test("rich text uses the same text section in scalar display and render content") {
    val value = CellValue.RichText(RichText.plain("abc"))
    val format = NumFmt.Custom("0;-0;0;\"T:\"@")
    assertEquals(NumFmtFormatter.formatValue(value, format), "T:abc")
    assertEquals(RenderUtils.renderedContent(value, format).text, "T:abc")
  }

  test("TEXT numeric-string coercion is independent of rich-text storage") {
    val cells = sheet
      .put(ref"A1", CellValue.Text("123"))
      .put(ref"A2", CellValue.RichText(RichText.plain("123")))
    assertEquals(cells.evaluateFormula("=TEXT(A1,\"0.00\")"), Right(CellValue.Text("123.00")))
    assertEquals(cells.evaluateFormula("=TEXT(A2,\"0.00\")"), Right(CellValue.Text("123.00")))
  }
