package com.tjclp.xl

import munit.FunSuite
import com.tjclp.xl.style.color.Color
import com.tjclp.xl.style.font.Font
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.richtext.RichText.{*, given}
import com.tjclp.xl.cell.{Cell, CellValue}
import com.tjclp.xl.macros.ref

/** Tests for rich text DSL and TextRun/RichText domain model */
class RichTextSpec extends FunSuite:

  // ========== TextRun Tests ==========

  test("TextRun: create plain run") {
    val run = TextRun("Hello")
    assertEquals(run.text, "Hello")
    assertEquals(run.font, None)
  }

  test("TextRun: bold modifier creates run with bold font") {
    val run = TextRun("Bold").bold
    assert(run.font.exists(_.bold), "Font should be bold")
  }

  test("TextRun: italic modifier creates run with italic font") {
    val run = TextRun("Italic").italic
    assert(run.font.exists(_.italic), "Font should be italic")
  }

  test("TextRun: underline modifier creates run with underline font") {
    val run = TextRun("Underline").underline
    assert(run.font.exists(_.underline), "Font should be underlined")
  }

  test("TextRun: red color modifier") {
    val run = TextRun("Red").red
    assert(run.font.exists(_.color.isDefined), "Font should have color")
  }

  test("TextRun: chained modifiers work") {
    val run = TextRun("Bold Red Underline").bold.red.underline
    run.font match
      case Some(f) =>
        assert(f.bold, "Should be bold")
        assert(f.underline, "Should be underlined")
        assert(f.color.isDefined, "Should have color")
      case None => fail("Font should be defined")
  }

  test("TextRun: size modifier") {
    val run = TextRun("Big").size(18.0)
    assert(run.font.exists(_.sizePt == 18.0), "Font size should be 18pt")
  }

  test("TextRun: fontFamily modifier") {
    val run = TextRun("Calibri Text").fontFamily("Calibri")
    assert(run.font.exists(_.name == "Calibri"), "Font family should be Calibri")
  }

  test("TextRun: withColor with custom color") {
    val purple = Color.fromHex("#800080").toOption.getOrElse(fail("Should parse color"))
    val run = TextRun("Purple").withColor(purple)
    assert(run.font.exists(_.color.contains(purple)), "Should have purple color")
  }

  // ========== String Extension DSL ==========

  test("String.bold extension creates bold run") {
    val run = "Bold".bold
    assert(run.font.exists(_.bold), "Should be bold")
  }

  test("String.italic extension creates italic run") {
    val run = "Italic".italic
    assert(run.font.exists(_.italic), "Should be italic")
  }

  test("String.red extension creates red run") {
    val run = "Red".red
    assert(run.font.exists(_.color.isDefined), "Should have color")
  }

  test("String extensions: chained") {
    val run = "Bold Red".bold.red
    run.font match
      case Some(f) =>
        assert(f.bold)
        assert(f.color.isDefined)
      case None => fail("Font should be defined")
  }

  test("String.size extension") {
    val run = "Big".size(20.0)
    assert(run.font.exists(_.sizePt == 20.0))
  }

  // ========== RichText Composition ==========

  test("RichText: single run") {
    val text = RichText(TextRun("Hello"))
    assertEquals(text.runs.size, 1)
    assertEquals(text.toPlainText, "Hello")
  }

  test("RichText: multiple runs") {
    val text = RichText(
      TextRun("Bold", Some(Font.default.withBold(true))),
      TextRun(" normal "),
      TextRun("Italic", Some(Font.default.withItalic(true)))
    )
    assertEquals(text.runs.size, 3)
    assertEquals(text.toPlainText, "Bold normal Italic")
  }

  test("RichText: concatenation with + operator") {
    val text1 = RichText(TextRun("Hello"))
    val text2 = RichText(TextRun(" World"))
    val combined = text1 + text2
    assertEquals(combined.runs.size, 2)
    assertEquals(combined.toPlainText, "Hello World")
  }

  test("RichText: add TextRun with +") {
    val text = RichText(TextRun("Hello"))
    val updated = text + TextRun(" World")
    assertEquals(updated.runs.size, 2)
  }

  test("RichText: DSL composition") {
    val text = "Bold".bold + " normal " + "Italic".italic
    assertEquals(text.runs.size, 3)
    assert(text.runs(0).font.exists(_.bold), "First run should be bold")
    assert(text.runs(1).font.isEmpty, "Second run should have no formatting")
    assert(text.runs(2).font.exists(_.italic), "Third run should be italic")
  }

  test("RichText: complex composition") {
    val text = "Error: ".red.bold + "File not found".underline + " (code: " + "404".bold + ")"
    assertEquals(text.runs.size, 5)
    assertEquals(text.toPlainText, "Error: File not found (code: 404)")
  }

  test("RichText.isPlainText: returns true for no formatting") {
    val text = RichText(TextRun("Plain"))
    assert(text.isPlainText)
  }

  test("RichText.isPlainText: returns false for formatted runs") {
    val text = RichText(TextRun("Bold").bold)
    assert(!text.isPlainText)
  }

  test("RichText.plain: creates plain text") {
    val text = RichText.plain("Hello")
    assertEquals(text.runs.size, 1)
    assert(text.isPlainText)
  }

  // ========== Given Conversions ==========

  test("String implicitly converts to TextRun") {
    val run: TextRun = "Hello"
    assertEquals(run.text, "Hello")
  }

  test("TextRun implicitly converts to RichText") {
    val text: RichText = TextRun("Hello")
    assertEquals(text.runs.size, 1)
  }

  test("Given conversions enable DSL composition") {
    // This compiles because String → TextRun → RichText
    val text: RichText = "Hello".bold + " World"
    assertEquals(text.runs.size, 2)
  }

  // ========== CellValue Integration ==========

  test("CellValue.RichText stores rich text") {
    val richText = "Bold".bold + " normal"
    val value = CellValue.RichText(richText)
    value match
      case CellValue.RichText(rt) => assertEquals(rt.toPlainText, "Bold normal")
      case other => fail(s"Expected RichText, got $other")
  }

  test("Cell with RichText value") {
    val richText = "Title".bold.size(18.0)
    val cellVal = Cell(ref"A1", CellValue.RichText(richText))
    assertEquals(cellVal.value, CellValue.RichText(richText))
  }

  // ========== Color Shortcuts ==========

  test("TextRun: all color shortcuts work") {
    val red = "Red".red
    val green = "Green".green
    val blue = "Blue".blue
    val black = "Black".black
    val white = "White".white

    assert(red.font.exists(_.color.isDefined))
    assert(green.font.exists(_.color.isDefined))
    assert(blue.font.exists(_.color.isDefined))
    assert(black.font.exists(_.color.isDefined))
    assert(white.font.exists(_.color.isDefined))
  }

  // ========== Realistic Use Cases ==========

  test("Financial model: positive/negative with colors") {
    val positive = "+12.5%".green.bold
    val negative = "-5.2%".red.bold
    val neutral = "0.0%"

    val text = "Q1: ".bold + positive + " | Q2: ".bold + negative + " | Q3: ".bold + neutral
    assert(text.runs.size == 6, "Should have 6 runs") // Q1:, +12.5%, |Q2:, -5.2%, |Q3:, 0.0%
    assertEquals(text.toPlainText, "Q1: +12.5% | Q2: -5.2% | Q3: 0.0%")
  }

  test("Report header: mixed sizes and styles") {
    val title = "Annual Report ".size(18.0).bold + "2025".size(18.0).italic
    assertEquals(title.runs.size, 2)
    assert(title.runs(0).font.exists(f => f.bold && f.sizePt == 18.0), "First run should be bold 18pt")
    assert(title.runs(1).font.exists(f => f.italic && f.sizePt == 18.0), "Second run should be italic 18pt")
  }
