package com.tjclp.xl.sheets

import java.time.LocalDate

import com.tjclp.xl.{*, given}
import com.tjclp.xl.sheets.styleSyntax.getCellStyle
import munit.FunSuite

/**
 * GH-465: runtime twins of the transparent-inline `put`/`style`/`merge`/`comment` forms and of
 * `Workbook.apply(String…)`.
 *
 * The transparent forms specialize on literal-ness: a string literal returns `Sheet`, the very same
 * call with a `val` returns `XLResult[Sheet]` — and nothing at the call site says so (the top
 * scripting footgun). The twins spell `XLResult` in a plain non-inline signature, parse through the
 * `RefType` parser (so sheet-qualified refs are recognized and rejected explicitly instead of
 * failing as garbage), and reuse the literal forms' value/style paths byte-for-byte.
 */
class RuntimeTwinsSpec extends FunSuite:

  /** Defeat constant folding so the transparent forms take their runtime branch. */
  private def dynamic(s: String): String = List(s).mkString

  private val qualifiedMessage =
    "sheet-qualified refs are not accepted by Sheet.putAt/styleAt/mergeAt/commentAt; " +
      "use wb.update(sheet, ...)"

  private val bold = CellStyle.default.bold
  private val note = Comment.plainText("check this", Some("gen"))

  // ========== putAt ==========

  test("Sheet.named(s).flatMap(_.putAt(r, 1)) equals the literal form for valid strings") {
    val s = dynamic("Data")
    val r = dynamic("B2")
    val viaTwins: XLResult[Sheet] = Sheet.named(s).flatMap(_.putAt(r, 1))
    assertEquals(viaTwins, Right(Sheet("Data").put("B2", 1)): XLResult[Sheet])
  }

  test("putAt applies the literal put's codec semantics (inferred formats, style registry)") {
    val base = Sheet("Codecs")
    val date = LocalDate.of(2025, 1, 15)
    val viaTwins = for
      a <- base.putAt(dynamic("A1"), BigDecimal("1000.50"))
      b <- a.putAt(dynamic("B1"), date)
      c <- b.putAt(dynamic("C1"), "text")
      d <- c.putAt(dynamic("D1"), fx"=A1*2")
    yield d
    val literal = base
      .put(ref"A1", BigDecimal("1000.50"))
      .put(ref"B1", date)
      .put(ref"C1", "text")
      .put(ref"D1", fx"=A1*2")
    assertEquals(viaTwins, Right(literal): XLResult[Sheet])
    // Codec-inferred formats really landed (not just structural equality of two empty registries)
    viaTwins.foreach { sheet =>
      assertEquals(sheet.getCellStyle(ref"A1").map(_.numFmt), Some(NumFmt.Decimal))
      assertEquals(sheet.getCellStyle(ref"B1").map(_.numFmt), Some(NumFmt.Date))
    }
  }

  test("putAt with an explicit style equals the literal styled put (codec NumFmt merge)") {
    val date = LocalDate.of(2025, 1, 15)
    val viaTwin = Sheet("Styled").putAt(dynamic("A1"), date, bold)
    val literal = Sheet("Styled").put(ref"A1", date, bold)
    assertEquals(viaTwin, Right(literal): XLResult[Sheet])
    viaTwin.foreach { sheet =>
      val style = sheet.getCellStyle(ref"A1")
      assertEquals(style.map(_.font.bold), Some(true))
      assertEquals(style.map(_.numFmt), Some(NumFmt.Date)) // explicit General yields to codec Date
    }
  }

  test("putAt overwrites an existing cell the same way the literal put does") {
    val base = Sheet("Over").put(ref"A1", "old").put(ref"B1", 1)
    assertEquals(
      base.putAt(dynamic("A1"), 42),
      Right(base.put(ref"A1", 42)): XLResult[Sheet]
    )
  }

  test("putAt rejects a range where a single cell is required") {
    val ref = dynamic("A1:B2")
    assertEquals(
      Sheet("S").putAt(ref, 1),
      Left(XLError.InvalidCellRef("A1:B2", "expected a single cell")): XLResult[Sheet]
    )
    assertEquals(
      Sheet("S").putAt(ref, 1, bold),
      Left(XLError.InvalidCellRef("A1:B2", "expected a single cell")): XLResult[Sheet]
    )
  }

  // ========== styleAt ==========

  test("styleAt on a cell equals the literal style (withCellStyle)") {
    val base = Sheet("St").put(ref"A1", "x")
    val viaTwin = base.styleAt(dynamic("A1"), bold)
    assertEquals(viaTwin, Right(base.style(ref"A1", bold)): XLResult[Sheet])
    viaTwin.foreach(sheet => assertEquals(sheet.getCellStyle(ref"A1"), Some(bold)))
  }

  test("styleAt on a range styles every cell (withRangeStyle), creating blanks as needed") {
    val base = Sheet("St").put(ref"A1", "x").put(ref"C2", "y")
    val range = ref"A1:C2"
    val viaTwin = base.styleAt(dynamic("A1:C2"), bold)
    assertEquals(viaTwin, Right(base.style(range, bold)): XLResult[Sheet])
    viaTwin.foreach { sheet =>
      range.cells.foreach { cell =>
        assertEquals(sheet.getCellStyle(cell), Some(bold), s"unstyled: ${cell.toA1}")
      }
      assertEquals(sheet.cells.size, 6)
    }
  }

  // ========== mergeAt ==========

  test("mergeAt merges the range like the literal merge") {
    val base = Sheet("M").put(ref"A1", "title")
    val viaTwin = base.mergeAt(dynamic("A1:C1"))
    assertEquals(viaTwin, Right(base.merge(ref"A1:C1")): XLResult[Sheet])
    viaTwin.foreach(sheet => assertEquals(sheet.mergedRanges, Set(ref"A1:C1")))
  }

  test("mergeAt rejects a single cell (a merge needs a range, like the literal form)") {
    assertEquals(
      Sheet("M").mergeAt(dynamic("A1")),
      Left(XLError.InvalidRange("A1", "expected a range like A1:B2")): XLResult[Sheet]
    )
  }

  // ========== commentAt ==========

  test("commentAt adds the comment like the literal comment") {
    val base = Sheet("C").put(ref"B2", 1)
    val viaTwin = base.commentAt(dynamic("B2"), note)
    assertEquals(viaTwin, Right(base.comment(ref"B2", note)): XLResult[Sheet])
    viaTwin.foreach(sheet => assertEquals(sheet.getComment(ref"B2"), Some(note)))
  }

  test("commentAt rejects a range") {
    assertEquals(
      Sheet("C").commentAt(dynamic("A1:B2"), note),
      Left(XLError.InvalidCellRef("A1:B2", "expected a single cell")): XLResult[Sheet]
    )
  }

  // ========== shared parsing contract ==========

  test("every twin rejects a sheet-qualified ref with InvalidReference") {
    val sheet = Sheet("Q")
    val expected: XLResult[Sheet] = Left(XLError.InvalidReference(qualifiedMessage))
    val cells = List("Sales!A1", "'Q1 Sales'!A1")
    val ranges = List("Sales!A1:B2", "'Q1 Sales'!A1:B2")
    (cells ++ ranges).foreach { q =>
      assertEquals(sheet.putAt(dynamic(q), 1), expected, q)
      assertEquals(sheet.putAt(dynamic(q), 1, bold), expected, q)
      assertEquals(sheet.styleAt(dynamic(q), bold), expected, q)
      assertEquals(sheet.mergeAt(dynamic(q)), expected, q)
      assertEquals(sheet.commentAt(dynamic(q), note), expected, q)
    }
  }

  test("every twin reports garbage as InvalidCellRef / InvalidRange naming the input") {
    val sheet = Sheet("G")
    List("", "NOT A REF!!!", "A0", "XFE1", "A1048577", "1A").foreach { bad =>
      sheet.putAt(dynamic(bad), 1) match
        case Left(XLError.InvalidCellRef(r, _)) => assertEquals(r, bad)
        case other => fail(s"putAt('$bad'): expected InvalidCellRef, got $other")
      sheet.commentAt(dynamic(bad), note) match
        case Left(XLError.InvalidCellRef(r, _)) => assertEquals(r, bad)
        case other => fail(s"commentAt('$bad'): expected InvalidCellRef, got $other")
      sheet.styleAt(dynamic(bad), bold) match
        case Left(XLError.InvalidCellRef(r, _)) => assertEquals(r, bad)
        case other => fail(s"styleAt('$bad'): expected InvalidCellRef, got $other")
    }
    List("A1:", ":B2", "A1:B2:C3", "A1:XFE9", "nope:nope").foreach { bad =>
      sheet.mergeAt(dynamic(bad)) match
        case Left(XLError.InvalidRange(r, _)) => assertEquals(r, bad)
        case other => fail(s"mergeAt('$bad'): expected InvalidRange, got $other")
      sheet.styleAt(dynamic(bad), bold) match
        case Left(XLError.InvalidRange(r, _)) => assertEquals(r, bad)
        case other => fail(s"styleAt('$bad'): expected InvalidRange, got $other")
    }
  }

  test("twins accept lower-case and reversed-corner inputs the RefType parser accepts") {
    val base = Sheet("L")
    assertEquals(base.putAt(dynamic("b2"), 1), Right(base.put(ref"B2", 1)): XLResult[Sheet])
    assertEquals(
      base.mergeAt(dynamic("c3:a1")), // reversed corners normalize
      Right(base.merge(ref"A1:C3")): XLResult[Sheet]
    )
  }

  // ========== Workbook.named ==========

  test("Workbook.named(name) builds a one-sheet workbook and agrees with the dynamic apply") {
    val names = List("Sales", "Income Statement", "Q1:Q2", "", "X" * 32, "History")
    names.foreach { n =>
      val viaApply: XLResult[Workbook] = Workbook(dynamic(n))
      assertEquals(Workbook.named(n), viaApply, n)
    }
    Workbook.named("Sales") match
      case Right(wb) => assertEquals(wb.sheets.map(_.name.value), Vector("Sales"))
      case Left(err) => fail(s"expected Right, got Left($err)")
    Workbook.named("Q1:Q2") match
      case Left(XLError.InvalidSheetName(name, _)) => assertEquals(name, "Q1:Q2")
      case other => fail(s"expected Left(InvalidSheetName), got $other")
  }

  test("Workbook.named(\"A\", \"B\") has two sheets in order") {
    Workbook.named("A", "B") match
      case Right(wb) => assertEquals(wb.sheets.map(_.name.value), Vector("A", "B"))
      case Left(err) => fail(s"expected Right, got Left($err)")
    Workbook.named("A", "B", "C", "D") match
      case Right(wb) => assertEquals(wb.sheets.map(_.name.value), Vector("A", "B", "C", "D"))
      case Left(err) => fail(s"expected Right, got Left($err)")
  }

  test("Workbook.named(\"A\", \"A\") is Left(DuplicateSheet)") {
    assertEquals(Workbook.named("A", "A"), Left(XLError.DuplicateSheet("A")): XLResult[Workbook])
    assertEquals(
      Workbook.named("A", "B", "A"),
      Left(XLError.DuplicateSheet("A")): XLResult[Workbook]
    )
    // Excel compares sheet names case-insensitively: "data" repeats "Data"
    assertEquals(
      Workbook.named("Data", "data"),
      Left(XLError.DuplicateSheet("data")): XLResult[Workbook]
    )
  }

  test("Workbook.named validates every name, first failure wins in argument order") {
    Workbook.named("A", "Q1:Q2", "History") match
      case Left(XLError.InvalidSheetName(name, _)) => assertEquals(name, "Q1:Q2")
      case other => fail(s"expected Left(InvalidSheetName(Q1:Q2)), got $other")
    Workbook.named("A", "B", "") match
      case Left(XLError.InvalidSheetName(name, _)) => assertEquals(name, "")
      case other => fail(s"expected Left(InvalidSheetName('')), got $other")
    Workbook.named("A", "A", "Q1:Q2") match
      case Left(XLError.DuplicateSheet(name)) => assertEquals(name, "A")
      case other => fail(s"expected Left(DuplicateSheet(A)), got $other")
  }

  test("Workbook.named multi-sheet agrees with the dynamic apply for valid distinct names") {
    val (a, b, c) = (dynamic("Sales"), dynamic("Marketing"), dynamic("Finance"))
    val viaApply: XLResult[Workbook] = Workbook(a, b, c)
    assertEquals(Workbook.named(a, b, c), viaApply)
  }
