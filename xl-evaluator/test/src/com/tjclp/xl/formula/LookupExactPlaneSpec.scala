package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

/**
 * The exact lookup plane of MATCH, XMATCH, XLOOKUP, VLOOKUP and HLOOKUP, and rich text as text,
 * against results desktop Excel (Microsoft 365 for Mac) calculated on this grid (#724, #775, #693).
 *
 * A1:A16 = "apple", 10, "10", TRUE, 1, blank, "", `=""`, rich text "kiwi", `=NA()`, "a*c", "abc",
 * "BANANA", 0, FALSE, "1", and B1:B16 = "r1".."r16". A20:P20 holds the same keys across a row with
 * C21:… = "c1".."c16" below them. D1:E2 = pear/e-pear, kiwi/e-kiwi; G1:I5 is a 2-D block; K1:K4 =
 * "kiwi", rich "kiwi", "KIWI", "pear"; L1:L4 = rich "mango", "apple", 5, "zebra"; U1 = rich "kiwi".
 */
class LookupExactPlaneSpec extends FunSuite:

  private def rich(s: String): CellValue =
    CellValue.RichText(RichText(TextRun(s.take(2)).bold, TextRun(s.drop(2))))

  private val keys: Vector[Option[CellValue]] = Vector(
    Some(CellValue.Text("apple")),
    Some(CellValue.Number(10)),
    Some(CellValue.Text("10")),
    Some(CellValue.Bool(true)),
    Some(CellValue.Number(1)),
    None,
    Some(CellValue.Text("")),
    Some(CellValue.Formula("\"\"", Some(CellValue.Text("")))),
    Some(rich("kiwi")),
    Some(CellValue.Formula("NA()", Some(CellValue.Error(CellError.NA)))),
    Some(CellValue.Text("a*c")),
    Some(CellValue.Text("abc")),
    Some(CellValue.Text("BANANA")),
    Some(CellValue.Number(0)),
    Some(CellValue.Bool(false)),
    Some(CellValue.Text("1"))
  )

  private val block: Vector[Vector[CellValue]] = Vector(
    Vector(CellValue.Number(1), CellValue.Number(2), CellValue.Number(3)),
    Vector(CellValue.Text("fig"), CellValue.Text("x"), CellValue.Number(5)),
    Vector(CellValue.Number(6), CellValue.Number(7), CellValue.Number(8)),
    Vector(CellValue.Text("y"), CellValue.Number(9), CellValue.Text("z")),
    Vector(CellValue.Number(10), CellValue.Number(11), CellValue.Number(12))
  )

  private val sheet: Sheet =
    val withKeys = keys.zipWithIndex.foldLeft(Sheet("Data")) { case (s, (value, i)) =>
      val withResults = s
        .put(ref"B1".shift(0, i), CellValue.Text(s"r${i + 1}"))
        .put(ref"A21".shift(i, 0), CellValue.Text(s"c${i + 1}"))
      value.fold(withResults)(v =>
        withResults.put(ref"A1".shift(0, i), v).put(ref"A20".shift(i, 0), v)
      )
    }
    val withBlock = block.zipWithIndex.foldLeft(withKeys) { case (s, (row, r)) =>
      row.zipWithIndex.foldLeft(s) { case (acc, (v, c)) => acc.put(ref"G1".shift(c, r), v) }
    }
    withBlock
      .put(ref"D1", CellValue.Text("pear"))
      .put(ref"E1", CellValue.Text("e-pear"))
      .put(ref"D2", CellValue.Text("kiwi"))
      .put(ref"E2", CellValue.Text("e-kiwi"))
      .put(ref"K1", CellValue.Text("kiwi"))
      .put(ref"K2", rich("kiwi"))
      .put(ref"K3", CellValue.Text("KIWI"))
      .put(ref"K4", CellValue.Text("pear"))
      .put(ref"L1", rich("mango"))
      .put(ref"L2", CellValue.Text("apple"))
      .put(ref"L3", CellValue.Number(5))
      .put(ref"L4", CellValue.Text("zebra"))
      .put(ref"U1", rich("kiwi"))
      .put(ref"V1", CellValue.Text("abc"))
      .put(ref"V2", CellValue.Text("=abc"))
      .put(ref"W1", CellValue.Text("*"))
      .put(ref"W2", CellValue.Text("a"))
      .put(ref"X1", CellValue.Text("a~"))
      .put(ref"X2", CellValue.Text("a"))
      .put(ref"Y1", CellValue.Text("~"))
      .put(ref"Y2", CellValue.Text("a"))
      .put(ref"S1", CellValue.Text("abc"))
      .put(ref"S2", CellValue.Text("x"))
      .put(ref"Z1", CellValue.Number(10))
      .put(ref"Z2", CellValue.Text("10"))

  private def check(formula: String, expected: CellValue)(implicit loc: munit.Location): Unit =
    assertEquals(sheet.evaluateFormula(formula), Right(expected), formula)

  private def num(n: Long): CellValue = CellValue.Number(BigDecimal(n))
  private def text(s: String): CellValue = CellValue.Text(s)
  private val na = CellValue.Error(CellError.NA)

  /** An array formula's spilled column, top to bottom, as each cell's text. */
  private def spill(formula: String): Vector[String] =
    val (out, range) = sheet
      .evaluateArrayFormula(formula, ref"AA1")
      .fold(err => fail(s"$formula: $err"), identity)
    range.cellsRowMajor.toVector.map { at =>
      out(at).value match
        case CellValue.RichText(rt) => rt.toPlainText
        case CellValue.Text(s) => s
        case CellValue.Number(n) => n.toString
        case other => other.toString
    }

  // ===== #693: one exact matcher, so text and numbers never cross =====

  test("#693: exact VLOOKUP and HLOOKUP match by the key's own type") {
    check("=VLOOKUP(\"10\",A1:B16,2,FALSE)", text("r3"))
    check("=VLOOKUP(10,A1:B16,2,FALSE)", text("r2"))
    check("=VLOOKUP(1,A1:B16,2,FALSE)", text("r5"))
    check("=VLOOKUP(\"1\",A1:B16,2,FALSE)", text("r16"))
    check("=VLOOKUP(TRUE,A1:B16,2,FALSE)", text("r4"))
    check("=VLOOKUP(FALSE,A1:B16,2,FALSE)", text("r15"))
    check("=VLOOKUP(0,A1:B16,2,FALSE)", text("r14"))
    check("=VLOOKUP(\"banana\",A1:B16,2,FALSE)", text("r13"))
    check("=VLOOKUP(\"10\",Z1:Z2,1,FALSE)", text("10"))
    check("=HLOOKUP(\"10\",A20:P21,2,FALSE)", text("c3"))
    check("=HLOOKUP(1,A20:P21,2,FALSE)", text("c5"))
    check("=HLOOKUP(TRUE,A20:P21,2,FALSE)", text("c4"))
    check("=HLOOKUP(\"1\",A20:P21,2,FALSE)", text("c16"))
  }

  test("#693: MATCH, XMATCH and XLOOKUP keep their type rule") {
    check("=MATCH(\"10\",A1:A16,0)", num(3))
    check("=MATCH(10,A1:A16,0)", num(2))
    check("=MATCH(1,A1:A16,0)", num(5))
    check("=MATCH(TRUE,A1:A16,0)", num(4))
    check("=MATCH(0,A1:A16,0)", num(14))
    check("=MATCH(\"1\",A1:A16,0)", num(16))
    check("=SUM(MATCH({\"10\",10},A1:A16,0))", num(5))
    check("=XMATCH(\"10\",A1:A16)", num(3))
    check("=XMATCH(1,A1:A16)", num(5))
    check("=XMATCH(FALSE,A1:A16)", num(15))
    check("=XLOOKUP(\"1\",A1:A16,B1:B16,\"nf\")", text("r16"))
    check("=XLOOKUP(TRUE,A1:A16,B1:B16,\"nf\")", text("r4"))
  }

  // ===== #693: MATCH type 0 and exact VLOOKUP/HLOOKUP apply wildcards =====

  test("#693: exact MATCH, VLOOKUP and HLOOKUP apply * and ?") {
    check("=MATCH(\"a*\",A1:A16,0)", num(1))
    check("=MATCH(\"a?c\",A1:A16,0)", num(11))
    check("=VLOOKUP(\"a*\",A1:B16,2,FALSE)", text("r1"))
    check("=VLOOKUP(\"1?\",A1:B16,2,FALSE)", text("r3"))
    check("=HLOOKUP(\"a*\",A20:P21,2,FALSE)", text("c1"))
    check("=HLOOKUP(\"k?wi\",A20:P21,2,FALSE)", text("c9"))
    check("=COUNT(MATCH({\"a*\",\"x\"},S1:S2,0))", num(2))
  }

  test("#693: a wildcard matches text cells only, empty text included, never a blank") {
    check("=MATCH(\"*\",A2:A16,0)", num(2))
    check("=MATCH(\"1*\",A1:A16,0)", num(3))
    check("=VLOOKUP(\"*\",A2:B16,2,FALSE)", text("r3"))
    check("=MATCH(\"*\",A6:A16,0)", num(2))
  }

  test("#693: exact MATCH and VLOOKUP read ~ with the lookup escape grammar") {
    check("=MATCH(\"a~*c\",A1:A16,0)", num(11))
    check("=VLOOKUP(\"a~*c\",A1:B16,2,FALSE)", text("r11"))
    check("=MATCH(\"~*\",W1:W2,0)", num(1))
    check("=MATCH(\"~a~bc\",A1:A16,0)", num(12))
    // a trailing lone ~ drops: "a~" finds "a", and a lone "~" is the empty pattern
    check("=MATCH(\"a~\",X1:X2,0)", num(2))
    check("=VLOOKUP(\"~\",Y1:Y2,1,FALSE)", na)
  }

  test("#693: XMATCH and XLOOKUP match_mode 0 compare the lookup value as written") {
    check("=XMATCH(\"a*\",A1:A16)", na)
    check("=XMATCH(\"a*c\",A1:A16)", num(11))
    check("=XLOOKUP(\"a*\",A1:A16,B1:B16,\"nf\")", text("nf"))
    check("=XLOOKUP(\"k*\",A1:A16,B1:B16,\"nf\",2)", text("r9"))
  }

  test("#693: a leading = in a match_mode 2 lookup value is literal") {
    check("=XMATCH(\"=a*\",V1:V2,2)", num(2))
  }

  test("#693: empty text finds empty text, never a blank cell") {
    check("=MATCH(\"\",A1:A16,0)", num(7))
    check("=MATCH(\"\",A8:A16,0)", num(1))
    check("=VLOOKUP(\"\",A1:B16,2,FALSE)", text("r7"))
    check("=XLOOKUP(\"\",A1:A16,B1:B16,\"nf\")", text("r7"))
    check("=MATCH(A7,A1:A16,0)", num(7))
  }

  test("#693: a blank lookup value in MATCH, VLOOKUP and HLOOKUP is the number 0") {
    check("=MATCH(A6,A1:A16,0)", num(14))
    check("=VLOOKUP(A6,A1:B16,2,FALSE)", text("r14"))
    check("=HLOOKUP(A6,A20:P21,2,FALSE)", text("c14"))
  }

  test("#693: an error key is skipped and an error lookup value propagates") {
    check("=MATCH(\"abc\",A10:A12,0)", num(3))
    check("=MATCH(A10,A1:A16,0)", na)
  }

  // ===== #724: MATCH over a 2-D range =====

  test("#724: MATCH over more than one row and column is #N/A in every mode") {
    check("=MATCH(\"fig\",G1:I5,0)", na)
    check("=MATCH(3,G1:I5,0)", na)
    check("=MATCH(3,G1:I5)", na)
    check("=MATCH(3,G1:I5,-1)", na)
    check("=MATCH(3,G1:I1,0)", num(3))
    check("=MATCH(1,G1:G1,0)", num(1))
    check("=INDEX(G1:I5,MATCH(\"fig\",G1:G5,0),1)", text("fig"))
  }

  test("#724: XMATCH and XLOOKUP over 2-D stay #VALUE!") {
    check("=XMATCH(3,G1:I5)", CellValue.Error(CellError.Value))
    check("=XLOOKUP(3,G1:I5,G1:I5,\"nf\")", CellValue.Error(CellError.Value))
  }

  // ===== #775: rich text counts as text =====

  test("#775: a rich-text key matches its plain text in every exact lookup") {
    check("=MATCH(\"kiwi\",A1:A16,0)", num(9))
    check("=MATCH(\"kiwi\",U1,0)", num(1))
    check("=XMATCH(\"kiwi\",A1:A16)", num(9))
    check("=XLOOKUP(\"kiwi\",A1:A16,B1:B16,\"nf\")", text("r9"))
    check("=VLOOKUP(\"kiwi\",A1:B16,2,FALSE)", text("r9"))
    check("=HLOOKUP(\"kiwi\",A20:P21,2,FALSE)", text("c9"))
    check("=MATCH(\"k*\",A1:A16,0)", num(9))
  }

  test("#775: a rich-text lookup value matches as text") {
    check("=MATCH(U1,D1:D2,0)", num(2))
    check("=XLOOKUP(U1,D1:D2,E1:E2,\"nf\")", text("e-kiwi"))
    check("=VLOOKUP(U1,D1:E2,2,FALSE)", text("e-kiwi"))
  }

  test("#775: ISTEXT is TRUE and ISNONTEXT FALSE for rich text") {
    check("=ISTEXT(A9)", CellValue.Bool(true))
    check("=ISTEXT(I20)", CellValue.Bool(true))
    check("=ISNONTEXT(A9)", CellValue.Bool(false))
    check("=T(A9)", text("kiwi"))
  }

  test("#775: ISNONTEXT is TRUE for a blank and FALSE for empty text") {
    check("=ISNONTEXT(A6)", CellValue.Bool(true))
    check("=ISNONTEXT(A7)", CellValue.Bool(false))
    check("=ISNONTEXT(A8)", CellValue.Bool(false))
    check("=ISNONTEXT(A2)", CellValue.Bool(true))
    check("=ISNONTEXT(A10)", CellValue.Bool(true))
    check("=ISTEXT(A7)", CellValue.Bool(true))
    check("=ISTEXT(A8)", CellValue.Bool(true))
  }

  test("#775: UNIQUE reads rich text as its plain text, case-insensitively") {
    check("=ROWS(UNIQUE(K1:K4))", num(2))
    assertEquals(spill("=UNIQUE(K1:K4)"), Vector("kiwi", "pear"))
  }

  test("#775: SORT places rich text among the text") {
    assertEquals(spill("=SORT(L1:L4)"), Vector("5", "apple", "mango", "zebra"))
    assertEquals(spill("=SORT(L1:L4,1,-1)"), Vector("zebra", "mango", "apple", "5"))
  }
