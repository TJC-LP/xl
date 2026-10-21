package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

/**
 * Criteria and match_mode 2 lookups against results desktop Excel (Microsoft 365 for Mac)
 * calculated on this grid (#774, #773, #780, #731).
 *
 * A1:A14 = "abc", "a*", "*", 4, 5, TRUE, blank, `=""`, rich text "ab", `=NA()`, "4", "a?c", "a~b",
 * 45, and B1:B14 = 2^(row-1), so a SUMIF over B is the bitmask of the rows a criterion matches.
 * C1:C3 = "abc", "a*", "*" and D1:D3 = "r1", "r2", "r3"; E1:E3 = 4, "x", 7; F1 = the text
 * "1E+999999999", F2 = 1E+300; G1:H3 = [4, blank; "a", 4; `=""`, TRUE].
 */
class CriteriaExcelParitySpec extends FunSuite:

  private val columnA: Vector[Option[CellValue]] = Vector(
    Some(CellValue.Text("abc")),
    Some(CellValue.Text("a*")),
    Some(CellValue.Text("*")),
    Some(CellValue.Number(4)),
    Some(CellValue.Number(5)),
    Some(CellValue.Bool(true)),
    None,
    Some(CellValue.Formula("\"\"", Some(CellValue.Text("")))),
    Some(CellValue.RichText(RichText(TextRun("a"), TextRun("b")))),
    Some(CellValue.Formula("NA()", Some(CellValue.Error(CellError.NA)))),
    Some(CellValue.Text("4")),
    Some(CellValue.Text("a?c")),
    Some(CellValue.Text("a~b")),
    Some(CellValue.Number(45))
  )

  private val sheet: Sheet =
    val withA = columnA.zipWithIndex.foldLeft(Sheet("Data")) { case (s, (value, i)) =>
      val withB = s.put(ref"B1".shift(0, i), CellValue.Number(BigDecimal(2).pow(i)))
      value.fold(withB)(withB.put(ref"A1".shift(0, i), _))
    }
    withA
      .put(ref"C1", CellValue.Text("abc"))
      .put(ref"C2", CellValue.Text("a*"))
      .put(ref"C3", CellValue.Text("*"))
      .put(ref"D1", CellValue.Text("r1"))
      .put(ref"D2", CellValue.Text("r2"))
      .put(ref"D3", CellValue.Text("r3"))
      .put(ref"E1", CellValue.Number(4))
      .put(ref"E2", CellValue.Text("x"))
      .put(ref"E3", CellValue.Number(7))
      .put(ref"F1", CellValue.Text("1E+999999999"))
      .put(ref"F2", CellValue.Number(BigDecimal("1E+300")))
      .put(ref"G1", CellValue.Number(4))
      .put(ref"G2", CellValue.Text("a"))
      .put(ref"H2", CellValue.Number(4))
      .put(ref"G3", CellValue.Formula("\"\"", Some(CellValue.Text(""))))
      .put(ref"H3", CellValue.Bool(true))

  private def check(formula: String, expected: CellValue)(implicit loc: munit.Location): Unit =
    assertEquals(sheet.evaluateFormula(formula), Right(expected), formula)

  private def num(n: Long): CellValue = CellValue.Number(BigDecimal(n))

  /** The SUMIF bitmask of the 1-based rows of A1:A14 a criterion matches. */
  private def rows(rs: Int*): CellValue = num(rs.map(r => 1L << (r - 1)).sum)

  private val na = CellValue.Error(CellError.NA)

  // ===== #774: XMATCH/XLOOKUP match_mode 2 read the lookup value with the escape grammar =====

  test("#774: an escaped wildcard in match_mode 2 finds the literal character") {
    check("=XMATCH(\"~*\",C1:C3,2)", num(3))
    check("=XMATCH(\"a~*\",C1:C3,2)", num(2))
    check("=XMATCH(\"a*\",C1:C3,2)", num(1))
    check("=XLOOKUP(\"~*\",C1:C3,D1:D3,\"nf\",2)", CellValue.Text("r3"))
    check("=XLOOKUP(\"a~*\",C1:C3,D1:D3,\"nf\",2,-1)", CellValue.Text("r2"))
    check("=XMATCH(\"a~?c\",A1:A14,2)", num(12))
    check("=SUM(XMATCH({\"~*\",\"a~*\"},C1:C3,2))", num(5))
  }

  test("#774: match_mode 2 drops a ~ before any character, wildcard or not") {
    check("=XMATCH(\"a~~b\",A1:A14,2)", num(13))
    check("=XMATCH(\"~4\",A1:A14,2)", num(11))
    // a lone trailing ~ stays a literal tilde ("a~" is in no cell)
    check("=XMATCH(\"a~\",A1:A14,2)", na)
  }

  test("#774: match_mode 0 compares the lookup value as written") {
    check("=XMATCH(\"a~~b\",A1:A14,0)", na)
    check("=XMATCH(\"~*\",C1:C3,0)", na)
  }

  test("#774/#780: match_mode 2 wildcards match text cells only") {
    check("=XMATCH(4,A1:A14,2)", num(4))
    check("=XMATCH(\"4\",A1:A14,2)", num(11))
    check("=XMATCH(\"4*\",A1:A14,2)", num(11))
    check("=XMATCH(\"*\",A4:A14,2)", num(5))
    check("=XMATCH(\"t*\",A1:A14,2)", na)
  }

  // ===== #773: "=op" reads op as a bare operand =====

  test("#773: =op applies wildcards like the bare operand") {
    check("=SUMIF(A1:A14,\"a*\",B1:B14)", rows(1, 2, 9, 12, 13))
    check("=SUMIF(A1:A14,\"=a*\",B1:B14)", rows(1, 2, 9, 12, 13))
    check("=COUNTIF(A1:A14,\"=a~*\")", num(1))
    check("=SUMIF(A1:A14,\"=?\",B1:B14)", rows(3, 11))
    check("=SUMIF(A1:A14,\"=4*\",B1:B14)", rows(11))
    check("=SUMIF(A1:A14,\"=*\",B1:B14)", rows(1, 2, 3, 8, 9, 11, 12, 13))
    check("=SUMIF(A1:A14,\"=a~~b*\",B1:B14)", rows(13))
    check("=COUNTIFS(A1:A14,\"=a*\",B1:B14,\">2\")", num(3))
    check("=SUM(COUNTIF(A1:A14,{\"=a*\",\"a*\",\"=~*\"}))", num(11))
  }

  test("#773: =op with no * or ? compares as written") {
    check("=SUMIF(A1:A14,\"=4\",B1:B14)", rows(4, 11))
    check("=SUMIF(A1:A14,\"=TRUE\",B1:B14)", rows(6))
    check("=SUMIF(A1:A14,\"=a~~b\",B1:B14)", num(0))
    check("=SUMIF(A1:A14,\"=a~b\",B1:B14)", rows(13))
    check("=SUMIF(A1:A14,\"a~~b\",B1:B14)", num(0))
    check("=SUMIF(A1:A14,\"a~b\",B1:B14)", rows(13))
    check("=SUMIF(A1:A14,\"4\",B1:B14)", rows(4, 11))
    check("=SUMIF(A1:A14,4,B1:B14)", rows(4, 11))
  }

  test("#773: a bare = matches truly empty cells only; \"\" also matches empty text") {
    check("=SUMIF(A1:A14,\"=\",B1:B14)", rows(7))
    check("=SUMIF(A1:A14,\"\",B1:B14)", rows(7, 8))
  }

  test("#773: in an operand with * or ?, ~ escapes any character") {
    check("=SUMIF(A1:A14,\"~*\",B1:B14)", rows(3))
    check("=SUMIF(A1:A14,\"a~~b*\",B1:B14)", rows(13))
    check("=SUMIF(A1:A14,\"a~b*\",B1:B14)", rows(1, 9))
    check("=SUMIF(A1:A14,\"~a~b*\",B1:B14)", rows(1, 9))
  }

  // ===== #780: wildcard criteria match text cells only =====

  test("#780: wildcard criteria skip numbers, booleans, errors and blanks") {
    check("=SUMIF(A1:A14,\"*\",B1:B14)", rows(1, 2, 3, 8, 9, 11, 12, 13))
    check("=COUNTIF(A1:A14,\"*\")", num(8))
    check("=SUMIF(A1:A14,\"?\",B1:B14)", rows(3, 11))
    check("=SUMIF(A1:A14,\"4*\",B1:B14)", rows(11))
    check("=SUMIF(A1:A14,\"T*\",B1:B14)", num(0))
    check("=COUNTIF(G1:H3,\"*\")", num(2))
    check("=MAXIFS(B1:B14,A1:A14,\"*\")", num(4096))
  }

  test("#780: a negated wildcard matches every non-text cell, errors included") {
    check("=SUMIF(A1:A14,\"<>*\",B1:B14)", rows(4, 5, 6, 7, 10, 14))
    check("=SUMIF(A1:A14,\"<>4*\",B1:B14)", rows(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14))
    check("=SUMIF(A1:A14,\"<>a*\",B1:B14)", rows(3, 4, 5, 6, 7, 8, 10, 11, 14))
  }

  test("#780: an extreme-exponent criterion is text and never throws") {
    check("=COUNTIF(F1:F2,\"1E+999999999\")", num(1))
    check("=COUNTIF(F1:F2,\"1E*\")", num(1))
    check("=SUMIF(A1:A14,\"=1E+999999999\",B1:B14)", num(0))
    check("=COUNTIF(F1:F2,\"<>1E+999999999\")", num(1))
  }

  // ===== #731: "<>n" matches every cell but a number equal to n =====

  test("#731: <>n counts blanks, text, numeric text, booleans and errors") {
    check("=SUMIF(A1:A14,\"<>4\",B1:B14)", rows(1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14))
    check("=COUNTIF(A1:A14,\"<>4\")", num(13))
    check("=COUNTIF(E:E,\"<>4\")", num(1048575))
    check("=COUNTIF(G1:H3,\"<>4\")", num(4))
    check("=COUNTIF(A10,\"<>4\")", num(1))
    check("=COUNTIFS(A1:A14,\"<>4\",A1:A14,\"<>5\")", num(12))
    check("=SUMIF(A1:A14,\"<>\"&\"4\",B1:B14)", rows(1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14))
    check("=SUMIF(A1:A14,\"<>5\",B1:B14)", rows(1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14))
  }

  test("#731: every <> criterion counts error cells") {
    check("=SUMIF(A1:A14,\"<>abc\",B1:B14)", rows(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14))
    check("=SUMIF(A1:A14,\"<>TRUE\",B1:B14)", rows(1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12, 13, 14))
    check("=SUMIF(A1:A14,\"<>~*\",B1:B14)", rows(1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14))
    check("=SUMIF(A1:A14,\"<>a~~b\",B1:B14)", rows((1 to 14)*))
    check("=SUMIF(A1:A14,\"<>=4\",B1:B14)", rows((1 to 14)*))
  }

  test("#731: a bare <> matches every cell but a truly empty one") {
    check("=SUMIF(A1:A14,\"<>\",B1:B14)", rows(1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 14))
  }
