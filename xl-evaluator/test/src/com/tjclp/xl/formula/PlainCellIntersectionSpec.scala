package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * A plain formula cell (a `<f>` without `t="array"`) is a legacy formula to Excel: Excel 365 opens
 * it with implicit intersection wherever the legacy evaluation applied it (it shows `@` there), and
 * LibreOffice computes the same values. A reference in a value position — an operand, a scalar
 * argument, a criterion, an IF condition, the CHOOSE index, the whole formula — reads the cell in
 * the formula's own row (a column) or column (a row), and is `#VALUE!` where the formula's row or
 * column does not cross it. An aggregate's argument takes a reference whole, but an expression in
 * it is evaluated as a value. IF and CHOOSE return references.
 *
 * Every expected value below is LibreOffice 24.2's recalculation of the same plain `<f>` cells (the
 * cell each formula sits in is part of the case). Rows where LibreOffice is known to differ from
 * Excel (named formulas, `SUM(ROW(r))`, a 2-D intersection across sheets, an error criterion) are
 * pinned in the tests after the table, to Excel's documented behaviour.
 */
class PlainCellIntersectionSpec extends FunSuite:

  private def N(n: String): CellValue = CellValue.Number(BigDecimal(n))
  private def T(s: String): CellValue = CellValue.Text(s)
  private def B(b: Boolean): CellValue = CellValue.Bool(b)
  private def E(code: String): CellValue =
    CellValue.Error(CellError.parse(code).fold(msg => fail(msg), identity))

  private val a = Vector("1", "-2", "3.5", "0", "10", "7", "-4.25", "100", "2", "5")
  private val c = Vector(
    "apple",
    "banana",
    "cherry",
    "date",
    "elder",
    "fig",
    "grape",
    "honeydew",
    "iceberg",
    "jam"
  )
  private val d = Vector(1, 2, 3, 1, 2, 3, 1, 2, 3, 1)

  /**
   * Sheet1: A1:A10 numbers, B1:B10 = 2, C1:C10 text, D1:D10 an index column, A12:E12 = 1..5;
   * Other!A1:A10 = 10..100; the names nmRef (a reference), nmExpr and nmAbs (formulas).
   */
  private val book: Workbook =
    val sheet1 = (0 until 10).foldLeft(Sheet(SheetName.unsafe("Sheet1"))) { (s, i) =>
      s.put(ARef.from0(0, i), N(a(i)))
        .put(ARef.from0(1, i), N("2"))
        .put(ARef.from0(2, i), T(c(i)))
        .put(ARef.from0(3, i), CellValue.Number(BigDecimal(d(i))))
    }
    val withRow = (0 until 5).foldLeft(sheet1) { (s, j) =>
      s.put(ARef.from0(j, 11), CellValue.Number(BigDecimal(j + 1)))
    }
    val other = (0 until 10).foldLeft(Sheet(SheetName.unsafe("Other"))) { (s, i) =>
      s.put(ARef.from0(0, i), CellValue.Number(BigDecimal((i + 1) * 10)))
    }
    Workbook(Vector(withRow, other))
      .withDefinedName("nmRef", "Sheet1!$A$1:$A$10")
      .withDefinedName("nmExpr", "Sheet1!$A$1:$A$10*2")
      .withDefinedName("nmAbs", "ABS(Sheet1!$A$1:$A$10)")

  /** `formula` as a plain formula cell at `sheet!at`, evaluated at its own position. */
  private def plain(sheet: String, at: String, formula: String): Either[String, CellValue] =
    val ref = ARef.parse(at).fold(msg => fail(msg), identity)
    book.sheets.find(_.name.value == sheet) match
      case None => Left(s"no sheet $sheet")
      case Some(target) =>
        val placed = target.put(ref, CellValue.Formula(formula, None))
        placed.evaluateCell(ref, Clock.system, Some(book.put(placed))).left.map(_.message)

  private val libreOffice: List[(String, String, String, CellValue)] = List(
    ("Sheet1", "H5", "A1:A10", N("10")),
    ("Sheet1", "H20", "A1:A10", E("#VALUE!")),
    ("Sheet1", "I5", "A1:A10*2", N("20")),
    ("Sheet1", "I20", "A1:A10*2", E("#VALUE!")),
    ("Sheet1", "J5", "-A1:A10", N("-10")),
    ("Sheet1", "J20", "-A1:A10", E("#VALUE!")),
    ("Sheet1", "K5", "A1:A10&\"x\"", T("10x")),
    ("Sheet1", "K20", "A1:A10&\"x\"", E("#VALUE!")),
    ("Sheet1", "L5", "A1:A10>2", B(true)),
    ("Sheet1", "L20", "A1:A10>2", E("#VALUE!")),
    ("Sheet1", "M5", "IF(A1:A10>2,\"big\",\"small\")", T("big")),
    ("Sheet1", "M20", "IF(A1:A10>2,\"big\",\"small\")", E("#VALUE!")),
    ("Sheet1", "N5", "CHOOSE(D1:D10,\"x\",\"y\",\"z\")", T("y")),
    ("Sheet1", "N20", "CHOOSE(D1:D10,\"x\",\"y\",\"z\")", E("#VALUE!")),
    ("Sheet1", "O5", "ABS(A1:A10)", N("10")),
    ("Sheet1", "O20", "ABS(A1:A10)", E("#VALUE!")),
    ("Sheet1", "P5", "ROUND(A1:A10/7,1)", N("1.4")),
    ("Sheet1", "P20", "ROUND(A1:A10/7,1)", E("#VALUE!")),
    ("Sheet1", "Q5", "ABS(A1:A10+B1:B10)", N("12")),
    ("Sheet1", "Q20", "ABS(A1:A10+B1:B10)", E("#VALUE!")),
    ("Sheet1", "R5", "SUM(A1:A10*B1:B10)", N("20")),
    ("Sheet1", "R20", "SUM(A1:A10*B1:B10)", E("#VALUE!")),
    ("Sheet1", "S5", "SUM(ABS(A1:A10))", N("10")),
    ("Sheet1", "S20", "SUM(ABS(A1:A10))", E("#VALUE!")),
    ("Sheet1", "T5", "MAX(ABS(A1:A10))", N("10")),
    ("Sheet1", "T20", "MAX(ABS(A1:A10))", E("#VALUE!")),
    ("Sheet1", "U5", "AVERAGE(A1:A10*2)", N("20")),
    ("Sheet1", "U20", "AVERAGE(A1:A10*2)", E("#VALUE!")),
    ("Sheet1", "V5", "SUM(A1:A10)", N("122.25")),
    ("Sheet1", "V20", "SUM(A1:A10)", N("122.25")),
    ("Sheet1", "X5", "COUNTIF(A1:A10,\">\"&A1:A10)", N("1")),
    ("Sheet1", "X20", "COUNTIF(A1:A10,\">\"&A1:A10)", E("#VALUE!")),
    ("Sheet1", "Y5", "SUMIF(A1:A10,\">\"&B1:B10)", N("125.5")),
    ("Sheet1", "Y20", "SUMIF(A1:A10,\">\"&B1:B10)", E("#VALUE!")),
    ("Sheet1", "Z5", "SUMPRODUCT(A1:A10*B1:B10)", N("244.5")),
    ("Sheet1", "Z20", "SUMPRODUCT(A1:A10*B1:B10)", N("244.5")),
    ("Sheet1", "AA5", "SUMPRODUCT(ABS(A1:A10))", N("134.75")),
    ("Sheet1", "AA20", "SUMPRODUCT(ABS(A1:A10))", N("134.75")),
    ("Sheet1", "AB5", "SUMPRODUCT(IF(A1:A10>2,1,0))", N("5")),
    ("Sheet1", "AB20", "SUMPRODUCT(IF(A1:A10>2,1,0))", N("5")),
    ("Sheet1", "AC5", "SUMPRODUCT(--(A1:A10>2))", N("5")),
    ("Sheet1", "AC20", "SUMPRODUCT(--(A1:A10>2))", N("5")),
    ("Sheet1", "AG5", "Other!A1:A10*2", N("100")),
    ("Sheet1", "AG20", "Other!A1:A10*2", E("#VALUE!")),
    ("Sheet1", "AH5", "A:A*2", N("20")),
    ("Sheet1", "AH20", "A:A*2", N("0")),
    ("Sheet1", "AI5", "INDEX(A1:A10,0)*1", N("10")),
    ("Sheet1", "AI20", "INDEX(A1:A10,0)*1", E("#VALUE!")),
    ("Sheet1", "AJ5", "OFFSET(A1,0,0,10,1)*1", N("10")),
    ("Sheet1", "AJ20", "OFFSET(A1,0,0,10,1)*1", E("#VALUE!")),
    ("Sheet1", "AK5", "INDIRECT(\"A1:A10\")*1", N("10")),
    ("Sheet1", "AK20", "INDIRECT(\"A1:A10\")*1", E("#VALUE!")),
    ("Sheet1", "AL5", "SUM(INDEX(A1:A10,0))", N("122.25")),
    ("Sheet1", "AL20", "SUM(INDEX(A1:A10,0))", N("122.25")),
    ("Sheet1", "AM5", "SUM(OFFSET(A1,0,0,10,1))", N("122.25")),
    ("Sheet1", "AM20", "SUM(OFFSET(A1,0,0,10,1))", N("122.25")),
    ("Sheet1", "AP5", "nmRef", N("10")),
    ("Sheet1", "AP20", "nmRef", E("#VALUE!")),
    ("Sheet1", "AQ5", "nmRef*2", N("20")),
    ("Sheet1", "AQ20", "nmRef*2", E("#VALUE!")),
    ("Sheet1", "AR5", "SUM(ABS(nmRef))", N("10")),
    ("Sheet1", "AR20", "SUM(ABS(nmRef))", E("#VALUE!")),
    ("Sheet1", "AS5", "ISNUMBER(A1:A10)", B(true)),
    ("Sheet1", "AS20", "ISNUMBER(A1:A10)", B(false)),
    ("Sheet1", "AT5", "N(A1:A10)", N("10")),
    ("Sheet1", "AT20", "N(A1:A10)", E("#VALUE!")),
    ("Sheet1", "AU5", "LEN(C1:C10)", N("5")),
    ("Sheet1", "AU20", "LEN(C1:C10)", E("#VALUE!")),
    ("Sheet1", "AV5", "TEXT(A1:A10,\"0\")", T("10")),
    ("Sheet1", "AV20", "TEXT(A1:A10,\"0\")", E("#VALUE!")),
    ("Sheet1", "AW5", "AND(A1:A10>0)", B(true)),
    ("Sheet1", "AW20", "AND(A1:A10>0)", E("#VALUE!")),
    ("Sheet1", "AX5", "AND(A1:A10)", B(false)),
    ("Sheet1", "AX20", "AND(A1:A10)", B(false)),
    ("Sheet1", "AZ5", "MIN(IF(A1:A10>2,A1:A10,\"\"))", N("-4.25")),
    ("Sheet1", "AZ20", "MIN(IF(A1:A10>2,A1:A10,\"\"))", E("#VALUE!")),
    ("Sheet1", "BA5", "SUMPRODUCT(IF(A1:A10>2,A1:A10,0))", N("125.5")),
    ("Sheet1", "BA20", "SUMPRODUCT(IF(A1:A10>2,A1:A10,0))", N("125.5")),
    ("Sheet1", "BB5", "IF(A1:A10>2,A1:A10*10,0)", N("100")),
    ("Sheet1", "BB20", "IF(A1:A10>2,A1:A10*10,0)", E("#VALUE!")),
    ("Sheet1", "BC5", "SUM(IF(A1:A10>2,1,0))", N("1")),
    ("Sheet1", "BC20", "SUM(IF(A1:A10>2,1,0))", E("#VALUE!")),
    ("Sheet1", "BD5", "SUM(LEN(C1:C10))", N("5")),
    ("Sheet1", "BD20", "SUM(LEN(C1:C10))", E("#VALUE!")),
    ("Sheet1", "BE5", "SUMPRODUCT(LEN(C1:C10))", N("52")),
    ("Sheet1", "BE20", "SUMPRODUCT(LEN(C1:C10))", N("52")),
    ("Sheet1", "BF5", "ROW(A1:A10)", N("1")),
    ("Sheet1", "BF20", "ROW(A1:A10)", N("1")),
    ("Sheet1", "BL5", "SUM(A1:A10*2)", N("20")),
    ("Sheet1", "BL20", "SUM(A1:A10*2)", E("#VALUE!")),
    ("Sheet1", "BM5", "COUNT(1/A1:A10)", N("1")),
    ("Sheet1", "BM20", "COUNT(1/A1:A10)", N("0")),
    ("Sheet1", "BN5", "IFERROR(1/A1:A10,\"e\")", N("0.1")),
    ("Sheet1", "BN20", "IFERROR(1/A1:A10,\"e\")", T("e")),
    ("Sheet1", "BO5", "VLOOKUP(D1:D10,D1:D3,1,FALSE)", N("2")),
    ("Sheet1", "BO20", "VLOOKUP(D1:D10,D1:D3,1,FALSE)", E("#VALUE!")),
    ("Sheet1", "BP5", "MATCH(A1:A10,A1:A10,0)", N("5")),
    ("Sheet1", "BP20", "MATCH(A1:A10,A1:A10,0)", E("#VALUE!")),
    ("Sheet1", "BQ5", "SUMPRODUCT(MAX(ABS(A1:A10)))", N("100")),
    ("Sheet1", "BQ20", "SUMPRODUCT(MAX(ABS(A1:A10)))", N("100")),
    ("Sheet1", "BS5", "SUM(IFERROR(A1:A10*1,0))", N("10")),
    ("Sheet1", "BS20", "SUM(IFERROR(A1:A10*1,0))", N("0")),
    ("Sheet1", "BT5", "ISERROR(A1:A10)", B(false)),
    ("Sheet1", "BT20", "ISERROR(A1:A10)", B(true)),
    ("Sheet1", "BU5", "A1:A10+0=A1:A10", B(true)),
    ("Sheet1", "BU20", "A1:A10+0=A1:A10", E("#VALUE!")),
    ("Sheet1", "C25", "A12:E12*2", N("6")),
    ("Sheet1", "H25", "A12:E12*2", E("#VALUE!")),
    ("Sheet1", "C26", "A12:E12", N("3")),
    ("Sheet1", "H26", "A12:E12", E("#VALUE!")),
    ("Sheet1", "C27", "SUM(ABS(A12:E12))", N("3")),
    ("Sheet1", "H27", "SUM(ABS(A12:E12))", E("#VALUE!")),
    ("Other", "B6", "Sheet1!A1:A10*2", N("14")),
    ("Other", "G6", "Sheet1!A1:A10*2", N("14")),
    ("Sheet1", "H5", "SUM(IF(TRUE,A1:A10,0))", N("122.25")),
    ("Sheet1", "H20", "SUM(IF(TRUE,A1:A10,0))", N("122.25")),
    ("Sheet1", "I5", "IF(TRUE,A1:A10,0)", N("10")),
    ("Sheet1", "I20", "IF(TRUE,A1:A10,0)", E("#VALUE!")),
    ("Sheet1", "J5", "SUM(IF(A1:A10>2,A1:A10,0))", N("122.25")),
    ("Sheet1", "J20", "SUM(IF(A1:A10>2,A1:A10,0))", E("#VALUE!")),
    ("Sheet1", "K5", "SUM(IF(A1:A10>2,A1:A10*2,0))", N("20")),
    ("Sheet1", "K20", "SUM(IF(A1:A10>2,A1:A10*2,0))", E("#VALUE!")),
    ("Sheet1", "L5", "SUM(IF(A1:A10>50,A1:A10,0))", N("0")),
    ("Sheet1", "L20", "SUM(IF(A1:A10>50,A1:A10,0))", E("#VALUE!")),
    ("Sheet1", "M5", "SUM(CHOOSE(1,A1:A10,B1:B10))", N("122.25")),
    ("Sheet1", "M20", "SUM(CHOOSE(1,A1:A10,B1:B10))", N("122.25")),
    ("Sheet1", "O5", "SUMPRODUCT(IFERROR(1/A1:A10,0))", N("1.50327731092437")),
    ("Sheet1", "O20", "SUMPRODUCT(IFERROR(1/A1:A10,0))", N("1.50327731092437")),
    ("Sheet1", "P5", "SUMPRODUCT(IF(TRUE,ABS(A1:A10),0))", N("134.75")),
    ("Sheet1", "P20", "SUMPRODUCT(IF(TRUE,ABS(A1:A10),0))", N("134.75")),
    ("Sheet1", "Q5", "SUMPRODUCT(nmExpr)", N("244.5")),
    ("Sheet1", "Q20", "SUMPRODUCT(nmExpr)", N("244.5")),
    ("Sheet1", "T5", "SUMPRODUCT(nmAbs)", N("134.75")),
    ("Sheet1", "T20", "SUMPRODUCT(nmAbs)", N("134.75")),
    ("Sheet1", "U5", "SUM(nmRef)", N("122.25")),
    ("Sheet1", "U20", "SUM(nmRef)", N("122.25")),
    ("Sheet1", "V5", "SUM(nmRef*B1:B10)", N("20")),
    ("Sheet1", "V20", "SUM(nmRef*B1:B10)", E("#VALUE!")),
    ("Sheet1", "W5", "COUNTIF(A1:A10,A1:A10)", N("1")),
    ("Sheet1", "W20", "COUNTIF(A1:A10,A1:A10)", E("#VALUE!")),
    ("Sheet1", "X5", "SUM(COUNTIF(A1:A10,\">\"&A1:A10))", N("1")),
    ("Sheet1", "X20", "SUM(COUNTIF(A1:A10,\">\"&A1:A10))", E("#VALUE!")),
    ("Sheet1", "Y5", "SUMPRODUCT(COUNTIF(A1:A10,\">\"&A1:A10))", N("45")),
    ("Sheet1", "Y20", "SUMPRODUCT(COUNTIF(A1:A10,\">\"&A1:A10))", N("45")),
    ("Sheet1", "Z5", "MAX(A1:A10*1)", N("10")),
    ("Sheet1", "Z20", "MAX(A1:A10*1)", E("#VALUE!")),
    ("Sheet1", "AA5", "SUM(A1:A10,B1:B10*1)", N("124.25")),
    ("Sheet1", "AA20", "SUM(A1:A10,B1:B10*1)", E("#VALUE!")),
    ("Sheet1", "AB5", "SUM((A1:A10))", N("122.25")),
    ("Sheet1", "AB20", "SUM((A1:A10))", N("122.25")),
    ("Sheet1", "AD5", "SUM(-A1:A10)", N("-10")),
    ("Sheet1", "AD20", "SUM(-A1:A10)", E("#VALUE!")),
    ("Sheet1", "AF5", "AVERAGE(IF(A1:A10>2,A1:A10,\"\"))", N("12.225")),
    ("Sheet1", "AF20", "AVERAGE(IF(A1:A10>2,A1:A10,\"\"))", E("#VALUE!")),
    ("Sheet1", "AG5", "CONCATENATE(C1:C10,\"!\")", T("elder!")),
    ("Sheet1", "AG20", "CONCATENATE(C1:C10,\"!\")", E("#VALUE!")),
    ("Sheet1", "AH5", "IFERROR(A1:A10,0)", N("10")),
    ("Sheet1", "AH20", "IFERROR(A1:A10,0)", N("0")),
    ("Sheet1", "AI5", "SUMPRODUCT(ISNUMBER(A1:A10)*1)", N("10")),
    ("Sheet1", "AI20", "SUMPRODUCT(ISNUMBER(A1:A10)*1)", N("10")),
    ("Sheet1", "AJ5", "SUM(N(A1:A10))", N("10")),
    ("Sheet1", "AJ20", "SUM(N(A1:A10))", E("#VALUE!")),
    ("Sheet1", "AK5", "SUMPRODUCT(N(A1:A10))", N("122.25")),
    ("Sheet1", "AK20", "SUMPRODUCT(N(A1:A10))", N("122.25")),
    ("Sheet1", "AL5", "OR(A1:A10>50)", B(false)),
    ("Sheet1", "AL20", "OR(A1:A10>50)", E("#VALUE!")),
    ("Sheet1", "AM5", "NOT(A1:A10>2)", B(false)),
    ("Sheet1", "AM20", "NOT(A1:A10>2)", E("#VALUE!"))
  )

  /**
   * Row 4, where A4 = 0 makes `A1:A10>2` FALSE, from a second LibreOffice run: an IF that selects
   * `""` hands MIN, AVERAGE and SUM a text value, which is #VALUE! (COUNTA counts it); a logical
   * value typed into an aggregate is 1 or 0 and COUNT counts it.
   */
  private val libreOfficeRow4: List[(String, String, String, CellValue)] = List(
    ("Sheet1", "H4", "MIN(IF(A1:A10>2,A1:A10,\"\"))", E("#VALUE!")),
    ("Sheet1", "I4", "AVERAGE(IF(A1:A10>2,A1:A10,\"\"))", E("#VALUE!")),
    ("Sheet1", "J4", "COUNTA(IF(A1:A10>2,A1:A10,\"\"))", N("1")),
    ("Sheet1", "K4", "SUM(A1:A10>2)", N("0")),
    ("Sheet1", "K5", "SUM(A1:A10>2)", N("1")),
    ("Sheet1", "K20", "SUM(A1:A10>2)", E("#VALUE!")),
    ("Sheet1", "L4", "COUNT(A1:A10>2)", N("1")),
    ("Sheet1", "L20", "COUNT(A1:A10>2)", N("0")),
    ("Sheet1", "M4", "COUNT(A1:A10*1)", N("1")),
    ("Sheet1", "M4", "COUNTA(A1:A10*1)", N("1")),
    ("Sheet1", "M20", "COUNTA(A1:A10*1)", N("1")),
    ("Sheet1", "N5", "MEDIAN(A1:A10*1)", N("10")),
    ("Sheet1", "O5", "AVERAGE(A1:A10,B1:B10*1)", N("11.2954545454545454545")),
    ("Sheet1", "P5", "MAX(A1:A10,A1:A10*10)", N("100")),
    ("Sheet1", "Q4", "AND(A1:A10>-5,TRUE)", B(true)),
    ("Sheet1", "Q20", "AND(A1:A10>-5,TRUE)", E("#VALUE!")),
    ("Sheet1", "R5", "OR(A1:A10=10)", B(true)),
    ("Sheet1", "S4", "NOT(A1:A10)", B(true)),
    ("Sheet1", "S5", "NOT(A1:A10)", B(false)),
    ("Sheet1", "T4", "SUM(CHOOSE(2,A1:A10,B1:B10))", N("20")),
    ("Sheet1", "U4", "SUM(CHOOSE(D1:D10,A1:A10,B1:B10,A1:A10))", N("122.25")),
    ("Sheet1", "U5", "SUM(CHOOSE(D1:D10,A1:A10,B1:B10,A1:A10))", N("20")),
    ("Sheet1", "V4", "SUM(IF(A1:A10>2,A1:A10,B1:B10))", N("20")),
    ("Sheet1", "V5", "SUM(IF(A1:A10>2,A1:A10,B1:B10))", N("122.25")),
    ("Sheet1", "W4", "IF(TRUE,A1:A10,0)", N("0")),
    ("Sheet1", "X5", "CHOOSE(2,A1:A10,B1:B10)", N("2")),
    ("Sheet1", "Y5", "INDEX(A1:B10,0,2)*1", N("2")),
    ("Sheet1", "Z5", "SUM(INDEX(A1:B10,0,2))", N("20")),
    ("Sheet1", "AA5", "SUM(INDIRECT(\"A1:A10\")*1)", N("10")),
    ("Sheet1", "AB5", "SUM(INDEX(A1:A10,0)*B1:B10)", N("20")),
    ("Sheet1", "AC4", "COUNTIF(A1:A10,\"<\"&A1:A10)", N("2")),
    ("Sheet1", "AC5", "COUNTIF(A1:A10,\"<\"&A1:A10)", N("8")),
    ("Sheet1", "AD5", "SUMIF(A1:A10,\">\"&MAX(A1:A10*0.5))", N("117")),
    ("Sheet1", "AE20", "SUMIF(A1:A10,\">\"&AVERAGE(A1:A10))", N("100")),
    ("Sheet1", "AF5", "SUM(A:A*B:B)", N("20")),
    ("Sheet1", "AF20", "SUM(A:A*B:B)", N("0")),
    ("Sheet1", "AG20", "ISBLANK(A:A)", B(true))
  )

  /** LibreOffice prints 15 significant digits: numbers agree to 1e-12. */
  private def agree(obtained: Either[String, CellValue], expected: CellValue): Unit =
    (obtained, expected) match
      case (Right(CellValue.Number(got)), CellValue.Number(want)) =>
        assert((got - want).abs < BigDecimal("1e-12"), s"$got is not $want")
      case _ => assertEquals(obtained, Right(expected))

  (libreOffice ++ libreOfficeRow4).foreach { (sheet, at, formula, expected) =>
    test(s"=$formula at $sheet!$at is $expected") {
      agree(plain(sheet, at, formula), expected)
    }
  }
