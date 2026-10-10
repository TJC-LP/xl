package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * #709: a text value in a typed scalar slot — a numeric, integer or date argument, the IF/IFS
 * condition, NOT's operand, a computed AND/OR value — is Excel's `#VALUE!`, an error value the cell
 * caches and IFERROR/ISERROR/ERROR.TYPE see, never a host failure that leaves the cell uncached.
 * Numeric text keeps Excel's rule: `"5"` is 5 in a numeric, integer or date slot, and `#VALUE!` in
 * a logical one.
 *
 * The table is LibreOffice 24.2's recalculation of the same plain `<f>` cells (the cell each
 * formula sits in is part of the case). The rows where Excel and LibreOffice differ, and the shapes
 * LibreOffice cannot compute, are pinned after it, each with its reason.
 */
class TextScalarSlotSpec extends FunSuite:

  private def N(n: String): CellValue = CellValue.Number(BigDecimal(n))
  private def T(s: String): CellValue = CellValue.Text(s)
  private def B(b: Boolean): CellValue = CellValue.Bool(b)
  private val Value: CellValue = CellValue.Error(CellError.Value)

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
  private val g: Vector[CellValue] =
    Vector(N("3"), N("-4"), N("5"), T("n/a"), N("-6"), N("7"), N("8"), N("-9"), N("10"), N("11"))

  /**
   * Sheet1: A1:A10 numbers, B1:B10 = 2, C1:C10 text, G1:G10 numbers with the placeholder "n/a" in
   * G4, F1 "5", F2 "2024-01-15", F3 "TRUE", F4 " 5 " (all text).
   */
  private val book: Workbook =
    val sheet1 = (0 until 10).foldLeft(Sheet(SheetName.unsafe("Sheet1"))) { (s, i) =>
      s.put(ARef.from0(0, i), N(a(i)))
        .put(ARef.from0(1, i), N("2"))
        .put(ARef.from0(2, i), T(c(i)))
        .put(ARef.from0(6, i), g(i))
    }
    val withCells = sheet1
      .put(ref"F1", T("5"))
      .put(ref"F2", T("2024-01-15"))
      .put(ref"F3", T("TRUE"))
      .put(ref"F4", T(" 5 "))
    Workbook(Vector(withCells))

  /** `formula` as a plain formula cell at `Sheet1!at`, evaluated at its own position. */
  private def plain(at: String, formula: String): Either[String, CellValue] =
    val ref = ARef.parse(at).fold(msg => fail(msg), identity)
    book.sheets.headOption match
      case None => Left("no sheet")
      case Some(target) =>
        val placed = target.put(ref, CellValue.Formula(formula, None))
        placed.evaluateCell(ref, Clock.system, Some(book.put(placed))).left.map(_.message)

  private def agree(at: String, formula: String, expected: CellValue): Unit =
    assertEquals(plain(at, formula), Right(expected), s"$at: =$formula")

  private val libreOffice: List[(String, String, CellValue)] = List(
    // the issue's table
    ("J1", "ABS(C5)", Value),
    ("J2", "IF(C6,1,0)", Value),
    ("J3", "NOT(C7)", Value),
    ("J4", "ROUND(C8,0)", Value),
    ("J5", "IFERROR(ABS(C9),\"caught\")", T("caught")),
    // integer, date and numeric slots over a text cell
    ("J6", "DATE(2024,C5,1)", Value),
    ("J7", "YEAR(C5)", Value),
    ("J8", "MONTH(C5)", Value),
    ("J9", "DAY(C1)", Value),
    ("J10", "SQRT(C5)", Value),
    ("J11", "LEFT(\"abc\",C5)", Value),
    ("J12", "MID(\"abcdef\",C1,2)", Value),
    ("J13", "ROUND(1.5,C1)", Value),
    ("J14", "POWER(C1,2)", Value),
    ("J15", "MOD(C1,2)", Value),
    ("J16", "EDATE(DATE(2024,1,1),C5)", Value),
    ("J17", "EDATE(C1,1)", Value),
    ("J18", "EOMONTH(C1,0)", Value),
    ("J19", "CHOOSE(C5,1,2)", Value),
    // the operators over a single text cell
    ("J20", "C5+0", Value),
    ("J21", "C1*1", Value),
    ("J22", "-C1", Value),
    ("J23", "C1^2", Value),
    ("J24", "C1%", Value),
    // the error value is visible to the error functions
    ("J25", "ISERROR(ABS(C5))", B(true)),
    ("J26", "ERROR.TYPE(ABS(C5))", N("3")),
    ("J27", "ERROR.TYPE(IF(C6,1,0))", N("3")),
    ("J28", "SUM(ABS(C5))", Value),
    // literals, computed text and references INDIRECT or a sheet name resolve
    ("J29", "ABS(\"abc\")", Value),
    ("J30", "IF(\"abc\",1,0)", Value),
    ("J31", "NOT(\"abc\")", Value),
    ("J32", "ROUND(\"abc\",0)", Value),
    ("J33", "YEAR(\"abc\")", Value),
    ("J34", "ABS(\"\")", Value),
    ("J35", "IF(\"\",1,0)", Value),
    ("J36", "YEAR(\"\")", Value),
    ("J37", "DATE(2024,\"\",1)", Value),
    ("J38", "LEFT(\"abc\",\"x\")", Value),
    ("J39", "ABS(C5&\"\")", Value),
    ("J40", "IF(C5&\"\",1,0)", Value),
    ("J41", "ABS(INDIRECT(\"C5\"))", Value),
    ("J42", "IF(INDIRECT(\"C5\"),1,0)", Value),
    ("J43", "NOT(INDIRECT(\"C5\"))", Value),
    ("J44", "ABS(Sheet1!C5)", Value),
    ("J45", "IF(Sheet1!C6,1,0)", Value),
    // computed AND/OR values: a text value is #VALUE! (a reference's text is ignored)
    ("J46", "AND(C1&\"\")", Value),
    ("J47", "OR(C1&\"\")", Value),
    ("J48", "AND(TRUE,\"abc\")", Value),
    ("J49", "OR(FALSE,\"abc\")", Value),
    ("J50", "AND({\"a\",TRUE})", Value),
    ("J51", "OR({\"a\",FALSE})", Value),
    ("J52", "AND(IF({1,0},\"a\",TRUE))", Value),
    ("J53", "SUMPRODUCT(--AND(C1:C2&\"\"))", Value),
    ("J54", "AND(C1)", Value),
    ("J55", "AND(C1,TRUE)", B(true)),
    // numeric text is its number in a numeric, integer or date slot
    ("J56", "ABS(F1)", N("5")),
    ("J57", "F1+0", N("5")),
    ("J58", "ABS(F4)", N("5")),
    ("J59", "MONTH(DATE(2024,F1,1))", N("5")),
    ("J60", "YEAR(F1)", N("1900")),
    ("J61", "YEAR(\"45000\")", N("2023")),
    ("J62", "ABS(\"5\")", N("5")),
    ("J63", "IF(F3,1,0)", N("1")),
    // a plain cell intersects a column: the text row is #VALUE!, the others their value
    ("H3", "ABS(G1:G10)", N("5")),
    ("H4", "ABS(G1:G10)", Value),
    ("H5", "ABS(G1:G10)", N("6")),
    ("I1", "IF(C1:C10,1,0)", Value),
    ("I3", "MONTH(DATE(2024,G1:G10,1))", N("5")),
    ("I4", "DATE(2024,G1:G10,1)", Value),
    ("I5", "ROUND(G1:G10,0)", N("-6")),
    ("I6", "NOT(G1:G10)", B(false))
  )

  test("#709: a text value in a typed scalar slot is #VALUE!, as LibreOffice computes it") {
    val diffs = libreOffice.flatMap { case (at, formula, expected) =>
      val got = plain(at, formula)
      Option.when(got != Right(expected))(s"$at =$formula: expected $expected, got $got")
    }
    assert(diffs.isEmpty, diffs.mkString("\n"))
  }

  test("#709: Excel's logical slots refuse numeric and boolean-looking text where LO converts") {
    // Excel converts only the text TRUE/FALSE in a logical slot and never converts text to a number
    // there; LibreOffice reads "5" as 5 (IF and NOT give 1 and FALSE). Likewise "TRUE" is not a
    // number to Excel (`="TRUE"+0` is #VALUE!) where LibreOffice reads it as 1.
    agree("K1", "IF(F1,1,0)", Value)
    agree("K2", "NOT(\"5\")", Value)
    agree("K3", "IF(\"5\",1,0)", Value)
    agree("K4", "ABS(F3)", Value)
  }

  test("#709: IFS and LET carry the same #VALUE! (LibreOffice 24.2 lacks both)") {
    agree("K5", "IFS(C5,1,TRUE,2)", Value)
    agree("K6", "LET(x,C5,ABS(x))", Value)
    agree("K7", "LET(x,C5,IF(x,1,0))", Value)
    agree("K8", "IFERROR(LET(x,C5,ABS(x)),\"caught\")", T("caught"))
  }

  test("#709: host failures in a typed slot stay loud") {
    assert(plain("K9", "ABS(Missing!A1)").isLeft, "a missing sheet is not #VALUE!")
    assert(plain("K10", "IF(Missing!A1,1,0)").isLeft, "a missing sheet is not #VALUE!")
  }
