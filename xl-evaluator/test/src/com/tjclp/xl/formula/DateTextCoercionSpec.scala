package com.tjclp.xl.formula

import munit.FunSuite

import java.time.LocalDate

import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.DateText
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import com.tjclp.xl.workbooks.{Workbook, WorkbookMetadata}

/**
 * #953: DATEVALUE, and date text read as its serial wherever a single value coerces text to a
 * number (VALUE, the arithmetic operators, numeric, integer and date arguments), in Excel's US
 * forms. Excel values are desktop Excel's (US locale) as given in the issue; 12/31/2028 is serial
 * 47118.
 */
class DateTextCoercionSpec extends FunSuite:

  private val sheet = Sheet("Test")
    .put(ref"A1", CellValue.Text("12/31/2028"))
    .put(ref"A2", CellValue.Text("31-Dec-2028"))
    .put(ref"A3", CellValue.Text("not a date"))
    .put(ref"A4", CellValue.DateTime(LocalDate.of(2028, 12, 31).atStartOfDay()))
    .put(ref"A5", CellValue.Number(BigDecimal(47118)))

  private val Dec31: CellValue = CellValue.Number(BigDecimal(47118))

  private def eval(formula: String): Either[?, CellValue] = sheet.evaluateFormula(formula)

  private def assertNumber(formula: String, expected: BigDecimal)(implicit
    loc: munit.Location
  ): Unit =
    eval(formula) match
      case Right(CellValue.Number(n)) => assertEquals(n, expected, formula)
      case other => fail(s"$formula: expected $expected, got $other")

  private def assertError(formula: String, code: CellError)(implicit
    loc: munit.Location
  ): Unit =
    assertEquals(eval(formula), Right(CellValue.Error(code)), formula)

  // ===== The issue's table =====

  test("#953: DATEVALUE(\"12/31/2028\") is 47118") {
    assertNumber("=DATEVALUE(\"12/31/2028\")", 47118)
  }

  test("#953: VALUE(\"12/31/2028\") is 47118") {
    assertNumber("=VALUE(\"12/31/2028\")", 47118)
  }

  test("#953: \"12/31/2028\"+0 is 47118") {
    assertNumber("=\"12/31/2028\"+0", 47118)
  }

  test("#953: a text cell holding a date reads as its serial in arithmetic (=A1+0)") {
    assertNumber("=A1+0", 47118)
    assertNumber("=A2+0", 47118)
  }

  test("#953: DATEVALUE reads the ISO and day-month-name forms") {
    assertNumber("=DATEVALUE(\"2028-12-31\")", 47118)
    assertNumber("=DATEVALUE(\"31-Dec-2028\")", 47118)
  }

  // ===== The forms covered =====

  private val forms: List[(String, BigDecimal)] = List(
    "12/31/2028" -> 47118,
    "1/2/2028" -> 46754,
    "01/02/2028" -> 46754,
    "12-31-2028" -> 47118,
    "12/31/28" -> 47118,
    "1/2/29" -> 47120,
    "1/2/30" -> 10960,
    "2028-12-31" -> 47118,
    "2028/12/31" -> 47118,
    "2028-1-2" -> 46754,
    "31-Dec-2028" -> 47118,
    "31-dec-28" -> 47118,
    "31 Dec 2028" -> 47118,
    "31 December 2028" -> 47118,
    "Dec 31, 2028" -> 47118,
    "December 31, 2028" -> 47118,
    "Dec-2028" -> 47088,
    "December 2028" -> 47088,
    "  12/31/2028  " -> 47118,
    "1/1/1900" -> 1,
    "2/28/1900" -> 59,
    "2/29/1900" -> 60,
    "3/1/1900" -> 61,
    "12/31/9999" -> 2958465
  )

  forms.foreach { case (text, serial) =>
    test(s"#953: DATEVALUE and VALUE read \"$text\" as $serial") {
      assertNumber(s"=DATEVALUE(\"$text\")", serial)
      assertNumber(s"=VALUE(\"$text\")", serial)
      assertNumber(s"=\"$text\"+0", serial)
      assertEquals(DateText.serial(text), Some(serial), text)
    }
  }

  // ===== Refusals =====

  private val notDates: List[String] = List(
    "not a date",
    "",
    "13/1/2028",
    "2/30/2028",
    "2/29/2027",
    "0/1/2028",
    "1/0/2028",
    "12/31/1899",
    "1/1/10000",
    "31-Foo-2028",
    "47118"
  )

  notDates.foreach { text =>
    test(s"#953: DATEVALUE(\"$text\") is #VALUE!") {
      assertError(s"=DATEVALUE(\"$text\")", CellError.Value)
      assertEquals(DateText.serial(text), None, text)
    }
  }

  test("#953: DATEVALUE takes text only: a number, a date cell and a logical are #VALUE!") {
    assertError("=DATEVALUE(47118)", CellError.Value)
    assertError("=DATEVALUE(A4)", CellError.Value)
    assertError("=DATEVALUE(A5)", CellError.Value)
    assertError("=DATEVALUE(TRUE)", CellError.Value)
  }

  test("#953: DATEVALUE over a text cell reads it; over non-date text it is #VALUE!") {
    assertEquals(eval("=DATEVALUE(A1)"), Right(Dec31))
    assertEquals(eval("=DATEVALUE(A2)"), Right(Dec31))
    assertError("=DATEVALUE(A3)", CellError.Value)
  }

  test("#953: IFERROR sees DATEVALUE's #VALUE!") {
    assertEquals(eval("=IFERROR(DATEVALUE(\"x\"),-1)"), Right(CellValue.Number(BigDecimal(-1))))
  }

  test("#953: DATEVALUE lifts over an array") {
    assertEquals(
      eval("=SUM(DATEVALUE({\"12/31/2028\",\"1/1/1900\"}))"),
      Right(CellValue.Number(BigDecimal(47119)))
    )
  }

  // ===== Other single-value coercions share the grammar =====

  test("#953: date text is its serial in numeric, integer and date arguments") {
    assertNumber("=ABS(\"12/31/2028\")", 47118)
    assertNumber("=\"12/31/2028\"-\"12/1/2028\"", 30)
    assertNumber("=-\"12/31/2028\"", -47118)
    assertNumber("=YEAR(\"12/31/2028\")", 2028)
    assertNumber("=MONTH(A1)", 12)
    assertNumber("=DAY(\"31-Dec-2028\")", 31)
    assertEquals(eval("=LEFT(\"abcdef\",\"1/2/1900\")"), Right(CellValue.Text("ab")))
  }

  test("#953: TEXT formats date text as the date it spells") {
    assertEquals(eval("=TEXT(\"12/31/2028\",\"yyyy\")"), Right(CellValue.Text("2028")))
    assertEquals(eval("=TEXT(A1,\"mmm yyyy\")"), Right(CellValue.Text("Dec 2028")))
    assertEquals(eval("=TEXT(A2,\"0\")"), Right(CellValue.Text("47118")))
    assertEquals(eval("=TEXT(A3,\"0\")"), Right(CellValue.Text("not a date")))
  }

  test("#953: non-date text stays #VALUE! in arithmetic and VALUE") {
    assertError("=A3+0", CellError.Value)
    assertError("=VALUE(\"31-Foo-2028\")", CellError.Value)
  }

  // Not modeled (docs/LIMITATIONS.md): Excel reads a time, and a date without a year as one in the
  // current year; xl refuses them so a result never depends on the clock or a time grammar.
  test("#953: not modeled: time text and year-less dates are #VALUE! (Excel reads them)") {
    List("12/31/2028 10:30", "10:30", "12/31", "31-Dec", "Dec-28").foreach { text =>
      assertError(s"=DATEVALUE(\"$text\")", CellError.Value)
      assertError(s"=\"$text\"+0", CellError.Value)
    }
  }

  test("#953: a range's date text is still skipped by SUM (range folds are out of scope)") {
    assertEquals(eval("=SUM(A1:A2)"), Right(CellValue.Number(BigDecimal(0))))
  }

  // ===== #956: XIRR and XNPV read date text in the dates range =====

  // -1000 on 1/1/2028 and 1100 on 12/31/2028, 365 days apart: XIRR is 0.1 and XNPV at 10% is 0.
  private def flows(second: CellValue): Sheet =
    Sheet("Flows")
      .put(ref"A1", CellValue.Number(BigDecimal(-1000)))
      .put(ref"A2", CellValue.Number(BigDecimal(1100)))
      .put(ref"B1", CellValue.DateTime(LocalDate.of(2028, 1, 1).atStartOfDay()))
      .put(ref"B2", second)

  private def numberOf(sheet: Sheet, formula: String): BigDecimal =
    sheet.evaluateFormula(formula) match
      case Right(CellValue.Number(n)) => n
      case other => fail(s"$formula: expected a number, got $other")

  private val dateTextCells: List[(String, CellValue)] = List(
    "plain text 12/31/2028" -> CellValue.Text("12/31/2028"),
    "plain text 31-Dec-2028" -> CellValue.Text("31-Dec-2028"),
    "plain text 2028-12-31" -> CellValue.Text("2028-12-31"),
    "a formula returning the text" -> CellValue.Formula("IF(1=2,\"x\",\"12/31/2028\")", None)
  )

  dateTextCells.foreach { case (label, cell) =>
    test(s"#956: XIRR and XNPV read date text in the dates range ($label)") {
      val withText = flows(cell)
      val withDate = flows(CellValue.DateTime(LocalDate.of(2028, 12, 31).atStartOfDay()))
      val rate = numberOf(withText, "=XIRR(A1:A2,B1:B2)")
      assertEqualsDouble(rate.toDouble, 0.1, 1e-9)
      assertEquals(rate, numberOf(withDate, "=XIRR(A1:A2,B1:B2)"))
      val npv = numberOf(withText, "=XNPV(0.1,A1:A2,B1:B2)")
      assertEqualsDouble(npv.toDouble, 0.0, 1e-9)
      assertEquals(npv, numberOf(withDate, "=XNPV(0.1,A1:A2,B1:B2)"))
    }
  }

  // A 1904-system workbook stores its date cells as 1904 serials, which the evaluator reads as 1900
  // ones (docs/LIMITATIONS.md); date text there must decode on that same shifted footing, or a range
  // mixing the two gets a spurious 1,462-day gap. 2028-01-01 is 1904 serial 45291, 2028-12-31 45656.
  test("#956: in a 1904 workbook, date text and numeric dates mix on one epoch") {
    def book(second: CellValue): (Sheet, Workbook) =
      val sheet = Sheet("Flows")
        .put(ref"A1", CellValue.Number(BigDecimal(-1000)))
        .put(ref"A2", CellValue.Number(BigDecimal(1100)))
        .put(ref"B1", CellValue.Number(BigDecimal(45291)))
        .put(ref"B2", second)
      (sheet, Workbook(sheet).copy(metadata = WorkbookMetadata(date1904 = true)))
    def evalIn(second: CellValue, formula: String): BigDecimal =
      val (sheet, wb) = book(second)
      sheet.evaluateFormula(formula, workbook = Some(wb)) match
        case Right(CellValue.Number(n)) => n
        case other => fail(s"$formula: expected a number, got $other")
    val numeric = CellValue.Number(BigDecimal(45656))
    val text = CellValue.Text("12/31/2028")
    assertEqualsDouble(evalIn(text, "=XIRR(A1:A2,B1:B2)").toDouble, 0.1, 1e-9)
    assertEquals(evalIn(text, "=XIRR(A1:A2,B1:B2)"), evalIn(numeric, "=XIRR(A1:A2,B1:B2)"))
    assertEquals(evalIn(text, "=XNPV(0.1,A1:A2,B1:B2)"), evalIn(numeric, "=XNPV(0.1,A1:A2,B1:B2)"))
  }

  test("#956: other text in the dates range is still #VALUE!") {
    List(CellValue.Text("not a date"), CellValue.Text("12/31"), CellValue.Text("47118")).foreach {
      cell =>
        val sheet = flows(cell)
        assertEquals(
          sheet.evaluateFormula("=XIRR(A1:A2,B1:B2)"),
          Right(CellValue.Error(CellError.Value)),
          cell.toString
        )
        assertEquals(
          sheet.evaluateFormula("=XNPV(0.1,A1:A2,B1:B2)"),
          Right(CellValue.Error(CellError.Value)),
          cell.toString
        )
    }
  }
