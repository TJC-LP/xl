package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-476: the parity gaps that fail LOUD — host errors that abort whole-book recalc on ordinary
 * banker files, as distinct from the silent-wrong-value class (#467/#488).
 *
 *   - SEARCH absent (FIND exists but is case-sensitive; SEARCH is the banker idiom)
 *   - N() and HYPERLINK() absent → UnknownFunction
 *   - `=IF(1=1,+S2!G1,0)`: unary plus on a cross-sheet ref inside a function argument escaped
 *     PolyRef resolution and hit the "Unresolved SheetPolyRef" programming-error arm
 *   - VALUE("None") raised a host error where Excel caches #VALUE!
 */
class LoudParityGapsSpec extends FunSuite:

  private def sheetWith(name: String, cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(Sheet(SheetName.unsafe(name))) { case (s, (ref, v)) => s.put(ref, v) }

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  // A1 = "Widget Corporation": o at 9 and 12, so start_num is observable
  private val text = Sheet("T")
    .put(ARef.from0(0, 0), CellValue.Text("Widget Corporation"))
    .put(ARef.from0(0, 1), CellValue.Number(BigDecimal("42.5")))
    .put(ARef.from0(0, 2), CellValue.Bool(true))

  // ========== SEARCH ==========

  test("GH-476: SEARCH is case-insensitive (where FIND is not)") {
    assertEquals(text.evaluateFormula("""=SEARCH("widget", A1)"""), Right(num(1)))
    assertEquals(text.evaluateFormula("""=SEARCH("CORP", A1)"""), Right(num(8)))
  }

  test("GH-476: FIND stays case-sensitive") {
    text.evaluateFormula("""=FIND("widget", A1)""") match
      case Right(CellValue.Error(CellError.Value)) => ()
      case Left(_) => ()
      case other => fail(s"FIND must not match 'widget' in 'Widget Corp', got $other")
  }

  test("GH-476: SEARCH honours start_num") {
    assertEquals(text.evaluateFormula("""=SEARCH("o", A1)"""), Right(num(9)))
    assertEquals(text.evaluateFormula("""=SEARCH("o", A1, 10)"""), Right(num(12)))
  }

  test("GH-476: SEARCH start_num past the end is a cached #VALUE!") {
    assertEquals(
      text.evaluateFormula("""=SEARCH("o", A1, 99)"""),
      Right(CellValue.Error(CellError.Value))
    )
  }

  test("GH-476: SEARCH supports Excel wildcards") {
    assertEquals(text.evaluateFormula("""=SEARCH("W?dget", A1)"""), Right(num(1)))
    assertEquals(text.evaluateFormula("""=SEARCH("Wid*rp", A1)"""), Right(num(1)))
  }

  test("GH-476: SEARCH miss is a cached #VALUE!, not a host error") {
    assertEquals(
      text.evaluateFormula("""=SEARCH("zzz", A1)"""),
      Right(CellValue.Error(CellError.Value))
    )
  }

  test("GH-476: SEARCH composes with the IFERROR/MID idiom") {
    assertEquals(
      text.evaluateFormula("""=IFERROR(SEARCH("zzz", A1), 0)"""),
      Right(num(0))
    )
  }

  // ========== N ==========

  test("GH-476: N() coerces per Excel's table") {
    assertEquals(text.evaluateFormula("=N(5)"), Right(num(5)))
    assertEquals(text.evaluateFormula("=N(TRUE)"), Right(num(1)))
    assertEquals(text.evaluateFormula("=N(FALSE)"), Right(num(0)))
    assertEquals(text.evaluateFormula("""=N("hello")"""), Right(num(0)))
    assertEquals(text.evaluateFormula("=N(A1)"), Right(num(0)))
    assertEquals(text.evaluateFormula("=N(A2)"), Right(CellValue.Number(BigDecimal("42.5"))))
    assertEquals(text.evaluateFormula("=N(A3)"), Right(num(1)))
    assertEquals(text.evaluateFormula("=N(Z9)"), Right(num(0)))
  }

  test("GH-476: N() of a date is its Excel serial") {
    assertEquals(text.evaluateFormula("=N(DATE(2026,1,1))"), Right(num(46023)))
  }

  test("GH-476: N() propagates direct and referenced Excel errors") {
    val errors = Sheet("Errors")
      .put(ref"A1", CellValue.Error(CellError.Div0))
      .put(ref"A2", CellValue.Error(CellError.NA))

    assertEquals(errors.evaluateFormula("=N(1/0)"), Right(CellValue.Error(CellError.Div0)))
    assertEquals(errors.evaluateFormula("=N(A1)"), Right(CellValue.Error(CellError.Div0)))
    assertEquals(errors.evaluateFormula("=N(A2)"), Right(CellValue.Error(CellError.NA)))
  }

  test("GH-476: N() propagates an error cached by a referenced formula") {
    val errors = Sheet("Errors")
      .put(
        ref"A1",
        CellValue.Formula("=1/0", Some(CellValue.Error(CellError.Div0)))
      )

    assertEquals(errors.evaluateFormula("=N(A1)"), Right(CellValue.Error(CellError.Div0)))
  }

  // ========== HYPERLINK ==========

  test("GH-476: HYPERLINK displays the friendly name when given") {
    assertEquals(
      text.evaluateFormula("""=HYPERLINK("https://x.test","Open")"""),
      Right(CellValue.Text("Open"))
    )
  }

  test("GH-476: HYPERLINK with no friendly name displays the link") {
    assertEquals(
      text.evaluateFormula("""=HYPERLINK("https://x.test")"""),
      Right(CellValue.Text("https://x.test"))
    )
  }

  // ========== unary plus on a cross-sheet ref inside a function argument ==========

  private val book: Workbook =
    Workbook(
      Vector(
        sheetWith("Main", ARef.from0(0, 0) -> num(1)),
        sheetWith("S2", ARef.from0(6, 0) -> num(11)) // G1 = 11
      )
    )

  private def main: Sheet = book.sheets.headOption.getOrElse(fail("no Main sheet"))

  test("GH-476: +Sheet!Ref inside a function argument resolves (was Unresolved SheetPolyRef)") {
    assertEquals(
      main.evaluateFormula("=IF(1=1,+S2!G1,0)", workbook = Some(book)),
      Right(num(11))
    )
  }

  test("GH-476: bare +Sheet!Ref and arithmetic over it keep working") {
    assertEquals(main.evaluateFormula("=+S2!G1", workbook = Some(book)), Right(num(11)))
    assertEquals(main.evaluateFormula("=+S2!G1*2", workbook = Some(book)), Right(num(22)))
  }

  test("GH-476: +Ref (same sheet) inside a function argument resolves too") {
    val s = Sheet("Main").put(ARef.from0(0, 0), num(7))
    assertEquals(s.evaluateFormula("=IF(1=1,+A1,0)"), Right(num(7)))
    assertEquals(s.evaluateFormula("=SUM(+A1,1)"), Right(num(8)))
  }

  // ========== VALUE error parity ==========

  test("GH-476: VALUE of unparseable text is a cached #VALUE!, not a host error") {
    assertEquals(
      text.evaluateFormula("""=VALUE("None")"""),
      Right(CellValue.Error(CellError.Value))
    )
  }

  test("GH-476: VALUE's #VALUE! is catchable by IFERROR") {
    assertEquals(text.evaluateFormula("""=IFERROR(VALUE("None"), 0)"""), Right(num(0)))
  }

  test("GH-476: VALUE still parses the Excel-numeric-text forms") {
    assertEquals(text.evaluateFormula("""=VALUE("$1,234")"""), Right(num(1234)))
    assertEquals(
      text.evaluateFormula("""=VALUE("(500)")"""),
      Right(CellValue.Number(BigDecimal(-500)))
    )
  }
