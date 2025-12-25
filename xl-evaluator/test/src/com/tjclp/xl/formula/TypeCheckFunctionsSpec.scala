package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.addressing.SheetName
import munit.FunSuite

/**
 * Comprehensive tests for type-check functions: ISNUMBER, ISTEXT, ISBLANK, ISERR.
 *
 * ISNUMBER - Returns TRUE if value is numeric
 * ISTEXT   - Returns TRUE if value is a text string
 * ISBLANK  - Returns TRUE if cell is empty (NOT for empty string)
 * ISERR    - Returns TRUE if error (except #N/A)
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class TypeCheckFunctionsSpec extends FunSuite:
  val emptySheet = new Sheet(name = SheetName.unsafe("Test"))

  /** Helper to create sheet with cells */
  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  /** Helper to evaluate a boolean formula */
  def evalBool(formula: String, sheet: Sheet): Either[String, Boolean] =
    FormulaParser.parse(formula) match
      case Right(expr) =>
        Evaluator.eval(expr, sheet) match
          case Right(value: Boolean) => Right(value)
          case Right(other) => Left(s"Expected Boolean, got: $other")
          case Left(err) => Left(s"Eval error: $err")
      case Left(err) => Left(s"Parse error: $err")

  // ===== ISNUMBER Tests =====

  test("ISNUMBER: returns true for numeric cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(42))
    val result = evalBool("=ISNUMBER(A1)", sheet)
    assertEquals(result, Right(true))
  }

  test("ISNUMBER: returns true for decimal number") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal("3.14159")))
    val result = evalBool("=ISNUMBER(A1)", sheet)
    assertEquals(result, Right(true))
  }

  test("ISNUMBER: returns false for text cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("hello"))
    val result = evalBool("=ISNUMBER(A1)", sheet)
    assertEquals(result, Right(false))
  }

  test("ISNUMBER: returns false for empty cell") {
    val sheet = emptySheet
    val result = evalBool("=ISNUMBER(A1)", sheet)
    assertEquals(result, Right(false))
  }

  test("ISNUMBER: returns true for formula with numeric result") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(10),
      ref"A2" -> CellValue.Formula("=A1*2", Some(CellValue.Number(20)))
    )
    val result = evalBool("=ISNUMBER(A2)", sheet)
    assertEquals(result, Right(true))
  }

  test("ISNUMBER: returns false for boolean cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Bool(true))
    val result = evalBool("=ISNUMBER(A1)", sheet)
    assertEquals(result, Right(false))
  }

  test("ISNUMBER: returns false for error cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.Div0))
    val result = evalBool("=ISNUMBER(A1)", sheet)
    assertEquals(result, Right(false))
  }

  // ===== ISTEXT Tests =====

  test("ISTEXT: returns true for text cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("hello world"))
    val result = evalBool("=ISTEXT(A1)", sheet)
    assertEquals(result, Right(true))
  }

  test("ISTEXT: returns true for empty string") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text(""))
    val result = evalBool("=ISTEXT(A1)", sheet)
    assertEquals(result, Right(true))
  }

  test("ISTEXT: returns false for numeric cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(123))
    val result = evalBool("=ISTEXT(A1)", sheet)
    assertEquals(result, Right(false))
  }

  test("ISTEXT: returns false for empty cell") {
    val sheet = emptySheet
    val result = evalBool("=ISTEXT(A1)", sheet)
    assertEquals(result, Right(false))
  }

  test("ISTEXT: returns true for formula with text result") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Text("Hello"),
      ref"A2" -> CellValue.Formula("=A1", Some(CellValue.Text("Hello")))
    )
    val result = evalBool("=ISTEXT(A2)", sheet)
    assertEquals(result, Right(true))
  }

  test("ISTEXT: returns false for boolean cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Bool(false))
    val result = evalBool("=ISTEXT(A1)", sheet)
    assertEquals(result, Right(false))
  }

  // ===== ISBLANK Tests =====

  test("ISBLANK: returns true for empty cell") {
    val sheet = emptySheet
    val result = evalBool("=ISBLANK(A1)", sheet)
    assertEquals(result, Right(true))
  }

  test("ISBLANK: returns false for cell with number") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(0))
    val result = evalBool("=ISBLANK(A1)", sheet)
    assertEquals(result, Right(false))
  }

  test("ISBLANK: returns false for cell with text") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("data"))
    val result = evalBool("=ISBLANK(A1)", sheet)
    assertEquals(result, Right(false))
  }

  test("ISBLANK: returns false for empty string (NOT blank)") {
    // Key Excel behavior: empty string is NOT considered blank
    val sheet = sheetWith(ref"A1" -> CellValue.Text(""))
    val result = evalBool("=ISBLANK(A1)", sheet)
    assertEquals(result, Right(false))
  }

  test("ISBLANK: returns false for zero") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(0))
    val result = evalBool("=ISBLANK(A1)", sheet)
    assertEquals(result, Right(false))
  }

  test("ISBLANK: returns false for error cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.Value))
    val result = evalBool("=ISBLANK(A1)", sheet)
    assertEquals(result, Right(false))
  }

  // ===== ISERR Tests =====

  test("ISERR: returns true for #DIV/0!") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.Div0))
    val result = evalBool("=ISERR(A1)", sheet)
    assertEquals(result, Right(true))
  }

  test("ISERR: returns true for #VALUE!") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.Value))
    val result = evalBool("=ISERR(A1)", sheet)
    assertEquals(result, Right(true))
  }

  test("ISERR: returns true for #REF!") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.Ref))
    val result = evalBool("=ISERR(A1)", sheet)
    assertEquals(result, Right(true))
  }

  test("ISERR: returns true for #NAME?") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.Name))
    val result = evalBool("=ISERR(A1)", sheet)
    assertEquals(result, Right(true))
  }

  test("ISERR: returns true for #NUM!") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.Num))
    val result = evalBool("=ISERR(A1)", sheet)
    assertEquals(result, Right(true))
  }

  test("ISERR: returns FALSE for #N/A (key difference from ISERROR)") {
    // This is the key difference between ISERR and ISERROR
    // ISERR returns FALSE for #N/A, while ISERROR returns TRUE
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.NA))
    val result = evalBool("=ISERR(A1)", sheet)
    assertEquals(result, Right(false))
  }

  test("ISERR: returns false for numeric cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(100))
    val result = evalBool("=ISERR(A1)", sheet)
    assertEquals(result, Right(false))
  }

  test("ISERR: returns false for text cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("data"))
    val result = evalBool("=ISERR(A1)", sheet)
    assertEquals(result, Right(false))
  }

  test("ISERR: returns false for empty cell") {
    val sheet = emptySheet
    val result = evalBool("=ISERR(A1)", sheet)
    assertEquals(result, Right(false))
  }

  // ===== Comparison: ISERR vs ISERROR =====

  test("ISERR vs ISERROR: both return true for #DIV/0!") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.Div0))
    assertEquals(evalBool("=ISERR(A1)", sheet), Right(true))
    assertEquals(evalBool("=ISERROR(A1)", sheet), Right(true))
  }

  test("ISERR vs ISERROR: differ on #N/A") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.NA))
    // ISERR excludes #N/A
    assertEquals(evalBool("=ISERR(A1)", sheet), Right(false))
    // ISERROR includes all errors
    assertEquals(evalBool("=ISERROR(A1)", sheet), Right(true))
  }

  // ===== Round-trip Tests =====

  test("ISNUMBER: parse/print round-trip") {
    val formula = "=ISNUMBER(A1)"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight, s"Failed to parse: $formula")
    val printed = FormulaPrinter.print(parsed.toOption.get)
    assertEquals(printed, formula)
  }

  test("ISTEXT: parse/print round-trip") {
    val formula = "=ISTEXT(B2)"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight, s"Failed to parse: $formula")
    val printed = FormulaPrinter.print(parsed.toOption.get)
    assertEquals(printed, formula)
  }

  test("ISBLANK: parse/print round-trip") {
    val formula = "=ISBLANK(C3)"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight, s"Failed to parse: $formula")
    val printed = FormulaPrinter.print(parsed.toOption.get)
    assertEquals(printed, formula)
  }

  test("ISERR: parse/print round-trip") {
    val formula = "=ISERR(D4)"
    val parsed = FormulaParser.parse(formula)
    assert(parsed.isRight, s"Failed to parse: $formula")
    val printed = FormulaPrinter.print(parsed.toOption.get)
    assertEquals(printed, formula)
  }
