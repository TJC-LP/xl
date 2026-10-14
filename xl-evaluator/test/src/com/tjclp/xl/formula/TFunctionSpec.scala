package com.tjclp.xl.formula

import java.time.LocalDateTime

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.eval.{ArrayResult, Evaluator}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-684: T(value). Text (a literal, a text cell, a rich-text cell) returns itself; every other
 * value returns "" (Microsoft, "T function"); a formula cell answers T of its cached value; an
 * error propagates. Each expected value below is also what LibreOffice 24.2 caches for the same
 * formula.
 */
class TFunctionSpec extends FunSuite:

  private val name = SheetName.unsafe("Sheet1")

  private def text(s: String): CellValue = CellValue.Text(s)
  private val empty: CellValue = text("")

  private val sheet: Sheet = List[(ARef, CellValue)](
    ref"A1" -> text("OldLabelAlpha"),
    ref"A2" -> CellValue.RichText("Bold".bold + " plain"),
    ref"A3" -> text(""),
    ref"B1" -> CellValue.Number(42),
    ref"B2" -> CellValue.Bool(true),
    ref"B3" -> CellValue.DateTime(LocalDateTime.of(2026, 1, 1, 0, 0)),
    ref"B4" -> CellValue.Error(CellError.Div0),
    ref"B5" -> CellValue.Error(CellError.NA),
    ref"C1" -> CellValue.Formula("A1&\"x\"", Some(text("OldLabelAlphax"))),
    ref"C2" -> CellValue.Formula("1+1", Some(CellValue.Number(2))),
    ref"C3" -> CellValue.Formula("1/0", Some(CellValue.Error(CellError.Div0))),
    ref"C4" -> CellValue.Formula("A2", Some(CellValue.RichText(RichText.plain("cached rich")))),
    ref"E11" -> text("named text")
  ).foldLeft(Sheet(name)) { case (s, (at, v)) => s.put(at, v) }

  private val book: Workbook = Workbook(sheet)
    .withDefinedName("BlankName", "Sheet1!$E$10")
    .withDefinedName("TextName", "Sheet1!$E$11")

  private def evalT(formula: String): Either[String, CellValue] =
    book.evaluateFormula(formula, name).left.map(_.message)

  test("GH-684: T returns text as itself") {
    assertEquals(evalT("=T(\"lit\")"), Right(text("lit")))
    assertEquals(evalT("=T(A1)"), Right(text("OldLabelAlpha")))
    assertEquals(evalT("=T(\"\")"), Right(empty))
    assertEquals(evalT("=T(A3)"), Right(empty), "a cell holding \"\" is text")
  }

  test("GH-684: T returns a rich-text cell's plain text") {
    assertEquals(evalT("=T(A2)"), Right(text("Bold plain")))
  }

  test("GH-684: T returns \"\" for numbers, booleans and dates") {
    assertEquals(evalT("=T(42)"), Right(empty))
    assertEquals(evalT("=T(TRUE)"), Right(empty))
    assertEquals(evalT("=T(DATE(2026,1,1))"), Right(empty))
    assertEquals(evalT("=T(B1)"), Right(empty))
    assertEquals(evalT("=T(B2)"), Right(empty))
    assertEquals(evalT("=T(B3)"), Right(empty))
  }

  test("GH-684: T returns \"\" for an empty cell, not 0") {
    assertEquals(evalT("=T(E9)"), Right(empty))
    assertEquals(evalT("=LEN(T(E9))"), Right(CellValue.Number(0)))
    assertEquals(evalT("=ISTEXT(T(5))"), Right(CellValue.Bool(true)))
  }

  test("GH-684: T of a formula cell is T of its cached value") {
    assertEquals(evalT("=T(C1)"), Right(text("OldLabelAlphax")))
    assertEquals(evalT("=T(C2)"), Right(empty))
    assertEquals(evalT("=T(C4)"), Right(text("cached rich")))
  }

  test("GH-684: T propagates errors: literal, error cell, cached error") {
    assertEquals(evalT("=T(1/0)"), Right(CellValue.Error(CellError.Div0)))
    assertEquals(evalT("=T(NA())"), Right(CellValue.Error(CellError.NA)))
    assertEquals(evalT("=T(B4)"), Right(CellValue.Error(CellError.Div0)))
    assertEquals(evalT("=T(B5)"), Right(CellValue.Error(CellError.NA)))
    assertEquals(evalT("=T(C3)"), Right(CellValue.Error(CellError.Div0)))
    assertEquals(evalT("=IFERROR(T(B4),\"caught\")"), Right(text("caught")))
  }

  test("GH-684: T through a defined name reads the named cell") {
    assertEquals(evalT("=T(BlankName)"), Right(empty))
    assertEquals(evalT("=T(TextName)"), Right(text("named text")))
  }

  test("GH-684: T over a range spills element-wise in an array context") {
    val expr = FormulaParser.parse("=T(A1:C1)").fold(e => fail(e.toString), identity)
    assertEquals[Either[Any, Any], Either[Any, Any]](
      Evaluator.arrayInstance.eval(expr, sheet, Clock.system, Some(book), None),
      Right(ArrayResult(Vector(Vector(text("OldLabelAlpha"), empty, text("OldLabelAlphax")))))
    )
  }

  test("GH-684: T over a range in a plain cell is implicitly intersected") {
    def plain(at: ARef): Either[String, CellValue] =
      val placed = sheet.put(at, CellValue.Formula("T(A1:A3)", None))
      placed.evaluateCell(at, Clock.system, Some(book.put(placed))).left.map(_.message)
    assertEquals(plain(ref"H1"), Right(text("OldLabelAlpha")))
    assertEquals(plain(ref"H2"), Right(text("Bold plain")))
    assertEquals(plain(ref"H9"), Right(CellValue.Error(CellError.Value)))
  }

  test("GH-684: T as a name prefix still parses cell and column references") {
    val withT = book.put(sheet.put(ref"T1", text("in T1")).put(ref"T2", CellValue.Number(3)))
    def at(formula: String): Either[String, CellValue] =
      withT.evaluateFormula(formula, name).left.map(_.message)
    assertEquals(at("=T1"), Right(text("in T1")))
    assertEquals(at("=T(T1)"), Right(text("in T1")))
    assertEquals(at("=SUM(T:T)"), Right(CellValue.Number(3)))
  }

  test("GH-684: recalculation computes T() results (LibreOffice-saved probe)") {
    // The cells of the LibreOffice-saved probe in the #684 comment, written with no cached values:
    // a recalculation must compute every value LibreOffice cached, where it used to fail each cell
    // with "Unknown function 'T'". D7 and D8 read T() cells through the dependency graph.
    def formula(expr: String): CellValue =
      CellValue.Formula(expr, None, FormulaKind.Normal())
    val probe = Workbook(
      Sheet(name)
        .put(ref"A1", text("OldLabelAlpha"))
        .put(ref"B1", CellValue.Number(42))
        .put(ref"D1", formula("T(A1)"))
        .put(ref"D2", formula("T(B1)"))
        .put(ref"D3", formula("T(E9)"))
        .put(ref"D4", formula("T(BlankName)"))
        .put(ref"D5", formula("T(1/0)"))
        .put(ref"D6", formula("T(\"lit\")"))
        .put(ref"D7", formula("T(D1)&\"!\""))
        .put(ref"D8", formula("T(D5)"))
    ).withDefinedName("BlankName", "Sheet1!$E$10")
    val result = probe.recalculate(Clock.system)
    assertEquals(result.errors, Vector.empty)
    val expected = List(
      ref"D1" -> text("OldLabelAlpha"),
      ref"D2" -> empty,
      ref"D3" -> empty,
      ref"D4" -> empty,
      ref"D5" -> CellValue.Error(CellError.Div0),
      ref"D6" -> text("lit"),
      ref"D7" -> text("OldLabelAlpha!"),
      ref"D8" -> CellValue.Error(CellError.Div0)
    )
    val recalculated = result.workbook(name).fold(e => fail(e.message), identity)
    expected.foreach { (at, value) =>
      recalculated(at).value match
        case CellValue.Formula(_, cached, _) => assertEquals(cached, Some(value), at.toA1)
        case other => fail(s"${at.toA1}: expected a formula, got $other")
    }
  }
