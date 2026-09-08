package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.EvalFormulaSupport
import munit.FunSuite

/**
 * W2.1 (ADR-017 §2.12), the formula half: `EvalFormulaSupport` is the evaluator behind the core
 * `Edit` interpreter — `validate` parses, `shift` is `FormulaOps.shift`, the structural edits are
 * `StructuralEditor.*Checked`, `renameSheet` is `SheetRenamer.rename` — and
 * `import com.tjclp.xl.{*, given}` supplies it as the `given FormulaSupport`, so the edits the
 * text-only support refuses (a drag, an insert, a rename) apply here and rewrite the formulas of
 * every sheet.
 */
class EvalFormulaSupportSpec extends FunSuite:

  private val Data = SheetName.unsafe("Data")
  private val Other = SheetName.unsafe("Other")
  private val Renamed = SheetName.unsafe("Renamed")

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def f(text: String): CellValue = CellValue.Formula(text, None)
  private def formulaText(v: CellValue): String = v match
    case CellValue.Formula(expr, _, _) => expr
    case other => other.toString

  private val book: Workbook = Workbook(
    Sheet(Data).put(ref"A1", num(10)).put(ref"A2", num(20)).put(ref"B1", f("A1*2")),
    Sheet(Other).put(ref"A1", f("Data!A1+1"))
  )

  test("validate parses (leading = optional) and reports a broken formula as FormulaError") {
    assertEquals(EvalFormulaSupport.validate("=SUM(A1:A3)"), Right(()))
    assertEquals(EvalFormulaSupport.validate("A1*2"), Right(()))
    assertEquals(EvalFormulaSupport.validate("  =A1 "), Right(()))
    EvalFormulaSupport.validate("=SUM(A1:A3") match
      case Left(XLError.FormulaError(formula, _)) => assertEquals(formula, "=SUM(A1:A3")
      case other => fail(s"expected FormulaError, got $other")
    assert(EvalFormulaSupport.validate("").isLeft)
  }

  test("shift is FormulaOps.shift: relative references move, anchors hold, broken text refuses") {
    assertEquals(EvalFormulaSupport.shift("=A1+$B$1", 1, 2), Right("=B3+$B$1"))
    assertEquals(EvalFormulaSupport.shift("A1", 0, 1), Right("A2"))
    assert(EvalFormulaSupport.shift("=A1+", 1, 1).isLeft)
  }

  test("the structural edits rewrite references on every sheet and refuse before mutating") {
    val inserted = EvalFormulaSupport.insertRows(book, Data, Row.from0(0), 1)
    assertEquals(inserted.flatMap(_(Data)).map(_(ref"A2").value), Right(num(10)))
    assertEquals(inserted.flatMap(_(Data)).map(s => formulaText(s(ref"B2").value)), Right("A2*2"))
    assertEquals(
      inserted.flatMap(_(Other)).map(s => formulaText(s(ref"A1").value)),
      Right("Data!A2+1")
    )
    val deleted = EvalFormulaSupport.deleteRows(book, Data, Row.from0(0), 1)
    assertEquals(deleted.flatMap(_(Data)).map(_(ref"A1").value), Right(num(20)))
    // Other!A1 referenced the deleted Data!A1: that reference is #REF!, the formula survives (GH-629)
    assertEquals(
      deleted.flatMap(_(Other)).map(s => formulaText(s(ref"A1").value)),
      Right("#REF!+1")
    )
    val widened = EvalFormulaSupport.insertCols(book, Data, Column.from0(0), 2)
    assertEquals(widened.flatMap(_(Data)).map(s => formulaText(s(ref"D1").value)), Right("C1*2"))
    assertEquals(
      widened.flatMap(_(Other)).map(s => formulaText(s(ref"A1").value)),
      Right("Data!C1+1")
    )
    val narrowed = EvalFormulaSupport.deleteCols(book, Data, Column.from0(1), 1)
    assertEquals(narrowed.flatMap(_(Data)).map(_.contains(ref"B1")), Right(false))
    // StructuralEditor alone no-ops on an unknown sheet; the seam refuses it by name
    assertEquals(
      EvalFormulaSupport.insertRows(book, SheetName.unsafe("Nope"), Row.from0(0), 1),
      Left(XLError.SheetNotFound("Nope"))
    )
    assertEquals(
      EvalFormulaSupport.deleteCols(book, SheetName.unsafe("Nope"), Column.from0(0), 1),
      Left(XLError.SheetNotFound("Nope"))
    )
  }

  test("renameSheet is SheetRenamer.rename: the tab and every reference to it move") {
    val renamed = EvalFormulaSupport.renameSheet(book, Data, Renamed)
    assertEquals(renamed.map(_.sheetNames.map(_.value)), Right(Seq("Renamed", "Other")))
    assertEquals(
      renamed.flatMap(_(Other)).map(s => formulaText(s(ref"A1").value)),
      Right("Renamed!A1+1")
    )
    assert(EvalFormulaSupport.renameSheet(book, Data, Other).isLeft)
  }

  test(
    "import com.tjclp.xl.{*, given} supplies EvalFormulaSupport and the interpreter runs on it"
  ) {
    assert(summon[FormulaSupport] eq EvalFormulaSupport)
    val r = book.editIn(Data)(
      Edit.DragFormula(Area(None, ref"C1:C2"), "=A1*3", ref"C1", None),
      Edit.Fill(Area(None, CellRange(ref"B1", ref"B1")), ref"B1:B2", Edit.FillDir.Down),
      Edit.InsertRows(None, Row.from0(0), 1),
      Edit.RenameSheet(Data, Renamed)
    )
    val data = r.flatMap(_(Renamed)).fold(e => fail(e.message), identity)
    assertEquals(formulaText(data(ref"C2").value), "A2*3")
    assertEquals(formulaText(data(ref"C3").value), "A3*3")
    assertEquals(formulaText(data(ref"B3").value), "A3*2")
    val other = r.flatMap(_(Other)).fold(e => fail(e.message), identity)
    assertEquals(formulaText(other(ref"A1").value), "Renamed!A2+1")
    // the text-only support refuses the same insert with the capability it lacks
    book.editIn(Data)(Edit.InsertRows(None, Row.from0(0), 1))(using FormulaSupport.textOnly) match
      case Left(XLError.EditFailed(1, "insert-rows", XLError.UnsupportedCapability(_, _, _))) => ()
      case other => fail(s"expected an UnsupportedCapability refusal, got $other")
  }
