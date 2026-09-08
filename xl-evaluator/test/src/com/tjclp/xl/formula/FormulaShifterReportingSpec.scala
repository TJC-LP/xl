package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.eval.EvalFormulaSupport
import com.tjclp.xl.ops.FormulaSupport
import munit.FunSuite

/**
 * GH-628: a shift that writes `#REF!` for an off-grid or deleted reference reports which references
 * it voided, spelled as they were, so a caller can warn without grepping the text for a `#REF!` the
 * formula may legitimately have contained already.
 */
class FormulaShifterReportingSpec extends FunSuite:

  private def parsed(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(e => fail(e.toString), identity)

  private def drag(formula: String, colDelta: Int, rowDelta: Int): (String, Vector[String]) =
    val shifted = FormulaShifter.shiftReporting(parsed(formula), colDelta, rowDelta)
    (FormulaPrinter.print(shifted.expr), shifted.voided)

  test("a drag reports each reference it carried off the grid, in formula order") {
    assertEquals(drag("=A1+B3", 0, -1), ("=#REF!+B2", Vector("A1")))
    assertEquals(drag("=A1+A2", 0, -2), ("=#REF!+#REF!", Vector("A1", "A2")))
    assertEquals(drag("=SUM(A1:A2)+$A$1", 0, -2), ("=SUM(#REF!)+$A$1", Vector("A1:A2")))
    assertEquals(drag("=COUNTIF($A:$A,B1)", 0, -1), ("=COUNTIF($A:$A, #REF!)", Vector("B1")))
    assertEquals(drag("=Data!B2*XFD1", 1, 0), ("=Data!C2*#REF!", Vector("XFD1")))
    assertEquals(drag("=SUM(Data!A1:B1)", -1, 0), ("=SUM(#REF!)", Vector("Data!A1:B1")))
  }

  test("a drag that voids nothing reports nothing, and a #REF! already present is not a report") {
    assertEquals(drag("=A1+B3", 1, 1), ("=B2+C4", Vector.empty))
    assertEquals(drag("=#REF!+B3", 0, 1), ("=#REF!+B4", Vector.empty))
    assertEquals(FormulaShifter.shiftReporting(parsed("=A1"), 0, 0).voided, Vector.empty)
  }

  test("a structural delete reports the references it voided and keeps the formula") {
    def deleteRows(formula: String, at: Int, count: Int): (String, Vector[String]) =
      val shifted = FormulaShifter.shiftStructuralReporting(
        parsed(formula),
        shiftLocal = true,
        editedSheet = "S",
        isRow = true,
        at = at,
        delta = -count
      )
      (FormulaPrinter.print(shifted.expr), shifted.voided)
    assertEquals(deleteRows("=A1+B3", 0, 1), ("=#REF!+B2", Vector("A1")))
    assertEquals(
      deleteRows("=SUM(A1:A2)+SUM(A5:A6)", 0, 2),
      ("=SUM(#REF!)+SUM(A3:A4)", Vector("A1:A2"))
    )
    assertEquals(deleteRows("=SUM(A1:A10)", 0, 2), ("=SUM(A1:A8)", Vector.empty))
    assertEquals(deleteRows("=A5", 0, 1), ("=A4", Vector.empty))
  }

  test("FormulaOps and EvalFormulaSupport carry the report through the text form") {
    assertEquals(
      FormulaOps.shiftReporting("=A1+B3", 0, -1),
      Right(FormulaSupport.Shifted("=#REF!+B2", Vector("A1")))
    )
    assertEquals(
      FormulaOps.shiftReporting("A1+B3", 0, 1),
      Right(FormulaSupport.Shifted("A2+B4", Vector.empty))
    )
    assertEquals(
      EvalFormulaSupport.shiftReporting("=SUM(A1:A2)", 0, -1),
      Right(FormulaSupport.Shifted("=SUM(#REF!)", Vector("A1:A2")))
    )
    // the default FormulaSupport reports nothing but refuses just like shift
    assert(FormulaSupport.textOnly.shiftReporting("=A1", 0, 1).isLeft)
  }
