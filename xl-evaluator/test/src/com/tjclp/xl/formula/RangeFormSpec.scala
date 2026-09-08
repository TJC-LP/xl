package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.Column
import com.tjclp.xl.formula.printer.{FormulaOps, FormulaPrinter, FormulaShifter}
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAllNoShrink

/**
 * GH-612: the `RangeForm` carried on the AST's range nodes — its classification agrees with
 * `CellRange.parse` by construction, and an inconsistent hand-built pairing is treated as the
 * corner range it addresses (the type cannot enforce the precondition, so the printer and both
 * shifters must stay truthful).
 */
class RangeFormSpec extends ScalaCheckSuite:

  private def parsedRange(text: String): CellRange =
    CellRange.parse(text).fold(e => fail(s"$text should parse: $e"), identity)

  private val genPart: Gen[String] =
    for
      dollar <- Gen.oneOf("", "$")
      body <- Gen.oneOf(
        Gen.choose(0, Column.MaxIndex0).map(Column.from0(_).toLetter),
        Gen.choose(1, 1048576).map(_.toString),
        for
          c <- Gen.choose(0, 40)
          r <- Gen.choose(1, 60)
        yield s"${Column.from0(c).toLetter}$r"
      )
    yield dollar + body

  property(
    "RangeForm.of agrees with CellRange.parse: Columns implies a full column, Rows a full row"
  ) {
    forAllNoShrink(genPart, genPart) { (start, end) =>
      val text = s"$start:$end"
      CellRange.parse(text) match
        case Left(_) => true // a mixed spelling such as A:3 is not a range; nothing to classify
        case Right(range) =>
          RangeForm.of(text, range) match
            case RangeForm.Columns => range.isFullColumn
            case RangeForm.Rows => range.isFullRow
            case RangeForm.Cells => !range.isFullColumn && !range.isFullRow
    }
  }

  test(
    "RangeForm.of: the spelled form wins; a corner spelling over every row/column canonicalises"
  ) {
    assertEquals(RangeForm.of("A:C", parsedRange("A:C")), RangeForm.Columns)
    assertEquals(RangeForm.of("$3:$10", parsedRange("$3:$10")), RangeForm.Rows)
    assertEquals(RangeForm.of("A1:A1048576", parsedRange("A1:A1048576")), RangeForm.Columns)
    assertEquals(RangeForm.of("A1:XFD1", parsedRange("A1:XFD1")), RangeForm.Rows)
    assertEquals(RangeForm.of("A1:XFD1048576", parsedRange("A1:XFD1048576")), RangeForm.Rows)
    assertEquals(RangeForm.of("A2:A1048576", parsedRange("A2:A1048576")), RangeForm.Cells)
    // the no-colon branch: a single-cell "range" is a corner range
    assertEquals(RangeForm.of("A1", parsedRange("A1")), RangeForm.Cells)
  }

  test("an inconsistent hand-built node prints its corners, never a widened A:B") {
    val corner = CellRange(ref"A1", ref"B2")
    assertEquals(FormulaPrinter.print(TExpr.RangeRef(corner, RangeForm.Columns)), "=A1:B2")
    assertEquals(FormulaPrinter.print(TExpr.RangeRef(corner, RangeForm.Rows)), "=A1:B2")
    assertEquals(
      FormulaPrinter.print(TExpr.SheetRange(SheetName.unsafe("S"), corner, RangeForm.Columns)),
      "=S!A1:B2"
    )
    // a consistent pairing still prints its form
    assertEquals(
      FormulaPrinter.print(TExpr.RangeRef(parsedRange("A:B"), RangeForm.Columns)),
      "=A:B"
    )
    assertEquals(RangeForm.Columns.actualFor(corner), RangeForm.Cells)
    assertEquals(RangeForm.Columns.actualFor(parsedRange("A:B")), RangeForm.Columns)
  }

  test("an inconsistent hand-built node drags and restructures like the corner range it is") {
    val node = TExpr.RangeRef(CellRange(ref"A1", ref"B2"), RangeForm.Columns)
    assertEquals(FormulaPrinter.print(FormulaShifter.shift(node, 1, 1)), "=B2:C3")
    val inserted = FormulaShifter.shiftStructural(node, shiftLocal = true, "S", isRow = true, 0, 2)
    assertEquals(FormulaPrinter.print(inserted), "=A3:B4")
    val deleted = FormulaShifter.shiftStructural(node, shiftLocal = true, "S", isRow = true, 0, -2)
    // both rows deleted: #REF!, not "still every row"
    assertEquals(FormulaPrinter.print(deleted), "=#REF!")
  }

  // Every reference sits at least four cells from the grid's edges (the summed deltas below reach
  // ±4), and no range mixes a relative corner with an anchored one: when a relative corner
  // overtakes an anchored one the normalised print swaps the `$` (`E:$E` right then left is
  // `$E:E`), which is not invertible — Excel has the same asymmetry — so the law is stated for
  // on-grid, non-crossing shifts only.
  private val onGridFormulas = List(
    "=COUNTIF($A:$A,E5)+SUM(5:5)+E5",
    "=SUM(E:H)*Sheet1!5:8",
    "=SUM([2]Book1!E:E)+[2]Book1!E5",
    "=VLOOKUP(E5,E:H,2,FALSE)&\"x\"",
    "=SUM($E$5:$G$9)+SUM(E5:G9)"
  )

  property("commutativity: shifting twice equals shifting once by the summed deltas, on the grid") {
    val delta = Gen.choose(-2, 2)
    forAllNoShrink(Gen.oneOf(onGridFormulas), delta, delta, delta, delta) { (f, c1, r1, c2, r2) =>
      val twice = FormulaOps.shift(f, c1, r1).flatMap(FormulaOps.shift(_, c2, r2))
      assertEquals(twice, FormulaOps.shift(f, c1 + c2, r1 + r2))
      true
    }
  }
