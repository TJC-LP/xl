package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError
import com.tjclp.xl.formula.eval.CellEvalError
import com.tjclp.xl.formula.graph.DependencyGraph
import munit.FunSuite

/**
 * GH-280: internal diagnostic formatters quote cell-ref-shaped sheet names.
 *
 * After GH-263, user-facing formula PRINTING quotes sheet names that read as cell references —
 * but diagnostics still built bare strings, so an error about a sheet literally named `A1` read
 * `A1!B2` (ambiguous: is that a ref chain?). These are never re-parsed; the quoting is purely for
 * unambiguous human reading via SheetName.quoteForFormula.
 */
class DiagnosticQuotingSpec extends FunSuite:

  private val trickySheet = new Sheet(name = SheetName.unsafe("S"))

  test("GH-280: cross-sheet REF without workbook context quotes the sheet name") {
    // assert on the refStr portion of the message (the formula echo at the end would
    // contain the quoted text regardless)
    trickySheet.evaluateFormula("='A1'!B2") match
      case Left(err) =>
        assert(
          err.message.contains("reference 'A1'!B2 requires"),
          s"expected quoted refStr in: ${err.message}"
        )
      case Right(v) => fail(s"expected Left(missing workbook), got $v")
  }

  test("GH-280: cross-sheet RANGE without workbook context quotes the sheet name") {
    trickySheet.evaluateFormula("=SUM('A1'!B2:B5)") match
      case Left(err) =>
        assert(
          err.message.contains("range 'A1'!B2:B5 requires"),
          s"expected quoted refStr in: ${err.message}"
        )
      case Right(v) => fail(s"expected Left(missing workbook), got $v")
  }

  test("GH-280: bare cross-sheet range misuse quotes the sheet name") {
    trickySheet.evaluateFormula("='A1'!B2:B5") match
      case Left(err) =>
        assert(
          err.message.contains("range 'A1'!B2:B5 "),
          s"expected quoted refStr in: ${err.message}"
        )
      case Right(v) => fail(s"expected Left(range misuse), got $v")
  }

  test("GH-280: QualifiedRef.toString quotes cell-ref-shaped names, leaves safe names bare") {
    assertEquals(
      DependencyGraph.QualifiedRef(SheetName.unsafe("A1"), ref"B2").toString,
      "'A1'!B2"
    )
    assertEquals(
      DependencyGraph.QualifiedRef(SheetName.unsafe("Data"), ref"B2").toString,
      "Data!B2"
    )
    assertEquals(
      DependencyGraph.QualifiedRef(SheetName.unsafe("My Sheet"), ref"B2").toString,
      "'My Sheet'!B2"
    )
  }

  test("GH-280: RangeLocation.toA1 quotes cross-sheet names (diagnostic rendering)") {
    val range = CellRange(ref"B2", ref"B5")
    assertEquals(TExpr.RangeLocation.CrossSheet(SheetName.unsafe("A1"), range).toA1, "'A1'!B2:B5")
    assertEquals(TExpr.RangeLocation.CrossSheet(SheetName.unsafe("Data"), range).toA1, "Data!B2:B5")
    assertEquals(TExpr.RangeLocation.Local(range).toA1, "B2:B5")
  }

  test("GH-280: CellEvalError.render quotes the sheet name") {
    val err = CellEvalError(SheetName.unsafe("A1"), ref"B2", XLError.FormulaError("=X", "boom"))
    assert(err.render.startsWith("'A1'!B2:"), s"expected quoted prefix in: ${err.render}")
    val safe = CellEvalError(SheetName.unsafe("Data"), ref"B2", XLError.FormulaError("=X", "boom"))
    assert(safe.render.startsWith("Data!B2:"), s"expected bare prefix in: ${safe.render}")
  }

  test("GH-280: printWithTypes (TExpr debug rendering) quotes sheet-qualified shapes") {
    val expr = FormulaParser.parse("='A1'!B2+1").fold(e => fail(e.toString), identity)
    val debug = FormulaPrinter.printWithTypes(expr)
    assert(debug.contains("'A1'"), s"expected quoted sheet name in: $debug")
  }
