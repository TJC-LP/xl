package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.display.FormulaDisplayStrategy
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.FormulaFormatting
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

import java.time.LocalDate

/**
 * Tests for number-format inheritance on formula cells (GH-184).
 *
 * Covers FormulaFormatting.inferFormatFromReferences (category rules, first-reference-wins,
 * cross-sheet resolution) and the opt-in Sheet.putFormulaInheriting extension.
 */
class FormulaFormattingSpec extends FunSuite:
  private val base = new Sheet(name = SheetName.unsafe("Test"))

  private val currency = CellStyle.default.withNumFmt(NumFmt.Currency)
  private val percent = CellStyle.default.withNumFmt(NumFmt.Percent)
  private val percentDecimal = CellStyle.default.withNumFmt(NumFmt.PercentDecimal)
  private val text = CellStyle.default.withNumFmt(NumFmt.Text)
  private val accountingCode = """_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)"""
  private val accounting = CellStyle.default.withNumFmt(NumFmt.Custom(accountingCode))

  private def parseOrFail(formula: String): TExpr[?] =
    FormulaParser.parse(formula) match
      case Right(expr) => expr
      case Left(err) => fail(s"Parse failed for '$formula': $err")

  private def numFmtOf(sheet: Sheet, ref: ARef): Option[NumFmt] =
    sheet.cells.get(ref).flatMap(_.styleId).flatMap(sheet.styleRegistry.get).map(_.numFmt)

  private def styleOf(sheet: Sheet, ref: ARef): Option[CellStyle] =
    sheet.cells.get(ref).flatMap(_.styleId).flatMap(sheet.styleRegistry.get)

  // ===== inferFormatFromReferences =====

  test("single currency reference inherits Currency") {
    val sheet = base.put(ref"B2", 100, currency)
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=B2*2"), sheet)
    assertEquals(inferred, Some(NumFmt.Currency))
  }

  test("multiple references with identical format inherit that format") {
    val sheet = base
      .put(ref"B2", 1000000, currency)
      .put(ref"B3", 600000, currency)
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=B2-B3"), sheet)
    assertEquals(inferred, Some(NumFmt.Currency))
  }

  test("mixed categories (currency + percent) yield General (None)") {
    val sheet = base
      .put(ref"B2", 100, currency)
      .put(ref"B3", BigDecimal("0.25"), percent)
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=B2*B3"), sheet)
    assertEquals(inferred, None)
  }

  test("mixed categories (currency + date) yield General (None)") {
    val sheet = base
      .put(ref"B2", 100, currency)
      .put(ref"B3", LocalDate.of(2026, 1, 15)) // codec auto-infers Date format
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=B2+B3"), sheet)
    assertEquals(inferred, None)
  }

  test("mixed categories (text + currency) yield General (None)") {
    val sheet = base
      .put(ref"B2", "label", text)
      .put(ref"B3", 100, currency)
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=B2&B3"), sheet)
    assertEquals(inferred, None)
  }

  test("same category, different precision: first reference wins (Percent then PercentDecimal)") {
    val sheet = base
      .put(ref"B2", BigDecimal("0.25"), percent)
      .put(ref"B3", BigDecimal("0.125"), percentDecimal)
    val forward = FormulaFormatting.inferFormatFromReferences(parseOrFail("=B2+B3"), sheet)
    assertEquals(forward, Some(NumFmt.Percent))
    val reverse = FormulaFormatting.inferFormatFromReferences(parseOrFail("=B3+B2"), sheet)
    assertEquals(reverse, Some(NumFmt.PercentDecimal))
  }

  test("custom accounting code categorizes as currency: first reference wins") {
    val sheet = base
      .put(ref"B2", 100, accounting)
      .put(ref"B3", 200, currency)
    val forward = FormulaFormatting.inferFormatFromReferences(parseOrFail("=B2+B3"), sheet)
    assertEquals(forward, Some(NumFmt.Custom(accountingCode)))
    val reverse = FormulaFormatting.inferFormatFromReferences(parseOrFail("=B3+B2"), sheet)
    assertEquals(reverse, Some(NumFmt.Currency))
  }

  test("date-formatted reference inherits Date") {
    val sheet = base.put(ref"B2", LocalDate.of(2026, 1, 15)) // codec auto-infers Date format
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=B2+1"), sheet)
    assertEquals(inferred, Some(NumFmt.Date))
  }

  test("formula with no references (=1+2) yields None") {
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=1+2"), base)
    assertEquals(inferred, None)
  }

  test("non-existent reference yields None without error") {
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=Z99*2"), base)
    assertEquals(inferred, None)
  }

  test("references without any number format yield None") {
    val sheet = base.put(ref"B2", 42) // Int infers no format (General)
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=B2*2"), sheet)
    assertEquals(inferred, None)
  }

  test("unformatted references are skipped: format comes from the formatted one") {
    val sheet = base
      .put(ref"B2", 42) // General - skipped
      .put(ref"B3", 100, currency)
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=B2+B3"), sheet)
    assertEquals(inferred, Some(NumFmt.Currency))
  }

  test("range reference inherits the common format (=SUM(A1:A3))") {
    val sheet = base
      .put(ref"A1", 10, currency)
      .put(ref"A2", 20, currency)
      .put(ref"A3", 30, currency)
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=SUM(A1:A3)"), sheet)
    assertEquals(inferred, Some(NumFmt.Currency))
  }

  test("full-column aggregate is bounded by the used range (=SUM(A:A))") {
    val sheet = base
      .put(ref"A1", 10, currency)
      .put(ref"A2", 20, currency)
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=SUM(A:A)"), sheet)
    assertEquals(inferred, Some(NumFmt.Currency))
  }

  test("range with mixed categories yields None") {
    val sheet = base
      .put(ref"A1", 10, currency)
      .put(ref"A2", BigDecimal("0.5"), percent)
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=SUM(A1:A2)"), sheet)
    assertEquals(inferred, None)
  }

  test("cross-sheet reference resolves through the workbook") {
    val data = new Sheet(name = SheetName.unsafe("Data")).put(ref"B2", 250, currency)
    val summary = new Sheet(name = SheetName.unsafe("Summary"))
    val wb = Workbook(data, summary)
    val inferred =
      FormulaFormatting.inferFormatFromReferences(parseOrFail("=Data!B2*2"), summary, Some(wb))
    assertEquals(inferred, Some(NumFmt.Currency))
  }

  test("cross-sheet aggregate resolves through the workbook (=SUM(Data!B1:B3))") {
    val data = new Sheet(name = SheetName.unsafe("Data"))
      .put(ref"B1", 1, currency)
      .put(ref"B2", 2, currency)
      .put(ref"B3", 3, currency)
    val summary = new Sheet(name = SheetName.unsafe("Summary"))
    val wb = Workbook(data, summary)
    val inferred =
      FormulaFormatting.inferFormatFromReferences(parseOrFail("=SUM(Data!B1:B3)"), summary, Some(wb))
    assertEquals(inferred, Some(NumFmt.Currency))
  }

  test("cross-sheet reference without workbook context yields None") {
    val summary = new Sheet(name = SheetName.unsafe("Summary"))
    val inferred = FormulaFormatting.inferFormatFromReferences(parseOrFail("=Data!B2*2"), summary)
    assertEquals(inferred, None)
  }

  // ===== putFormulaInheriting =====

  test("putFormulaInheriting writes the formula and inherits Currency (GH-184 repro)") {
    val sheet = base
      .put(ref"B2", 1000000, currency)
      .put(ref"B3", 600000, currency)
    sheet.putFormulaInheriting(ref"B4", "=B2-B3") match
      case Right(updated) =>
        updated.cells.get(ref"B4").map(_.value) match
          case Some(CellValue.Formula(expr, _)) => assertEquals(expr, "=B2-B3")
          case other => fail(s"Expected formula cell at B4, got $other")
        assertEquals(numFmtOf(updated, ref"B4"), Some(NumFmt.Currency))
      case Left(err) => fail(s"putFormulaInheriting failed: $err")
  }

  test("putFormulaInheriting leaves style untouched when categories are mixed") {
    val sheet = base
      .put(ref"B2", 100, currency)
      .put(ref"B3", BigDecimal("0.25"), percent)
    sheet.putFormulaInheriting(ref"B4", "=B2*B3") match
      case Right(updated) =>
        assertEquals(numFmtOf(updated, ref"B4"), None)
        updated.cells.get(ref"B4").map(_.value) match
          case Some(CellValue.Formula(expr, _)) => assertEquals(expr, "=B2*B3")
          case other => fail(s"Expected formula cell at B4, got $other")
      case Left(err) => fail(s"putFormulaInheriting failed: $err")
  }

  test("putFormulaInheriting preserves an existing non-General target format (Excel parity)") {
    val sheet = base
      .put(ref"B2", 100, currency)
      .put(ref"B4", 0, percent) // target already explicitly formatted
    sheet.putFormulaInheriting(ref"B4", "=B2*2") match
      case Right(updated) => assertEquals(numFmtOf(updated, ref"B4"), Some(NumFmt.Percent))
      case Left(err) => fail(s"putFormulaInheriting failed: $err")
  }

  test("putFormulaInheriting merges the inherited format into the target's existing style") {
    val bold = CellStyle.default.bold
    val sheet = base
      .put(ref"B2", 100, currency)
      .put(ref"B4", 0, bold) // bold but General: format slot is free to inherit
    sheet.putFormulaInheriting(ref"B4", "=B2*2") match
      case Right(updated) =>
        assertEquals(styleOf(updated, ref"B4"), Some(bold.withNumFmt(NumFmt.Currency)))
      case Left(err) => fail(s"putFormulaInheriting failed: $err")
  }

  test("putFormulaInheriting returns Left on parse error") {
    val result = base.putFormulaInheriting(ref"B4", "=SUM((")
    assert(result.isLeft, s"Expected parse failure, got $result")
  }

  test("putFormulaInheriting normalizes a missing leading equals sign") {
    val sheet = base.put(ref"B2", 100, currency)
    sheet.putFormulaInheriting(ref"B4", "B2*2") match
      case Right(updated) =>
        updated.cells.get(ref"B4").map(_.value) match
          case Some(CellValue.Formula(expr, _)) => assertEquals(expr, "=B2*2")
          case other => fail(s"Expected formula cell at B4, got $other")
        assertEquals(numFmtOf(updated, ref"B4"), Some(NumFmt.Currency))
      case Left(err) => fail(s"putFormulaInheriting failed: $err")
  }

  test("putFormulaInheriting workbook overload inherits across sheets") {
    val data = new Sheet(name = SheetName.unsafe("Data")).put(ref"B2", 250, currency)
    val summary = new Sheet(name = SheetName.unsafe("Summary"))
    val wb = Workbook(data, summary)
    summary.putFormulaInheriting(ref"A1", "=Data!B2*2", wb) match
      case Right(updated) => assertEquals(numFmtOf(updated, ref"A1"), Some(NumFmt.Currency))
      case Left(err) => fail(s"putFormulaInheriting failed: $err")
  }

  test("putFormulaInheriting without workbook still writes cross-sheet formula, no inherit") {
    val summary = new Sheet(name = SheetName.unsafe("Summary"))
    summary.putFormulaInheriting(ref"A1", "=Data!B2*2") match
      case Right(updated) =>
        updated.cells.get(ref"A1").map(_.value) match
          case Some(CellValue.Formula(expr, _)) => assertEquals(expr, "=Data!B2*2")
          case other => fail(s"Expected formula cell at A1, got $other")
        assertEquals(numFmtOf(updated, ref"A1"), None)
      case Left(err) => fail(s"putFormulaInheriting failed: $err")
  }

  test("putFormulaInheriting end-to-end: recalculate caches value, display shows $400,000.00") {
    val sheet = base
      .put(ref"B2", 1000000, currency)
      .put(ref"B3", 600000, currency)
    sheet.putFormulaInheriting(ref"B4", "=B2-B3") match
      case Right(withFormula) =>
        val recalced = Workbook(withFormula).recalculate()
        assert(recalced.errors.isEmpty, s"recalculate errors: ${recalced.errors}")
        val outSheet = recalced.workbook.sheets.headOption.getOrElse(fail("workbook has no sheet"))
        outSheet.cells.get(ref"B4").map(_.value) match
          case Some(CellValue.Formula(_, Some(CellValue.Number(n)))) =>
            assertEquals(n, BigDecimal(400000))
          case other => fail(s"Expected cached formula value at B4, got $other")
        assertEquals(numFmtOf(outSheet, ref"B4"), Some(NumFmt.Currency))
        given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
        assertEquals(outSheet.displayCell(ref"B4").formatted, "$400,000.00")
      case Left(err) => fail(s"putFormulaInheriting failed: $err")
  }
