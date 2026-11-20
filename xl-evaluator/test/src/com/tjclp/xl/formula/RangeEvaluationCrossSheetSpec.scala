package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.RangeEvalResult
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{CalcPr, Workbook}
import munit.FunSuite

/**
 * #891: the window evaluation behind `view --eval` computes the window's precedents on every sheet,
 * not only its own. A window on one sheet whose formulas read another sheet shows what a full
 * recalculation would put in those cells, even when the other sheet's caches are stale (after
 * `put --no-recalc`, say); until #891 the other sheet's caches were read as they stood.
 *
 * Every book here holds caches that are stale against its constants, so a value read off a cache
 * differs from a value computed.
 */
class RangeEvaluationCrossSheetSpec extends FunSuite:

  private def num(n: BigDecimal): CellValue = CellValue.Number(n)
  private def formula(text: String): CellValue = CellValue.Formula(text)
  private def cached(text: String, value: BigDecimal): CellValue =
    CellValue.Formula(text, Some(num(value)))
  private def at(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)
  private def range(a1: String): CellRange = CellRange.parse(a1).fold(fail(_), identity)

  private def sheet(name: String, cells: (String, CellValue)*): Sheet =
    cells.foldLeft(Sheet(SheetName.unsafe(name)))((s, kv) => s.put(at(kv._1), kv._2))

  /** The window `a1` of the book's first sheet. */
  private def window(wb: Workbook, a1: String): RangeEvalResult =
    wb.sheets(0).evaluateForRangePerCell(range(a1), Clock.system, Some(wb))

  private def located(result: RangeEvalResult): Vector[String] =
    result.failures.map(e => s"${e.sheet.value}!${e.ref.toA1}")

  /** What a full recalculation computes for the first sheet's cells `refs`. */
  private def recalculated(wb: Workbook, refs: String*): Map[ARef, CellValue] =
    val values = wb.recalculate().evaluated.getOrElse(wb.sheets(0).name, Map.empty)
    refs.map(at).flatMap(ref => values.get(ref).map(ref -> _)).toMap

  /**
   * `S!A1` was 1 when the caches were written and is 5 now. `T!B1` reads `S!A1` and `S!F1`; `S!H1`
   * reads `T!B1`: a chain from the window's sheet through another and back.
   */
  private def chain(b1: CellValue): Workbook =
    Workbook(sheet("S", "A1" -> num(5), "F1" -> cached("A1*2", 2), "H1" -> cached("T!B1*10", 30)))
      .put(sheet("T", "B1" -> b1))

  test("a precedent on another sheet is computed, not read off its stale cache") {
    val wb = chain(cached("S!A1+S!F1", 3))
    val result = window(wb, "H1:H1")
    assertEquals(result.values, Map(at("H1") -> num(150)))
    assert(result.isClean, result.summary)
    assertEquals(result.values, recalculated(wb, "H1"))
  }

  test("the chain back to the window's sheet: the cell behind the other sheet is computed too") {
    // F1 lies outside the window and is reached only through T!B1
    val wb = chain(cached("S!A1+S!F1", 3))
    assertEquals(window(wb, "F1:H1").values, Map(at("F1") -> num(10), at("H1") -> num(150)))
    assertEquals(window(wb, "F1:H1").values, recalculated(wb, "F1", "H1"))
  }

  test("an uncached precedent on another sheet is computed in order") {
    val result = window(chain(formula("S!A1+S!F1")), "H1:H1")
    assertEquals(result.values, Map(at("H1") -> num(150)))
    assert(result.isClean, result.summary)
  }

  test("only the window's cells are values: a same-address cell of another sheet is not") {
    val wb = Workbook(sheet("S", "A1" -> num(5), "B1" -> cached("T!B1+1", 3)))
      .put(sheet("T", "B1" -> cached("S!A1*2", 2), "C1" -> cached("B1+1", 3)))
    // S!B1 = T!B1 + 1 with T!B1 = 10; T!C1 is no precedent of the window and no value of it
    assertEquals(window(wb, "A1:C1").values, Map(at("B1") -> num(11)))
  }

  test("a formula of another sheet the window does not read is left alone, failing or not") {
    val wb = Workbook(sheet("S", "A1" -> num(5), "H1" -> cached("T!B1*10", 30)))
      .put(sheet("T", "B1" -> cached("S!A1*3", 3), "Z9" -> formula("NOSUCHFN(1)")))
    val result = window(wb, "H1:H1")
    assertEquals(result.values, Map(at("H1") -> num(150)))
    assert(result.isClean, result.summary)
  }

  test("a precedent on another sheet that fails is reported there and blocks the window's cell") {
    val wb = Workbook(sheet("S", "H1" -> cached("T!B1+1", 4), "C1" -> cached("7+1", 0)))
      .put(sheet("T", "B1" -> cached("NOSUCHFN(1)", 3)))
    val result = window(wb, "A1:H1")
    // H1 is not computed off T!B1's cache: it keeps what the file holds, and the cause is named
    assertEquals(result.values, Map(at("C1") -> num(8)))
    assertEquals(located(result), Vector("T!B1"))
    assertEquals(result.blocked, Vector(at("H1")))
    assert(result.summary.startsWith("T!B1: "), result.summary)
  }

  test("failures list the window's sheet first, then other sheets") {
    val wb = Workbook(
      sheet("S", "A1" -> formula("NOSUCHFN(2)"), "H1" -> cached("T!B1+1", 4))
    ).put(sheet("T", "B1" -> formula("NOSUCHFN(1)"))).put(sheet("A", "A1" -> num(1)))
    assertEquals(located(window(wb, "A1:H1")), Vector("S!A1", "T!B1"))
  }

  test("a dynamic reader on another sheet: that sheet's formulas are computed with it") {
    // T!B1 reads T!C1 through INDIRECT, which the static closure cannot see; T!C1 is stale
    val wb = Workbook(sheet("S", "A1" -> num(5), "H1" -> cached("T!B1*10", 30))).put(
      sheet("T", "B1" -> cached("INDIRECT(\"C1\")+S!A1", 3), "C1" -> cached("S!A1*2", 2))
    )
    val result = window(wb, "H1:H1")
    assertEquals(result.values, Map(at("H1") -> num(150)))
    assertEquals(result.values, recalculated(wb, "H1"))
  }

  // ---- cycles report as they did -----------------------------------------------------------------

  test("a cycle through another sheet fails the window's member and blocks its reader") {
    val wb = Workbook(
      sheet("S", "A1" -> cached("T!A1+1", 1), "B1" -> cached("A1*2", 2), "C1" -> cached("7+1", 0))
    ).put(sheet("T", "A1" -> cached("S!A1+1", 1)))
    val result = window(wb, "A1:C1")
    assertEquals(result.values, Map(at("C1") -> num(8)))
    assertEquals(located(result), Vector("S!A1"))
    assert(result.failures.forall(_.error.message.contains("Circular reference")))
    assertEquals(result.blocked, Vector(at("B1")))
  }

  test("a cycle wholly on another sheet blocks the window's reader of it") {
    val wb = Workbook(sheet("S", "H1" -> cached("T!X1*2", 4), "C1" -> cached("7+1", 0)))
      .put(sheet("T", "X1" -> cached("Y1+1", 2), "Y1" -> cached("X1+1", 1)))
    val result = window(wb, "A1:H1")
    assertEquals(result.values, Map(at("C1") -> num(8)))
    // the members are on the other sheet: none is listed (as before #891), the reader is blocked
    assertEquals(result.failures, Vector.empty)
    assertEquals(result.blocked, Vector(at("H1")))
  }

  test("an iterating book keeps a cross-sheet cycle's converged caches and computes off them") {
    val book = Workbook(sheet("S", "A1" -> cached("T!A1/2+3", 4.5), "B1" -> cached("A1*2", 0)))
      .put(sheet("T", "A1" -> cached("S!A1/2+1", 3.25)))
    val wb =
      book.copy(metadata = book.metadata.copy(calcPr = Some(CalcPr(iterativeCalculation = true))))
    val result = window(wb, "A1:B1")
    assertEquals(result.values, Map(at("B1") -> num(9)))
    assert(result.isClean, result.summary)
  }
