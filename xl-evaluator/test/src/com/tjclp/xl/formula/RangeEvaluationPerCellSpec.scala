package com.tjclp.xl.formula

import java.time.{LocalDate, LocalDateTime}

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{CellEvalError, RangeEvalResult}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * `evaluateForRangePerCell` — the total, per-cell range evaluation behind `view --eval`: a cell
 * that cannot evaluate is reported with its location instead of failing the whole range, its
 * transitive dependents inside the evaluation closure are blocked (never computed from a stale
 * cache), and cycles fail only when they sit inside the closure.
 */
class RangeEvaluationPerCellSpec extends FunSuite:

  private val name = SheetName.unsafe("Sheet1")
  private def num(n: BigDecimal): CellValue = CellValue.Number(n)
  private def formula(text: String): CellValue = CellValue.Formula(text)
  private def cached(text: String, value: CellValue): CellValue =
    CellValue.Formula(text, Some(value))
  private def range(a1: String): CellRange =
    CellRange.parse(a1).fold(e => fail(s"bad range $a1: $e"), identity)

  /** The Weaver data block: B2:D4 = 5,-3,1 / -2,4,-6 / 7,1,2 (SUM = 9). */
  private val data: Sheet = List[(ARef, CellValue)](
    ref"B2" -> num(5),
    ref"C2" -> num(-3),
    ref"D2" -> num(1),
    ref"B3" -> num(-2),
    ref"C3" -> num(4),
    ref"D3" -> num(-6),
    ref"B4" -> num(7),
    ref"C4" -> num(1),
    ref"D4" -> num(2)
  ).foldLeft(Sheet(name)) { case (s, (at, v)) => s.put(at, v) }

  private def withCells(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(data) { case (s, (at, v)) => s.put(at, v) }

  /** F2 cannot evaluate (a missing sheet); F3 has a stale cache; G2 reads F2; G3 reads F3. */
  private val degraded: Sheet = withCells(
    ref"F2" -> formula("Missing!A1+1"),
    ref"F3" -> cached("B2*2", num(99)),
    ref"F4" -> formula("SUM(B2:D4)"),
    ref"G2" -> cached("F2+1", num(7)),
    ref"G3" -> formula("F3+1")
  )

  test("a clean sheet evaluates exactly what the fail-fast evaluateForRange evaluates") {
    val clean = withCells(
      ref"F3" -> cached("B2*2", num(99)),
      ref"F4" -> formula("SUM(B2:D4)"),
      ref"G3" -> formula("F3+1"),
      ref"K1" -> formula("B4*10"),
      ref"G4" -> formula("K1+1")
    )
    val window = range("A1:G4")
    val perCell = clean.evaluateForRangePerCell(window, Clock.system, None)
    assert(perCell.isClean, perCell.summary)
    assertEquals(Right(perCell.values), clean.evaluateForRange(window))
    assertEquals(perCell.values.get(ref"F3"), Some(num(10)))
    assertEquals(perCell.values.get(ref"G4"), Some(num(71)))
    // K1 is a precedent outside the window: evaluated, but not reported
    assertEquals(perCell.values.get(ref"K1"), None)
    assertEquals(perCell.failures, Vector.empty)
    assertEquals(perCell.blocked, Vector.empty)
  }

  test("one failing cell leaves the others live, and is reported with its location") {
    val result = degraded.evaluateForRangePerCell(range("A1:G4"), Clock.system, None)
    assertEquals(result.values.get(ref"F3"), Some(num(10)))
    assertEquals(result.values.get(ref"F4"), Some(num(9)))
    assertEquals(result.values.get(ref"G3"), Some(num(11)), "G3 reads the live F3, not 99")
    assertEquals(result.failures.map(f => (f.sheet, f.ref)), Vector((name, ref"F2")))
    val message = result.failures.headOption.map(_.error.message).getOrElse("")
    assert(message.contains("Missing!A1+1"), message)
    assert(!result.isClean)
  }

  test("a failing cell blocks its dependents, which never read its stale cache") {
    val result = degraded.evaluateForRangePerCell(range("A1:G4"), Clock.system, None)
    assertEquals(result.blocked, Vector(ref"G2"))
    assertEquals(result.values.get(ref"G2"), None)
    assertEquals(result.values.get(ref"F2"), None)
  }

  test("the one-argument overload uses the system clock and no workbook") {
    assertEquals(
      degraded.evaluateForRangePerCell(range("A1:G4")),
      degraded.evaluateForRangePerCell(range("A1:G4"), Clock.system, None)
    )
  }

  test("the workbook context resolves cross-sheet references per cell") {
    val other = Sheet(SheetName.unsafe("Other")).put(ref"A1", num(41))
    val sheet = withCells(ref"F2" -> formula("Other!A1+1"), ref"F3" -> formula("Nope!A1"))
    val wb = Workbook(sheet).put(other)
    val result = sheet.evaluateForRangePerCell(range("F2:F3"), Clock.system, Some(wb))
    assertEquals(result.values, Map(ref"F2" -> num(42)))
    assertEquals(result.failures.map(_.ref), Vector(ref"F3"))
  }

  test("a cycle inside the window fails its members and blocks their dependents") {
    val sheet = withCells(
      ref"H2" -> formula("H3+1"),
      ref"H3" -> formula("H2+1"),
      ref"H4" -> formula("H2*2"),
      ref"F2" -> formula("B2+1")
    )
    val result = sheet.evaluateForRangePerCell(range("A1:H4"), Clock.system, None)
    assertEquals(result.failures.map(_.ref), Vector(ref"H2", ref"H3"))
    assert(result.failures.forall(_.error.message.contains("Circular reference")))
    assertEquals(result.blocked, Vector(ref"H4"))
    assertEquals(result.values, Map(ref"F2" -> num(6)))
  }

  test("a cycle outside the window's closure no longer fails the view") {
    val sheet = withCells(
      ref"J2" -> formula("J3+1"),
      ref"J3" -> formula("J2+1"),
      ref"F2" -> formula("B2+1")
    )
    val result = sheet.evaluateForRangePerCell(range("A1:G4"), Clock.system, None)
    assert(result.isClean, result.summary)
    assertEquals(result.values, Map(ref"F2" -> num(6)))
    // the fail-fast method keeps its whole-sheet cycle check
    assert(sheet.evaluateForRange(range("A1:G4")).isLeft)
  }

  test("failures and blocked cells are listed in row-major order") {
    val sheet = withCells(
      ref"G2" -> formula("Nope!A1"),
      ref"F3" -> formula("Nope!A2"),
      ref"B5" -> formula("Nope!A3"),
      ref"C6" -> formula("B5+1"),
      ref"A6" -> formula("G2+1"),
      ref"H3" -> formula("F3+1")
    )
    val result = sheet.evaluateForRangePerCell(range("A1:H6"), Clock.system, None)
    assertEquals(result.failures.map(_.ref), Vector(ref"G2", ref"F3", ref"B5"))
    assertEquals(result.blocked, Vector(ref"H3", ref"A6", ref"C6"))
  }

  test("summary: first three failures, the rest counted, then the blocked dependents") {
    val sheet = withCells(
      ref"F2" -> formula("Nope!A1"),
      ref"F3" -> formula("Nope!A2"),
      ref"F4" -> formula("Nope!A3"),
      ref"F5" -> formula("Nope!A4"),
      ref"G2" -> formula("F2+1"),
      ref"G3" -> formula("F3+1")
    )
    val result = sheet.evaluateForRangePerCell(range("A1:G5"), Clock.system, None)
    val shown = result.failures.take(3).map(_.render).mkString("; ")
    assertEquals(
      result.summary,
      s"$shown; … and 1 more; 2 dependent formulas not evaluated (G2, G3)"
    )
    assert(result.summary.startsWith("Sheet1!F2: "), result.summary)
    val single = degraded.evaluateForRangePerCell(range("A1:G4"), Clock.system, None)
    assertEquals(
      single.summary,
      s"${single.failures.map(_.render).mkString}; 1 dependent formula not evaluated (G2)"
    )
  }

  test("summary of a clean evaluation counts the window's formulas") {
    val clean = withCells(ref"F3" -> formula("B2*2"), ref"F4" -> formula("SUM(B2:D4)"))
    assertEquals(clean.evaluateForRangePerCell(range("A1:G4")).summary, "Evaluated 2 formulas")
    assertEquals(
      clean.evaluateForRangePerCell(range("F3:F3")).summary,
      "Evaluated 1 formula"
    )
    assertEquals(
      data.evaluateForRangePerCell(range("A1:G4")),
      RangeEvalResult(Map.empty, Vector.empty, Vector.empty)
    )
  }

  test("an internal defect in one cell is contained to that cell") {
    val boom = new Clock:
      def today(): LocalDate = throw new IllegalStateException("clock down")
      def now(): LocalDateTime = throw new IllegalStateException("clock down")
    val sheet = withCells(ref"F2" -> formula("TODAY()+1"), ref"F3" -> formula("B2*2"))
    val result = sheet.evaluateForRangePerCell(range("A1:G4"), boom, None)
    assertEquals(result.values, Map(ref"F3" -> num(10)))
    assertEquals(result.failures.map(_.ref), Vector(ref"F2"))
    val message = result.failures.headOption.map(_.error.message).getOrElse("")
    assert(message.contains("internal evaluator defect"), message)
  }

  test("an Excel error value is a value, not a failure") {
    val sheet = withCells(ref"F2" -> formula("1/0"), ref"F3" -> formula("F2+1"))
    val result = sheet.evaluateForRangePerCell(range("F2:F3"))
    assert(result.isClean, result.summary)
    assertEquals(result.values.get(ref"F2"), Some(CellValue.Error(CellError.Div0)))
    assertEquals(result.values.get(ref"F3"), Some(CellValue.Error(CellError.Div0)))
  }

  test("dynamic references widen the evaluation to every formula, reporting the window's only") {
    val sheet = withCells(
      ref"K1" -> formula("B2*100"),
      ref"F2" -> formula("INDIRECT(\"K1\")+1"),
      ref"K2" -> formula("Nope!A1"),
      ref"Z7" -> formula("Z8+1"),
      ref"Z8" -> formula("Z7+1"),
      ref"Z9" -> formula("Z7*2")
    )
    val result = sheet.evaluateForRangePerCell(range("F2:F2"))
    assertEquals(result.values, Map(ref"F2" -> num(501)))
    // K2's failure and the Z7/Z8 cycle (with its dependent Z9) are evaluated — the INDIRECT makes
    // every formula a target — but none feeds the window, so none is reported
    assert(result.isClean, result.summary)
  }

  test("a window formula reading an outside failure through INDIRECT reports the failure itself") {
    val sheet = withCells(
      ref"F2" -> formula("INDIRECT(\"K2\")+1"),
      ref"K2" -> cached("Nope!A1", num(1)),
      ref"Z7" -> formula("Z8+1"),
      ref"Z8" -> formula("Z7+1")
    )
    val result = sheet.evaluateForRangePerCell(range("F2:F2"))
    assertEquals(result.failures.map(_.ref), Vector(ref"F2"))
    assertEquals(result.values, Map.empty[ARef, CellValue], "F2 must not read K2's stale cache")
    assertEquals(result.blocked, Vector.empty[ARef])
  }

  test("a dynamic reader of a failed cell re-derives it and fails, never reading its stale cache") {
    // K1 reads F2 through INDIRECT, which the static graph cannot see, so K1 is not blocked by
    // F2's failure: only the stripped cache makes K1 re-derive F2 instead of computing 7 + 1.
    val sheet = withCells(
      ref"F2" -> cached("Missing!A1", num(7)),
      ref"K1" -> cached("INDIRECT(\"F2\")+1", num(8)),
      ref"K2" -> cached("F2+100", num(107))
    )
    val result = sheet.evaluateForRangePerCell(range("F1:K2"))
    assertEquals(result.failures.map(_.ref), Vector(ref"K1", ref"F2"))
    assertEquals(result.blocked, Vector(ref"K2"))
    assertEquals(result.values.get(ref"K1"), None, "K1 must not be the stale 7 + 1")
    assertEquals(result.values, Map.empty[ARef, CellValue])
  }

  test("a dynamic reader of a cycle member fails, never reading the member's stale cache") {
    val sheet = withCells(
      ref"H2" -> cached("H3+1", num(3)),
      ref"H3" -> cached("H2+1", num(4)),
      ref"J1" -> cached("INDIRECT(\"H2\")*1", num(3)),
      ref"F2" -> formula("B2+1")
    )
    val result = sheet.evaluateForRangePerCell(range("F1:J3"))
    assertEquals(result.failures.map(_.ref), Vector(ref"J1", ref"H2", ref"H3"))
    assertEquals(result.values.get(ref"J1"), None, "J1 must not be H2's stale cache")
    assertEquals(result.values, Map(ref"F2" -> num(6)))
  }

  test("CellEvalError locations render sheet-qualified") {
    val result = degraded.evaluateForRangePerCell(range("F2:F2"))
    val first: Option[CellEvalError] = result.failures.headOption
    assert(first.exists(_.render.startsWith("Sheet1!F2: ")), first.toString)
  }
