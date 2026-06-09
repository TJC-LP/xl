package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * Tests for Workbook.recalculate (0.11.0): total whole-workbook recalculation with per-cell
 * error reporting, cycle isolation, and automatic cross-sheet context.
 *
 * Also covers the latent bugs in the pre-0.11.0 withCachedFormulas this replaces: one failing
 * formula uncached the ENTIRE sheet, and cross-sheet formulas always failed (no workbook
 * context), silently uncaching their sheet.
 */
class RecalcSpec extends FunSuite:

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def formula(expr: String): CellValue = CellValue.Formula(expr, None)
  private val a1 = ARef.from0(0, 0)
  private val a2 = ARef.from0(0, 1)
  private val a3 = ARef.from0(0, 2)
  private val a4 = ARef.from0(0, 3)

  private def cached(wb: Workbook, sheetName: String, ref: ARef): Option[CellValue] =
    wb.sheets
      .find(_.name.value == sheetName)
      .flatMap(_.cells.get(ref))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(v)) => v }

  test("clean workbook: every formula evaluates, caches, and isClean holds"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, num(10))
      .put(a2, formula("=A1*2"))
      .put(a3, formula("=A2+5"))
    val result = Workbook(sheet).recalculate()
    assert(result.isClean)
    assertEquals(cached(result.workbook, "S", a2), Some(num(20)))
    assertEquals(cached(result.workbook, "S", a3), Some(num(25)))
    assertEquals(result.toEither.map(_.sheets.size), Right(1))

  test("sibling survival: one failing formula doesn't uncache the rest of the sheet"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, num(10))
      .put(a2, formula("=A1*2"))
      .put(a3, formula("=NOSUCHFN(A1)"))
    val result = Workbook(sheet).recalculate()
    assertEquals(cached(result.workbook, "S", a2), Some(num(20)), "sibling must stay cached")
    assertEquals(cached(result.workbook, "S", a3), None, "failed cell must stay uncached")
    assertEquals(result.errors.map(_.ref), Vector(a3))
    assert(!result.isClean)

  test("cycle isolation: participants reported circular, acyclic remainder still evaluates"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, formula("=A2+1")) // A1 -> A2
      .put(a2, formula("=A1+1")) // A2 -> A1 (cycle)
      .put(a3, num(7))
      .put(a4, formula("=A3*3")) // independent of the cycle
    val result = Workbook(sheet).recalculate()
    assertEquals(cached(result.workbook, "S", a4), Some(num(21)), "acyclic cell must evaluate")
    assertEquals(cached(result.workbook, "S", a1), None)
    assertEquals(cached(result.workbook, "S", a2), None)
    assertEquals(result.errors.map(_.ref).toSet, Set(a1, a2))
    assert(result.errors.forall(_.error.message.contains("Circular")))

  test("dependents of a cycle are reported as blocked, not evaluated"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, formula("=A2+1"))
      .put(a2, formula("=A1+1")) // cycle A1<->A2
      .put(a3, formula("=A1*2")) // depends on the cycle
    val result = Workbook(sheet).recalculate()
    assertEquals(cached(result.workbook, "S", a3), None)
    val byRef = result.errors.map(e => e.ref -> e.error.message).toMap
    assert(byRef(a3).contains("Blocked"))
    assertEquals(result.errors.size, 3)

  test("self-referencing formula is circular"):
    val sheet = Sheet(SheetName.unsafe("S")).put(a1, formula("=A1+1"))
    val result = Workbook(sheet).recalculate()
    assertEquals(result.errors.map(_.ref), Vector(a1))
    assert(result.errors.forall(_.error.message.contains("Circular")))

  test("cross-sheet formulas recalculate and cache (latent pre-0.11.0 bug)"):
    val data = Sheet(SheetName.unsafe("Data")).put(a1, num(10)).put(a2, num(20))
    val summary = Sheet(SheetName.unsafe("Summary"))
      .put(a1, formula("=SUM(Data!A1:A2)"))
      .put(a2, formula("=A1*2")) // intra-sheet, downstream of the cross-sheet result
    val result = Workbook(data, summary).recalculate()
    assert(result.isClean, s"expected clean, got: ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, "Summary", a1), Some(num(30)))
    assertEquals(cached(result.workbook, "Summary", a2), Some(num(60)))

  test("withCachedFormulas ≡ recalculate(_).workbook"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(a1, num(10))
      .put(a2, formula("=A1*2"))
      .put(a3, formula("=NOSUCHFN(A1)"))
    val wb = Workbook(sheet)
    val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 6, 9))
    assertEquals(wb.withCachedFormulas(clock), wb.recalculate(clock).workbook)

  test("evaluated map exposes computed values per sheet"):
    val sheet = Sheet(SheetName.unsafe("S")).put(a1, num(2)).put(a2, formula("=A1+3"))
    val result = Workbook(sheet).recalculate()
    assertEquals(
      result.evaluated.get(SheetName.unsafe("S")).flatMap(_.get(a2)),
      Some(num(5))
    )

  test("CellEvalError.render is location-qualified"):
    val sheet = Sheet(SheetName.unsafe("S")).put(a1, formula("=A1"))
    val result = Workbook(sheet).recalculate()
    assert(result.errors.headOption.exists(_.render.startsWith("S!A1:")))

  test("cross-sheet cycle stays total: caught by depth guard, reported per cell (not Tarjan)"):
    // Cycle spanning sheets is invisible to the per-sheet Tarjan pass — pin the graceful path:
    // both cells error (recursion guard), nothing throws, the acyclic remainder still evaluates.
    val s1 = Sheet(SheetName.unsafe("S1"))
      .put(a1, formula("=S2!A1+1"))
      .put(a2, num(5))
      .put(a3, formula("=A2*2")) // acyclic, must survive
    val s2 = Sheet(SheetName.unsafe("S2")).put(a1, formula("=S1!A1+1"))
    val result = Workbook(s1, s2).recalculate()
    assertEquals(cached(result.workbook, "S1", a3), Some(num(10)))
    val errorRefs = result.errors.map(e => (e.sheet.value, e.ref)).toSet
    assert(errorRefs.contains(("S1", a1)), s"S1!A1 should error, got: ${result.errors.map(_.render)}")
    assert(errorRefs.contains(("S2", a1)), s"S2!A1 should error, got: ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, "S1", a1), None)
    assertEquals(cached(result.workbook, "S2", a1), None)

  test("stack safety: 100k-deep dependency chain neither overflows nor reports cycles"):
    // Graph built directly (no parsing) — pins the iterative Tarjan engine. The recursive
    // strongConnect this replaces overflowed around ~10k frames, violating recalculate's
    // totality promise on deep-but-legal generated models.
    val n = 100000
    val deps = (0 until n).map { i =>
      val node = ARef.from0(0, i)
      val next = if i + 1 < n then Set(ARef.from0(0, i + 1)) else Set.empty[ARef]
      node -> next
    }.toMap
    val graph = com.tjclp.xl.formula.graph.DependencyGraph(deps, Map.empty)
    assertEquals(com.tjclp.xl.formula.graph.DependencyGraph.cyclicNodes(graph), Set.empty[ARef])
    assert(com.tjclp.xl.formula.graph.DependencyGraph.detectCycles(graph).isRight)

  test("stack safety: 100k-node cycle is fully collected without overflow"):
    val n = 100000
    val deps = (0 until n).map { i =>
      ARef.from0(0, i) -> Set(ARef.from0(0, (i + 1) % n)) // last points back to first
    }.toMap
    val graph = com.tjclp.xl.formula.graph.DependencyGraph(deps, Map.empty)
    assertEquals(com.tjclp.xl.formula.graph.DependencyGraph.cyclicNodes(graph).size, n)
    assert(com.tjclp.xl.formula.graph.DependencyGraph.detectCycles(graph).isLeft)

  test("recalculate end-to-end on a 3k-deep formula chain stays total and clean"):
    val n = 3000
    val sheet = (1 until n).foldLeft(Sheet(SheetName.unsafe("Deep")).put(a1, num(1))) {
      (s, i) => s.put(ARef.from0(0, i), formula(s"=A$i+1"))
    }
    val result = Workbook(sheet).recalculate()
    assert(result.isClean, s"expected clean, got ${result.errors.take(3).map(_.render)}")
    assertEquals(
      result.evaluated.values.headOption.flatMap(_.get(ARef.from0(0, n - 1))),
      Some(num(n))
    )
