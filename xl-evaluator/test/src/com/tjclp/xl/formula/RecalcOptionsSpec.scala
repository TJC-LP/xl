package com.tjclp.xl.formula

import java.time.LocalDate

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.FormulaKind
import com.tjclp.xl.formula.eval.{
  DependentRecalculation,
  IterativeCalc,
  IterativeMode,
  RecalcOptions,
  RecalcResult
}
import com.tjclp.xl.workbooks.CalcPr
import munit.FunSuite

/**
 * ADR-017 §2.8: ONE recalculation primitive driven by an options record. `recalculate(options)`
 * must reproduce every existing overload it subsumes (the @targetName lattice stays as forwarders),
 * `recalculateAfterEdit` wraps the private after-edit seam, `recalculateUncached` computes only
 * cache-less formulas, and `RecalcResult.summary` prints byte-for-byte what the CLI prints.
 */
class RecalcOptionsSpec extends FunSuite:

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def formula(text: String, cached: Option[CellValue] = None): CellValue =
    CellValue.Formula(text, cached)
  private val S = SheetName.unsafe("S")
  private val T = SheetName.unsafe("T")

  private def assertSameResult(actual: RecalcResult, expected: RecalcResult): Unit =
    assertEquals(actual.errors, expected.errors)
    assertEquals(actual.evaluated, expected.evaluated)
    assertEquals(actual.workbook, expected.workbook)
    assertEquals(actual.converged, expected.converged)
    assertEquals(actual.iterationsUsed, expected.iterationsUsed)
    assertEquals(actual.cycles, expected.cycles)

  private def cellValue(wb: Workbook, sheet: String, ref: ARef): CellValue =
    wb.sheets.find(_.name.value == sheet).getOrElse(fail(s"missing sheet $sheet"))(ref).value

  private def acyclic: Workbook =
    Workbook(
      Sheet("S").put(ref"A1", num(2)).put(ref"A2", formula("A1*2")).put(ref"A3", formula("A2+1")),
      Sheet("T").put(ref"A1", formula("S!A3*10"))
    )

  /** The canonical convergent pair (IterativeRecalcSpec): B1 = 100/0.95, A1 = B1/10. */
  private def circular: Workbook =
    Workbook(Sheet("S").put(ref"A1", formula("B1*0.1")).put(ref"B1", formula("100+A1/2")))

  private val declared =
    CalcPr(
      iterativeCalculation = true,
      maxIterations = Some(100),
      maxChange = Some(BigDecimal("0.001"))
    )

  test("RecalcOptions.default carries the documented defaults") {
    // Clock.system / Rng.system are capability instances compared by identity, so the record is
    // pinned field by field rather than as a whole.
    for opts <- Vector(RecalcOptions.default, RecalcOptions()) do
      assertEquals(opts.iterative, IterativeMode.FromCalcPr: IterativeMode)
      assertEquals(opts.parallelism, 1)
      assertEquals(opts.seedTables, false)
  }

  test("recalculate(RecalcOptions()) ≡ recalculate() on an acyclic book") {
    assertSameResult(acyclic.recalculate(RecalcOptions()), acyclic.recalculate())
    assertSameResult(acyclic.recalculate(RecalcOptions.default), acyclic.recalculate())
  }

  test(
    "FromCalcPr on an iterate-declared book ≡ recalculate(Clock.system, IterativeCalc.fromCalcPr(cp))"
  ) {
    val wb = circular.withCalcPr(declared)
    val viaOptions = wb.recalculate(RecalcOptions())
    assertSameResult(viaOptions, wb.recalculate(Clock.system, IterativeCalc.fromCalcPr(declared)))
    assert(viaOptions.isClean, viaOptions.errors.map(_.render).mkString("; "))
    assert(viaOptions.iterationsUsed > 0)
  }

  test(
    "FromCalcPr on a book WITHOUT the declaration is the default cycle-isolating recalculation"
  ) {
    val viaOptions = circular.recalculate(RecalcOptions())
    assertSameResult(viaOptions, circular.recalculate())
    assert(!viaOptions.isClean)
  }

  test("IterativeMode.Off reports a declared cycle as an error") {
    val result =
      circular.withCalcPr(declared).recalculate(RecalcOptions(iterative = IterativeMode.Off))
    assert(result.errors.exists(_.error.message.contains("Circular reference")), result.errors)
    assertSameResult(result, circular.withCalcPr(declared).recalculate())
  }

  test("IterativeMode.Force fixpoints an undeclared cycle") {
    val force = IterativeMode.Force(IterativeCalc(100, BigDecimal("0.001")))
    val result = circular.recalculate(RecalcOptions(iterative = force))
    assert(result.isClean && result.converged && result.iterationsUsed > 0)
    assertSameResult(result, circular.recalculate(IterativeCalc(100, BigDecimal("0.001"))))
  }

  test("parallelism = 4 ≡ sequential (wide independent region, above the wave cutoff)") {
    val wide = (0 until 40).foldLeft(Sheet("S").put(ref"A1", num(3))) { (s, i) =>
      s.put(ARef.from0(1, i), formula(s"A1*${i + 1}"))
    }
    val wb = Workbook(wide)
    assertSameResult(
      wb.recalculate(RecalcOptions(parallelism = 4)),
      wb.recalculate(RecalcOptions())
    )
    assertSameResult(wb.recalculate(RecalcOptions(parallelism = 4)), wb.recalculateParallel(4))
    assertSameResult(wb.recalculate(RecalcOptions(parallelism = 4)), wb.recalculate())
  }

  test("a declared iteration wins over parallelism (recalcHonoringCalcPr's rule)") {
    val wb = circular.withCalcPr(declared)
    assertSameResult(
      wb.recalculate(RecalcOptions(parallelism = 4)),
      wb.recalculate(Clock.system, IterativeCalc.fromCalcPr(declared))
    )
  }

  test("recalculateUncached leaves cached cells byte-identical and fills uncached ones") {
    // 999 is deliberately WRONG: only a recalculation of B1 could turn it into 4, so the
    // assertion distinguishes "cache preserved" from "cache recomputed to the same value".
    val stale = Some(num(999))
    val wb = Workbook(
      Sheet("S")
        .put(ref"A1", num(2))
        .put(ref"B1", formula("A1*2", stale))
        .put(ref"C1", formula("B1+1"))
        .put(ref"D1", formula("A1+1", Some(num(3))))
        .put(ref"E1", formula("C1*2")),
      Sheet("T").put(ref"A1", formula("S!C1+S!B1"))
    )
    val result = wb.recalculateUncached(RecalcOptions())
    assert(result.isClean, result.errors.map(_.render).mkString("; "))
    // GH-468 doctrine: a cached cell is never rewritten, even when its cache is wrong.
    assertEquals(cellValue(result.workbook, "S", ref"B1"), formula("A1*2", stale))
    assertEquals(cellValue(result.workbook, "S", ref"D1"), formula("A1+1", Some(num(3))))
    // Uncached cells compute from their inputs' caches AS THEY ARE (999 + 1), in dependency order.
    assertEquals(cellValue(result.workbook, "S", ref"C1"), formula("B1+1", Some(num(1000))))
    assertEquals(cellValue(result.workbook, "S", ref"E1"), formula("C1*2", Some(num(2000))))
    assertEquals(cellValue(result.workbook, "T", ref"A1"), formula("S!C1+S!B1", Some(num(1999))))
    assertEquals(result.evaluated(S), Map(ref"C1" -> num(1000), ref"E1" -> num(2000)))
    assertEquals(result.evaluated(T), Map(ref"A1" -> num(1999)))
  }

  test("recalculateUncached on a fully cached book returns it unchanged") {
    val wb = Workbook(
      Sheet("S").put(ref"A1", num(2)).put(ref"B1", formula("A1*2", Some(num(999))))
    )
    val result = wb.recalculateUncached(RecalcOptions())
    assert(result.isClean)
    assertEquals(result.workbook, wb)
    assertEquals(result.evaluated(S), Map.empty[ARef, CellValue])
  }

  test("recalculateUncached reports an uncached cycle member instead of guessing") {
    val wb = Workbook(
      Sheet("S")
        .put(ref"A1", formula("B1*0.1"))
        .put(ref"B1", formula("100+A1/2"))
        .put(ref"C1", formula("1+1"))
    )
    val result = wb.recalculateUncached(RecalcOptions())
    assert(result.errors.exists(e => e.ref == ref"A1" && e.error.message.contains("Circular")))
    assertEquals(cellValue(result.workbook, "S", ref"C1"), formula("1+1", Some(num(2))))
  }

  test(
    "recalculateAfterEdit(sheet, refs, opts) ≡ DependentRecalculation.recalculateAfterEdit(wb, sheet, refs, opts.clock)"
  ) {
    val wb = Workbook(
      Sheet("S").put(ref"A1", num(2)).put(ref"B1", formula("A1*2", Some(num(2)))),
      Sheet("T")
        .put(ref"A1", formula("S!B1+1", Some(num(3))))
        .put(ref"A2", formula("TODAY()", Some(num(1))))
    )
    val opts = RecalcOptions(clock = Clock.fixedDate(LocalDate.of(2026, 1, 1)))
    val viaOptions = wb.recalculateAfterEdit(S, Set(ref"A1"), opts)
    assertSameResult(
      viaOptions,
      DependentRecalculation.recalculateAfterEdit(wb, S, Set(ref"A1"), opts.clock)
    )
    // The cone: B1 and T!A1 refresh; the unrelated volatile T!A2 keeps its cache.
    assertEquals(cellValue(viaOptions.workbook, "S", ref"B1"), formula("A1*2", Some(num(4))))
    assertEquals(cellValue(viaOptions.workbook, "T", ref"A1"), formula("S!B1+1", Some(num(5))))
    assertEquals(cellValue(viaOptions.workbook, "T", ref"A2"), formula("TODAY()", Some(num(1))))
  }

  test("recalculateAfterEdit on an iterate-declared book falls back to the full recalculation") {
    val wb = circular.withCalcPr(declared).put(Sheet("T").put(ref"A1", num(1)))
    assertSameResult(
      wb.recalculateAfterEdit(T, Set(ref"A1"), RecalcOptions()),
      wb.recalculate(RecalcOptions())
    )
  }

  test("summary: the clean forms the CLI prints") {
    val one = Workbook(Sheet("S").put(ref"A1", num(5)).put(ref"C1", formula("1/0"))).recalculate()
    assertEquals(one.summary, "Recalculated 1 formula (1 error value)")
    val two = Workbook(
      Sheet("S").put(ref"A1", num(5)).put(ref"B1", formula("A1*2")).put(ref"C1", formula("B1+1"))
    ).recalculate()
    assertEquals(two.summary, "Recalculated 2 formulas")
    val twoErrors = Workbook(
      Sheet("S").put(ref"A1", formula("1/0")).put(ref"B1", formula("NA()"))
    ).recalculate()
    assertEquals(twoErrors.summary, "Recalculated 2 formulas (2 error values)")
  }

  test("summary: host failures list the first three rendered errors, then an ellipsis") {
    val failing = Workbook(
      Sheet("S")
        .put(ref"A1", num(5))
        .put(ref"B1", formula("A1*2"))
        .put(ref"C1", formula("NOSUCHFN(A1)"))
    ).recalculate()
    val rendered = failing.errors.headOption.map(_.render).getOrElse(fail("expected one error"))
    assertEquals(failing.summary, s"Recalculated 1 formula; 1 error ($rendered)")
    val many = (0 until 5)
      .foldLeft(Sheet("S")) { (s, i) => s.put(ARef.from0(0, i), formula("NOSUCHFN(1)")) }
    val manyResult = Workbook(many).recalculate()
    val shown = manyResult.errors.take(3).map(_.render).mkString("; ")
    assertEquals(manyResult.summary, s"Recalculated 0 formulas; 5 errors ($shown; ...)")
  }

  test("summary: the iterative verdicts") {
    val force = IterativeMode.Force(IterativeCalc(100, BigDecimal("0.001")))
    val converged = circular.recalculate(RecalcOptions(iterative = force))
    assertEquals(
      converged.summary,
      s"Recalculated 2 formulas; converged in ${converged.iterationsUsed} iterative round(s)"
    )
    val oscillator = Workbook(Sheet("S").put(ref"A1", formula("1-B1")).put(ref"B1", formula("A1")))
    val exhausted = oscillator.recalculate(
      RecalcOptions(iterative = IterativeMode.Force(IterativeCalc(10, BigDecimal("0.001"))))
    )
    assertEquals(
      exhausted.summary,
      "Recalculated 2 formulas; WARNING: iterative calculation exhausted 10 round(s) without converging (last values kept)"
    )
  }

  test("seedTables = true seeds data-table interiors after the recalculation") {
    val base = Sheet("DT")
      .put(ref"B1", num(10))
      .put(ref"B2", num(20))
      .put(ref"C4", formula("B1*B2"))
      .put(ref"D4", num(1))
      .put(ref"E4", num(2))
      .put(ref"C5", num(100))
      .put(ref"C6", num(200))
    val authored = base.dataTable(ref"D5:E6", ref"B1", ref"B2").fold(e => fail(e.message), identity)
    val wb = Workbook(authored)
    val seeded = wb.recalculate(RecalcOptions(seedTables = true))
    assert(seeded.isClean, seeded.errors.map(_.render).mkString("; "))
    cellValue(seeded.workbook, "DT", ref"D5") match
      case CellValue.Formula(_, cached, _: FormulaKind.DataTable) =>
        assertEquals(cached, Some(num(100))) // B1 <- 1, B2 <- 100
      case other => fail(s"expected a seeded record, got $other")
    assertEquals(cellValue(seeded.workbook, "DT", ref"E6"), num(400))
    assertEquals(cellValue(seeded.workbook, "DT", ref"C4"), formula("B1*B2", Some(num(200))))
    // ≡ recalculate() then seedDataTables()
    val manual = wb.recalculate().workbook.seedDataTables().fold(e => fail(e.message), identity)
    assertEquals(seeded.workbook, manual)
    // Off by default: the record stays uncached.
    cellValue(wb.recalculate(RecalcOptions()).workbook, "DT", ref"D5") match
      case CellValue.Formula(_, cached, _: FormulaKind.DataTable) => assertEquals(cached, None)
      case other => fail(s"expected an unseeded record, got $other")
  }
