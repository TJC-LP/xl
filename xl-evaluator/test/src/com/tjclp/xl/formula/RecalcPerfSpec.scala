package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError, Evaluator, IterativeCalc, RangeOperand}
import com.tjclp.xl.formula.functions.{EvalContext, FunctionSpecsBase}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.FunSuite

/**
 * GH-346: recalculate() and formula eval on recursive multi-branch cross-sheet chains (the LBO
 * debt-schedule shape) must run in time linear in the cell count, not the path count.
 *
 * Pre-fix, an uncached formula reference was recursively re-evaluated once per PATH (no
 * memoization), and whole-workbook recalculation ordered cells per sheet (not globally), so
 * cross-sheet precedents were always uncached when hit. On this shape — each period's beginning
 * balance read by three formulas, chained backwards period over period — evaluation cost grew ~3^n:
 * at n=24 that is ~10^11 evaluations, hours of CPU. Post-fix both paths complete in milliseconds.
 * The structural pins below are the gates; the suite's 30 s `munitTimeout` is the only backstop for
 * a return to exponential or quadratic work, which no counter can express and which a wall-clock
 * budget equal to that timeout never decided (#843).
 */
class RecalcPerfSpec extends FunSuite:

  private val Periods = 24

  private def num(bd: BigDecimal): CellValue = CellValue.Number(bd)
  private def formula(expr: String): CellValue = CellValue.Formula(expr, None)

  // 0-based (col, row) helpers for the schedule columns; rows are periods (row i = period i+1)
  private def at(col: Int, period: Int): ARef = ARef.from0(col, period - 1)
  private val ColB = 1 // Debt: beginning balance / Forecast: EBITDA
  private val ColD = 3 // Debt: interest
  private val ColF = 5 // Debt: cash available for sweep
  private val ColH = 7 // Debt: sweep
  private val ColJ = 9 // Debt: ending balance

  /**
   * Debt-schedule shape (every formula uncached):
   *   - Forecast!B{i}: EBITDA, 100 growing 10%/period
   *   - Debt!B{i}: beginning balance (period 1: opening 1000; else prior ending =J{i-1})
   *   - Debt!D{i}: interest on beginning balance (=B{i}*0.08)
   *   - Debt!F{i}: cash available (=Forecast!B{i}-D{i}) — cross-sheet
   *   - Debt!H{i}: sweep (=MIN(F{i},B{i}))
   *   - Debt!J{i}: ending balance (=B{i}-H{i})
   *   - Returns!B1: terminal ending + terminal EBITDA — cross-sheet
   */
  private def buildWorkbook(n: Int): Workbook =
    val forecast = (2 to n).foldLeft(
      Sheet(SheetName.unsafe("Forecast")).put(at(ColB, 1), num(BigDecimal(100)))
    ) { (s, i) =>
      s.put(at(ColB, i), formula(s"=B${i - 1}*1.1"))
    }
    val debt = (1 to n).foldLeft(Sheet(SheetName.unsafe("Debt"))) { (s, i) =>
      val withBeginning =
        if i == 1 then s.put(at(ColB, 1), num(BigDecimal(1000)))
        else s.put(at(ColB, i), formula(s"=J${i - 1}"))
      withBeginning
        .put(at(ColD, i), formula(s"=B$i*0.08"))
        .put(at(ColF, i), formula(s"=Forecast!B$i-D$i"))
        .put(at(ColH, i), formula(s"=MIN(F$i,B$i)"))
        .put(at(ColJ, i), formula(s"=B$i-H$i"))
    }
    val returns = Sheet(SheetName.unsafe("Returns"))
      .put(ARef.from0(1, 0), formula(s"=Debt!J$n+Forecast!B$n"))
    Workbook(forecast, debt, returns)

  /** Reference values computed with the same BigDecimal operations the evaluator uses. */
  private case class Terminals(ebitda: BigDecimal, ending: BigDecimal)
  private def expectedTerminals(n: Int): Terminals =
    val init = (BigDecimal(100), BigDecimal(1000))
    val (ebitda, ending) = (1 to n).foldLeft(init) { case ((e, beg), i) =>
      val ebitda = if i == 1 then e else e * BigDecimal("1.1")
      val interest = beg * BigDecimal("0.08")
      val cash = ebitda - interest
      val sweep = cash.min(beg)
      val ending = beg - sweep
      (ebitda, ending)
    }
    // The fold threads (current ebitda, ending balance); ending of period i is beginning of i+1
    Terminals(ebitda, ending)

  private def cached(wb: Workbook, sheetName: String, ref: ARef): Option[CellValue] =
    wb.sheets
      .find(_.name.value == sheetName)
      .flatMap(_.cells.get(ref))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(v), _) => v }

  test("GH-346: whole-workbook recalculate() on a recursive schedule is linear and correct"):
    val wb = buildWorkbook(Periods)
    val result = wb.recalculate()
    assert(result.isClean, s"expected clean, got: ${result.errors.take(5).map(_.render)}")
    val Terminals(ebitda, ending) = expectedTerminals(Periods)
    assertEquals(cached(result.workbook, "Debt", at(ColJ, Periods)), Some(num(ending)))
    assertEquals(cached(result.workbook, "Returns", ARef.from0(1, 0)), Some(num(ending + ebitda)))

  test("GH-346: single-shot eval of a deep uncached cross-sheet cell is memoized, not per-path"):
    val wb = buildWorkbook(Periods)
    val evaluated = wb.evaluateFormula(s"=Debt!J$Periods+0", "Returns")
    assertEquals(evaluated, Right(num(expectedTerminals(Periods).ending)))

  test("GH-346: cross-sheet aggregate over uncached formula cells computes their values"):
    // SUM over ANOTHER sheet's formula cells: with per-sheet ordering the aggregate read the
    // original uncached snapshot (values silently skipped); the global order must guarantee
    // computed values regardless of sheet position in the workbook.
    val b1 = ARef.from0(1, 0)
    val b2 = ARef.from0(1, 1)
    val b3 = ARef.from0(1, 2)
    val data = Sheet(SheetName.unsafe("Data"))
      .put(b1, num(BigDecimal(10)))
      .put(b2, formula("=B1*2"))
      .put(b3, formula("=B2+5"))
    val summary = Sheet(SheetName.unsafe("Summary")).put(b1, formula("=SUM(Data!B1:B3)"))
    // Summary placed FIRST so a sheet-ordered pass cannot accidentally see Data computed
    val result = Workbook(summary, data).recalculate()
    assert(result.isClean, s"expected clean, got: ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, "Summary", b1), Some(num(BigDecimal(55))))

  /**
   * GH-492 tripwire: a long chain of cyclic components (each with its OWN `maxIter` budget under
   * the condensation walk) plus a long acyclic tail. The worst case is `Σ maxIter × |SCCᵢ|`
   * member-evaluations, identical to the pre-GH-492 joint `rounds × |core|`; this catches both a
   * regression to that worst case being hit in practice and any per-component overlay overhead that
   * scales badly with the component count. If it ever bites, the fix is a shared remaining-budget
   * counter threaded through the walk, NOT a new configuration parameter.
   */
  private val Years = 20
  private val TailLength = 300

  private def chainedCycles: Workbook =
    // Columns A..E per year row: A beg (acyclic), {B interest, C pay, E end} cyclic, D fee
    // (acyclic dependent). Year i+1's beg reads year i's end, so the condensation is a 20-link
    // chain of 3-member SCCs separated by acyclic bridge cells.
    val model = (1 to Years).foldLeft(Sheet(SheetName.unsafe("Model"))) { (s, i) =>
      val beg = if i == 1 then num(BigDecimal(1000)) else formula(s"=E${i - 1}")
      s.put(ARef.from0(0, i - 1), beg)
        .put(ARef.from0(1, i - 1), formula(s"=(A$i+E$i)/2*0.08"))
        .put(ARef.from0(2, i - 1), formula(s"=20-B$i"))
        .put(ARef.from0(4, i - 1), formula(s"=A$i-C$i"))
        .put(ARef.from0(3, i - 1), formula(s"=B$i*0.05"))
    }
    val tail = (2 to TailLength).foldLeft(
      Sheet(SheetName.unsafe("Tail")).put(ARef.from0(0, 0), formula(s"=Model!E$Years"))
    )((s, k) => s.put(ARef.from0(0, k - 1), formula(s"=A${k - 1}+1")))
    Workbook(model, tail)

  test("GH-492: a 20-component SCC chain with a long acyclic tail stays inside the budget"):
    val wb = chainedCycles
    val result = wb.recalculate(IterativeCalc(100, BigDecimal("1E-9")))
    assert(result.isClean, s"expected clean, got: ${result.errors.take(5).map(_.render)}")
    assert(result.converged, s"expected convergence, got ${result.unconverged.map(_.render)}")
    assertEquals(result.cycles.size, Years)
    // Every component solved on its own budget, so none of them ran anywhere near maxIter.
    assert(
      result.cycles.forall(_.rounds < 30),
      s"a component burned its budget: ${result.cycles.map(_.rounds)}"
    )

  /**
   * GH-518 tripwire: Kahn's core (`kahnOrder`, shared by `topologicalSort` and
   * `qualifiedTopologicalSort`) must run in time and allocation linear in the graph, not quadratic.
   * The pre-fix formulation appended each node to an immutable-List accumulator (`acc :+ node` — a
   * full copy per node) and rebuilt the pending queue per node (`rest ++ newlyZero`): on this
   * 100k-cell chain that is ~5e9 cons-cell allocations, minutes of GC against a 30s budget, while
   * the linear version finishes in milliseconds.
   */
  private val ChainCells = 100000

  test("GH-518: topologicalSort on a 100k-cell chain is linear, ordered, and inside the budget"):
    val refs = (0 until ChainCells).map(i => ARef.from0(0, i)).toVector
    val dependencies: Map[ARef, Set[ARef]] =
      (1 until ChainCells)
        .map(i => refs(i) -> Set(refs(i - 1)))
        .toMap
        .updated(refs(0), Set.empty[ARef])
    val dependents: Map[ARef, Set[ARef]] =
      (1 until ChainCells).map(i => refs(i - 1) -> Set(refs(i))).toMap
    val graph = DependencyGraph(dependencies, dependents)
    val sorted = DependencyGraph.topologicalSort(graph)
    sorted match
      case Right(order) =>
        assertEquals(order.size, ChainCells)
        assertEquals(order.headOption, Some(refs(0)))
        assertEquals(order.lastOption, Some(refs(ChainCells - 1)))
      case Left(circular) => fail(s"expected acyclic order, got: $circular")

  test("transitive dependent discovery is linear on a 100k-cell chain"):
    val refs = (0 until ChainCells).map(i => ARef.from0(0, i)).toVector
    val first = refs(0)
    val last = refs(ChainCells - 1)
    val dependents: Map[ARef, Set[ARef]] =
      (1 until ChainCells).map(i => refs(i - 1) -> Set(refs(i))).toMap
    val graph = DependencyGraph(Map.empty, dependents)

    val reached = DependencyGraph.transitiveDependents(graph, Set(first))

    assertEquals(reached.size, ChainCells - 1)
    assert(!reached.contains(first))
    assert(reached.contains(last))

  /**
   * GH-537 tripwire: a cycle member that fails every round (a host failure, not an error value)
   * used to pin its component at `maxIter` — `converged` demanded a `Right` from every member, so
   * the loop replayed an identical round 100,000 times. Once a round reproduces the previous one
   * exactly, every further round is a deterministic replay (pinned clock, no fresh randomness): the
   * engine must stop there and report the component as stalled, in a handful of rounds.
   */
  test("GH-537: a permanently failing member stalls its component instead of burning maxIter"):
    val wb = Workbook(
      Sheet(SheetName.unsafe("S"))
        .put(ARef.from0(0, 0), formula("=B1*0.5+10"))
        .put(ARef.from0(1, 0), formula("=A1*0.5+Nowhere!A1"))
    )
    val result = wb.recalculate(IterativeCalc(100000, BigDecimal("1E-9")))
    assert(!result.converged, "a failing member never converges")
    assert(
      result.cycles.forall(c => c.stalled && c.rounds <= 3),
      s"expected an early stall, got ${result.cycles.map(_.render)}"
    )

  /**
   * GH-537 tripwire: `dynamicCells(workbook)` classified every defined name once PER SHEET and
   * parsed its definition on every visit — sheets × names work (34× on a 96k-name field book) for
   * names whose verdict cannot depend on the reading sheet at all. A name with no sheet-scoped
   * variant and a name-free definition is classified once, and a parse memo keyed on the definition
   * text bounds parsing by the number of DISTINCT definitions, not sheets × names.
   *
   * The pin is STRUCTURAL: the pre-fix algorithm ran this fixture in well under a second, so no
   * wall-clock budget can tell the two apart. `dynamicCellsTraced` counts the parses and the name
   * classifications the run performs; they must be linear in the names, not in sheets × names.
   */
  private val NameSheets = 40
  private val StaticNames = 5000

  test(
    "GH-537: dynamicCells on 40 sheets x 5,000 static names parses and classifies each name once"
  ):
    val sheets = (1 to NameSheets).map { s =>
      Sheet(SheetName.unsafe(s"S$s"))
        .put(ARef.from0(0, 0), num(BigDecimal(1)))
        .put(ARef.from0(1, 0), formula("=name1*2+name4999"))
    }.toVector
    val names = (1 to StaticNames).map { i =>
      DefinedName(s"name$i", s"S1!$$A$$$i:$$A$$${i + 9}*1.05+SUM(S1!$$B$$1:$$B$$50)")
    }.toVector
    val base = Workbook(sheets)
    val wb = base.copy(metadata = base.metadata.copy(definedNames = names))
    val trace = DependencyGraph.dynamicCellsTraced(wb)
    assertEquals(trace.cells, Set.empty[DependencyGraph.QualifiedRef])
    assertEquals(DependencyGraph.dynamicCells(wb), trace.cells)
    // One parse per distinct definition and one classification per name — the cell formulas are
    // never parsed, since no name is dynamic and the candidate pre-filter skips them. The
    // sheets × names algorithm did 200,000 of each.
    assertEquals(trace.parses, StaticNames, "parses must not scale with the sheet count")
    assertEquals(trace.classifications, StaticNames, "classifications must not scale with sheets")
    assert(trace.parses < NameSheets * StaticNames / 10, s"${trace.parses} parses: sheets x names")

  // GH-736: #750 already gives finite computed references a populated box and an implicit
  // Empty tail. Pin the actual shared reference reader's work, not a wall-clock estimate.
  private val template = (2 to 201).foldLeft(Sheet(SheetName.unsafe("Template"))) { (s, row) =>
    s.put(ARef.from1(1, row), num(BigDecimal(1)))
  }

  private def templateArray(sheet: Sheet, text: String): ArrayResult =
    val parsed = FormulaParser.parse(text).fold(e => fail(e.toString), identity)
    Evaluator.arrayInstance.eval(parsed, sheet, Clock.system, Some(Workbook(sheet)), None) match
      case Right(array: ArrayResult) => array
      case other => fail(s"expected an array for $text: $other")

  test("GH-736: finite computed-reference reader reads 200 cells at both template sizes") {
    final class MeasuredReader extends FunctionSpecsBase:
      var reads = 0
      override protected def rangeCellReader(
        target: Sheet,
        ctx: EvalContext
      ): ARef => Either[EvalError, CellValue] = at =>
        reads += 1
        Right(target(at).value)
      def values(range: CellRange): Either[EvalError, ArrayResult] =
        val ctx = EvalContext(
          template,
          Clock.system,
          None,
          [A] => (_: TExpr[A]) => Left(EvalError.EvalFailed("unused", None)),
          _ => Left(EvalError.EvalFailed("unused", None))
        )
        boundedReferenceValues(RangeOperand(template, range), ctx)
    Vector(10000, 100000).foreach { last =>
      val reader = new MeasuredReader
      val range = CellRange(ref"A2", ARef.from1(1, last))
      val array = reader.values(range).fold(e => fail(e.toString), identity)
      assertEquals(reader.reads, 200)
      assertEquals((array.rows, array.cols), (last - 1, 1))
      assertEquals(array.retainedBoxCells, 200L)
      assertEquals(array(199, 0), num(BigDecimal(1)))
      assertEquals(array(200, 0), CellValue.Empty)
      assertEquals(array(last - 2, 0), CellValue.Empty)
    }
  }

  test("GH-736: INDEX keeps its declared shape while SUM folds only the populated box") {
    assertEquals(
      template.evaluateFormula("=SUM(INDEX($A$2:$A$100000,0))"),
      Right(num(BigDecimal(200)))
    )
    assertEquals(
      template.evaluateFormula("=ROWS(INDEX($A$2:$A$100000,0))"),
      Right(num(BigDecimal(99999)))
    )
    val array = templateArray(template, "=INDEX($A$2:$A$100000,0)")
    assertEquals((array.rows, array.cols), (99999, 1))
    assertEquals(array.retainedBoxCells, 200L)
    assertEquals(array.tally.size, 201)
  }

  test("GH-736: finite reference arithmetic observes blank padding and full constant alignment") {
    val sheet = Sheet(SheetName.unsafe("Small")).put(ref"A2", num(BigDecimal(2)))
    val array = templateArray(sheet, "=INDEX(A2:A5,0)+{10;20;30;40}")
    assertEquals(array.dense, Vector(12, 20, 30, 40).map(n => Vector(num(BigDecimal(n)))))
    val (spilled, range) = sheet
      .evaluateArrayFormula("=INDEX(A2:A5,0)", ref"C1")
      .fold(e => fail(e.toString), identity)
    assertEquals((range.height, range.width), (4, 1))
    assertEquals(
      range.cells.map(at => spilled(at).value).toVector,
      Vector(num(BigDecimal(2)), CellValue.Empty, CellValue.Empty, CellValue.Empty)
    )
  }

  test("GH-736: XLOOKUP and colon preserve finite widths and blank tails") {
    val sheet = Sheet(SheetName.unsafe("Lookup"))
      .put(ref"A1", num(BigDecimal(7)))
      .put(ref"B1", num(BigDecimal(3)))
    val lookup = templateArray(sheet, "=XLOOKUP(7,A1:A2,B1:M2)")
    assertEquals((lookup.rows, lookup.cols), (1, 12))
    assertEquals(lookup.retainedBoxCells, 1L)
    assertEquals(lookup(0, 11), CellValue.Empty)
    val colon = templateArray(template, "=A2:INDEX(A1:A100000,100000)")
    assertEquals((colon.rows, colon.cols), (99999, 1))
    assertEquals(colon.retainedBoxCells, 200L)
    assertEquals(colon(99998, 0), CellValue.Empty)
  }

  test("GH-736: sparse computed references preserve leading blanks and first errors") {
    val sheet = Sheet(SheetName.unsafe("Errors"))
      .put(ref"A4", CellValue.Error(CellError.Div0))
      .put(ref"B3", CellValue.Error(CellError.NA))
    val array = templateArray(sheet, "=INDEX(A2:B100000,0,0)")
    assertEquals(array(0, 0), CellValue.Empty)
    assertEquals(array(1, 1), CellValue.Error(CellError.NA))
    assertEquals(
      sheet.evaluateFormula("=SUM(INDEX(A2:B100000,0,0))"),
      Right(CellValue.Error(CellError.NA))
    )
    assertEquals(
      Sheet(SheetName.unsafe("Empty")).evaluateFormula("=SUM(INDEX(A2:A100000,0))"),
      Right(num(BigDecimal(0)))
    )
  }

  test("GH-736: recalculation includes late formula precedents inside finite computed ranges") {
    val data = template.put(ref"A50000", formula("=B1*2")).put(ref"B1", num(BigDecimal(3)))
    val out = Sheet(SheetName.unsafe("Out"))
      .put(ref"A1", formula("=SUM(INDEX(Template!A2:A100000,0))"))
    Vector(3, 5).foreach { input =>
      val book = Workbook(out, data.put(ref"B1", num(BigDecimal(input))))
      val result = book.recalculate()
      assert(result.isClean, result.errors.toString)
      assertEquals(cached(result.workbook, "Out", ref"A1"), Some(num(BigDecimal(200 + 2 * input))))
    }
  }
