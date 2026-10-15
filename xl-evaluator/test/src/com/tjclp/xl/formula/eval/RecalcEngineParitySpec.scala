package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.formula.graph.CompactGraphGens
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * #767: `RecalcEngine.Compact` and `RecalcEngine.Legacy` give the same `RecalcResult` on every path
 * — sequential (with a seeded RAND, whose draws follow evaluation order), parallel,
 * `recalculateUncached`, and iterative under both schemes — and the order is the canonical one:
 * error vectors in grid order, parallel equal to sequential.
 */
class RecalcEngineParitySpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(80)

  private val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 2))

  private def same(compact: RecalcResult, legacy: RecalcResult, label: String): Prop =
    Prop.all(
      Prop(
        compact.errors == legacy.errors
      ) :| s"$label errors\n${compact.errors}\n${legacy.errors}",
      Prop(compact.evaluated == legacy.evaluated) :| s"$label evaluated",
      Prop(compact.workbook == legacy.workbook) :| s"$label workbook",
      Prop(compact.errorReasons == legacy.errorReasons) :| s"$label reasons",
      Prop(compact.converged == legacy.converged) :| s"$label converged",
      Prop(compact.iterationsUsed == legacy.iterationsUsed) :| s"$label rounds",
      Prop(compact.cycles == legacy.cycles) :| s"$label cycles"
    )

  /** The random graph books plus a few seeded RAND readers, so evaluation order is observable. */
  private val genBook: Gen[Workbook] =
    for
      wb <- CompactGraphGens.genWorkbook
      rands <- Gen.listOfN(4, Gen.zip(Gen.choose(0, 3), Gen.choose(1, 10)))
    yield rands.foldLeft(wb) { case (book, (c, r)) =>
      val name = SheetName.unsafe("Mid")
      book.sheets.find(_.name == name).fold(book) { sheet =>
        val ref = ARef.from0(c + 4, r - 1)
        book.put(sheet.put(ref, CellValue.Formula(s"RAND()+A$r")))
      }
    }

  private def run(engine: RecalcEngine, options: RecalcOptions)(wb: Workbook): RecalcResult =
    recalculateCounted(wb, options, RecalcCounters.disabled, engine)

  property("sequential with a seeded Rng: Compact equals Legacy") {
    Prop.forAll(genBook) { wb =>
      def opts = RecalcOptions(clock = clock, rng = Rng.seeded(7L), iterative = IterativeMode.Off)
      same(run(RecalcEngine.Compact, opts)(wb), run(RecalcEngine.Legacy, opts)(wb), "sequential")
    }
  }

  property("parallel: Compact equals Legacy, and equals the sequential pass") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val par = RecalcOptions(clock = clock, iterative = IterativeMode.Off, parallelism = 3)
      val seq = par.copy(parallelism = 1)
      val compact = run(RecalcEngine.Compact, par)(wb)
      Prop.all(
        same(compact, run(RecalcEngine.Legacy, par)(wb), "parallel"),
        same(compact, run(RecalcEngine.Compact, seq)(wb), "parallel vs sequential")
      )
    }
  }

  /** Tall books: the depth-0 wave alone usually exceeds the parallel path's sequential cutoff. */
  private val genTall: Gen[Workbook] = CompactGraphGens.genWorkbookOf(40)

  /** The widest wave a parallel recalculation of `wb` would run (its cyclic core removed). */
  private def widestWave(wb: Workbook): Int =
    val graph = RecalcGraph.build(wb, RecalcEngine.Compact, RecalcCounters.disabled)
    val core = graph.cyclicCore
    val removed = core ++ graph.dependents(core, Set.empty)
    graph.order(removed).fold(_ => 0, o => o.waves(o.order).map(_.length).maxOption.getOrElse(0))

  test("the tall generator reaches the parallel path: waves wider than its cutoff") {
    val books =
      (1 to 20).flatMap(i => genTall(Gen.Parameters.default, org.scalacheck.rng.Seed(i.toLong)))
    assert(books.count(widestWave(_) >= 16) >= books.size / 2, books.map(widestWave).toString)
  }

  property("parallel over tall books: Compact equals Legacy and the sequential pass") {
    Prop.forAll(genTall) { wb =>
      val par = RecalcOptions(clock = clock, iterative = IterativeMode.Off, parallelism = 4)
      val seq = par.copy(parallelism = 1)
      val compact = run(RecalcEngine.Compact, par)(wb)
      Prop.all(
        same(compact, run(RecalcEngine.Legacy, par)(wb), "parallel"),
        same(compact, run(RecalcEngine.Compact, seq)(wb), "parallel vs sequential")
      )
    }
  }

  property("recalculateUncached: Compact equals Legacy") {
    Prop.forAll(genBook) { wb =>
      def opts = RecalcOptions(clock = clock, rng = Rng.seeded(7L), iterative = IterativeMode.Off)
      same(
        recalculateUncachedCounted(wb, opts, RecalcCounters.disabled, RecalcEngine.Compact),
        recalculateUncachedCounted(wb, opts, RecalcCounters.disabled, RecalcEngine.Legacy),
        "uncached"
      )
    }
  }

  property("iterative, Gauss-Seidel and Jacobi: Compact equals Legacy") {
    Prop.forAll(genBook, Gen.oneOf(IterationScheme.GaussSeidel, IterationScheme.Jacobi)) {
      (wb, scheme) =>
        def opts = RecalcOptions(
          clock = clock,
          rng = Rng.seeded(11L),
          iterative = IterativeMode.Force(IterativeCalc(25, BigDecimal("0.001"), scheme = scheme))
        )
        same(run(RecalcEngine.Compact, opts)(wb), run(RecalcEngine.Legacy, opts)(wb), s"$scheme")
    }
  }

  // ---- the canonical order, observed --------------------------------------------------------

  private def sheet(name: String, cells: (String, CellValue)*): Sheet =
    Sheet(
      SheetName.unsafe(name),
      cells = cells.map { (a1, v) =>
        val ref = ARef.parse(a1).getOrElse(ARef.from0(0, 0))
        ref -> Cell(ref, v)
      }.toMap
    )

  private def fx(text: String): CellValue = CellValue.Formula(text)

  test("cycle and blocked errors list in grid order: A2 before A10, sheet name first") {
    val wb = Workbook(
      Vector(
        sheet("b", "A10" -> fx("A10+1"), "A2" -> fx("A2+1"), "C1" -> fx("A10")),
        sheet("B", "A3" -> fx("A3*2"), "B1" -> fx("A3"))
      )
    )
    Vector(RecalcEngine.Compact, RecalcEngine.Legacy).foreach { engine =>
      val opts = RecalcOptions(clock = clock, iterative = IterativeMode.Off)
      val errors = run(engine, opts)(wb).errors.map(e => s"${e.sheet.value}!${e.ref.toA1}")
      // circular first, then blocked; each in (sheet name, row, column) order
      assertEquals(errors, Vector("B!A3", "b!A2", "b!A10", "B!B1", "b!C1"), engine.toString)
    }
  }

  test("a seeded RAND draws in grid order whatever the cells' insertion order") {
    val cells = (1 to 12).reverse.map(r => s"A$r" -> fx("RAND()"))
    val wb = Workbook(Vector(sheet("S", cells*)))
    val opts = RecalcOptions(clock = clock, rng = Rng.seeded(3L), iterative = IterativeMode.Off)
    val result = run(RecalcEngine.Compact, opts)(wb)
    val rng = Rng.seeded(3L)
    val expected = (1 to 12).map(_ => rng.nextDouble())
    val actual = (1 to 12).map { r =>
      result.evaluated(SheetName.unsafe("S"))(ARef.from0(0, r - 1)) match
        case CellValue.Number(n) => n.toDouble
        case other => fail(s"A$r: $other")
    }
    assertEquals(actual.toVector, expected.toVector)
  }

  test("parallel waves re-list errors in the sequential order") {
    // 40 columns: row 3 fails reading a computed row-2 cell (depth 1), row 4 fails reading a
    // constant (depth 0); grid order puts row 3 first, the wave fold puts row 4 first
    val cols = 40
    val cells = (0 until cols).flatMap { c =>
      val col = ARef.from0(c, 0).toA1.takeWhile(_.isLetter)
      Vector(
        s"${col}1" -> CellValue.Number(BigDecimal(c)),
        s"${col}2" -> fx(s"${col}1*2"),
        s"${col}3" -> fx(s"NOSUCHFN(${col}2)"),
        s"${col}4" -> fx(s"NOSUCHFN(${col}1)")
      )
    }
    val wb = Workbook(Vector(sheet("P", cells*)))
    val seq =
      run(RecalcEngine.Compact, RecalcOptions(clock = clock, iterative = IterativeMode.Off))(wb)
    val par = run(
      RecalcEngine.Compact,
      RecalcOptions(clock = clock, iterative = IterativeMode.Off, parallelism = 4)
    )(wb)
    assertEquals(par.errors, seq.errors)
    val listed = seq.errors.map(e => QualifiedRef(e.sheet, e.ref))
    assertEquals(listed, listed.sortBy(QualifiedRef.gridKey))
  }
