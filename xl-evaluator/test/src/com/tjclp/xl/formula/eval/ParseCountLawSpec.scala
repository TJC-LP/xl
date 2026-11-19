package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.eval.DataTableSeeder.*
import com.tjclp.xl.formula.eval.RecalcCounters.Counter
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.formula.graph.CompactGraphGens
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{CalcPr, Workbook}
import munit.ScalaCheckSuite
import org.scalacheck.Prop

/**
 * #891: one parse of each formula per operation, beyond the full recalculation
 * ([[RecalcCountersSpec]] and [[PassCacheLawSpec]] hold it there). An operation builds one
 * recalculation graph, whose build parses every formula cell once, and everything after it (the
 * evaluation, the edit cone, the audit's classification, the data-table what-if passes) reads those
 * parses.
 *
 * `FormulaParses` counts parses of formula cell texts: the graph build's, and every later parse of
 * a cell the build's parse could not serve. Each law is asserted on a book of F formulas and on one
 * ten times larger, and the table laws at two interior sizes as well, so the count is shown to be F
 * and not a constant, and not a function of how many cells a what-if pass evaluates. F is the
 * number of formula cells that are graph nodes: a data-table record is none.
 *
 * Defined-name bodies are not formula cells: each consumer parses the bodies it reads once
 * (`NameParses` counts the evaluation's).
 */
class ParseCountLawSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(60)

  private val s = SheetName.unsafe("S")
  private val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 2))
  private val options =
    RecalcOptions(clock = clock, rng = Rng.seeded(5L), iterative = IterativeMode.Off)

  private def r(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)
  private def n(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def f(text: String): CellValue = CellValue.Formula(text)

  /**
   * `rows` rows of a model every formula of which moves with `J1`: `A` holds constants, `B = A*J1`
   * and `C = B + SUM($A$1:A)`. 2 x `rows` formulas.
   */
  private def model(rows: Int): Sheet =
    (1 to rows).foldLeft(Sheet(s).put(r("J1"), n(5))) { (sheet, i) =>
      sheet
        .put(r(s"A$i"), n(i))
        .put(r(s"B$i"), f(s"A$i*$$J$$1"))
        .put(r(s"C$i"), f(s"B$i+SUM($$A$$1:A$i)"))
    }

  /**
   * [[model]] with a one-input data table of `cells` interior cells below the corner `M1`, which
   * sums the model's column C: the table's what-if cone is the whole model, re-derived for every
   * interior cell. One formula more than the model (the corner); the records are not graph nodes.
   */
  private def tableBook(
    rows: Int,
    cells: Int,
    corner: Int => String = rows => s"SUM(C1:C$rows)+J1"
  ): Workbook =
    val interior = CellRange.parse(s"M2:M${1 + cells}").fold(fail(_), identity)
    val kind: FormulaKind.DataTable =
      FormulaKind.DataTable(interior, dt2D = false, dtr = false, r1 = Some(r("J1")), r2 = None)
    val base = model(rows).put(r("M1"), f(corner(rows)))
    Workbook((2 to 1 + cells).foldLeft(base) { (sheet, i) =>
      sheet.put(r(s"L$i"), n(i)).put(r(s"M$i"), CellValue.dataTable(kind, None): CellValue)
    })

  private val sizes = Vector(10, 100)
  private val interiors = Vector(3, 30)

  /** What a counted run parsed, with the size of the graph it built. */
  private def counted(run: RecalcCounters => Any): (Long, Long) =
    val counters = RecalcCounters.create()
    run(counters)
    val snap = counters.snapshot
    (snap(Counter.FormulaParses), snap(Counter.GraphNodes))

  private def settled(wb: Workbook): Workbook = wb.recalculate(options).workbook

  private def seeded(wb: Workbook): Workbook =
    settled(wb).seedDataTablesReport(clock, None).fold(e => fail(e.message), _.workbook)

  // ---- the after-edit pass -----------------------------------------------------------------------

  test("the after-edit pass parses each formula of the book once, whatever the cone") {
    sizes.foreach { rows =>
      val book = settled(Workbook(model(rows)))
      val formulas = 2L * rows
      // J1 reaches every formula; A{rows} reaches two
      Vector("J1" -> 2 * rows, s"A$rows" -> 2).foreach { (a1, cone) =>
        val edited = book.put(book.sheets(0).put(r(a1), n(50)))
        val counters = RecalcCounters.create()
        val result = DependentRecalculation.recalculateAfterEdit(
          edited,
          s,
          Set(r(a1)),
          clock,
          new Evaluator.AggregateMemo(counters = counters),
          Rng.seeded(1L)
        )
        assertEquals(result.evaluated.valuesIterator.map(_.size).sum, cone, s"$rows rows, $a1")
        assertEquals(counters.snapshot(Counter.FormulaParses), formulas, s"$rows rows, $a1")
        assertEquals(counters.snapshot(Counter.GraphNodes), formulas, s"$rows rows, $a1")
      }
    }
  }

  test("a formula written by the edit is parsed once, as its new text") {
    val book = settled(Workbook(model(10)))
    val edited = book.put(book.sheets(0).put(r("B3"), f("A3*100")))
    val counters = RecalcCounters.create()
    val result = DependentRecalculation.recalculateAfterEdit(
      edited,
      s,
      Set(r("B3")),
      clock,
      new Evaluator.AggregateMemo(counters = counters),
      Rng.seeded(1L)
    )
    assertEquals(result.evaluated(s)(r("B3")), n(300))
    assertEquals(counters.snapshot(Counter.FormulaParses), 20L)
  }

  // ---- the audit ---------------------------------------------------------------------------------

  test("the audit parses each formula once, with and without the data-table check") {
    for rows <- sizes; cells <- interiors do
      val book = seeded(tableBook(rows, cells))
      val formulas = 2L * rows + 1
      Vector(true, false).foreach { check =>
        val (parses, nodes) = counted(WorkbookAudit.of(book, clock, check, _))
        assertEquals(nodes, formulas, s"$rows rows, $cells cells, tables $check")
        assertEquals(parses, formulas, s"$rows rows, $cells cells, tables $check")
      }
  }

  test("a counted audit is the audit") {
    val book = seeded(tableBook(10, 3))
    Vector(true, false).foreach { check =>
      assertEquals(
        WorkbookAudit.of(book, clock, check, RecalcCounters.create()),
        WorkbookAudit.of(book, clock, check)
      )
    }
  }

  // ---- data-table seeding ------------------------------------------------------------------------

  test("seeding parses each formula once, however many interior cells re-derive the cone") {
    for rows <- sizes; cells <- interiors do
      val book = settled(tableBook(rows, cells))
      val formulas = 2L * rows + 1
      val (parses, nodes) = counted(seedDataTablesReportCounted(book, clock, None, _))
      assertEquals(nodes, formulas, s"$rows rows, $cells cells")
      assertEquals(parses, formulas, s"$rows rows, $cells cells")
  }

  test("the seeding the counters observe computes the table") {
    val book = settled(tableBook(10, 3))
    val report = seedDataTablesReportCounted(book, clock, None, RecalcCounters.create())
    assertEquals(report.warnings, Vector.empty)
    // C(i) = i*J1 + i(i+1)/2, so the corner is 55*J1 + 220 + J1: at J1 = 2, 3, 4
    assertEquals(interior(report, "M2", "M3", "M4"), Vector(332, 388, 444).map(v => Some(n(v))))
    assertEquals(Right(report), book.seedDataTablesReport(clock, None))
  }

  test("seeding a table behind a reference cycle parses each formula once (iterated lane)") {
    // X1 and Y1 converge on each other and read the input; the corner reads X1 and the model
    def cyclic(rows: Int, cells: Int): Workbook =
      val wb = tableBook(rows, cells, rows => s"X1+SUM(C1:C$rows)")
      val sheet = wb.sheets(0).put(r("X1"), f("Y1/2+J1")).put(r("Y1"), f("X1/2+1"))
      wb.put(sheet)
        .copy(metadata = wb.metadata.copy(calcPr = Some(CalcPr(iterativeCalculation = true))))
    val iterative = Some(IterativeCalc(100, 0.001))
    for rows <- sizes; cells <- interiors do
      val book = cyclic(rows, cells).recalculate(clock, IterativeCalc(100, 0.001)).workbook
      val formulas = 2L * rows + 3
      val counters = RecalcCounters.create()
      val report = seedDataTablesReportCounted(book, clock, iterative, counters)
      assertEquals(report.warnings, Vector.empty, s"$rows rows, $cells cells")
      assertEquals(counters.snapshot(Counter.FormulaParses), formulas, s"$rows rows, $cells cells")
  }

  // ---- recalculation followed by a second consumer of its graph ---------------------------------

  test("`recalc --tables` (recalculate, then seed its result) parses each formula once") {
    for rows <- sizes; cells <- interiors do
      val book = tableBook(rows, cells)
      val formulas = 2L * rows + 1
      val (parses, nodes) = counted(recalculateThenSeedTablesCounted(book, clock, None, 1, _))
      assertEquals(nodes, formulas, s"$rows rows, $cells cells")
      assertEquals(parses, formulas, s"$rows rows, $cells cells")
      // run back to back as two operations, each builds its own graph
      val (apart, _) = counted { counters =>
        val result = recalculateCounted(book, options, counters)
        seedDataTablesReportCounted(result.workbook, clock, None, counters)
      }
      assertEquals(apart, 2 * formulas, s"$rows rows, $cells cells, apart")
  }

  test("`recalculate` with `seedTables` parses each formula once") {
    for rows <- sizes; cells <- interiors do
      val book = tableBook(rows, cells)
      val (parses, _) = counted(recalculateCounted(book, options.copy(seedTables = true), _))
      assertEquals(parses, 2L * rows + 1, s"$rows rows, $cells cells")
  }

  test("a recalculation with its edit cone (batch, name edits, iterating books) parses once") {
    sizes.foreach { rows =>
      val book = Workbook(model(rows))
      val seeds = Set(QualifiedRef(s, r("J1")))
      val (parses, nodes) = counted(recalculateWithEditConeCounted(book, clock, None, seeds, _))
      assertEquals(nodes, 2L * rows, s"$rows rows")
      assertEquals(parses, 2L * rows, s"$rows rows")
    }
  }

  test("a lookup narrowed on a what-if input costs seeding a second graph: 2F, by design") {
    // the corner's column selector is the table's input. Recalculation reads the one strip the
    // selector picks today; seeding substitutes the selector, so it sets that graph aside and
    // builds one with the lookup's whole table (RecalcGraph.narrowsOnWhatIfInput)
    def lookupBook(rows: Int): Workbook =
      val interior = CellRange.parse("M2:M3").fold(fail(_), identity)
      val kind: FormulaKind.DataTable =
        FormulaKind.DataTable(interior, dt2D = false, dtr = false, r1 = Some(r("K1")), r2 = None)
      val sheet = model(rows)
        .put(r("K1"), n(2))
        .put(r("M1"), f(s"VLOOKUP(1,A1:C$rows,K1,FALSE)"))
        .put(r("L2"), n(2))
        .put(r("L3"), n(3))
        .put(r("M2"), CellValue.dataTable(kind, None): CellValue)
        .put(r("M3"), CellValue.dataTable(kind, None): CellValue)
      Workbook(sheet)
    sizes.foreach { rows =>
      val book = lookupBook(rows)
      val formulas = 2L * rows + 1
      val counters = RecalcCounters.create()
      val (_, report) = recalculateThenSeedTablesCounted(book, clock, None, 1, counters)
      assertEquals(counters.snapshot(Counter.FormulaParses), 2 * formulas, s"$rows rows")
      // B1 = 1*5 and C1 = 5+1: the substituted selector reads the other strip
      assertEquals(interior(report, "M2", "M3"), Vector(5, 6).map(v => Some(n(v))))
      val (audit, _) = counted(WorkbookAudit.of(report.workbook, clock, true, _))
      assertEquals(audit, 2 * formulas, s"$rows rows, audit")
    }
  }

  /** The values a seeding run left in the interior cells `refs`: a record's cache, or a value. */
  private def interior(report: DataTableSeedReport, refs: String*): Vector[Option[CellValue]] =
    refs.toVector.map { a1 =>
      report.workbook.sheets(0).cells.get(r(a1)).map(_.value) match
        case Some(CellValue.Formula(_, cache, _)) => cache
        case other => other
    }

  // ---- sharing the graph changes no result -------------------------------------------------------

  property("recalculate-then-seed over one graph equals the two public calls back to back") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val (result, report) = recalculateThenSeedTables(wb, clock, None, 1)
      val apart = wb.recalculate(clock)
      Prop.all(
        Prop(result == apart) :| "the recalculation",
        Prop(
          Right(report) == apart.workbook
            .seedDataTablesReport(clock, declaredIterative(apart.workbook))
        ) :| "the seeding report"
      )
    }
  }

  property("recalculate-with-cone over one graph equals `recalculate` and `editCone`") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val seeds = wb.sheets.iterator.flatMap { sheet =>
        sheet.cells.keysIterator.take(2).map(QualifiedRef(sheet.name, _))
      }.toSet
      val (result, cone) = recalculateWithEditCone(wb, clock, None, seeds)
      Prop.all(
        Prop(result == wb.recalculate(clock)) :| "the recalculation",
        Prop(cone == DependentRecalculation.editCone(wb, seeds)) :| "the cone"
      )
    }
  }

  /**
   * A two-input table over G5:I6 whose row input `K1` is the column selector of the corner's lookup
   * into A10:D11 (every cell of which moves with the column input `J1`), unseeded.
   */
  private def selectorBook(corner: String): Workbook =
    val kind: FormulaKind.DataTable = FormulaKind.DataTable(
      CellRange.parse("G5:I6").fold(fail(_), identity),
      dt2D = true,
      dtr = false,
      r1 = Some(r("K1")),
      r2 = Some(r("J1"))
    )
    val cells = Seq(
      "A10" -> n(1),
      "B10" -> f("J1*1"),
      "C10" -> f("J1*10"),
      "D10" -> f("J1*100"),
      "A11" -> n(2),
      "B11" -> f("J1*2"),
      "C11" -> f("J1*20"),
      "D11" -> f("J1*200"),
      "J1" -> n(5),
      "K1" -> n(2),
      "F4" -> f(corner),
      "G4" -> n(2),
      "H4" -> n(3),
      "I4" -> n(4),
      "F5" -> n(1),
      "F6" -> n(3),
      "G5" -> (CellValue.dataTable(kind, None): CellValue)
    )
    Workbook(cells.foldLeft(Sheet(s))((sheet, kv) => sheet.put(r(kv._1), kv._2)))

  test("a table whose input selects the corner's lookup strip seeds the same over a shared graph") {
    // Excel: interior(row j, column k) is the lookup's column k of row 2 at J1 = j
    val expected = Vector(2, 20, 200, 6, 60, 600).map(v => Some(n(v)))
    Vector("VLOOKUP(2,A10:D11,K1,FALSE)", "VLOOKUP(2,A10:D11,K1+0,FALSE)", "INDEX(A10:D11,2,K1)")
      .foreach { corner =>
        val book = selectorBook(corner)
        val (result, report) = recalculateThenSeedTables(book, clock, None, 1)
        assertEquals(report.warnings, Vector.empty, corner)
        assertEquals(interior(report, "G5", "H5", "I5", "G6", "H6", "I6"), expected, corner)
        val apart = book.recalculate(clock)
        assertEquals(result, apart, corner)
        assertEquals(Right(report), apart.workbook.seedDataTablesReport(clock, None), corner)
        assertEquals(
          recalculateCounted(
            book,
            options.copy(seedTables = true),
            RecalcCounters.disabled
          ).workbook,
          report.workbook,
          corner
        )
      }
  }

  test("the generated books of the sharing property do seed tables") {
    val books = (1 to 60).flatMap { i =>
      CompactGraphGens.genWorkbook(
        org.scalacheck.Gen.Parameters.default,
        org.scalacheck.rng.Seed(i.toLong)
      )
    }
    val seededTables = books.count { wb =>
      val settledBook = wb.recalculate(clock).workbook
      settledBook.seedDataTablesReport(clock, None).exists(_.workbook != settledBook)
    }
    assert(seededTables >= 30, s"$seededTables of ${books.size} books changed under seeding")
  }

  test("sharing the graph changes nothing on the table books, iterating ones included") {
    val iterating = tableBook(10, 3).copy(metadata =
      tableBook(10, 3).metadata.copy(calcPr = Some(CalcPr(iterativeCalculation = true)))
    )
    val calc = Some(IterativeCalc(100, 0.001))
    Vector((tableBook(10, 3), None, 1), (tableBook(100, 30), None, 4), (iterating, calc, 1))
      .foreach { (wb, iterative, parallelism) =>
        val (result, report) = recalculateThenSeedTables(wb, clock, iterative, parallelism)
        val apart = iterative match
          case Some(c) => wb.recalculate(clock, c)
          case None if parallelism > 1 => wb.recalculateParallel(clock, parallelism)
          case None => wb.recalculate(clock)
        assertEquals(result, apart)
        assertEquals(
          Right(report),
          apart.workbook.seedDataTablesReport(clock, declaredIterative(apart.workbook))
        )
      }
  }
