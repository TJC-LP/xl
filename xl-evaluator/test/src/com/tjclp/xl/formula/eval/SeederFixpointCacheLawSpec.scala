package com.tjclp.xl.formula.eval

import scala.util.Random

import com.tjclp.xl.CellRange
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * #891: the data-table seeder's iterated lane fixpoints a cycle once per interior cell, each
 * fixpoint over its own shared criteria cache ([[Evaluator.AggregateMemo.fixpointCache]]). Law: a
 * seeding run gives the same workbook and the same warnings as the lane the seeder ran before it
 * shared the cache, a plain `Evaluator.instance` with `columnCache = false`, reached through the
 * test-only [[DataTableSeeder.seedReportWith]].
 *
 * The what-if substitution is what makes the seeder's use of the cache its own case: between two
 * interior cells the input cell changes, and with it every cone formula downstream of it, so a
 * criteria column may hold different values in each fixpoint. A cache that outlived one interior
 * cell would serve the previous substitution's column.
 *
 * Seeds are fixed; each test asserts what its cases exercised (cells seeded, fixpoints that did not
 * converge, slices built), so the law cannot go vacuous unnoticed.
 */
class SeederFixpointCacheLawSpec extends FunSuite:

  private val s = SheetName.unsafe("S")
  private def at(c: Int, r: Int): ARef = ARef.from0(c, r)
  private def a1(text: String): ARef = ARef.parse(text).fold(fail(_), identity)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def range(text: String): CellRange = CellRange.parse(text).fold(fail(_), identity)

  /** The lane before #891: no memo, no shared cache. */
  private val plain: DataTableSeeder.FixpointLane =
    DataTableSeeder.FixpointLane((rng, _) => Evaluator.instance(rng), columnCache = false)

  /** What one seeding run on the production lane did. */
  private final case class Seeded(
    report: DataTableSeedReport,
    // every cache a region memo carried, one per fixpoint
    caches: Vector[PassCache]
  ):
    def slices: Int = caches.iterator.map(_.builtSlices.size).sum

  /** Seed `wb` on the production lane, recording the caches its fixpoints used. */
  private def seedShared(wb: Workbook, iterative: IterativeCalc): Seeded =
    val caches = Vector.newBuilder[PassCache]
    val shared = DataTableSeeder.FixpointLane.Shared
    val recording = shared.copy(roundEvaluator = (rng, memo) =>
      caches ++= memo.passCache
      shared.roundEvaluator(rng, memo))
    val report = DataTableSeeder.seedReportWith(wb, Clock.system, Some(iterative), recording)
    Seeded(report, caches.result().distinctBy(System.identityHashCode))

  /**
   * Both lanes agree on `wb`: the interior cell by cell (for a readable failure), every warning,
   * then the whole workbook. Returns the production lane's run.
   */
  private def assertLanesAgree(
    wb: Workbook,
    interior: CellRange,
    iterative: IterativeCalc,
    clue: => String
  ): Seeded =
    val shared = seedShared(wb, iterative)
    val unshared = DataTableSeeder.seedReportWith(wb, Clock.system, Some(iterative), plain)
    def values(report: DataTableSeedReport): Vector[(String, CellValue)] =
      val sheet = report.workbook(s).fold(e => fail(e.message), identity)
      interior.cellsRowMajor.map(ref => ref.toA1 -> sheet.cells.valueAt(ref)).toVector
    assertEquals(values(shared.report), values(unshared), clue)
    assertEquals(shared.report.warnings, unshared.warnings, clue)
    assertEquals(shared.report.workbook, unshared.workbook, clue)
    shared

  private def seededValues(report: DataTableSeedReport, interior: CellRange): Vector[BigDecimal] =
    val sheet = report.workbook(s).fold(e => fail(e.message), identity)
    interior.cellsRowMajor.toVector.flatMap { ref =>
      sheet.cells.valueAt(ref) match
        case CellValue.Number(n) => Some(n)
        case CellValue.Formula(_, Some(CellValue.Number(n)), _) => Some(n)
        case _ => None
    }

  // ----- generated circular data tables -----

  /**
   * A 1-D column-oriented table K10:K13 (axis J10:J13, source K9) over a four-member cycle G1:G4 of
   * criteria calls. Column A holds text keys, B and D numbers, E numbers mixed with cached and
   * uncached formulas, some of them downstream of the input or of a member. The input is D3 (a cell
   * INSIDE a criteria column) or H1.
   */
  private def gen(rnd: Random): Sheet =
    val rows = 40
    val input = if rnd.nextBoolean() then "D3" else "H1"
    val inputAbs = "$" + input.take(1) + "$" + input.drop(1)
    val base = (0 until rows).foldLeft(Sheet(s)) { (sheet, r) =>
      val e: CellValue = rnd.nextInt(10) match
        case 0 => CellValue.Formula(s"$inputAbs*2+B${r + 1}", None)
        case 1 => CellValue.Formula(s"$inputAbs+1", Some(num(rnd.nextInt(5))))
        case 2 => CellValue.Formula(s"G1+B${r + 1}", None)
        case 3 => CellValue.Formula(s"B${r + 1}*2", Some(num(rnd.nextInt(5))))
        case _ => num(rnd.nextInt(5))
      sheet
        .put(at(0, r), CellValue.Text(Vector("x", "y", "z")(rnd.nextInt(3))))
        .put(at(1, r), num(rnd.nextInt(50)))
        .put(at(3, r), num(rnd.nextInt(5)))
        .put(at(4, r), if rnd.nextInt(3) == 0 then e else num(rnd.nextInt(5)))
    }
    def crit(numeric: Boolean): String =
      if !numeric then Vector("\"x\"", "\"y\"", "\"<>z\"")(rnd.nextInt(3))
      else Vector("\">0\"", "\"<3\"", "2", inputAbs, "\">\"&G2", "G3")(rnd.nextInt(6))
    def critRange(): (String, Boolean) = rnd.nextInt(6) match
      case 0 => ("A1:A40", false)
      case 1 => ("D1:D40", true)
      case 2 => ("E1:E40", true)
      case 3 => ("D17:D40", true)
      case 4 => ("G1:G40", true)
      case _ => ("E17:E40", true)
    def term(): String =
      val (r1, n1) = critRange()
      val sum = Vector("B", "D", "E")(rnd.nextInt(3))
      val (lo, hi) = if r1.contains("17") then (17, 40) else (1, 40)
      rnd.nextInt(9) match
        case 4 => "SUM(E1:E40)"
        case 5 => "AVERAGE(G1:G40)"
        case 6 => "MAX(D:D)+COUNT(E:E)"
        case 7 => s"COUNTIF(E:E,${crit(true)})+SUMIF(D:D,${crit(true)},B:B)"
        case 8 => "SUMPRODUCT(B1:B40,E1:E40)/100"
        case 0 => s"SUMIFS($sum$lo:$sum$hi,$r1,${crit(n1)})"
        case 1 => s"COUNTIF($r1,${crit(n1)})"
        case 2 => s"SUMIF($r1,${crit(n1)},$sum$lo:$sum$hi)"
        case _ =>
          val keys = if lo == 17 then "A17:A40" else "A1:A40"
          s"COUNTIFS($r1,${crit(n1)},$keys,${crit(false)})"
    val members = (0 until 4).map { i =>
      val next = (i + 1) % 4
      at(6, i) -> s"(${term()}+${term()})/1000+G${next + 1}*0.3+$inputAbs*0.01"
    }
    val withMembers = members.foldLeft(base)((sheet, m) => sheet.put(m._1, CellValue.Formula(m._2)))
    val source =
      if rnd.nextBoolean() then "G1+SUMIF(E1:E40,\">1\")/100" else "G2*2+COUNTIF(D1:D40,\">1\")"
    val kind: FormulaKind.DataTable = FormulaKind.DataTable(
      ref = range("K10:K13"),
      dt2D = false,
      dtr = false,
      r1 = Some(a1(input)),
      r2 = None
    )
    withMembers
      .put(a1("H1"), num(1))
      .put(a1("K9"), CellValue.Formula(source))
      .put(a1("J10"), num(1))
      .put(a1("J11"), num(2))
      .put(a1("J12"), num(3))
      .put(a1("J13"), num(4))
      .put(a1("K10"), CellValue.dataTable(kind, None))

  private val schemes = List(IterationScheme.Jacobi, IterationScheme.GaussSeidel)

  test("#891: generated circular data tables seed alike with and without the shared cache") {
    val interior = range("K10:K13")
    var seededCells = 0
    var slices = 0
    var notConverged = 0
    var skipped = 0
    (0 until Cases).foreach { i =>
      val seed = 5000 + i
      val wb = Workbook(gen(new Random(seed)))
      // every third case stops after three rounds, so fixpoints end unconverged too
      val maxIter = if i % 3 == 2 then 3 else 200
      schemes.foreach { scheme =>
        val iterative = IterativeCalc(maxIter, BigDecimal("1E-9"), scheme = scheme)
        val run = assertLanesAgree(wb, interior, iterative, s"seed $seed, $scheme, $maxIter rounds")
        val cells = seededValues(run.report, interior).size
        seededCells += cells
        slices += run.slices
        // one fixpoint, so one cache, per interior cell: none outlives its substitution
        assertEquals(run.caches.size, interior.cellsRowMajor.size, s"seed $seed, $scheme")
        run.report.warnings.foreach {
          case SeedTableWarning.NotConverged(_, _, combinations, _) => notConverged += combinations
          case SeedTableWarning.Skipped(_, _, cells, _) => skipped += cells
          case _ => ()
        }
      }
    }
    def clue = s"seededCells=$seededCells slices=$slices notConverged=$notConverged " +
      s"skipped=$skipped"
    assert(seededCells > Cases * schemes.size * 3, clue)
    assert(slices > 0, clue)
    assert(notConverged > 0, clue)
  }

  private val Cases = 30

  // ----- fixed cases, with the values worked out by hand -----

  // Rows 1..40: A a text key, B the row number, D the 0-based row modulo 5 (0, 1, 2, 3, 4, 0, …).
  private val grid: Sheet = (0 until 40).foldLeft(Sheet(s)) { (sheet, r) =>
    sheet
      .put(at(0, r), CellValue.Text(if r % 2 == 0 then "x" else "y"))
      .put(at(1, r), num(r + 1))
      .put(at(3, r), num(r % 5))
  }

  private val settings: List[IterativeCalc] =
    schemes.map(scheme => IterativeCalc(200, BigDecimal("1E-9"), scheme = scheme))

  /** The interior's seeded numbers against `expected`, row-major; the fixpoint stops at 1E-9. */
  private def assertInterior(
    run: Seeded,
    interior: CellRange,
    expected: Vector[BigDecimal],
    clue: => String
  ): Unit =
    val actual = seededValues(run.report, interior)
    assertEquals(actual.size, expected.size, s"$clue: $actual")
    actual.zip(expected).foreach { (a, e) =>
      assert((a - e).abs < BigDecimal("1E-6"), s"$clue: $actual, expected $expected")
    }
    assertEquals(run.report.warnings, Vector.empty, clue)

  /** What each fixpoint's cache holds at `ref`'s row of `ref`'s column, in interior order. */
  private def sliced(run: Seeded, ref: ARef): Vector[Option[CellValue]] =
    run.caches.map(
      _.builtSlices.find(_.col == ref.col.index0).map(_.valueAt(ref.row.index0))
    )

  // G1 = a + 0.3 G2 and G2 = b + 0.3 G1 give G1 + G2 = (a + b) / 0.7
  private def sumOfCycle(a: BigDecimal, b: BigDecimal): BigDecimal = (a + b) / BigDecimal("0.7")

  test("#891: the input cell inside a criteria column: each interior cell slices its own value") {
    // input D3; the cycle's criteria ranges are column D, which holds no member and no formula
    val kind: FormulaKind.DataTable =
      FormulaKind.DataTable(range("K10:K13"), dt2D = false, dtr = false, Some(a1("D3")), None)
    val axis = Vector(0, 3, 7, 9)
    val sheet = axis.zipWithIndex
      .foldLeft(grid)((acc, v) => acc.put(at(9, 9 + v._2), num(v._1)))
      .put(a1("G1"), CellValue.Formula("SUMIFS(B1:B40,D1:D40,\">2\")/1000+G2*0.3"))
      .put(a1("G2"), CellValue.Formula("COUNTIF(D1:D40,D3)/100+G1*0.3"))
      .put(a1("K9"), CellValue.Formula("G1+G2"))
      .put(a1("K10"), CellValue.dataTable(kind, None))
    // B over D > 2: rows with D = 3 sum to 172, with D = 4 to 180, and row 3 (B = 3) joins when
    // the substituted D3 is above 2. Cells of D equal to D3: eight for 0 and for 3 plus D3 itself,
    // D3 alone for 7 and 9.
    val expected = Vector(
      sumOfCycle(BigDecimal("0.352"), BigDecimal("0.09")),
      sumOfCycle(BigDecimal("0.355"), BigDecimal("0.09")),
      sumOfCycle(BigDecimal("0.355"), BigDecimal("0.01")),
      sumOfCycle(BigDecimal("0.355"), BigDecimal("0.01"))
    )
    settings.foreach { iterative =>
      val run = assertLanesAgree(Workbook(sheet), kind.ref, iterative, iterative.scheme.toString)
      assertInterior(run, kind.ref, expected, iterative.scheme.toString)
      // column D is sliced in every fixpoint, and each slice holds that cell's substitution
      assertEquals(sliced(run, a1("D3")), axis.map(v => Option(num(v))))
    }
  }

  test("#891: an uncached cone formula inside a criteria column is sliced once it is resolved") {
    // input H1; E1:E10 are uncached formulas downstream of it, E11:E40 are 0; the cycle's
    // criteria ranges are column E
    val kind: FormulaKind.DataTable =
      FormulaKind.DataTable(range("K10:K13"), dt2D = false, dtr = false, Some(a1("H1")), None)
    val axis = Vector(0, 3, 7, 9)
    val withE = (0 until 40).foldLeft(grid) { (acc, r) =>
      acc.put(at(4, r), if r < 10 then CellValue.Formula(s"$$H$$1+B${r + 1}", None) else num(0))
    }
    val sheet = axis.zipWithIndex
      .foldLeft(withE)((acc, v) => acc.put(at(9, 9 + v._2), num(v._1)))
      .put(a1("H1"), num(1))
      .put(a1("G1"), CellValue.Formula("COUNTIF(E1:E40,\">10\")/100+G2*0.3"))
      .put(a1("G2"), CellValue.Formula("SUMIFS(B1:B40,E1:E40,\"<=10\")/1000+G1*0.3"))
      .put(a1("K9"), CellValue.Formula("G1+G2"))
      .put(a1("K10"), CellValue.dataTable(kind, None))
    // E1:E10 = v + 1 .. v + 10: v of them are above 10, and the B of the others sum to
    // 1 + … + (10 - v); B11:B40 (765) always count, their E being 0
    val expected = Vector(
      sumOfCycle(BigDecimal("0.00"), BigDecimal("0.820")),
      sumOfCycle(BigDecimal("0.03"), BigDecimal("0.793")),
      sumOfCycle(BigDecimal("0.07"), BigDecimal("0.771")),
      sumOfCycle(BigDecimal("0.09"), BigDecimal("0.766"))
    )
    settings.foreach { iterative =>
      val run = assertLanesAgree(Workbook(sheet), kind.ref, iterative, iterative.scheme.toString)
      assertInterior(run, kind.ref, expected, iterative.scheme.toString)
      // E1 = H1 + 1 under each substitution, read from that fixpoint's slice
      assertEquals(sliced(run, a1("E1")), axis.map(v => Option(num(v + 1))))
    }
  }

  test("#891: a cone formula that cannot be resolved keeps its criteria column unsliced") {
    // as above, but E5 reads a sheet the workbook does not hold: no substitution resolves it, so
    // it reaches every fixpoint uncached, and a slice of column E would read it as blank
    val kind: FormulaKind.DataTable =
      FormulaKind.DataTable(range("K10:K13"), dt2D = false, dtr = false, Some(a1("H1")), None)
    val withE = (0 until 40).foldLeft(grid) { (acc, r) =>
      acc.put(at(4, r), if r < 10 then CellValue.Formula(s"$$H$$1+B${r + 1}", None) else num(0))
    }
    val sheet = Vector(0, 3, 7, 9).zipWithIndex
      .foldLeft(withE)((acc, v) => acc.put(at(9, 9 + v._2), num(v._1)))
      .put(a1("E5"), CellValue.Formula("$H$1+'NoSuchSheet'!A1", None))
      .put(a1("H1"), num(1))
      .put(a1("G1"), CellValue.Formula("COUNTIF(E1:E40,\">10\")/100+G2*0.3"))
      .put(a1("G2"), CellValue.Formula("SUMIFS(B1:B40,E1:E40,\"<=10\")/1000+G1*0.3"))
      .put(a1("K9"), CellValue.Formula("G1+G2"))
      .put(a1("K10"), CellValue.dataTable(kind, None))
    settings.foreach { iterative =>
      val run = assertLanesAgree(Workbook(sheet), kind.ref, iterative, iterative.scheme.toString)
      // the members fail on the scan that evaluates E5, so nothing is seeded, and says so
      assertEquals(seededValues(run.report, kind.ref), Vector.empty)
      val skipped = run.report.warnings.collect { case w: SeedTableWarning.Skipped => w.cells }
      val unresolved =
        run.report.warnings.collect { case w: SeedTableWarning.ConeUnresolved => w.refs }
      assertEquals(skipped, Vector(4))
      assertEquals(unresolved.flatten.size, 1, unresolved.toString)
      assertEquals(run.caches.size, 4)
      assertEquals(sliced(run, a1("E1")), Vector.fill(4)(Option.empty[CellValue]))
    }
  }

  test("#891: a 2-D table substitutes both inputs, one of them inside a criteria column") {
    // row input H1 (top axis K9:L9), column input D3 (left axis J10:J12), source in the corner J9
    val kind: FormulaKind.DataTable =
      FormulaKind.DataTable(
        range("K10:L12"),
        dt2D = true,
        dtr = false,
        Some(a1("H1")),
        Some(a1("D3"))
      )
    val sheet = grid
      .put(a1("H1"), num(1))
      .put(a1("K9"), num(1))
      .put(a1("L9"), num(2))
      .put(a1("J10"), num(0))
      .put(a1("J11"), num(3))
      .put(a1("J12"), num(7))
      .put(a1("G1"), CellValue.Formula("SUMIFS(B1:B40,D1:D40,\">2\")/1000+G2*0.3"))
      .put(a1("G2"), CellValue.Formula("COUNTIF(D1:D40,D3)/100*H1+G1*0.3"))
      .put(a1("J9"), CellValue.Formula("G1+G2"))
      .put(a1("K10"), CellValue.dataTable(kind, None))
    // as in the first case, with the count scaled by H1 (1, then 2) along each row
    val rows = Vector(("0.352", "0.09"), ("0.355", "0.09"), ("0.355", "0.01"))
    val expected =
      rows.flatMap((a, b) => Vector(1, 2).map(h => sumOfCycle(BigDecimal(a), BigDecimal(b) * h)))
    settings.foreach { iterative =>
      val run = assertLanesAgree(Workbook(sheet), kind.ref, iterative, iterative.scheme.toString)
      assertInterior(run, kind.ref, expected, iterative.scheme.toString)
      assertEquals(sliced(run, a1("D3")), Vector(0, 0, 3, 3, 7, 7).map(v => Option(num(v))))
    }
  }
