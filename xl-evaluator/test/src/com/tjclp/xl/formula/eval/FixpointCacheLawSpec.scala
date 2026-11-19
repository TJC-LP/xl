package com.tjclp.xl.formula.eval

import scala.util.Random

import com.tjclp.xl.CellRange
import com.tjclp.xl.addressing.{ARef, Column, Row, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * #891: the law of the criteria cache one iterative fixpoint shares
 * ([[Evaluator.AggregateMemo.fixpointCache]]). Over generated fixpoints, three runs give the same
 * [[WorkbookEvaluator.FixpointOutcome]] (results, rounds, converged, stalled, largest delta) under
 * Jacobi and under Gauss–Seidel:
 *
 *   - the shared cache (`jacobiFixpoint`'s default);
 *   - `columnCache = false`, the same evaluator with nothing shared;
 *   - a plain `Evaluator.instance`, which has no memo at all (what the data-table seeder's fixpoint
 *     ran before it shared the cache).
 *
 * The generator covers two sheets whose workbook order differs from their name order; whole-column
 * and windowed criteria ranges, with and without members in them; cached and uncached formulas
 * outside the fixpoint inside criteria columns; members as criterion values; defined names as
 * criteria ranges; an array-formula member; and members whose evaluation fails, so the held value
 * and the `stalled` exit run.
 *
 * Seeds are fixed, so the coverage asserted at the end of each test is exact, not probabilistic: a
 * change that stops the generator from building slices, or from failing members, fails the suite
 * instead of leaving the law vacuous.
 */
class FixpointCacheLawSpec extends FunSuite:

  private val rows = 40
  // workbook order ("b", "Z") is not name order ("Z" sorts before "b")
  private val names = Vector("b", "Z")
  private def at(c: Int, r: Int): ARef = ARef.from0(c, r)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def colL(c: Int): String = Column.from0(c).toLetter

  private final case class Case(
    wb: Workbook,
    members: List[(QualifiedRef, Int, String)],
    description: String
  )

  /**
   * One fixpoint: three to six members in columns C and F of either sheet, each the sum of three
   * criteria terms and a read of another member. Columns A (text keys), B, E (numbers) are
   * constants; C and F hold numbers and the members; D mixes constants with cached and uncached
   * formulas, some of which read a member. With `failing`, one member's formula also reads a sheet
   * the workbook does not hold, so its evaluation is a `Left` every round.
   */
  private def gen(rnd: Random, failing: Boolean): Case =
    val slots =
      for
        s <- 0 to 1
        (c, maxRow) <- List((2, 5), (5, 3))
        r <- 0 to maxRow
      yield (s, c, r)
    val memberSlots = rnd.shuffle(slots.toVector).take(3 + rnd.nextInt(4))
    def q(s: Int, c: Int, r: Int) = s"'${names(s)}'!${colL(c)}${r + 1}"

    // (first row, last row, whole column): rows are 1-based
    def shape(): (Int, Int, Boolean) = rnd.nextInt(5) match
      case 0 => (1, 40, false)
      case 1 => (17, 40, false)
      case 2 => (3, 30, false)
      case 3 => (20, 38, false)
      case _ => (0, 0, true)
    def rng(s: Int, c: Int, sh: (Int, Int, Boolean)): String =
      if sh._3 then s"'${names(s)}'!${colL(c)}:${colL(c)}"
      else s"'${names(s)}'!${colL(c)}${sh._1}:${colL(c)}${sh._2}"
    // a criteria range and the column it reads (which decides the criterion's type)
    def critRange(sh: (Int, Int, Boolean)): (String, Int) =
      rnd.nextInt(12) match
        case 0 if !sh._3 && sh._1 == 1 => ("nmKeys", 0)
        case 1 if !sh._3 && sh._1 == 1 => ("nmE", 4)
        case _ =>
          val c = rnd.nextInt(6)
          (rng(rnd.nextInt(2), c, sh), c)
    def criterion(col: Int): String =
      val m = memberSlots(rnd.nextInt(memberSlots.size))
      if col == 0 then
        rnd.nextInt(4) match
          case 0 => "\"x\""
          case 1 => "\"y\""
          case 2 => "\"<>z\""
          case _ => "\"z\""
      else
        rnd.nextInt(7) match
          case 0 => "\">0\""
          case 1 => "\"<5\""
          case 2 => "2"
          case 3 => q(m._1, m._2, m._3)
          case 4 => s"\">\"&${q(m._1, m._2, m._3)}"
          case 5 => "\"<>3\""
          case _ => "\">=1\""
    def sumRange(sh: (Int, Int, Boolean)): String = rng(rnd.nextInt(2), 1 + rnd.nextInt(5), sh)
    def term(): String =
      val sh = shape()
      val (r1, c1) = critRange(sh)
      val (r2, c2) = critRange(sh)
      rnd.nextInt(8) match
        case 0 => s"SUMIFS(${sumRange(sh)},$r1,${criterion(c1)})"
        case 1 => s"SUMIFS(${sumRange(sh)},$r1,${criterion(c1)},$r2,${criterion(c2)})"
        case 2 => s"COUNTIFS($r1,${criterion(c1)},$r2,${criterion(c2)})"
        case 3 => s"COUNTIF($r1,${criterion(c1)})"
        case 4 => s"SUMIF($r1,${criterion(c1)},${sumRange(sh)})"
        case 5 => s"IFERROR(AVERAGEIFS(${sumRange(sh)},$r1,${criterion(c1)}),0)"
        case 6 => s"MAXIFS(${sumRange(sh)},$r1,${criterion(c1)})"
        case _ => s"SUMIF($r1,${criterion(c1)})"

    val failingSlot = Option.when(failing)(memberSlots(rnd.nextInt(memberSlots.size)))
    val memberFormulas = memberSlots.map { case (s, c, r) =>
      val other = memberSlots(rnd.nextInt(memberSlots.size))
      val body = s"(${term()}+${term()}+${term()})/1000+${q(other._1, other._2, other._3)}*0.3+1"
      val text = if failingSlot.contains((s, c, r)) then s"=$body+$FailingRead" else s"=$body"
      (s, c, r, text)
    }
    val arrayMember = Option.when(rnd.nextInt(3) == 0)(memberSlots(0))

    val sheets = names.indices.toVector.map { s =>
      val base = (0 until rows).foldLeft(Sheet(SheetName.unsafe(names(s)))) { (sheet, r) =>
        val keyed = sheet
          .put(at(0, r), CellValue.Text(Vector("x", "y", "z")(rnd.nextInt(3))))
          .put(at(1, r), num(rnd.nextInt(50)))
          .put(at(2, r), num(rnd.nextInt(9)))
          .put(at(4, r), num(rnd.nextInt(5)))
          .put(at(5, r), num(rnd.nextInt(4)))
        val m = memberSlots(rnd.nextInt(memberSlots.size))
        val d: CellValue = rnd.nextInt(14) match
          case 0 => CellValue.Formula(s"=B${r + 1}*2", None)
          case 1 => CellValue.Formula(s"=B${r + 1}+1", Some(num(rnd.nextInt(7))))
          case 2 => CellValue.Formula(s"=${q(m._1, m._2, m._3)}*2", Some(num(rnd.nextInt(7))))
          case 3 if s == 1 => CellValue.Formula(s"=${q(m._1, m._2, m._3)}+B${r + 1}", None)
          case _ => num(rnd.nextInt(6))
        // half the time sheet "b" keeps column D free of formulas, so D can be sliced there
        if s == 0 && rnd.nextBoolean() then keyed.put(at(3, r), num(rnd.nextInt(6)))
        else keyed.put(at(3, r), d)
      }
      memberFormulas.filter(_._1 == s).foldLeft(base) { case (sheet, (_, c, r, text)) =>
        val kind: FormulaKind =
          if arrayMember.contains((s, c, r)) then
            FormulaKind.dynamicArray(CellRange(at(c, r), at(c, r)))
          else FormulaKind.Normal()
        sheet.put(at(c, r), CellValue.Formula(text, None, kind))
      }
    }
    val wb = Workbook(sheets)
      .withDefinedName("nmKeys", "'Z'!$A$1:$A$40")
      .withDefinedName("nmE", "'b'!$E$1:$E$40")
    val members = memberFormulas.toList.map { case (s, c, r, text) =>
      (QualifiedRef(SheetName.unsafe(names(s)), at(c, r)), s, text)
    }
    Case(wb, members, memberFormulas.map(m => s"${q(m._1, m._2, m._3)} ${m._4}").mkString("\n"))

  /** A read no workbook here can serve: the member's evaluation is a `Left`, not an error value. */
  private val FailingRead = "'NoSuchSheet'!A1"

  private enum Lane derives CanEqual:
    /** `jacobiFixpoint`'s default: the fixpoint's shared cache. */
    case Shared

    /** `columnCache = false`: the same evaluator, nothing shared. */
    case Unshared

    /** `Evaluator.instance`: no memo at all. */
    case Plain

    /** The shared lane over a cache the test built, with its own limits and live counters. */
    case Limited(cache: PassCache, counters: RecalcCounters)

  private def iterative(scheme: IterationScheme): IterativeCalc =
    IterativeCalc(40, BigDecimal("1E-9"), scheme = scheme)

  /** The outcome of `c` on `lane`, and every cache its region memos carried. */
  private def run(
    c: Case,
    scheme: IterationScheme,
    lane: Lane
  ): (WorkbookEvaluator.FixpointOutcome, Vector[Evaluator.AggregateMemo]) =
    val memos = Vector.newBuilder[Evaluator.AggregateMemo]
    val factory: (Rng, Evaluator.AggregateMemo) => Evaluator = lane match
      case Lane.Plain => (rng, _) => Evaluator.instance(rng)
      case Lane.Limited(cache, counters) =>
        (rng, _) =>
          // a region memo as `AggregateMemo.inFixpoint` makes it, counting its criteria calls
          val memo = new Evaluator.AggregateMemo(counters = counters, shared = Some(cache))
          memos += memo
          Evaluator.recalculationInstance(rng, memo)
      case _ =>
        (rng, memo) =>
          memos += memo
          Evaluator.recalculationInstance(rng, memo)
    val outcome = WorkbookEvaluator.jacobiFixpoint(
      c.wb,
      c.wb.sheets,
      c.members,
      iterative(scheme),
      Clock.system,
      Rng.system,
      factory,
      Map.empty,
      columnCache = lane != Lane.Unshared
    )
    (outcome, memos.result())

  /** The one cache every region memo of a shared run carries. */
  private def sharedCache(memos: Vector[Evaluator.AggregateMemo], clue: => String): PassCache =
    // a fresh memo per region (aggregate results are never shared) over one cache
    assertEquals(memos.map(System.identityHashCode).distinct.size, memos.size, clue)
    memos.flatMap(_.passCache).distinctBy(System.identityHashCode) match
      case Vector(only) => only
      case other => fail(s"${other.size} caches in one fixpoint\n$clue")

  private def wholeColumn(slice: ColumnSlice): Boolean =
    slice.from == 0 && slice.to == Row.MaxIndex0

  /** No slice may hold a row a member occupies: members are the cells a fixpoint changes. */
  private def assertNoMemberRowHeld(c: Case, slices: Vector[ColumnSlice], clue: => String): Unit =
    for
      slice <- slices
      (member, _, _) <- c.members
      if member.sheet == slice.sheet && member.ref.col.index0 == slice.col
    do
      val row = member.ref.row.index0
      val held = row >= slice.firstHeldRow && row < slice.firstHeldRow + slice.heldRows
      assert(!held, s"slice ${slice.id} holds member ${member.ref.toA1}\n$clue")

  private val schemes = List(IterationScheme.Jacobi, IterationScheme.GaussSeidel)

  /** What a run of cases exercised; asserted so the law cannot go vacuous unnoticed. */
  private final class Coverage:
    var runs = 0
    var runsWithSlices = 0
    var wholeColumnSlices = 0
    var boxSlices = 0
    var converged = 0
    var exhausted = 0
    var stalled = 0
    var failedMembers = 0

    def record(outcome: WorkbookEvaluator.FixpointOutcome, slices: Vector[ColumnSlice]): Unit =
      runs += 1
      if slices.nonEmpty then runsWithSlices += 1
      wholeColumnSlices += slices.count(wholeColumn)
      boxSlices += slices.count(s => !wholeColumn(s))
      if outcome.converged then converged += 1
      else if outcome.stalled then stalled += 1
      else exhausted += 1
      failedMembers += outcome.results.valuesIterator.count(_.isLeft)

    override def toString: String =
      s"runs=$runs runsWithSlices=$runsWithSlices wholeColumnSlices=$wholeColumnSlices " +
        s"boxSlices=$boxSlices converged=$converged exhausted=$exhausted stalled=$stalled " +
        s"failedMembers=$failedMembers"

  /** The three lanes agree on `c`; returns the shared lane's outcome and slices. */
  private def assertLanesAgree(
    c: Case,
    scheme: IterationScheme,
    seed: Int
  ): (WorkbookEvaluator.FixpointOutcome, Vector[ColumnSlice]) =
    def clue = s"seed $seed, $scheme\n${c.description}"
    val (shared, memos) = run(c, scheme, Lane.Shared)
    val (unshared, unsharedMemos) = run(c, scheme, Lane.Unshared)
    val (plain, _) = run(c, scheme, Lane.Plain)
    assertEquals(shared, unshared, s"shared cache != columnCache = false\n$clue")
    assertEquals(unshared, plain, s"recalculation evaluator != Evaluator.instance\n$clue")
    assertEquals(unsharedMemos.flatMap(_.passCache), Vector.empty, clue)
    val slices = sharedCache(memos, clue).builtSlices
    assertNoMemberRowHeld(c, slices, clue)
    (shared, slices)

  test("#891: shared cache == columnCache = false == Evaluator.instance, both schemes") {
    val coverage = new Coverage
    (0 until Cases).foreach { i =>
      val seed = 1000 + i
      val c = gen(new Random(seed), failing = false)
      schemes.foreach { scheme =>
        val (outcome, slices) = assertLanesAgree(c, scheme, seed)
        coverage.record(outcome, slices)
      }
    }
    // non-vacuity: most fixpoints sliced something, by both kinds of slice, and both exits ran
    assert(coverage.runsWithSlices * 2 > coverage.runs, coverage.toString)
    assert(coverage.wholeColumnSlices > 0, coverage.toString)
    assert(coverage.boxSlices > 0, coverage.toString)
    assert(coverage.converged > 0, coverage.toString)
    assert(coverage.exhausted > 0, coverage.toString)
    assertEquals(coverage.failedMembers, 0, coverage.toString)
  }

  test("#891: the lanes agree when a member fails every round (held value, stalled exit)") {
    val coverage = new Coverage
    (0 until FailingCases).foreach { i =>
      val seed = 7000 + i
      val c = gen(new Random(seed), failing = true)
      schemes.foreach { scheme =>
        val (outcome, slices) = assertLanesAgree(c, scheme, seed)
        // the failing member is a Left in the outcome, and no fixpoint with one converges
        assertEquals(outcome.results.valuesIterator.count(_.isLeft), 1, c.description)
        assert(!outcome.converged, c.description)
        coverage.record(outcome, slices)
      }
    }
    assertEquals(coverage.failedMembers, coverage.runs, coverage.toString)
    assert(coverage.stalled > 0, coverage.toString)
    assert(coverage.runsWithSlices > 0, coverage.toString)
  }

  // ----- the cache's budgets (PassCache.Limits) -----

  private val roomy: PassCache.Limits = Evaluator.AggregateMemo.FixpointLimits

  /** What the budget cases did under one `Limits`, summed over the cases. */
  private final case class Budgeted(
    slices: Long,
    heldRowsMax: Long,
    indexed: Long,
    scans: Long,
    fallbacks: Long,
    // slices left without an equality index once the fixpoint ended
    indexesRefused: Int
  )

  /**
   * Every budget case under `limits`, each compared with its `columnCache = false` outcome. The
   * cache is built as [[Evaluator.AggregateMemo.fixpointCache]] builds it, with live counters.
   */
  private def budgeted(limits: PassCache.Limits): Budgeted =
    val totals = RecalcCounters.create()
    var heldRowsMax = 0L
    var indexesRefused = 0
    (0 until BudgetCases).foreach { i =>
      val seed = 1000 + i
      val c = gen(new Random(seed), failing = i % 4 == 3)
      schemes.foreach { scheme =>
        val cache = new PassCache(limits, totals)
        cache.begin(Readiness.fixpoint(c.members.map(_._1)))
        val (limited, _) = run(c, scheme, Lane.Limited(cache, totals))
        val (unshared, _) = run(c, scheme, Lane.Unshared)
        def clue = s"seed $seed, $scheme, $limits\n${c.description}"
        assertEquals(limited, unshared, clue)
        val slices = cache.builtSlices
        assertNoMemberRowHeld(c, slices, clue)
        val held = slices.iterator.map(_.heldRows.toLong).sum
        assert(held <= limits.maxCachedCells, s"$held cells held\n$clue")
        heldRowsMax = math.max(heldRowsMax, held)
        indexesRefused += slices.count(cache.eqIndex(_).isEmpty)
      }
    }
    val snap = totals.snapshot
    import RecalcCounters.Counter
    Budgeted(
      snap(Counter.SlicesBuilt),
      heldRowsMax,
      snap(Counter.CriteriaIndexed),
      snap(Counter.CriteriaScan),
      snap(Counter.ScanFallbacks),
      indexesRefused
    )

  private lazy val unbounded: Budgeted = budgeted(roomy)

  test("#891: within its budgets a fixpoint's cache slices and indexes") {
    assert(unbounded.slices > 0, unbounded.toString)
    assert(unbounded.indexed > 0, unbounded.toString)
    assertEquals(unbounded.indexesRefused, 0, unbounded.toString)
  }

  test("#891: a fixpoint whose cache may hold no cells copies nothing and scans every call") {
    val capped = budgeted(roomy.copy(maxCachedCells = 0L))
    assertEquals(capped.slices, 0L, capped.toString)
    assertEquals(capped.indexed, 0L, capped.toString)
    assertEquals(capped.scans, unbounded.scans + unbounded.indexed, capped.toString)
    assert(capped.fallbacks > unbounded.fallbacks, s"$capped against $unbounded")
  }

  // the sheets hold 40 used rows: 40 cells is one whole column, 100 is two and a window
  List(40L, 100L).foreach { cells =>
    test(s"#891: a fixpoint whose cache may hold $cells cells scans the columns past the budget") {
      val capped = budgeted(roomy.copy(maxCachedCells = cells))
      // the budget bit (fewer slices, more fallbacks) without switching the cache off
      assert(capped.slices > 0 && capped.slices < unbounded.slices, s"$capped against $unbounded")
      assert(capped.heldRowsMax > cells / 2, capped.toString)
      assert(capped.indexed > 0 && capped.indexed < unbounded.indexed, capped.toString)
      assert(capped.fallbacks > unbounded.fallbacks, s"$capped against $unbounded")
    }
  }

  test("#891: a fixpoint whose indexes may hold one byte slices as before and indexes nothing") {
    val capped = budgeted(roomy.copy(maxIndexBytes = 1L))
    assertEquals(capped.slices, unbounded.slices, capped.toString)
    assertEquals(capped.fallbacks, unbounded.fallbacks, capped.toString)
    assertEquals(capped.indexed, 0L, capped.toString)
    assertEquals(capped.indexesRefused.toLong, capped.slices, capped.toString)
  }

  List(512L, 4096L).foreach { bytes =>
    test(s"#891: a fixpoint whose indexes may hold $bytes bytes indexes some columns only") {
      val capped = budgeted(roomy.copy(maxIndexBytes = bytes))
      assertEquals(capped.slices, unbounded.slices, capped.toString)
      assert(capped.indexed > 0 && capped.indexed < unbounded.indexed, s"$capped, $unbounded")
      assert(capped.indexesRefused > 0, capped.toString)
    }
  }

  private val Cases = 40
  private val BudgetCases = 8
  private val FailingCases = 20
