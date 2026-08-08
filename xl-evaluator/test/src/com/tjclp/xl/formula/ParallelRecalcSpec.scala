package com.tjclp.xl.formula

import java.time.{LocalDate, LocalDateTime}
import java.util.concurrent.atomic.{AtomicBoolean, AtomicInteger}
import java.util.concurrent.locks.LockSupport
import java.util.concurrent.{ConcurrentLinkedQueue, CountDownLatch, TimeUnit}

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.RecalcResult
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*

import scala.jdk.CollectionConverters.*

/**
 * GH-520: `recalculateParallel(n)` must be observationally IDENTICAL to `recalculate()` — same
 * workbook (caches included), same evaluated map, same error vector in the same order — for every
 * dependence the static graph can express. The wave partition guarantees it structurally (two
 * same-depth cells have no path between them, so neither can observe the other's write); these
 * gates pin the equality on the shapes most likely to break it: deep chains (waves below the
 * sequential cutoff), wide independent regions (the actually-parallel path), error-bearing books
 * (error ORDER is part of the contract), cycles (circular + blocked reporting precedes the pass),
 * and dynamic INDIRECT cells (excluded from waves, sequential evaluate-last).
 */
class ParallelRecalcSpec extends ScalaCheckSuite:

  private def num(d: Double): CellValue = CellValue.Number(BigDecimal(d))
  private def formula(expr: String): CellValue = CellValue.Formula(expr, None)

  private def assertSameResult(sequential: RecalcResult, parallel: RecalcResult): Unit =
    assertEquals(parallel.errors, sequential.errors)
    assertEquals(parallel.evaluated, sequential.evaluated)
    assertEquals(parallel.workbook, sequential.workbook)
    assertEquals(parallel.converged, sequential.converged)
    assertEquals(parallel.iterationsUsed, sequential.iterationsUsed)
    assertEquals(parallel.cycles, sequential.cycles)

  /**
   * 40 independent columns × 30 chained rows + a cross-sheet SUM roll-up: wide waves (40 ≥ cutoff).
   */
  private def wideGrid: Workbook =
    val cols = 40
    val rows = 30
    val data = (0 until cols).foldLeft(Sheet(SheetName.unsafe("Data"))) { (s0, c) =>
      (0 until rows).foldLeft(s0.put(ARef.from0(c, 0), num(100.0 + c))) { (s, r) =>
        if r == 0 then s
        else
          val colName = ARef.from0(c, r).toA1.takeWhile(_.isLetter)
          s.put(ARef.from0(c, r), formula(s"=$colName$r*1.01+SUM(Drivers!$$A$$1:$$A$$5)*0.001"))
      }
    }
    val drivers = (0 until 5).foldLeft(Sheet(SheetName.unsafe("Drivers"))) { (s, i) =>
      s.put(ARef.from0(0, i), num((i + 1) * 1.5))
    }
    val summary = Sheet(SheetName.unsafe("Summary"))
      .put(ARef.from0(0, 0), formula(s"=SUM(Data!A$rows:AN$rows)"))
    Workbook(drivers, data, summary)

  test("GH-520: wide independent grid — parallel(8) equals sequential, element for element"):
    val wb = wideGrid
    assertSameResult(wb.recalculate(), wb.recalculateParallel(8))

  test("GH-520: deep chain (waves below the cutoff) — parallel equals sequential"):
    val n = 200
    val chain =
      (2 to n).foldLeft(Sheet(SheetName.unsafe("Chain")).put(ARef.from0(0, 0), num(1.0))) {
        (s, i) => s.put(ARef.from0(0, i - 1), formula(s"=A${i - 1}*1.1+1"))
      }
    val wb = Workbook(chain)
    assertSameResult(wb.recalculate(), wb.recalculateParallel(4))

  test("GH-520: error-bearing book — errors, Excel errors and their ORDER are identical"):
    val cols = 20
    val s0 = Sheet(SheetName.unsafe("Errs"))
    // Row 0: constants; row 1: a wide wave where every third cell divides by zero and every
    // seventh calls an unknown function — both error planes, interleaved with healthy cells.
    val sheet = (0 until cols).foldLeft(s0) { (s, c) =>
      val colName = ARef.from0(c, 1).toA1.takeWhile(_.isLetter)
      val expr =
        if c % 3 == 0 then s"=$colName" + "1/0"
        else if c % 7 == 0 then s"=NOSUCHFN($colName" + "1)"
        else s"=$colName" + "1*2"
      s.put(ARef.from0(c, 0), num(c.toDouble)).put(ARef.from0(c, 1), formula(expr))
    }
    val wb = Workbook(sheet)
    val seq = wb.recalculate()
    val par = wb.recalculateParallel(8)
    assertSameResult(seq, par)
    assertEquals(par.excelErrors, seq.excelErrors)

  test("GH-520: cyclic core + blocked dependents — parallel equals sequential"):
    val a1 = ARef.from0(0, 0)
    val b1 = ARef.from0(1, 0)
    val c1 = ARef.from0(2, 0)
    // A1 <-> B1 cycle, C1 blocked downstream, plus a wide healthy region.
    val base = Sheet(SheetName.unsafe("Cyc"))
      .put(a1, formula("=B1+1"))
      .put(b1, formula("=A1+1"))
      .put(c1, formula("=A1*2"))
    val sheet = (0 until 20).foldLeft(base) { (s, c) =>
      s.put(ARef.from0(c, 2), num(c.toDouble))
        .put(ARef.from0(c, 3), formula(s"=${ARef.from0(c, 3).toA1.takeWhile(_.isLetter)}3+1"))
    }
    val wb = Workbook(sheet)
    assertSameResult(wb.recalculate(), wb.recalculateParallel(8))

  test("GH-520: INDIRECT bucket stays sequential — parallel equals sequential"):
    val base = Sheet(SheetName.unsafe("Dyn"))
      .put(ARef.from0(0, 0), num(7.0))
      .put(ARef.from0(1, 0), formula("=INDIRECT(\"A1\")*3"))
      .put(ARef.from0(2, 0), formula("=B1+1")) // static dependent of a dynamic cell — bucketed
    val sheet = (0 until 20).foldLeft(base) { (s, c) =>
      s.put(ARef.from0(c, 2), num(c.toDouble))
        .put(ARef.from0(c, 3), formula(s"=${ARef.from0(c, 3).toA1.takeWhile(_.isLetter)}3*2"))
    }
    val wb = Workbook(sheet)
    assertSameResult(wb.recalculate(), wb.recalculateParallel(8))

  test("GH-520: repeated parallel runs are deterministic"):
    val wb = wideGrid
    val first = wb.recalculateParallel(8)
    (1 to 3).foreach { _ =>
      assertSameResult(first, wb.recalculateParallel(8))
    }

  property("GH-520: random wide acyclic books preserve the full RecalcResult law"):
    val width = 20
    val cases = for
      constants <- Gen.listOfN(width, Gen.choose(-100, 100))
      factors <- Gen.listOfN(width, Gen.choose(-5, 5))
      offsets <- Gen.listOfN(width, Gen.choose(-20, 20))
      leftLinks <- Gen.listOfN(width, Gen.choose(0, width - 1))
      rightLinks <- Gen.listOfN(width, Gen.choose(0, width - 1))
    yield (constants, factors, offsets, leftLinks, rightLinks)

    forAll(cases) { case (constants, factors, offsets, leftLinks, rightLinks) =>
      val base = constants.zipWithIndex.foldLeft(Sheet(SheetName.unsafe("Generated"))) {
        case (sheet, (value, col)) => sheet.put(ARef.from0(col, 0), num(value.toDouble))
      }
      val firstLayer = (0 until width).foldLeft(base) { (sheet, col) =>
        val input = ARef.from0(col, 0).toA1
        sheet.put(
          ARef.from0(col, 1),
          formula(s"=$input*${factors(col)}+${offsets(col)}")
        )
      }
      val secondLayer = (0 until width).foldLeft(firstLayer) { (sheet, col) =>
        val left = ARef.from0(leftLinks(col), 1).toA1
        val right = ARef.from0(rightLinks(col), 1).toA1
        sheet.put(ARef.from0(col, 2), formula(s"=$left+$right"))
      }
      val wb = Workbook(secondLayer)

      assertSameResult(wb.recalculate(), wb.recalculateParallel(8))
    }

  test("GH-520: parallelism 1 and an absurd width both degrade safely"):
    val wb = wideGrid
    assertSameResult(wb.recalculate(), wb.recalculateParallel(1))
    assertSameResult(wb.recalculate(), wb.recalculateParallel(10000))

  test("GH-520: multi-hop defined name hiding INDIRECT stays out of parallel waves"):
    val sheetName = SheetName.unsafe("Dyn")
    val target = ARef.from0(0, 0)
    val reader = ARef.from0(19, 0)
    val staticNameReader = ARef.from0(20, 0)
    val staleTarget = CellValue.Formula("=1+1", Some(num(999)))
    val base = Sheet(sheetName).put(target, staleTarget)
    val sheet = (1 until 19)
      .foldLeft(base) { (s, col) =>
        s.put(ARef.from0(col, 0), formula(s"=${col + 1}+1"))
      }
      .put(reader, formula("=OuterDyn"))
      .put(staticNameReader, formula("=Case"))
    val wb = Workbook(sheet)
      .withDefinedName("InnerDyn", "INDIRECT(\"A1\")")
      .withDefinedName("OuterDyn", "InnerDyn")
      .withDefinedName("Case", "42")

    assertEquals(
      DependencyGraph.dynamicCells(wb),
      Set(QualifiedRef(sheetName, reader)),
      "only the name chain that reaches INDIRECT is dynamic; ordinary Case stays parallel-safe"
    )

    val sequential = wb.recalculate()
    val parallel = wb.recalculateParallel(8)
    assertSameResult(sequential, parallel)
    assertEquals(parallel.evaluated(sheetName)(reader), num(2))

  test("GH-520: dynamic-name analysis honors multi-hop sheet-scoped shadowing"):
    val calcName = SheetName.unsafe("Calc")
    val otherName = SheetName.unsafe("Other")
    val reader = ARef.from0(19, 0)
    val calc = Sheet(calcName)
      .put(ARef.from0(0, 0), num(7))
      .put(reader, formula("=Outer"))
    val other = Sheet(otherName)
      .put(ARef.from0(0, 0), num(9))
      .put(reader, formula("=Outer"))
    val base = Workbook(calc, other)
      .withDefinedName("Driver", "41")
      .withDefinedName("Outer", "Driver")
    val wb = base.copy(
      metadata = base.metadata.copy(
        definedNames = base.metadata.definedNames :+
          DefinedName("Driver", "OFFSET(A1, 0, 0)", localSheetId = Some(0))
      )
    )

    assertEquals(
      DependencyGraph.dynamicCells(wb),
      Set(QualifiedRef(calcName, reader)),
      "Calc's local Driver shadows the static workbook Driver only on Calc"
    )

    val result = wb.recalculateParallel(8)
    assert(result.isClean, result.errors.map(_.render).mkString("; "))
    assertEquals(result.evaluated(calcName)(reader), num(7))
    assertEquals(result.evaluated(otherName)(reader), num(41))

  test("GH-520: one pinned Clock generation is shared safely by every wave worker"):
    final class GuardClock extends Clock:
      val todayCalls = AtomicInteger(0)
      val nowCalls = AtomicInteger(0)
      val overlap = AtomicBoolean(false)
      private val active = AtomicBoolean(false)

      private def guarded[A](value: A): A =
        if !active.compareAndSet(false, true) then overlap.set(true)
        try
          LockSupport.parkNanos(2_000_000L)
          value
        finally active.set(false)

      def today(): LocalDate =
        todayCalls.incrementAndGet()
        guarded(LocalDate.of(2026, 8, 7))

      def now(): LocalDateTime =
        nowCalls.incrementAndGet()
        guarded(LocalDateTime.of(2026, 8, 7, 12, 30))

    val sheet = (0 until 20).foldLeft(Sheet(SheetName.unsafe("Clock"))) { (s, col) =>
      val expr = if col % 2 == 0 then "=TODAY()" else "=NOW()"
      s.put(ARef.from0(col, 0), formula(expr))
    }
    val wb = Workbook(sheet)
    val sequentialClock = new GuardClock
    val parallelClock = new GuardClock
    val sequential = wb.recalculate(sequentialClock)
    val parallel = wb.recalculateParallel(parallelClock, 8)

    assertSameResult(sequential, parallel)
    assertEquals(sequentialClock.todayCalls.get(), 1)
    assertEquals(sequentialClock.nowCalls.get(), 1)
    assertEquals(parallelClock.todayCalls.get(), 1)
    assertEquals(parallelClock.nowCalls.get(), 1)
    assert(!parallelClock.overlap.get(), "the caller's Clock must never be invoked concurrently")

  test("GH-520: pinned Clock forces only the volatile capability formulas request"):
    val clock = new Clock:
      def today(): LocalDate = throw new IllegalStateException("TODAY was not requested")
      def now(): LocalDateTime = LocalDateTime.of(2026, 8, 8, 12, 30)
    val sheet = (0 until 20).foldLeft(Sheet(SheetName.unsafe("NowOnly"))) { (s, col) =>
      s.put(ARef.from0(col, 0), formula("=NOW()"))
    }

    val result = Workbook(sheet).recalculateParallel(clock, 8)

    assert(result.isClean, result.errors.map(_.render).mkString("; "))

  test("GH-520: caller interruption cancels workers, restores interrupt, and stays total"):
    val started = CountDownLatch(1)
    val release = CountDownLatch(1)
    val finished = CountDownLatch(1)
    val interruptedAtReturn = AtomicBoolean(false)
    val results = new ConcurrentLinkedQueue[RecalcResult]()
    val clock = new Clock:
      def today(): LocalDate = LocalDate.of(2026, 8, 7)

      def now(): LocalDateTime =
        started.countDown()
        try
          release.await()
          LocalDateTime.of(2026, 8, 7, 12, 30)
        finally finished.countDown()

    val sheet = (0 until 20).foldLeft(Sheet(SheetName.unsafe("Interrupt"))) { (s, col) =>
      s.put(ARef.from0(col, 0), formula("=NOW()"))
    }
    val wb = Workbook(sheet)
    val caller = new Thread(
      () =>
        val result = wb.recalculateParallel(clock, 8)
        results.add(result)
        interruptedAtReturn.set(Thread.currentThread().isInterrupted())
      ,
      "parallel-recalc-interruption-test"
    )

    caller.start()
    assert(started.await(5, TimeUnit.SECONDS), "worker never entered the blocking Clock")
    caller.interrupt()
    caller.join(10_000L)
    release.countDown() // failure-path cleanup if an implementation regresses

    assert(!caller.isAlive, "interrupted recalculateParallel caller must return")
    assert(finished.await(5, TimeUnit.SECONDS), "the canceled Clock worker must terminate")
    assert(interruptedAtReturn.get(), "recalculateParallel must restore the caller interrupt")
    val result = Option(results.poll()).getOrElse(fail("interrupted recalc returned no result"))
    assertEquals(result.errors.size, 20)
    assert(result.errors.forall(_.error.message.contains("interrupted")))
    val liveWorkers = Thread.getAllStackTraces.keySet().asScala.filter { thread =>
      thread.isAlive && thread.getName.startsWith("xl-recalc-worker-")
    }
    assertEquals(liveWorkers.toVector, Vector.empty)

  test("GH-520: worker interruption is reported per cell and shuts down the pool"):
    val clock = new Clock:
      def today(): LocalDate = LocalDate.of(2026, 8, 7)
      def now(): LocalDateTime = throw new InterruptedException("injected worker failure")
    val sheet = (0 until 20).foldLeft(Sheet(SheetName.unsafe("WorkerFailure"))) { (s, col) =>
      s.put(ARef.from0(col, 0), formula("=NOW()"))
    }

    val result = Workbook(sheet).recalculateParallel(clock, 8)
    assertEquals(result.errors.size, 20)
    assert(result.errors.forall(_.error.message.contains("worker failed")))
    val liveWorkers = Thread.getAllStackTraces.keySet().asScala.filter { thread =>
      thread.isAlive && thread.getName.startsWith("xl-recalc-worker-")
    }
    assertEquals(liveWorkers.toVector, Vector.empty)
