package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.formula.eval.RecalcCounters.Counter
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/** GH-747: lookups and returned references share the aggregate pass's used-range memo. */
class WholeAxisUsedRangeSpec extends FunSuite:
  private val data = SheetName.unsafe("Data")
  private val out = SheetName.unsafe("Out")
  private val options = RecalcOptions(iterative = IterativeMode.Off)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def at(c: Int, r: Int): ARef = ARef.from0(c, r)

  private val source = Sheet(data)
    .put(at(0, 0), num(1))
    .put(at(0, 1), num(2))
    .put(at(1, 0), num(10))
    .put(at(1, 1), num(20))

  private def book(formulas: Vector[String], target: Sheet = source): Workbook =
    val output = formulas.zipWithIndex.foldLeft(Sheet(out)) { case (s, (f, r)) =>
      s.put(at(0, r), CellValue.Formula(f))
    }
    Workbook(target, output)

  private def check(wb: Workbook, expected: Vector[CellValue], scans: Long): Unit =
    Vector(1, 4).foreach { parallelism =>
      val counters = RecalcCounters.create()
      val result = recalculateCounted(wb, options.copy(parallelism = parallelism), counters)
      assert(result.isClean, result.errors.toString)
      assertEquals(expected.indices.map(r => result.evaluated(out)(at(0, r))).toVector, expected)
      assertEquals(counters.snapshot(Counter.UsedRangeScans), scans)
      val noCache = recalculateWithMemo(
        wb,
        options.copy(parallelism = parallelism),
        new Evaluator.AggregateMemo(cachesAggregates = false)
      )
      assertEquals(noCache.evaluated, result.evaluated)
      val uncached = recalculateUncachedCounted(
        wb,
        options.copy(parallelism = parallelism),
        RecalcCounters.create()
      )
      assertEquals(uncached.evaluated, result.evaluated)
    }

  Vector(
    "XLOOKUP(2,Data!A:A,Data!B:B)" -> num(20),
    "XMATCH(2,Data!A:A)" -> num(2),
    "MATCH(2,Data!A:A,0)" -> num(2),
    "VLOOKUP(2,Data!A:B,2,FALSE)" -> num(20),
    "HLOOKUP(10,Data!1:2,2,FALSE)" -> num(20),
    "LOOKUP(2,Data!A:A,Data!B:B)" -> num(20),
    "LOOKUP(2,Data!A:B)" -> num(20),
    "SUM(INDEX(Data!B:B,0))" -> num(30),
    "SUM(INDIRECT(\"Data!B:B\"))" -> num(30),
    "SUM(OFFSET(Data!B:B,0,0))" -> num(30),
    "AND(Data!A:A)" -> CellValue.Bool(true)
  ).foreach { case (formula, expected) =>
    test(s"GH-747: $formula shares one used-range scan across repeated calls") {
      check(book(Vector.fill(8)(formula)), Vector.fill(8)(expected), 1L)
    }
  }

  test("GH-747: mixed aggregates and lookups share one memo per sheet") {
    val other = Sheet(SheetName.unsafe("Other")).put(at(0, 0), num(9))
    val wb = book(
      Vector(
        "SUM(Data!B:B)",
        "MATCH(2,Data!A:A,0)",
        "SUM(INDEX(Data!B:B,0))",
        "MATCH(9,Other!A:A,0)"
      )
    )
    check(wb.put(other), Vector(30, 2, 30, 1).map(num), 2L)
  }

  test("GH-747: same-sheet formula-cache writes keep one bounds scan") {
    val sheet = (0 until 8).foldLeft(source) { (s, r) =>
      s.put(at(2, r), CellValue.Formula("XLOOKUP(2,A:A,B:B)"))
    }
    Vector(1, 4).foreach { parallelism =>
      val counters = RecalcCounters.create()
      val result =
        recalculateCounted(Workbook(sheet), options.copy(parallelism = parallelism), counters)
      assert(result.isClean, result.errors.toString)
      assertEquals(result.evaluated(data).values.toVector, Vector.fill(8)(num(20)))
      assertEquals(counters.snapshot(Counter.UsedRangeScans), 1L)
    }
  }

  test("GH-747: bounded references never scan sheet bounds") {
    val formulas = Vector(
      "XLOOKUP(2,Data!A1:A2,Data!B1:B2)",
      "MATCH(2,Data!A1:A2,0)",
      "SUM(INDEX(Data!B1:B2,0))",
      "AND(Data!A1:A2)"
    )
    check(book(formulas), Vector(num(20), num(2), num(30), CellValue.Bool(true)), 0L)
  }

  test("GH-747: empty whole axes cache an absent used range") {
    check(
      book(Vector.fill(8)("XLOOKUP(1,Data!A:A,Data!B:B,99)"), Sheet(data)),
      Vector.fill(8)(num(99)),
      1L
    )
  }

  test("GH-747: whole-axis returned arrays reuse the same bounds as lookup vectors") {
    val counters = RecalcCounters.create()
    val memo = new Evaluator.AggregateMemo(counters = counters)
    memo.passCache.foreach(_.begin(Readiness.build(Map.empty, Set.empty)))
    val evaluator = Evaluator.recalculationInstance(Rng.system, memo).withArrayResults
    val wb = Workbook(source)
    val expected = ArrayResult(Vector(Vector(num(10)), Vector(num(20))))
    Vector(
      "INDEX(B:B,0)" -> expected,
      "XLOOKUP(10,A1:B1,A:B)" -> expected,
      "INDIRECT(\"B:B\")" -> expected,
      "OFFSET(B:B,0,0)" -> expected,
      "INDEX(B5:B1048576,0)" -> ArrayResult.single(CellValue.Empty),
      "XLOOKUP(2,A:A,B:XFD)" -> ArrayResult.single(num(20))
    ).foreach { case (formula, array) =>
      val expr = FormulaParser.parse(formula).fold(e => fail(e.toString), identity)
      (0 until 8).foreach { _ =>
        val result = evaluator.eval(expr, source, workbook = Some(wb)).map(v => v: Any)
        assertEquals(result, Right(array): Either[EvalError, Any], formula)
      }
    }
    assertEquals(counters.snapshot(Counter.UsedRangeScans), 1L)
  }

  test("GH-747: lookup bounds refresh after structural changes and at the next pass") {
    val counters = RecalcCounters.create()
    val memo = new Evaluator.AggregateMemo(counters = counters)
    val cache = memo.passCache.getOrElse(fail("missing pass cache"))
    val evaluator = Evaluator.recalculationInstance(Rng.system, memo)
    val output = Sheet(out).put(at(0, 0), CellValue.Formula("MATCH(3,Data!A:A,0)"))
    def evaluate(target: Sheet): CellValue =
      SheetEvaluator
        .evaluateCellWithEvaluator(
          output,
          at(0, 0),
          evaluator,
          Clock.system,
          Some(Workbook(target, output))
        )
        .fold(e => fail(e.toString), identity)

    // A memo used outside a pass must never retain the bounds of an earlier sheet snapshot.
    val grown = source.put(at(0, 2), num(3))
    assertEquals(evaluate(source), CellValue.Error(CellError.NA))
    assertEquals(evaluate(grown), num(3))
    assertEquals(counters.snapshot(Counter.UsedRangeScans), 2L)
    cache.begin(Readiness.build(Map.empty, Set.empty))
    assertEquals(evaluate(source), CellValue.Error(CellError.NA))
    assertEquals(evaluate(source), CellValue.Error(CellError.NA))
    assertEquals(counters.snapshot(Counter.UsedRangeScans), 3L)
    cache.structureChanged(data)
    assertEquals(evaluate(grown), num(3))
    assertEquals(evaluate(grown), num(3))
    assertEquals(counters.snapshot(Counter.UsedRangeScans), 4L)
    cache.structureChanged(data)
    assertEquals(evaluate(source), CellValue.Error(CellError.NA))
    assertEquals(counters.snapshot(Counter.UsedRangeScans), 5L)
    cache.begin(Readiness.build(Map.empty, Set.empty))
    assertEquals(evaluate(grown), num(3))
    assertEquals(counters.snapshot(Counter.UsedRangeScans), 6L)
  }

  // Opt-in acceptance harness: 10k lookups over exactly 200k populated source cells. An early
  // match isolates bounds-scan cost from the separately linear key-search cost; no timing gate.
  if sys.env.get("XL_WHOLE_AXIS_PERF").contains("1") then
    test("GH-747: 10k whole-column lookups over 200k populated cells") {
      val dense = (0 until 10000).foldLeft(Sheet(data)) { (s, r) =>
        (0 until 20).foldLeft(s)((acc, c) => acc.put(at(c, r), num(c + 1)))
      }
      assertEquals(dense.cells.size, 200000)
      val wb = book(Vector.fill(10000)("XLOOKUP(1,Data!A:A,Data!B:B)"), dense)
      val counters = RecalcCounters.create()
      val start = System.nanoTime()
      val result = recalculateCounted(wb, options, counters)
      val elapsed = (System.nanoTime() - start).toDouble / 1e9
      println(
        f"GH-747: 10000 lookups / 200000 cells: $elapsed%.3f s; ${counters.snapshot(Counter.UsedRangeScans)} used-range scans"
      )
      assert(result.isClean, result.errors.toString)
      assertEquals(result.evaluated(out).size, 10000)
      assertEquals(result.evaluated(out).values.toSet, Set(num(2)))
      assertEquals(counters.snapshot(Counter.UsedRangeScans), 1L)
    }
