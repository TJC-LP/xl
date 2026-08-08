package com.tjclp.xl.formula

import java.util.concurrent.{Callable, Executors, TimeUnit}

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{EvalError, Evaluator, RecalcResult, SheetEvaluator}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

import scala.jdk.CollectionConverters.*

/** Calculation-generation cache laws for the common one-raw-range aggregate shape. */
class AggregateMemoSpec extends FunSuite:

  private val dataName = SheetName.unsafe("Data")
  private val inputsName = SheetName.unsafe("Inputs")
  private val outputName = SheetName.unsafe("Output")

  private def num(value: Int): CellValue = CellValue.Number(BigDecimal(value))
  private def formula(expr: String, cached: Option[CellValue] = None): CellValue =
    CellValue.Formula(expr, cached)

  private def evaluate(
    sheet: Sheet,
    ref: ARef,
    workbook: Workbook,
    evaluator: com.tjclp.xl.formula.eval.Evaluator
  ): Either[com.tjclp.xl.error.XLError, CellValue] =
    SheetEvaluator.evaluateCellWithEvaluator(
      sheet,
      ref,
      evaluator,
      Clock.system,
      Some(workbook)
    )

  private def assertSameResult(sequential: RecalcResult, parallel: RecalcResult): Unit =
    assertEquals(parallel.errors, sequential.errors)
    assertEquals(parallel.evaluated, sequential.evaluated)
    assertEquals(parallel.workbook, sequential.workbook)
    assertEquals(parallel.converged, sequential.converged)
    assertEquals(parallel.cycles, sequential.cycles)

  test("aggregate memo ignores anchors but distinguishes aggregator identity"):
    val data = Sheet(dataName)
      .put(ARef.from0(0, 0), num(1))
      .put(ARef.from0(0, 1), num(2))
      .put(ARef.from0(0, 2), num(3))
    val sumAnchored = ARef.from0(0, 0)
    val sumRelative = ARef.from0(1, 0)
    val average = ARef.from0(2, 0)
    val output = Sheet(outputName)
      .put(sumAnchored, formula("=SUM(Data!$A$1:$A$3)"))
      .put(sumRelative, formula("=SUM(Data!A1:A3)"))
      .put(average, formula("=AVERAGE(Data!A1:A3)"))
    val workbook = Workbook(data, output)
    val memo = new Evaluator.AggregateMemo
    val evaluator = Evaluator.recalculationInstance(Rng.system, memo)

    assertEquals(evaluate(output, sumAnchored, workbook, evaluator), Right(num(6)))
    assertEquals(evaluate(output, sumRelative, workbook, evaluator), Right(num(6)))
    assertEquals(evaluate(output, average, workbook, evaluator), Right(num(2)))
    assertEquals(
      memo.stats,
      Evaluator.AggregateMemoStats(hits = 1, fills = 2, bypasses = 0, entries = 2)
    )

  test("function-call and typed-node coercion modes cannot contaminate one another"):
    val sheet = Sheet(dataName).put(ref"A1", CellValue.Bool(true))
    val call = FormulaParser.parse("=SUM(A1:A1)") match
      case Right(expr) => expr
      case Left(error) => fail(s"failed to parse aggregate call: $error")
    val node = TExpr.Aggregate(
      "SUM",
      TExpr.RangeLocation.Local(CellRange(ref"A1", ref"A1"))
    )

    def evaluate(expr: TExpr[?], evaluator: Evaluator): Either[EvalError, String] =
      evaluator.eval(expr, sheet, Clock.system).map(_.toString)

    def run(
      first: TExpr[?],
      second: TExpr[?]
    ): (Either[EvalError, String], Either[EvalError, String]) =
      val memo = new Evaluator.AggregateMemo
      val evaluator = Evaluator.recalculationInstance(Rng.system, memo)
      val results = (evaluate(first, evaluator), evaluate(second, evaluator))
      assertEquals(memo.stats.fills, 2L)
      assertEquals(memo.stats.entries, 2)
      results

    assertEquals(run(call, node), (Right("0"), Right("1")))
    assertEquals(run(node, call), (Right("1"), Right("0")))

  test("same-generation snapshots reuse when writes leave the aggregated range unchanged"):
    val data = Sheet(dataName)
      .put(ARef.from0(0, 0), num(1))
      .put(ARef.from0(0, 1), num(2))
    val laterSnapshot = data.put(ARef.from0(1, 0), num(99))
    assert(!(laterSnapshot eq data))

    val outputRef = ARef.from0(0, 0)
    val output = Sheet(outputName).put(outputRef, formula("=SUM(Data!A1:A2)"))
    val memo = new Evaluator.AggregateMemo
    val evaluator = Evaluator.recalculationInstance(Rng.system, memo)

    assertEquals(evaluate(output, outputRef, Workbook(data, output), evaluator), Right(num(3)))
    assertEquals(
      evaluate(output, outputRef, Workbook(laterSnapshot, output), evaluator),
      Right(num(3))
    )
    assertEquals(
      memo.stats,
      Evaluator.AggregateMemoStats(hits = 1, fills = 1, bypasses = 0, entries = 1)
    )

  test("same-sheet sequential cache writes retain one shared aggregate entry"):
    val rows = 500
    val readers = 32
    val readerRefs = (1 to readers).map(col => ARef.from0(col, 0)).toVector
    val data = (1 to rows).foldLeft(Sheet(dataName)) { (sheet, row) =>
      sheet.put(ARef.from0(0, row - 1), num(1))
    }
    val initial = readerRefs.foldLeft(data) { (sheet, ref) =>
      sheet.put(ref, formula(s"=SUM(A1:A$rows)"))
    }
    val memo = new Evaluator.AggregateMemo
    val evaluator = Evaluator.recalculationInstance(Rng.system, memo)

    val (_, values) = readerRefs.foldLeft((initial, Vector.empty[CellValue])) {
      case ((sheet, acc), ref) =>
        evaluate(sheet, ref, Workbook(sheet), evaluator) match
          case Right(value) => (sheet.put(ref, value), acc :+ value)
          case Left(error) => fail(s"aggregate evaluation failed: $error")
    }

    assertEquals(values, Vector.fill(readers)(num(rows)))
    assertEquals(
      memo.stats,
      Evaluator.AggregateMemoStats(
        hits = readers - 1,
        fills = 1,
        bypasses = 0,
        entries = 1
      )
    )

  test("full-column cache keys follow changing effective used bounds"):
    val empty = Sheet(SheetName.unsafe("Empty"))
    val data = Sheet(dataName).put(
      ref"Z1000",
      formula("=IFERROR(COUNTBLANK(A:A)/0,Empty!A1)")
    )
    val output = Sheet(outputName).put(
      ref"A1",
      formula("=COUNTBLANK(Data!A:A)+0*Data!Z1000")
    )

    val result = Workbook(data, empty, output).recalculate()

    assert(result.isClean, result.errors.map(_.render).mkString("; "))
    assertEquals(result.evaluated(outputName)(ref"A1"), num(0))

  test("uncached formula ranges bypass; cached formula ranges may reuse"):
    val inputRef = ARef.from0(0, 0)
    val outputRef = ARef.from0(0, 0)
    val inputs1 = Sheet(inputsName).put(inputRef, num(1))
    val inputs2 = Sheet(inputsName).put(inputRef, num(2))
    val output = Sheet(outputName).put(outputRef, formula("=SUM(Data!A1:A1)"))

    // The same Data Sheet identity computes against two different workbook snapshots. Caching its
    // uncached formula would return the first workbook's value for the second workbook.
    val uncachedData = Sheet(dataName).put(inputRef, formula("=Inputs!A1"))
    val uncachedMemo = new Evaluator.AggregateMemo
    val uncachedEvaluator = Evaluator.recalculationInstance(Rng.system, uncachedMemo)
    assertEquals(
      evaluate(output, outputRef, Workbook(inputs1, uncachedData, output), uncachedEvaluator),
      Right(num(1))
    )
    assertEquals(
      evaluate(output, outputRef, Workbook(inputs2, uncachedData, output), uncachedEvaluator),
      Right(num(2))
    )
    assertEquals(
      uncachedMemo.stats,
      Evaluator.AggregateMemoStats(hits = 0, fills = 0, bypasses = 2, entries = 1)
    )

    // A cached formula's effective value is immutable within its Sheet snapshot and is safe.
    val cachedData = Sheet(dataName).put(inputRef, formula("=Inputs!A1", Some(num(7))))
    val cachedMemo = new Evaluator.AggregateMemo
    val cachedEvaluator = Evaluator.recalculationInstance(Rng.system, cachedMemo)
    assertEquals(
      evaluate(output, outputRef, Workbook(inputs1, cachedData, output), cachedEvaluator),
      Right(num(7))
    )
    assertEquals(
      evaluate(output, outputRef, Workbook(inputs2, cachedData, output), cachedEvaluator),
      Right(num(7))
    )
    assertEquals(
      cachedMemo.stats,
      Evaluator.AggregateMemoStats(hits = 1, fills = 1, bypasses = 0, entries = 1)
    )

  test("shared-range aggregate wave is sequential/parallel equivalent"):
    val rows = 500
    val readers = 32
    val data = (1 to rows).foldLeft(Sheet(dataName)) { (sheet, row) =>
      sheet.put(ARef.from0(0, row - 1), num(row))
    }
    val output = (0 until readers).foldLeft(Sheet(outputName)) { (sheet, col) =>
      val range = if col % 2 == 0 then "$A$1:$A$500" else "A1:A500"
      sheet.put(ARef.from0(col, 0), formula(s"=SUM(Data!$range)"))
    }
    val workbook = Workbook(data, output)
    val sequential = workbook.recalculate()
    val parallel = workbook.recalculateParallel(8)

    assertSameResult(sequential, parallel)
    val expected = num(rows * (rows + 1) / 2)
    assertEquals(parallel.evaluated(outputName).values.toSet, Set(expected))

  test("parallel readers single-flight one eligible range fold"):
    val rows = 5000
    val readers = 32
    val data = (1 to rows).foldLeft(Sheet(dataName)) { (sheet, row) =>
      sheet.put(ARef.from0(0, row - 1), num(1))
    }
    val outputRef = ARef.from0(0, 0)
    val output = Sheet(outputName).put(outputRef, formula(s"=SUM(Data!A1:A$rows)"))
    val workbook = Workbook(data, output)
    val memo = new Evaluator.AggregateMemo
    val evaluator = Evaluator.recalculationInstance(Rng.system, memo)
    val pool = Executors.newFixedThreadPool(8)
    val tasks = new java.util.ArrayList[Callable[Either[XLError, CellValue]]]()
    (0 until readers).foreach { _ =>
      tasks.add(() => evaluate(output, outputRef, workbook, evaluator))
    }

    try
      val results = pool.invokeAll(tasks).asScala.map(_.get()).toVector
      assertEquals(results, Vector.fill(readers)(Right(num(rows))))
      assertEquals(
        memo.stats,
        Evaluator.AggregateMemoStats(
          hits = readers - 1,
          fills = 1,
          bypasses = 0,
          entries = 1
        )
      )
    finally
      pool.shutdownNow()
      pool.awaitTermination(5, TimeUnit.SECONDS)
