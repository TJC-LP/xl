package com.tjclp.xl.formula.perf

import java.util.concurrent.TimeUnit

import scala.compiletime.uninitialized

import com.tjclp.xl.formula.eval.{
  IterativeMode,
  RecalcCounters,
  RecalcOptions,
  RecalcResult,
  WorkbookEvaluator
}
import com.tjclp.xl.workbooks.Workbook
import org.openjdk.jmh.annotations.*

/**
 * Whole-workbook recalculation of the synthetic loads book (#767), one recalculation per
 * invocation.
 *
 *   - `rows`: Loads data rows (1000 Summary and 500 x 14 Rate History formulas at every size)
 *   - `crit`: `helper` (three helper formula columns per row, criteria on them) or `raw` (no
 *     helpers, raw-column criteria)
 *   - `counters`: `off` is the public `recalculate` path (disabled counters); `on` measures. The
 *     ratio of the two is the cost of the instrumentation, which must stay under 1%.
 *
 * Today's engine needs ~13 GB of heap for `rows=5000 crit=helper` and runs minutes per invocation
 * at 10k+ rows, so the weekly trend job runs a subset (`.github/workflows/perf-trend.yml`). Lives
 * in the `com.tjclp.xl.formula.perf` package for the internal measured seam.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Array(Mode.SingleShotTime))
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 1)
@Measurement(iterations = 3)
@Fork(value = 1)
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.Null"))
class RecalcBenchmark {

  @Param(Array("5000", "10000", "20000"))
  var rows: Int = uninitialized

  @Param(Array("raw", "helper"))
  var crit: String = uninitialized

  @Param(Array("off", "on"))
  var counters: String = uninitialized

  var workbook: Workbook = uninitialized

  private val options = RecalcOptions(iterative = IterativeMode.Off)

  @Setup(Level.Trial)
  def setup(): Unit = {
    val helper = crit == "helper"
    val shape = LoadsBook.Shape(
      name = s"bench-$crit-$rows",
      rows = rows,
      aggs = 1000,
      rh = 500,
      helpers = helper,
      crit = if helper then LoadsBook.Crit.Helper else LoadsBook.Crit.Raw
    )
    workbook = LoadsBook.build(shape).workbook
  }

  @Benchmark
  def recalculate(): RecalcResult = {
    val c = if counters == "on" then RecalcCounters.create() else RecalcCounters.disabled
    WorkbookEvaluator.recalculateCounted(workbook, options, c)
  }
}
