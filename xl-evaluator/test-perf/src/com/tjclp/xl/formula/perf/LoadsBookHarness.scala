package com.tjclp.xl.formula.perf

import java.lang.management.{ManagementFactory, MemoryType}
import java.nio.file.{Files, Path}

import scala.jdk.CollectionConverters.*

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.{IterativeMode, RecalcCounters, RecalcOptions, RecalcResult}
import com.tjclp.xl.formula.eval.WorkbookEvaluator
import com.tjclp.xl.io.Excel
import com.tjclp.xl.workbooks.Workbook

/**
 * Runs one measured recalculation of a benchmark book (#767) and reads its oracle samples.
 *
 * Shared by [[LoadsBookMain]] and the perf specs so the numbers a spec gates on are the numbers the
 * harness prints.
 */
object LoadsBookHarness:

  /** What to run: the three entry points `Excel.writeChecked` and `recalculate` reach. */
  enum Mode derives CanEqual:
    /** `wb.recalculate(RecalcOptions(iterative = Off))`, sequential. */
    case Recalc

    /** `recalculate` with `parallelism = threads`. */
    case Parallel

    /** `wb.recalculateUncached(...)`: what `Excel.writeChecked` computes. */
    case Uncached

    /** `recalculateUncached` then `Excel.write`: `Excel.writeChecked`, with the write timed. */
    case Checked

  object Mode:
    def parse(s: String): Option[Mode] = Mode.values.find(_.toString.equalsIgnoreCase(s))

  /** One run: the result, the counter snapshot, its wall time and the JVM heap high-water mark. */
  final case class Run(
    result: RecalcResult,
    counters: RecalcCounters.Snapshot,
    wallSeconds: Double,
    heapPeakBytes: Long
  )

  /**
   * Run `mode` over `workbook`. `measured = false` passes the disabled counters, which is what the
   * public entry points do (the JMH benchmark compares the two). `Checked` writes to `out`, or to a
   * temporary file it deletes.
   */
  def run(
    workbook: Workbook,
    mode: Mode,
    measured: Boolean = true,
    threads: Int = 2,
    out: Option[Path] = None
  ): Run =
    val counters = if measured then RecalcCounters.create() else RecalcCounters.disabled
    val options = RecalcOptions(
      iterative = IterativeMode.Off,
      parallelism = if mode == Mode.Parallel then threads else 1
    )
    resetHeapPeaks()
    val start = System.nanoTime()
    val result = mode match
      case Mode.Recalc | Mode.Parallel =>
        WorkbookEvaluator.recalculateCounted(workbook, options, counters)
      case Mode.Uncached =>
        WorkbookEvaluator.recalculateUncachedCounted(workbook, options, counters)
      case Mode.Checked =>
        val recalculated = WorkbookEvaluator.recalculateUncachedCounted(workbook, options, counters)
        val path = out.getOrElse(Files.createTempFile("xl-perf-", ".xlsx"))
        val writeStart = System.nanoTime()
        Excel.write(recalculated.workbook, path.toString)
        counters.addNanos(RecalcCounters.Phase.Write, System.nanoTime() - writeStart)
        if out.isEmpty then Files.deleteIfExists(path)
        recalculated
    val wall = (System.nanoTime() - start).toDouble / 1e9
    Run(result, counters.snapshot, wall, heapPeak())

  /** The value of `sheet!ref` in `wb` as the oracle files spell it. */
  def sampleValue(wb: Workbook, qualified: String): String =
    val bang = qualified.lastIndexOf('!')
    val sheetName = qualified.take(bang).stripPrefix("'").stripSuffix("'")
    val cell = for
      sheet <- wb.sheets.find(_.name == SheetName.unsafe(sheetName))
      ref <- ARef.parse(qualified.drop(bang + 1)).toOption
      cell <- sheet.cells.get(ref)
    yield cell.value
    cell.fold("<none>")(show)

  /** Canonical text of a computed value: numbers without trailing zeros, text quoted. */
  def show(value: CellValue): String = value match
    case CellValue.Number(n) => n.bigDecimal.stripTrailingZeros.toPlainString
    case CellValue.Text(s) => "\"" + s + "\""
    case CellValue.Bool(b) => b.toString.toUpperCase(java.util.Locale.ROOT)
    case CellValue.Formula(_, Some(cached), _) => show(cached)
    case CellValue.Formula(_, None, _) => "<uncached>"
    case CellValue.Empty => "<empty>"
    case CellValue.Error(e) => e.toExcel
    case other => other.toString

  /** `ref<TAB>value` lines for every sample of `built`, read from `wb`. */
  def oracleLines(built: LoadsBook.Built, wb: Workbook): Vector[String] =
    built.sampleRefs.map(ref => s"$ref\t${sampleValue(wb, ref)}")

  /**
   * `Sheet!A1<TAB>value` for every formula cell of `wb`, sheet by sheet in workbook order, then row
   * by row: the full-book oracle (`loads-values-tiny.tsv`).
   */
  def valueLines(wb: Workbook): Vector[String] =
    wb.sheets.flatMap { sheet =>
      sheet.cells.toVector
        .sortBy((ref, _) => (ref.row.index0, ref.col.index0))
        .collect {
          case (ref, cell) if isFormula(cell.value) =>
            s"${sheet.name.value}!${ref.toA1}\t${show(cell.value)}"
        }
    }

  private def isFormula(value: CellValue): Boolean = value match
    case _: CellValue.Formula => true
    case _ => false

  /** Parse an oracle file's `ref<TAB>value` lines. */
  def parseOracle(text: String): Vector[(String, String)] =
    text.linesIterator
      .filter(_.nonEmpty)
      .map { line =>
        val tab = line.indexOf('\t')
        (line.take(tab), line.drop(tab + 1))
      }
      .toVector

  /** A fixed-width table of counters and phase seconds. */
  def report(label: String, run: Run): String =
    val snap = run.counters
    val counts =
      RecalcCounters.Counter.values.toVector.map(c => f"  ${c.toString}%-22s ${snap(c)}%,d")
    val phases =
      RecalcCounters.Phase.values.toVector.map(p =>
        f"  ${p.toString}%-22s ${snap.seconds(p)}%10.3f s"
      )
    (Vector(
      s"== $label",
      f"  wall                   ${run.wallSeconds}%10.3f s",
      f"  heap pool peaks (sum)  ${run.heapPeakBytes.toDouble / 1e9}%10.3f GB",
      s"  errors                 ${run.result.errors.size}",
      "  -- counters"
    ) ++ counts ++ Vector(
      "  -- phases (inclusive; Parse, Criteria and SheetPut nest in Eval)"
    ) ++ phases)
      .mkString("\n")

  private def heapPools =
    ManagementFactory.getMemoryPoolMXBeans.asScala.filter(_.getType == MemoryType.HEAP)

  private def resetHeapPeaks(): Unit = heapPools.foreach(_.resetPeakUsage())

  private def heapPeak(): Long = heapPools.map(_.getPeakUsage.getUsed).sum
