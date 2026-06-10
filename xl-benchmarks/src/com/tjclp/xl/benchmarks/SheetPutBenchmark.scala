package com.tjclp.xl.benchmarks

import com.tjclp.xl.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.codec.{CellWritable, CellWriter}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.Sheet
import org.openjdk.jmh.annotations.*
import org.openjdk.jmh.infra.Blackhole
import java.time.LocalDate
import java.util.concurrent.TimeUnit
import scala.compiletime.uninitialized
import scala.util.Random

/**
 * Microbenchmarks for Sheet.put overloads (GH-40).
 *
 * The codec-based put paths dispatch on the runtime value type via the master
 * `CellWriter[CellWritable]` pattern match (CellCodec.scala). These benchmarks measure whether that
 * dispatch is a real cost at 1000 cells with mixed value types
 * (String/Int/Double/Boolean/LocalDate):
 *
 *   - chainedSinglePuts: sheet.put(ref, value) folded per cell (dispatch + style inference each)
 *   - batchVarargsPut: sheet.put(updates*) — tuple seq built in setup, single bulk apply
 *   - preConstructedCellPuts: sheet.put(cell) folded per cell (type resolved ahead, no dispatch)
 *   - patchBatchPut: Patch.Batch of Patch.Put + applyPatch (the DSL path scripts use)
 *   - writerDispatchOnly: isolates the pattern match + codec write, no sheet machinery
 *
 * Note: the codec paths (1, 2) also auto-infer styles (LocalDate cells get NumFmt.Date), which the
 * CellValue paths (3, 4) skip by construction — writerDispatchOnly attributes how much of the
 * difference is dispatch vs style registration.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Array(Mode.AverageTime))
@OutputTimeUnit(TimeUnit.MICROSECONDS)
@Warmup(iterations = 3, time = 1, timeUnit = TimeUnit.SECONDS)
@Measurement(iterations = 5, time = 2, timeUnit = TimeUnit.SECONDS)
@Fork(value = 1)
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.Null"))
class SheetPutBenchmark {

  private val writer = CellWriter[CellWritable]

  var baseSheet: Sheet = uninitialized
  var mixedUpdates: Seq[(ARef, CellWritable)] = uninitialized
  var preConstructedCells: Seq[Cell] = uninitialized
  var putBatchPatch: Patch = uninitialized

  @Setup(Level.Trial)
  def setup(): Unit = {
    val random = new Random(42) // Fixed seed for reproducibility
    baseSheet = Sheet("PutBench")

    // 200 rows x 5 columns = 1000 cells with mixed value types (GH-40 scenario)
    mixedUpdates = (1 to 200).flatMap { row =>
      Seq[(ARef, CellWritable)](
        (ARef.from1(1, row), row), // Int
        (ARef.from1(2, row), s"User_${row % 100}"), // String
        (ARef.from1(3, row), random.nextDouble() * 10000), // Double
        (ARef.from1(4, row), row % 2 == 0), // Boolean
        (ARef.from1(5, row), LocalDate.of(2024, 1, 1).plusDays(row % 365)) // LocalDate
      )
    }

    // Same values with the type dispatch resolved ahead of time (in setup, not measured)
    preConstructedCells = mixedUpdates.map { case (ref, value) =>
      Cell(ref, writer.write(value)._1)
    }

    putBatchPatch = Patch.Batch(
      preConstructedCells.map(cell => Patch.Put(cell.ref, cell.value)).toVector
    )
  }

  /** Scenario 1: chained single puts — CellWriter dispatch + sheet copy per cell. */
  @Benchmark
  def chainedSinglePuts(): Sheet =
    mixedUpdates.foldLeft(baseSheet) { case (sheet, (ref, value)) =>
      sheet.put(ref, value)
    }

  /** Scenario 2: batch varargs put — one traversal, bulk cell apply (recommended batch API). */
  @Benchmark
  def batchVarargsPut(): Sheet =
    baseSheet.put(mixedUpdates*)

  /** Scenario 3: pre-constructed Cell puts — no dispatch, no style inference in the loop. */
  @Benchmark
  def preConstructedCellPuts(): Sheet =
    preConstructedCells.foldLeft(baseSheet)((sheet, cell) => sheet.put(cell))

  /** Scenario 4: Patch.Batch of Puts via applyPatch — the Patch monoid DSL path. */
  @Benchmark
  def patchBatchPut(): Sheet =
    Patch.applyPatch(baseSheet, putBatchPatch)

  /** Isolation: the CellWriter[CellWritable] runtime pattern match + codec write alone. */
  @Benchmark
  def writerDispatchOnly(bh: Blackhole): Unit =
    mixedUpdates.foreach { case (_, value) => bh.consume(writer.write(value)) }
}
