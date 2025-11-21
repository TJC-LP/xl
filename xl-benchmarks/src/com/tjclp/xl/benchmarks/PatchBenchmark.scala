package com.tjclp.xl.benchmarks

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, Column, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.patch.Patch
import org.openjdk.jmh.annotations.*
import java.util.concurrent.TimeUnit
import scala.compiletime.uninitialized

/**
 * Microbenchmarks for Patch operations.
 *
 * Tests:
 *   - Single cell update
 *   - Row update (1000 cells)
 *   - Column update (10000 cells)
 *   - Patch composition (Monoid operations)
 */
@State(Scope.Benchmark)
@BenchmarkMode(Array(Mode.AverageTime))
@OutputTimeUnit(TimeUnit.MICROSECONDS)
@Warmup(iterations = 3, time = 1, timeUnit = TimeUnit.SECONDS)
@Measurement(iterations = 5, time = 2, timeUnit = TimeUnit.SECONDS)
@Fork(value = 1)
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.Null"))
class PatchBenchmark {

  var baseSheet: Sheet = uninitialized
  var singleCellPatch: Patch = uninitialized
  var rowPatches: Seq[Patch] = uninitialized
  var columnPatches: Seq[Patch] = uninitialized

  @Setup(Level.Trial)
  def setup(): Unit = {
    // Create base sheet with some data
    baseSheet = BenchmarkUtils.generateSheet("Test", 1000, styled = false)

    // Single cell patch
    singleCellPatch = Patch.Put(ref"A1", CellValue.Number(BigDecimal(42)))

    // Row of patches (1000 cells in row 1)
    rowPatches = (1 to 1000).map { col =>
      Patch.Put(ARef.from1(col, 1), CellValue.Number(BigDecimal(col)))
    }

    // Column of patches (10000 cells in column A)
    columnPatches = (1 to 10000).map { row =>
      Patch.Put(ARef.from1(1, row), CellValue.Number(BigDecimal(row)))
    }
  }

  @Benchmark
  def singleCellUpdate(): Sheet = {
    // Measure single cell update overhead
    Patch.applyPatch(baseSheet, singleCellPatch).getOrElse(baseSheet)
  }

  @Benchmark
  def rowUpdate(): Sheet = {
    // Measure row update (1000 cells)
    val batchPatch = rowPatches.reduce(Patch.combine)
    Patch.applyPatch(baseSheet, batchPatch).getOrElse(baseSheet)
  }

  @Benchmark
  def columnUpdate(): Sheet = {
    // Measure column update (10000 cells)
    val batchPatch = columnPatches.reduce(Patch.combine)
    Patch.applyPatch(baseSheet, batchPatch).getOrElse(baseSheet)
  }

  @Benchmark
  def patchComposition(): Patch = {
    // Measure overhead of composing patches (Monoid operations)
    rowPatches.reduce(Patch.combine)
  }
}
