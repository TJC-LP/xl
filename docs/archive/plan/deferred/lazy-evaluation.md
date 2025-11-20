# Lazy Evaluation Architecture (Spark-Style)

> **Status**: ⏸ Deferred Indefinitely – Full Catalyst-style optimizer deemed overkill for Excel use cases.
> Builder pattern (Phase 1) may be reconsidered in future, but full optimizer (Phase 2-6) is not planned.
> See streaming-improvements.md for the prioritized alternative approach.
> **Archived**: 2025-11-20

## Status: SCOPED DOWN - Builder Pattern Only (Not Full Optimizer)

**Update 2025-11-11**: After technical review, **full Catalyst-style optimizer is overkill** for Excel use cases.

**Revised Scope**:
- ✅ **Builder pattern** for batching operations (1-2 weeks) - **RECOMMENDED**
- ❌ **Full optimizer** (predicate pushdown, cost-based, etc.) - **DEFERRED INDEFINITELY**
- ✅ **Streaming improvements** (P6.6, P6.7, P7.5) - **PRIORITIZED** (see streaming-improvements.md)

**Rationale**:
- Excel bottlenecks are **XML I/O, ZIP compression, and style indexing** (not query planning)
- Predicate pushdown has **minimal ROI** for sparse cell updates (not SQL-style filtering)
- Cost-based optimizer requires **heavy machinery** (statistics, plan enumeration) without commensurate gains
- Builder pattern captures **80-90% of benefits** with 10% of complexity

**Recommendation**: Implement **Phase 1 only** (builder pattern), skip Phase 2-6 (full optimizer)

See "Recommended Scope" section below for practical implementation.

---

## Original Vision (Full Optimizer - Now Considered Overkill)

Transform XL into a Spark-like library where operations are lazy by default and only execute when an action is triggered. This enables:

- **Query optimization**: Analyze and optimize the entire operation chain before execution
- **Constant memory**: Process millions of rows with O(1) memory via streaming
- **Conditional execution**: Build complex operation graphs that may not execute all branches
- **Beautiful APIs**: Fluent, composable transformations with explicit materialization points

## Motivation

### Current State (Eager Evaluation)

```scala
val sheet = Sheet(SheetName.unsafe("Data"))
  .put(cell"A1", "Title")      // Executes NOW, creates new Sheet
  .put(cell"A2", 100)           // Executes NOW, creates new Sheet
  .merge(range"A1:B1")          // Executes NOW, creates new Sheet

// 3 intermediate Sheets created immediately
// All Map operations executed even if some results unused
```

**Problems:**
- Intermediate allocations for every operation
- No opportunity for optimization (can't eliminate redundant operations)
- Memory grows linearly with dataset size
- Can't defer computation for conditional workflows

### Desired State (Lazy Evaluation)

```scala
val sheet = LazySheet("Data")
  .put(cell"A1", "Title")      // Builds logical plan (lazy)
  .put(cell"A2", 100)           // Builds logical plan (lazy)
  .merge(range"A1:B1")          // Builds logical plan (lazy)

// No computation yet! Just a description of operations

// Execute optimized plan on action:
sheet.write(path).unsafeRunSync()  // Optimizer runs, then streams to disk
```

**Benefits:**
- Zero intermediate allocations during planning
- Optimizer eliminates waste (consecutive puts → batched putAll)
- Constant memory with streaming execution
- Conditional workflows don't execute unused branches

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────┐
│                      LazySheet API                          │
│  (Transformations: put, merge, filter, limit)               │
└─────────────────────┬───────────────────────────────────────┘
                      │ Builds LogicalPlan
                      ▼
┌─────────────────────────────────────────────────────────────┐
│                    LogicalPlan ADT                          │
│  (Immutable tree of operations: Put, Merge, Filter, etc.)  │
└─────────────────────┬───────────────────────────────────────┘
                      │ On action (write/collect)
                      ▼
┌─────────────────────────────────────────────────────────────┐
│                    Optimizer Pipeline                        │
│  Pass 1: Batching (consecutive ops → bulk ops)              │
│  Pass 2: Dead Code Elimination (overwritten ops)            │
│  Pass 3: Predicate Pushdown (filter before load)            │
│  Pass 4: Cost-Based (statistics-driven reordering)          │
└─────────────────────┬───────────────────────────────────────┘
                      │ Produces optimized LogicalPlan
                      ▼
┌─────────────────────────────────────────────────────────────┐
│                    PhysicalPlan                             │
│  (Executable representation with streaming ops)             │
└─────────────────────┬───────────────────────────────────────┘
                      │ Execute
                      ▼
┌─────────────────────────────────────────────────────────────┐
│                  StreamingExecutor (fs2)                    │
│  (Constant memory, backpressure, chunked processing)        │
└─────────────────────┬───────────────────────────────────────┘
                      │ Produces Stream[F, Cell]
                      ▼
┌─────────────────────────────────────────────────────────────┐
│                      Actions                                 │
│  write(path), collect(), count, show(n)                     │
└─────────────────────────────────────────────────────────────┘
```

## Detailed Design

### 1. Logical Plan ADT

**Location:** `xl-core/src/com/tjclp/xl/plan/LogicalPlan.scala`

```scala
package com.tjclp.xl.plan

import com.tjclp.xl.api.*
import java.nio.file.Path
import cats.effect.IO

/**
 * Logical plan representing a sequence of operations on a sheet.
 *
 * This is an immutable tree structure that describes what to do, not how to do it.
 * Execution is deferred until an action (write, collect, count, show) is called.
 */
sealed trait LogicalPlan:
  // Transformations (lazy - build plan without executing)
  def put(ref: ARef, value: CellValue): LogicalPlan = Put(this, ref, value)
  def putAll(cells: Vector[Cell]): LogicalPlan = PutAll(this, cells)
  def merge(range: CellRange): LogicalPlan = Merge(this, range)
  def filter(predicate: Cell => Boolean): LogicalPlan = Filter(this, predicate)
  def limit(n: Int): LogicalPlan = Limit(this, n)
  def setStyle(ref: ARef, style: CellStyle): LogicalPlan = SetStyle(this, ref, style)

  // Actions (eager - trigger optimized execution)
  def collect(): MaterializedSheet = Executor.execute(this)
  def write(path: Path): IO[Unit] = Executor.writeStream(this, path)
  def count: Long = Executor.count(this)
  def show(n: Int = 20): String = Executor.show(this, n)
  def explain(): String = Optimizer.explainPlan(this)

  // Transform this plan by applying function to all nodes
  def transform(f: LogicalPlan => LogicalPlan): LogicalPlan

/**
 * Base case: Wrap an existing materialized sheet.
 */
case class BaseSheet(sheet: MaterializedSheet) extends LogicalPlan:
  def transform(f: LogicalPlan => LogicalPlan): LogicalPlan = f(this)

/**
 * Put a single cell value.
 */
case class Put(parent: LogicalPlan, ref: ARef, value: CellValue) extends LogicalPlan:
  def transform(f: LogicalPlan => LogicalPlan): LogicalPlan =
    f(Put(parent.transform(f), ref, value))

/**
 * Put multiple cells at once (bulk operation).
 * Optimizer will batch consecutive Put operations into PutAll.
 */
case class PutAll(parent: LogicalPlan, cells: Vector[Cell]) extends LogicalPlan:
  def transform(f: LogicalPlan => LogicalPlan): LogicalPlan =
    f(PutAll(parent.transform(f), cells))

/**
 * Merge a range of cells.
 */
case class Merge(parent: LogicalPlan, range: CellRange) extends LogicalPlan:
  def transform(f: LogicalPlan => LogicalPlan): LogicalPlan =
    f(Merge(parent.transform(f), range))

/**
 * Filter cells by predicate.
 * Optimizer will push filters close to data source.
 */
case class Filter(parent: LogicalPlan, predicate: Cell => Boolean) extends LogicalPlan:
  def transform(f: LogicalPlan => LogicalPlan): LogicalPlan =
    f(Filter(parent.transform(f), predicate))

/**
 * Limit number of cells.
 */
case class Limit(parent: LogicalPlan, n: Int) extends LogicalPlan:
  def transform(f: LogicalPlan => LogicalPlan): LogicalPlan =
    f(Limit(parent.transform(f), n))

/**
 * Set style for a cell.
 */
case class SetStyle(parent: LogicalPlan, ref: ARef, style: CellStyle) extends LogicalPlan:
  def transform(f: LogicalPlan => LogicalPlan): LogicalPlan =
    f(SetStyle(parent.transform(f), ref, style))
```

### 2. LazySheet API

**Location:** `xl-core/src/com/tjclp/xl/LazySheet.scala`

```scala
package com.tjclp.xl

import com.tjclp.xl.plan.*
import java.nio.file.Path
import cats.effect.IO

/**
 * Lazy sheet that builds a logical plan without executing operations.
 *
 * This is the primary API for working with sheets in lazy mode.
 * Operations are composable and form a computation graph that only executes
 * when an action (write, collect, count, show) is triggered.
 *
 * Example:
 * {{{
 * val sheet = LazySheet("Sales")
 *   .put(cell"A1", "Revenue")
 *   .put(cell"B1", 1000)
 *   .merge(range"A1:B1")
 *
 * // No computation yet!
 *
 * sheet.write(path).unsafeRunSync()  // Execute optimized plan, stream to disk
 * }}}
 */
case class LazySheet(plan: LogicalPlan):
  // Transformations (lazy)
  def put(ref: ARef, value: CellValue): LazySheet = LazySheet(plan.put(ref, value))
  def putAll(cells: Vector[Cell]): LazySheet = LazySheet(plan.putAll(cells))
  def put(ref: ARef, value: CellValue, style: CellStyle): LazySheet =
    LazySheet(plan.put(ref, value).setStyle(ref, style))
  def merge(range: CellRange): LazySheet = LazySheet(plan.merge(range))
  def filter(predicate: Cell => Boolean): LazySheet = LazySheet(plan.filter(predicate))
  def limit(n: Int): LazySheet = LazySheet(plan.limit(n))

  // Actions (eager - trigger execution)
  def collect(): MaterializedSheet = plan.collect()
  def write(path: Path): IO[Unit] = plan.write(path)
  def count: Long = plan.count
  def show(n: Int = 20): String = plan.show(n)
  def explain(): String = plan.explain()

object LazySheet:
  /**
   * Create a new lazy sheet with given name.
   * No materialized sheet is created until an action is triggered.
   */
  def apply(name: SheetName): LazySheet =
    LazySheet(BaseSheet(MaterializedSheet(name)))

  /**
   * Wrap an existing materialized sheet in lazy API.
   */
  def fromMaterialized(sheet: MaterializedSheet): LazySheet =
    LazySheet(BaseSheet(sheet))
```

### 3. Optimizer Pipeline

**Location:** `xl-core/src/com/tjclp/xl/optimizer/Optimizer.scala`

```scala
package com.tjclp.xl.optimizer

import com.tjclp.xl.plan.*

/**
 * Multi-pass query optimizer inspired by Spark's Catalyst optimizer.
 *
 * Applies 4 optimization passes in sequence:
 * 1. Batching: Combine consecutive operations (Put → PutAll)
 * 2. Dead Code Elimination: Remove overwritten operations
 * 3. Predicate Pushdown: Move filters close to data source
 * 4. Cost-Based: Reorder operations based on estimated cost
 */
object Optimizer:
  /**
   * Optimize a logical plan before execution.
   */
  def optimize(plan: LogicalPlan): LogicalPlan =
    val afterBatching = BatchingOptimizer(plan)
    val afterDeadCode = DeadCodeEliminator(afterBatching)
    val afterPushdown = PredicatePushdown(afterDeadCode)
    val afterCostBased = CostBasedOptimizer(afterPushdown)
    afterCostBased

  /**
   * Explain optimization steps for debugging.
   * Returns multi-line string showing:
   * - Original logical plan
   * - Each optimization pass result
   * - Final optimized plan
   * - Estimated cost savings
   */
  def explainPlan(plan: LogicalPlan): String =
    val original = formatPlan(plan, "Original Plan")
    val afterBatching = BatchingOptimizer(plan)
    val afterDeadCode = DeadCodeEliminator(afterBatching)
    val afterPushdown = PredicatePushdown(afterDeadCode)
    val afterCostBased = CostBasedOptimizer(afterPushdown)

    s"""
    |$original
    |
    |${formatPlan(afterBatching, "After Batching")}
    |
    |${formatPlan(afterDeadCode, "After Dead Code Elimination")}
    |
    |${formatPlan(afterPushdown, "After Predicate Pushdown")}
    |
    |${formatPlan(afterCostBased, "Optimized Plan")}
    |
    |Estimated Savings: ${estimateSavings(plan, afterCostBased)}
    """.stripMargin

  private def formatPlan(plan: LogicalPlan, title: String): String =
    s"== $title ==\n${planToString(plan, indent = 0)}"

  private def planToString(plan: LogicalPlan, indent: Int): String = ???

  private def estimateSavings(original: LogicalPlan, optimized: LogicalPlan): String = ???
```

#### 3.1 Batching Optimizer

**Location:** `xl-core/src/com/tjclp/xl/optimizer/BatchingOptimizer.scala`

```scala
package com.tjclp.xl.optimizer

import com.tjclp.xl.plan.*

/**
 * Combines consecutive operations into bulk operations.
 *
 * Examples:
 * - Put(Put(Put(...))) → PutAll([cell1, cell2, cell3])
 * - SetStyle(SetStyle(...)) → SetStyleBatch([...])
 *
 * Benefits:
 * - Reduces intermediate allocations (3 Sheets → 1 Sheet)
 * - Batched Map operations more efficient than individual updates
 * - Can process bulk operations in parallel
 */
object BatchingOptimizer:
  def apply(plan: LogicalPlan): LogicalPlan = plan.transform {
    case node @ Put(_, _, _) =>
      collectConsecutivePuts(node) match
        case Some((parent, cells)) if cells.size > 1 =>
          PutAll(parent, cells)
        case _ => node

    case node => node
  }

  /**
   * Collect all consecutive Put operations into a batch.
   * Returns None if there are no consecutive Puts to batch.
   */
  private def collectConsecutivePuts(plan: LogicalPlan): Option[(LogicalPlan, Vector[Cell])] =
    def collect(p: LogicalPlan, acc: Vector[Cell]): (LogicalPlan, Vector[Cell]) = p match
      case Put(parent, ref, value) =>
        collect(parent, Cell(ref, value) +: acc)
      case other =>
        (other, acc)

    collect(plan, Vector.empty) match
      case (parent, cells) if cells.size > 1 => Some((parent, cells))
      case _ => None
```

#### 3.2 Dead Code Eliminator

**Location:** `xl-core/src/com/tjclp/xl/optimizer/DeadCodeEliminator.scala`

```scala
package com.tjclp.xl.optimizer

import com.tjclp.xl.plan.*

/**
 * Removes operations that are overwritten by later operations.
 *
 * Examples:
 * - Put(A1, "x") then Put(A1, "y") → Only Put(A1, "y")
 * - Merge(A1:B1) then Merge(A1:B1) → Only one Merge
 * - SetStyle(A1, s1) then SetStyle(A1, s2) → Only SetStyle(A1, s2)
 *
 * Benefits:
 * - Eliminates wasted computation
 * - Reduces intermediate allocations
 * - Typical savings: 10-30% operation reduction
 */
object DeadCodeEliminator:
  def apply(plan: LogicalPlan): LogicalPlan = plan.transform {
    case Put(Put(parent, ref1, _), ref2, v2) if ref1 == ref2 =>
      // Second put overwrites first → eliminate first
      Put(parent, ref2, v2)

    case SetStyle(SetStyle(parent, ref1, _), ref2, s2) if ref1 == ref2 =>
      // Second style overwrites first → eliminate first
      SetStyle(parent, ref2, s2)

    case Merge(Merge(parent, range1), range2) if range1 == range2 =>
      // Duplicate merge → eliminate second
      Merge(parent, range1)

    case node => node
  }
```

#### 3.3 Predicate Pushdown

**Location:** `xl-core/src/com/tjclp/xl/optimizer/PredicatePushdown.scala`

```scala
package com.tjclp.xl.optimizer

import com.tjclp.xl.plan.*

/**
 * Moves filter operations close to data source.
 *
 * Benefits:
 * - Reduces data volume early in pipeline
 * - Avoids wasted computation on filtered-out cells
 * - Enables selective loading (future: read only matching cells from file)
 *
 * Example:
 * {{{
 * // Before:
 * Filter(Put(Put(BaseSheet, A1, "x"), A2, "y"), p)
 *
 * // After:
 * Put(Put(Filter(BaseSheet, p), A1, "x"), A2, "y")
 * // Filter applies before puts (if predicate independent of A1/A2)
 * }}}
 */
object PredicatePushdown:
  def apply(plan: LogicalPlan): LogicalPlan = plan.transform {
    case Filter(Put(parent, ref, value), predicate) =>
      // Can we move filter before put?
      if !predicateDependsOnCell(predicate, ref) then
        Put(Filter(parent, predicate), ref, value)
      else
        Filter(Put(parent, ref, value), predicate)

    case Filter(Merge(parent, range), predicate) =>
      // Can we move filter before merge?
      if !predicateDependsOnRange(predicate, range) then
        Merge(Filter(parent, predicate), range)
      else
        Filter(Merge(parent, range), predicate)

    case node => node
  }

  /**
   * Check if predicate depends on a specific cell.
   * Conservative analysis: returns true if unsure.
   */
  private def predicateDependsOnCell(p: Cell => Boolean, ref: ARef): Boolean =
    // TODO: Implement predicate analysis
    // For now, assume predicates are independent (optimistic)
    false

  private def predicateDependsOnRange(p: Cell => Boolean, range: CellRange): Boolean =
    false
```

#### 3.4 Cost-Based Optimizer

**Location:** `xl-core/src/com/tjclp/xl/optimizer/CostBasedOptimizer.scala`

```scala
package com.tjclp.xl.optimizer

import com.tjclp.xl.plan.*

/**
 * Statistics about a materialized sheet, used for cost-based optimization.
 */
case class Statistics(
  rowCount: Long,
  cellCount: Long,
  avgCellsPerRow: Double,
  distinctValues: Map[Column, Long],
  styleCount: Int
)

/**
 * Estimated cost of executing a plan.
 */
case class Cost(
  cpuOps: Long,      // Number of CPU operations
  memoryBytes: Long, // Peak memory usage
  ioBytes: Long      // Disk I/O volume
):
  def +(other: Cost): Cost = Cost(
    cpuOps + other.cpuOps,
    memoryBytes + other.memoryBytes,
    ioBytes + other.ioBytes
  )

  /** Total cost (weighted sum for comparison) */
  def total: Long = cpuOps + (memoryBytes / 1000) + (ioBytes / 1000)

/**
 * Cost-based optimizer that reorders operations based on estimated cost.
 *
 * Uses statistics collected from materialized sheets to estimate:
 * - CPU cost (number of operations)
 * - Memory cost (peak allocation)
 * - I/O cost (disk reads/writes)
 *
 * Selects execution plan with lowest total cost.
 */
object CostBasedOptimizer:
  /**
   * Optimize plan using cost-based analysis.
   * Requires statistics (typically collected during previous execution).
   */
  def apply(plan: LogicalPlan, stats: Option[Statistics] = None): LogicalPlan =
    // If no stats available, use default heuristics
    val s = stats.getOrElse(defaultStatistics)

    // Estimate cost of current plan
    val currentCost = estimateCost(plan, s)

    // Generate alternative plans by reordering operations
    val alternatives = generateAlternatives(plan)

    // Select plan with lowest cost
    val bestPlan = alternatives
      .map(p => (p, estimateCost(p, s)))
      .minByOption(_._2.total)
      .map(_._1)
      .getOrElse(plan)

    bestPlan

  /**
   * Estimate cost of executing a plan.
   */
  private def estimateCost(plan: LogicalPlan, stats: Statistics): Cost = plan match
    case BaseSheet(_) =>
      Cost(cpuOps = 0, memoryBytes = stats.cellCount * 100, ioBytes = 0)

    case Put(parent, _, _) =>
      val parentCost = estimateCost(parent, stats)
      parentCost + Cost(cpuOps = 10, memoryBytes = 100, ioBytes = 0)

    case PutAll(parent, cells) =>
      val parentCost = estimateCost(parent, stats)
      parentCost + Cost(cpuOps = cells.size * 5, memoryBytes = cells.size * 100, ioBytes = 0)

    case Filter(parent, _) =>
      val parentCost = estimateCost(parent, stats)
      // Assume filter reduces data by 50%
      val filteredCost = parentCost.copy(
        cpuOps = parentCost.cpuOps + stats.cellCount,
        memoryBytes = parentCost.memoryBytes / 2
      )
      filteredCost

    case Merge(parent, range) =>
      val parentCost = estimateCost(parent, stats)
      parentCost + Cost(cpuOps = range.size * 2, memoryBytes = 200, ioBytes = 0)

    case Limit(parent, n) =>
      val parentCost = estimateCost(parent, stats)
      parentCost.copy(
        memoryBytes = Math.min(parentCost.memoryBytes, n * 100)
      )

    case SetStyle(parent, _, _) =>
      val parentCost = estimateCost(parent, stats)
      parentCost + Cost(cpuOps = 5, memoryBytes = 50, ioBytes = 0)

  /**
   * Generate alternative execution plans by reordering operations.
   */
  private def generateAlternatives(plan: LogicalPlan): List[LogicalPlan] =
    // TODO: Implement plan generation
    // For now, return just the original plan
    List(plan)

  private val defaultStatistics = Statistics(
    rowCount = 1000,
    cellCount = 10000,
    avgCellsPerRow = 10.0,
    distinctValues = Map.empty,
    styleCount = 10
  )
```

### 4. Streaming Executor

**Location:** `xl-cats-effect/src/com/tjclp/xl/execution/StreamingExecutor.scala`

```scala
package com.tjclp.xl.execution

import cats.effect.{Async, IO}
import fs2.{Stream, Chunk}
import com.tjclp.xl.plan.*
import com.tjclp.xl.api.*
import java.nio.file.Path

/**
 * Executor that uses fs2 streams for constant-memory processing.
 *
 * Key features:
 * - O(1) memory regardless of dataset size
 * - Backpressure handling (never overwhelms downstream)
 * - Chunked processing (configurable chunk size)
 * - Parallel execution (configurable parallelism)
 *
 * Performance:
 * - 1M rows: ~4s write time with ~10MB memory
 * - 10M rows: ~40s write time with ~10MB memory
 * - Memory usage independent of dataset size
 */
object StreamingExecutor:
  /**
   * Execute logical plan as a stream of cells.
   * Stream is lazy and only materializes chunks as needed.
   */
  def executeStream[F[_]: Async](plan: LogicalPlan): Stream[F, Cell] = plan match
    case BaseSheet(sheet) =>
      // Stream cells from materialized sheet
      Stream.iterable(sheet.cells.values)

    case Put(parent, ref, value) =>
      // Append single cell to parent stream
      executeStream(parent) ++ Stream.emit(Cell(ref, value))

    case PutAll(parent, cells) =>
      // Append multiple cells to parent stream
      executeStream(parent) ++ Stream.emits(cells)

    case Filter(parent, predicate) =>
      // Filter cells in stream (lazy)
      executeStream(parent).filter(predicate)

    case Limit(parent, n) =>
      // Take first n cells from stream
      executeStream(parent).take(n.toLong)

    case Merge(parent, range) =>
      // Merges don't affect cell stream (metadata only)
      executeStream(parent)

    case SetStyle(parent, ref, style) =>
      // Update style for matching cell in stream
      executeStream(parent).map { cell =>
        if cell.ref == ref then cell.copy(styleId = Some(0)) // TODO: style registry
        else cell
      }

  /**
   * Count cells without materializing them.
   * More efficient than executeStream(...).compile.count because
   * it can short-circuit for some plan types.
   */
  def count[F[_]: Async](plan: LogicalPlan): F[Long] = plan match
    case BaseSheet(sheet) =>
      Async[F].pure(sheet.cells.size.toLong)

    case Put(parent, _, _) =>
      count(parent).map(_ + 1)

    case PutAll(parent, cells) =>
      count(parent).map(_ + cells.size)

    case Filter(parent, predicate) =>
      // Must materialize to count filtered results
      executeStream(parent).filter(predicate).compile.count

    case Limit(parent, n) =>
      count(parent).map(c => Math.min(c, n.toLong))

    case Merge(parent, _) | SetStyle(parent, _, _) =>
      count(parent)

  /**
   * Write plan to Excel file using streaming.
   * Uses constant memory regardless of file size.
   */
  def writeStream[F[_]: Async](plan: LogicalPlan, path: Path): F[Unit] =
    val optimized = Optimizer.optimize(plan)

    executeStream[F](optimized)
      .chunkN(1000)  // Process in chunks of 1000 cells
      .through(cellsToXml[F])
      .through(xmlToBytes[F])
      .through(fs2.io.file.Files[F].writeAll(fs2.io.file.Path.fromNioPath(path)))
      .compile
      .drain

  /**
   * Show first n cells as formatted table.
   * Useful for REPL debugging.
   */
  def show[F[_]: Async](plan: LogicalPlan, n: Int): F[String] =
    executeStream[F](plan)
      .take(n.toLong)
      .compile
      .toList
      .map(formatAsTable)

  /**
   * Convert stream of cells to XML elements.
   */
  private def cellsToXml[F[_]]: fs2.Pipe[F, Chunk[Cell], Chunk[scala.xml.Elem]] = ???

  /**
   * Convert XML elements to byte stream.
   */
  private def xmlToBytes[F[_]]: fs2.Pipe[F, Chunk[scala.xml.Elem], Byte] = ???

  /**
   * Format cells as ASCII table for REPL display.
   */
  private def formatAsTable(cells: List[Cell]): String = ???
```

**Streaming Configuration:**

```scala
/**
 * Configuration for streaming execution.
 */
case class StreamingConfig(
  chunkSize: Int = 1000,        // Process 1000 cells at a time
  maxBufferSize: Int = 10000,   // Buffer up to 10k cells
  parallelism: Int = 4           // Parallelize across 4 streams
)

object StreamingConfig:
  val default: StreamingConfig = StreamingConfig()
```

### 5. Executor (Non-Streaming)

**Location:** `xl-core/src/com/tjclp/xl/execution/Executor.scala`

```scala
package com.tjclp.xl.execution

import com.tjclp.xl.plan.*
import com.tjclp.xl.api.*
import cats.effect.IO
import java.nio.file.Path

/**
 * Basic executor that materializes entire sheet in memory.
 * Use this for small sheets or when you need immediate access to all cells.
 *
 * For large sheets (>100k cells), use StreamingExecutor instead.
 */
object Executor:
  /**
   * Execute logical plan and materialize full sheet.
   */
  def execute(plan: LogicalPlan): MaterializedSheet =
    val optimized = Optimizer.optimize(plan)
    executePhysical(optimized)

  private def executePhysical(plan: LogicalPlan): MaterializedSheet = plan match
    case BaseSheet(sheet) =>
      sheet

    case Put(parent, ref, value) =>
      executePhysical(parent).put(Cell(ref, value))

    case PutAll(parent, cells) =>
      executePhysical(parent).putAll(cells)

    case Merge(parent, range) =>
      executePhysical(parent).merge(range)

    case Filter(parent, predicate) =>
      val sheet = executePhysical(parent)
      val filtered = sheet.cells.values.filter(predicate)
      sheet.copy(cells = filtered.map(c => c.ref -> c).toMap)

    case Limit(parent, n) =>
      val sheet = executePhysical(parent)
      val limited = sheet.cells.values.take(n)
      sheet.copy(cells = limited.map(c => c.ref -> c).toMap)

    case SetStyle(parent, ref, style) =>
      val sheet = executePhysical(parent)
      val (registry, styleId) = sheet.styleRegistry.register(style)
      sheet
        .copy(styleRegistry = registry)
        .withCellStyle(ref, styleId)

  /**
   * Count cells by executing plan.
   */
  def count(plan: LogicalPlan): Long =
    execute(plan).cells.size.toLong

  /**
   * Show first n cells of plan.
   */
  def show(plan: LogicalPlan, n: Int): String =
    val sheet = execute(plan)
    val cells = sheet.cells.values.take(n).toList
    formatAsTable(cells)

  /**
   * Write plan to file (delegates to streaming for efficiency).
   */
  def writeStream(plan: LogicalPlan, path: Path): IO[Unit] =
    StreamingExecutor.writeStream[IO](plan, path)

  private def formatAsTable(cells: List[Cell]): String =
    // TODO: Implement ASCII table formatting
    cells.map(c => s"${c.ref.toA1}: ${c.value}").mkString("\n")
```

## API Migration Guide

### Breaking Changes

1. **`Sheet` renamed to `MaterializedSheet`** (internal use)
2. **`LazySheet` is now the public API**
3. **Must call `.collect()` to materialize**
4. **Actions return `IO[Unit]` instead of direct results**

### Before (Eager)

```scala
import com.tjclp.xl.api.*

val sheet = Sheet(SheetName.unsafe("Data"))
  .put(cell"A1", CellValue.Text("Title"))
  .put(cell"A2", CellValue.Number(100))
  .merge(range"A1:B1")

val workbook = Workbook(Vector(sheet))
ExcelIO.write(workbook, path)
```

### After (Lazy)

```scala
import com.tjclp.xl.api.*
import cats.effect.unsafe.implicits.global

// Option 1: Direct streaming write (recommended)
val sheet = LazySheet("Data")
  .put(cell"A1", "Title")
  .put(cell"A2", 100)
  .merge(range"A1:B1")

sheet.write(path).unsafeRunSync()  // Executes optimized plan, streams to disk

// Option 2: Materialize for testing/inspection
val materialized: MaterializedSheet = sheet.collect()
assert(materialized.get(cell"A1").isDefined)

// Option 3: Count without materializing
val cellCount: Long = sheet.count
println(s"Sheet has $cellCount cells")

// Option 4: Show preview (REPL)
println(sheet.show(10))  // Show first 10 cells
```

### Code Examples

#### Example 1: Large Dataset Processing

```scala
// Process 1M rows with constant memory
val sheet = LazySheet("BigData")
  .putAll(generateMillionCells())  // Lazy - not executed yet!
  .filter(_.value.isNumber)        // Lazy filter
  .limit(100000)                   // Lazy limit

// Only now does computation happen (streaming):
sheet.write(path).unsafeRunSync()  // ~10MB memory, 4s write time
```

#### Example 2: Conditional Workflows

```scala
def buildReport(config: ReportConfig): LazySheet =
  var sheet = LazySheet("Report")
    .put(cell"A1", config.title)

  if config.includeData then
    sheet = sheet.putAll(loadData())  // May not execute!

  if config.includeChart then
    sheet = sheet.merge(range"A1:D10")  // May not execute!

  sheet  // Returns plan, no execution yet

// Only execute the operations needed:
buildReport(ReportConfig(includeData = true, includeChart = false))
  .write(path)
  .unsafeRunSync()
```

#### Example 3: Query Optimization

```scala
val sheet = LazySheet("Optimizable")
  .put(cell"A1", "x")      // Operation 1
  .put(cell"A1", "y")      // Operation 2 (overwrites 1)
  .put(cell"A2", "z")      // Operation 3
  .merge(range"A1:B1")     // Operation 4

// View optimization:
println(sheet.explain())
// == Original Plan ==
// Merge [A1:B1]
//   Put [A2, z]
//     Put [A1, y]
//       Put [A1, x]
//         BaseSheet [Optimizable]
//
// == Optimized Plan ==
// Merge [A1:B1]
//   PutAll [A1->y, A2->z]  // Batched! First put eliminated!
//     BaseSheet [Optimizable]
//
// Estimated Savings: 33% operation reduction, 50% allocation reduction
```

## Implementation Phases

### Phase 1: Logical Plan Foundation (Week 1)

**Goal:** Basic lazy evaluation with collect() action

**Tasks:**
- [ ] Create LogicalPlan ADT with core operations (Put, PutAll, Merge, Filter)
- [ ] Create LazySheet API wrapping LogicalPlan
- [ ] Rename Sheet → MaterializedSheet
- [ ] Implement basic Executor.execute() (no optimization)
- [ ] Implement .collect() action
- [ ] Update 50 core tests to use LazySheet API
- [ ] Add plan equivalence tests (verify lazy produces same result as eager)

**Deliverable:** Working lazy evaluation with .collect()

### Phase 2: Basic Optimizations (Week 2)

**Goal:** Batching and dead code elimination

**Tasks:**
- [ ] Implement BatchingOptimizer (consecutive puts → putAll)
- [ ] Implement DeadCodeEliminator (remove overwritten operations)
- [ ] Implement OperationFusion (combine filters)
- [ ] Create Optimizer pipeline (3 passes)
- [ ] Add optimizer tests (verify correctness and savings)
- [ ] Add .explain() method to show optimization steps
- [ ] Measure performance improvement (expect 20-30% reduction)

**Deliverable:** Working optimizer with 3 passes

### Phase 3: Advanced Optimizations (Week 3)

**Goal:** Predicate pushdown and cost-based optimization

**Tasks:**
- [ ] Implement PredicatePushdown
- [ ] Implement StatisticsCollector
- [ ] Implement CostBasedOptimizer with cost estimation
- [ ] Add statistics tracking to MaterializedSheet
- [ ] Add optimizer configuration (enable/disable passes)
- [ ] Benchmark: measure optimization impact on various workloads
- [ ] Document optimization rules and when they apply

**Deliverable:** Full 4-pass optimizer with cost-based analysis

### Phase 4: Streaming Execution (Week 4)

**Goal:** Constant-memory processing with fs2

**Tasks:**
- [ ] Implement StreamingExecutor.executeStream()
- [ ] Implement StreamingExecutor.writeStream()
- [ ] Add chunking and backpressure handling
- [ ] Implement StreamingConfig
- [ ] Add streaming tests (verify constant memory with 1M+ cells)
- [ ] Benchmark: measure streaming performance vs eager
- [ ] Integrate with existing OOXML serializers

**Deliverable:** Streaming executor with O(1) memory

### Phase 5: Actions & Ergonomics (Week 5)

**Goal:** Complete action API and REPL ergonomics

**Tasks:**
- [ ] Implement .count action (optimized)
- [ ] Implement .show(n) action with ASCII table formatting
- [ ] Implement .write(path) action (streaming)
- [ ] Add .explain() with detailed optimization breakdown
- [ ] Add REPL helpers (pretty printing, summaries)
- [ ] Write migration guide
- [ ] Update all 263 tests to use new API
- [ ] Update README and documentation

**Deliverable:** Complete lazy evaluation system with all actions

### Phase 6: Polish & Release (Week 6)

**Goal:** Production-ready release

**Tasks:**
- [ ] Performance benchmarks (publish results)
- [ ] Memory profiling (verify O(1) with large datasets)
- [ ] Error handling improvements
- [ ] API documentation (Scaladoc)
- [ ] Tutorial examples
- [ ] Blog post: "Spark-style lazy evaluation in Scala"
- [ ] Release notes
- [ ] Version 2.0.0 release

**Deliverable:** XL 2.0 with lazy evaluation

## Performance Targets

### Before (Current Eager)

- **1M rows**: 6.2s write time, O(n) memory (~100MB for 1M cells)
- **10k unique styles**: 2.5s style indexing
- **Intermediate allocations**: 3 Sheets per 3 operations

### After (Optimized Lazy)

- **1M rows**: 4.0s write time (35% faster), O(1) memory (~10MB constant)
- **10k unique styles**: 0.8s style indexing (68% faster via batched deduplication)
- **Intermediate allocations**: 1 Sheet per collect() (66% reduction)
- **Query optimization**: 20-40% operation reduction on typical workloads

### Memory Scaling

| Dataset Size | Eager (Current) | Lazy (Streaming) |
|--------------|-----------------|------------------|
| 10k cells    | ~1 MB           | ~10 MB           |
| 100k cells   | ~10 MB          | ~10 MB           |
| 1M cells     | ~100 MB         | ~10 MB           |
| 10M cells    | ~1 GB           | ~10 MB           |

Lazy evaluation with streaming provides **constant memory** regardless of dataset size.

## Testing Strategy

### Test Categories

1. **Plan Equivalence Tests** (30 tests)
   - Verify lazy produces same result as eager
   - Test all operations (put, merge, filter, etc.)
   - Test complex operation chains

2. **Optimizer Tests** (40 tests)
   - Test each optimization pass independently
   - Test optimizer pipeline
   - Verify optimizations preserve semantics
   - Test optimization savings measurement

3. **Streaming Tests** (30 tests)
   - Verify constant memory with large datasets
   - Test backpressure handling
   - Test chunking configuration
   - Benchmark streaming vs eager

4. **Action Tests** (20 tests)
   - Test write() action (streaming)
   - Test collect() action (materialization)
   - Test count() action (optimized)
   - Test show() action (formatting)
   - Test explain() action (plan display)

5. **Integration Tests** (30 tests)
   - Test full workflows (read → transform → write)
   - Test large datasets (1M+ cells)
   - Test complex optimizations
   - Test error handling

### Example Tests

```scala
class LazySheetSpec extends munit.FunSuite:
  test("lazy put produces same result as eager put"):
    val lazy = LazySheet("Test")
      .put(cell"A1", "Hello")
      .collect()

    val eager = MaterializedSheet(SheetName.unsafe("Test"))
      .put(cell"A1", CellValue.Text("Hello"))

    assertEquals(lazy.cells, eager.cells)

  test("batching optimizer combines consecutive puts"):
    val plan = LazySheet("Test")
      .put(cell"A1", "A")
      .put(cell"A2", "B")
      .put(cell"A3", "C")
      .plan

    val optimized = Optimizer.optimize(plan)

    optimized match
      case PutAll(_, cells) =>
        assertEquals(cells.size, 3)
      case _ => fail("Expected PutAll")

  test("streaming processes 1M cells with constant memory"):
    val sheet = LazySheet("Big")
      .putAll((1 to 1_000_000).map(i => Cell(cell"A$i", CellValue.Number(i))).toVector)

    // Measure memory before and after
    val memBefore = Runtime.getRuntime.totalMemory - Runtime.getRuntime.freeMemory
    sheet.write(path).unsafeRunSync()
    val memAfter = Runtime.getRuntime.totalMemory - Runtime.getRuntime.freeMemory

    val memUsed = (memAfter - memBefore) / (1024 * 1024)  // MB
    assert(memUsed < 50, s"Used $memUsed MB (expected < 50 MB)")
```

## Open Questions

1. **Should we keep eager API alongside lazy?**
   - Pro: Easier migration, users can choose
   - Con: Two APIs to maintain, confusion

2. **Should filters be pushed all the way to file reading?**
   - Pro: Only load matching cells from disk
   - Con: Complex, requires OOXML streaming reader

3. **Should we support Spark-style RDD operations?**
   - map, flatMap, groupBy, reduce, etc.
   - Pro: Powerful, familiar to Spark users
   - Con: Complex, may not fit Excel model well

4. **How to handle style registry in streaming mode?**
   - Styles need deduplication across entire workbook
   - Streaming processes chunks independently
   - Need two-pass approach? Or sacrifice deduplication?

5. **Should actions return IO or direct values?**
   - Current plan: count → Long (direct), write → IO[Unit]
   - Alternative: All actions return IO (more consistent)

---

## Recommended Scope (Practical Implementation)

### What to Actually Build: Builder Pattern (1-2 Weeks)

After technical review, the recommendation is to implement **Phase 1 only** (builder pattern) and **skip the full optimizer** (Phase 2-6). Here's the practical, high-value implementation:

#### SheetBuilder API

**Location**: `xl-core/src/com/tjclp/xl/SheetBuilder.scala`

```scala
package com.tjclp.xl

import scala.collection.mutable

/**
 * Mutable builder for batching sheet operations.
 *
 * Accumulates put/merge/style operations in memory and applies them in a single pass
 * when build() is called. This eliminates intermediate Sheet allocations (30-50% speedup
 * for large batch operations).
 *
 * NOT thread-safe. Create one builder per thread.
 *
 * Example:
 * {{{
 * val builder = SheetBuilder("Sales")
 * (1 to 10000).foreach(i => builder.put(cell"A\$i", s"Row \$i"))
 * val sheet = builder.build  // Single allocation, ~30% faster than fold
 * }}}
 */
class SheetBuilder(name: SheetName):
  private val cellBuffer = mutable.Map[ARef, Cell]()
  private val styleBuffer = mutable.Map[ARef, CellStyle]()
  private val mergeBuffer = mutable.Set[CellRange]()

  /**
   * Queue a put operation (deferred).
   */
  def put(ref: ARef, value: CellValue): this.type =
    cellBuffer(ref) = Cell(ref, value)
    this

  /**
   * Queue a put with style (deferred).
   */
  def putStyled(ref: ARef, value: CellValue, style: CellStyle): this.type =
    cellBuffer(ref) = Cell(ref, value)
    styleBuffer(ref) = style
    this

  /**
   * Queue a merge operation (deferred).
   */
  def merge(range: CellRange): this.type =
    mergeBuffer += range
    this

  /**
   * Execute all queued operations and return materialized sheet.
   * This is the only point where computation happens.
   */
  def build: Sheet =
    // Start with base sheet
    var sheet = Sheet(name)

    // Apply all cells in one pass (single putAll)
    if cellBuffer.nonEmpty then
      sheet = sheet.putAll(cellBuffer.values)

    // Register and apply styles in one pass
    if styleBuffer.nonEmpty then
      var registry = sheet.styleRegistry
      val styled = cellBuffer.view.filterKeys(styleBuffer.contains).map { (ref, cell) =>
        val style = styleBuffer(ref)
        val (newReg, styleId) = registry.register(style)
        registry = newReg
        (ref, cell.copy(styleId = Some(styleId.value)))
      }
      sheet = sheet.copy(
        cells = sheet.cells ++ styled,
        styleRegistry = registry
      )

    // Apply merges
    if mergeBuffer.nonEmpty then
      sheet = sheet.copy(mergedRanges = sheet.mergedRanges ++ mergeBuffer)

    sheet

object SheetBuilder:
  def apply(name: SheetName): SheetBuilder = new SheetBuilder(name)
  def apply(name: String): SheetBuilder = new SheetBuilder(SheetName.unsafe(name))
```

#### Integration with putMixed

**Enhance builder to support codec-based API**:

```scala
import com.tjclp.xl.codec.syntax.*

class SheetBuilder(name: SheetName):
  // ... existing fields ...

  /**
   * Queue multiple typed puts with auto-inferred formatting (deferred).
   */
  def putMixed(updates: (ARef, Any)*): this.type =
    updates.foreach { case (ref, value) =>
      // Use CellCodec to encode value + infer style
      value match
        case v: String => put(ref, CellValue.Text(v))
        case v: Int => put(ref, CellValue.Number(BigDecimal(v)))
        case v: LocalDate =>
          putStyled(ref, CellValue.DateTime(v.atStartOfDay), inferDateStyle())
        case v: BigDecimal =>
          putStyled(ref, CellValue.Number(v), inferDecimalStyle())
        // ... more types
    }
    this
```

#### Benefits

- **30-50% faster** than `foldLeft` with `put` for large batches
- **Single allocation** instead of N intermediate sheets
- **Type-safe** with putMixed integration
- **Simple** - no optimizer complexity
- **Transparent** - easy to debug (just inspect buffer)

#### Estimated Work

- Implementation: 1-2 days
- Tests: 1 day (builder semantics, equivalence to eager, performance tests)
- Documentation: 1 day

**Total**: 3-4 days (vs 4-5 weeks for full optimizer)

---

### What to Skip: Full Optimizer (Phase 2-6)

**Deferred indefinitely** based on technical review:

#### Predicate Pushdown (Minimal ROI)
- **Problem**: Excel workloads are sparse cell updates, not SQL filtering
- **ROI**: Would save ~5-10% operations in typical workloads
- **Cost**: Complex predicate analysis, plan rewriting (2-3 weeks work)
- **Verdict**: Not worth complexity

#### Cost-Based Optimizer (Heavy Machinery)
- **Problem**: Requires statistics collection, plan enumeration, cost estimation
- **ROI**: Might save 10-20% operations for complex plans
- **Cost**: 3-4 weeks implementation + ongoing maintenance
- **Verdict**: Overkill for Excel (this is for SQL query planners)

#### LogicalPlan AST (Overhead Without Benefit)
- **Problem**: Building AST adds allocation overhead
- **ROI**: Only useful if optimizer passes provide value (they don't)
- **Cost**: 2-3 weeks to rewrite API
- **Verdict**: Builder pattern achieves same batching benefit without AST

---

### What to Prioritize Instead: Streaming Improvements

Focus on **actual bottlenecks** (see streaming-improvements.md):

1. **P6.6**: Fix streaming reader (2-3 days) - **CRITICAL**
   - Replace `readAllBytes()` with `fs2.io.readInputStream`
   - Achieve true O(1) memory for reads

2. **P6.7**: Compression defaults (1 day) - **QUICK WIN**
   - Default to DEFLATED (5-10x smaller files)
   - Configurable compression mode

3. **P7.5**: Two-phase streaming writer (3-4 weeks) - **HIGH VALUE**
   - Support SST and styles in streaming mode
   - Maintain O(1) memory with full features

**Total**: 4-5 weeks for streaming improvements vs 4-5 weeks for full optimizer

**Impact**: Streaming improvements provide **much higher ROI** than optimizer

---

## Revised Implementation Plan

### Phase 1: Builder Pattern (Recommended - 1 Week)

**Goal**: Reduce intermediate allocations for batch operations

**Tasks**:
- [ ] Implement `SheetBuilder` with mutable buffers
- [ ] Add `putMixed` integration for type-safe batching
- [ ] Add performance tests (compare to foldLeft baseline)
- [ ] Document usage patterns
- [ ] Add to README.md examples

**Deliverable**: 30-50% speedup for batch operations, simple API

**Timeline**: 3-4 days implementation + 1-2 days testing/docs = 1 week total

---

### Phase 2-6: Full Optimizer (NOT RECOMMENDED - Deferred)

**Original Plan**:
- Phase 2: Basic optimizations (batching, dead code elimination)
- Phase 3: Advanced optimizations (predicate pushdown, cost-based)
- Phase 4: Streaming execution
- Phase 5: Actions and ergonomics
- Phase 6: Polish

**Status**: **DEFERRED INDEFINITELY**

**Reason**: Complexity doesn't justify gains for Excel use cases. Focus on streaming improvements instead.

---

## Original Vision (Full Optimizer - Now Considered Overkill)

> **Note**: The sections below describe the original full-optimizer vision.
> This is preserved for reference but **not recommended for implementation**.
> See "Recommended Scope" above for practical guidance.

## Vision
