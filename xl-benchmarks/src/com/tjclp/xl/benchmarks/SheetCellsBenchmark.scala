package com.tjclp.xl.benchmarks

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.units.StyleId
import org.openjdk.jmh.annotations.*
import org.openjdk.jmh.infra.Blackhole
import java.util.concurrent.TimeUnit
import scala.compiletime.uninitialized
import scala.util.Random

/**
 * The cost of a sheet's cell storage, measured through the public `Sheet` API only (#891, PR B).
 *
 * Written so it compiles and runs unchanged before and after `Sheet.cells` moves from
 * `Map[ARef, Cell]` to `CellStore`: every operation goes through `sheet(ref)`, `sheet.put`,
 * `sheet.cells.get`, `sheet.cells.size`, `sheet.usedRange`, `sheet.getRange` and `==`. The one
 * exception is full iteration, which called `sheet.cells.valuesIterator` on the Map and calls
 * `sheet.cells.iterator` since the swap (the only line that changed).
 *
 * Shapes (`shape`), each holding `size` cells:
 *   - `dense20`: rows of 20 contiguous columns (a typical table)
 *   - `wide500`: rows of 500 contiguous columns (a wide model sheet)
 *   - `column`: one column, `size` rows
 *   - `sparse`: `size` distinct cells at random positions anywhere on Excel's grid
 *
 * Sizes 1k, 10k and 100k run by default; pass `-p size=1000000` for the 1M point.
 *
 * Every 20th cell carries a style and every 100th a hyperlink, so the style and link paths are
 * exercised the way a real book exercises them. All randomness uses fixed seeds.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Array(Mode.AverageTime))
@OutputTimeUnit(TimeUnit.MICROSECONDS)
@Warmup(iterations = 3, time = 1, timeUnit = TimeUnit.SECONDS)
@Measurement(iterations = 5, time = 2, timeUnit = TimeUnit.SECONDS)
@Fork(value = 1)
// JMH state needs `uninitialized` fields (Null) and the measured loops are plain while loops.
@SuppressWarnings(
  Array("org.wartremover.warts.Var", "org.wartremover.warts.While", "org.wartremover.warts.Null")
)
class SheetCellsBenchmark {
  import SheetCellsBenchmark.*

  @Param(Array("dense20", "wide500", "column", "sparse"))
  var shape: String = uninitialized

  @Param(Array("1000", "10000", "100000"))
  var size: Int = uninitialized

  /** The sheet under test, built once per trial. */
  var sheet: Sheet = uninitialized

  /** An independently built copy of `sheet` whose last row-major cell holds a different value. */
  var sheetOneOff: Sheet = uninitialized

  /** The sheet's cells in row-major order, and the same cells shuffled. */
  var sortedCells: Array[Cell] = uninitialized
  var shuffledCells: Array[Cell] = uninitialized

  /** `Probes` refs that hold a cell, and `Probes` refs that do not. */
  var hitRefs: Array[ARef] = uninitialized
  var missRefs: Array[ARef] = uninitialized

  /** 10k refs appended after the shape in row-major order, and 5k random refs within its box. */
  var appendRefs: Array[ARef] = uninitialized
  var randomRefs: Array[ARef] = uninitialized

  /** A ~1,000-position box in the middle of the shape. */
  var scanRange: CellRange = uninitialized

  /** The column the 50k-row scan walks (the shape's first column). */
  var scanCol: Int = 0

  var singlePutRef: ARef = uninitialized

  @Setup(Level.Trial)
  def setup(): Unit = {
    val positions = layout(shape, size, new Random(42))
    sortedCells = positions.zipWithIndex.map { case ((row, col), i) => cellAt(row, col, i) }
    shuffledCells = new Random(7).shuffle(sortedCells.toSeq).toArray
    sheet = build(sortedCells)

    // Built independently (not `sheet.put`), so equality cannot short-circuit on shared structure.
    val last = sortedCells.length - 1
    val oneOff = sortedCells.clone()
    oneOff(last) = oneOff(last).withValue(CellValue.Text("different"))
    sheetOneOff = build(oneOff)

    val random = new Random(99)
    hitRefs = Array.fill(Probes)(sortedCells(random.nextInt(sortedCells.length)).ref)
    val (maxRow, maxCol) = positions.foldLeft((0, 0)) { case ((mr, mc), (r, c)) =>
      (math.max(mr, r), math.max(mc, c))
    }
    // Misses sit inside the shape's bounding box (where a lookup would plausibly land) but avoid
    // every stored cell.
    val occupied = positions.toSet
    missRefs = Iterator
      .continually((random.nextInt(maxRow + 2), random.nextInt(maxCol + 2)))
      .filterNot(occupied.contains)
      .take(Probes)
      .map((r, c) => ARef.from0(c, r))
      .toArray

    val width = shape match
      case "dense20" => 20
      case "wide500" => 500
      case _ => 1
    // Appended rows stay on Excel's grid (the sparse shape already reaches its last rows).
    val appendStart = math.min(maxRow + 1, 1048576 - AppendCount / width - 1)
    appendRefs = Array.tabulate(AppendCount)(i => ARef.from0(i % width, appendStart + i / width))
    randomRefs = Array.fill(RandomPutCount)(
      ARef.from0(random.nextInt(maxCol + 1), random.nextInt(maxRow + 1))
    )

    val boxRows = math.max(1, 1000 / math.min(width, 10))
    val midRow = positions(positions.length / 2)._1
    scanRange = CellRange(
      ARef.from0(0, midRow),
      ARef.from0(math.min(width, 10) - 1, midRow + boxRows - 1)
    )
    scanCol = positions(0)._2
    singlePutRef = sortedCells(sortedCells.length / 2).ref
  }

  // ----- reads -----

  @Benchmark
  @OperationsPerInvocation(Probes)
  def getHit(bh: Blackhole): Unit = {
    var i = 0
    while i < hitRefs.length do
      bh.consume(sheet(hitRefs(i)))
      i += 1
  }

  @Benchmark
  @OperationsPerInvocation(Probes)
  def getMiss(bh: Blackhole): Unit = {
    var i = 0
    while i < missRefs.length do
      bh.consume(sheet.cells.get(missRefs(i)))
      i += 1
  }

  /** Full iteration (`valuesIterator` on the Map before the swap; nothing else changed). */
  @Benchmark
  def iterateAll(): Int = {
    val it = sheet.cells.iterator
    var n = 0
    while it.hasNext do if it.next().nonEmpty then n += 1
    n
  }

  /**
   * 50k single-cell reads down one column via `sheet(ref)`, the evaluator's column-scan pattern. On
   * the small and dense shapes most of these rows are past the data, so they are misses.
   */
  @Benchmark
  def columnScan50k(bh: Blackhole): Unit = {
    var row = 0
    while row < ColumnScanRows do
      bh.consume(sheet(ARef.from0(scanCol, row)).value)
      row += 1
  }

  @Benchmark
  def rangeScan(): Int = sheet.getRange(scanRange).size

  @Benchmark
  def usedRange(): Option[CellRange] = sheet.usedRange

  @Benchmark
  def size_(): Int = sheet.cells.size

  @Benchmark
  def equalityOneCellDiffers(): Boolean = sheet == sheetOneOff

  // ----- writes -----

  @Benchmark
  def singlePut(): Sheet = sheet.put(singlePutRef, CellValue.Number(BigDecimal(1)))

  /** 10k puts appended after the data, row-major (a writer filling a sheet). */
  @Benchmark
  def putChain10k(): Sheet = {
    var s = sheet
    var i = 0
    while i < appendRefs.length do
      s = s.put(appendRefs(i), CellValue.Number(BigDecimal(i)))
      i += 1
    s
  }

  /** 5k puts at random positions inside the shape's box, folded one at a time. */
  @Benchmark
  def randomPuts5k(): Sheet = {
    var s = sheet
    var i = 0
    while i < randomRefs.length do
      s = s.put(randomRefs(i), CellValue.Number(BigDecimal(i)))
      i += 1
    s
  }

  // ----- construction -----

  @Benchmark
  def buildSorted(): Sheet = build(sortedCells)

  @Benchmark
  def buildUnsorted(): Sheet = build(shuffledCells)
}

@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
object SheetCellsBenchmark {
  final val Probes = 1024
  final val AppendCount = 10000
  final val RandomPutCount = 5000
  final val ColumnScanRows = 50000

  /** Row-major (row, col) positions for `size` cells of `shape`. */
  def layout(shape: String, size: Int, random: Random): Array[(Int, Int)] = shape match
    case "dense20" => Array.tabulate(size)(i => (i / 20, i % 20))
    case "wide500" => Array.tabulate(size)(i => (i / 500, i % 500))
    case "column" => Array.tabulate(size)(i => (i, 0))
    case "sparse" =>
      val seen = scala.collection.mutable.HashSet.empty[(Int, Int)]
      while seen.size < size do seen += ((random.nextInt(1048576), random.nextInt(16384)))
      seen.toArray.sorted
    case other => throw new IllegalArgumentException(s"unknown shape $other")

  /** The i-th benchmark cell: mixed values, a style on every 20th, a link on every 100th. */
  def cellAt(row: Int, col: Int, i: Int): Cell = {
    val value = i % 4 match
      case 0 => CellValue.Number(BigDecimal(i))
      case 1 => CellValue.Text(s"t$i")
      case 2 => CellValue.Bool(i % 8 == 2)
      case _ => CellValue.Number(BigDecimal(i) / 7)
    Cell(
      ARef.from0(col, row),
      value,
      if i % 20 == 0 then Some(StyleId(i % 13)) else None,
      if i % 100 == 0 then Some(Hyperlink(s"https://example.com/$i")) else None
    )
  }

  /** A sheet built from `cells` by a fold of `put(cell)`, in the order given. */
  def build(cells: Array[Cell]): Sheet = {
    var s: Sheet = Sheet("Cells")
    var i = 0
    while i < cells.length do
      s = s.put(cells(i))
      i += 1
    s
  }
}
