package com.tjclp.xl.formula.eval

import java.util.concurrent.ConcurrentHashMap
import java.util.concurrent.atomic.{AtomicLong, AtomicReference}

import scala.collection.immutable.ArraySeq

import com.tjclp.xl.addressing.{ARef, CellRange, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.graph.FormulaParses
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}

/**
 * The slot a formula's top-level evaluation occupies in its recalculation pass (#767): its index in
 * the sequential order, or its wave under parallel waves. Only an evaluation at its own slot may
 * build or read a [[ColumnSlice]]; a precedent evaluated on demand, a dynamic-bucket member and
 * every evaluation outside a recalculation pass carry [[EvalPosition.None]].
 */
private[formula] opaque type EvalPosition = Int

private[formula] object EvalPosition:
  val None: EvalPosition = -1

  /** Slot `slot` (non-negative) of the pass's order. */
  def at(slot: Int): EvalPosition = if slot < 0 then None else slot

  extension (position: EvalPosition)
    def isDefined: Boolean = position >= 0
    def slot: Int = position

/**
 * #767: which formula cells of a recalculation pass are done before a given slot, by column.
 *
 * The pass builds it from the order it is about to run. `slots` holds every formula the pass
 * evaluates and the slot it evaluates at (sequential index or wave); `stripped` the formulas whose
 * cache the pass stripped and that it does not evaluate (cyclic and blocked cells). Formulas in
 * neither keep a cache the pass never changes (data tables; cached cells outside the targets of
 * `recalculateUncached`; cells outside an edit's cone), so they constrain nothing.
 *
 * A slice of rows `r0..r1` of one column may be built at slot `p` when every formula of the box
 * that the pass evaluates has a slot below `p` and none is stripped: those cells then hold this
 * pass's final values, and every other cell of the box holds a value the pass never changes. Each
 * column keeps its formula rows sorted, a sparse table of their slots (range maximum in O(1)) and a
 * prefix count of stripped rows.
 */
private[formula] final class Readiness private (
  columns: Map[(SheetName, Int), Readiness.Column]
):
  /** Whether rows `r0..r1` of `col` on `sheet` are final for a reader at `at`. */
  def allows(sheet: SheetName, col: Int, r0: Int, r1: Int, at: EvalPosition): Boolean =
    if !at.isDefined then false
    else if r0 > r1 then true
    else columns.get((sheet, col)).forall(_.allows(r0, r1, at.slot))

private[formula] object Readiness:

  /** One column's formula rows, their slots (-1: not evaluated) and stripped flags. */
  private final class Column(
    rows: IArray[Int],
    // levels(k)(i) = max slot over rows i until i + 2^k
    levels: IArray[IArray[Int]],
    // strippedBefore(i) = stripped rows among rows(0 until i)
    strippedBefore: IArray[Int]
  ):
    def allows(r0: Int, r1: Int, slot: Int): Boolean =
      val i = lowerBound(rows, r0)
      val j = lowerBound(rows, r1 + 1) - 1
      if i > j then true
      else strippedBefore(j + 1) - strippedBefore(i) == 0 && rangeMax(i, j) < slot

    private def rangeMax(i: Int, j: Int): Int =
      val k = 31 - Integer.numberOfLeadingZeros(j - i + 1)
      math.max(levels(k)(i), levels(k)(j - (1 << k) + 1))

  /** The first index whose row is at least `row` (`rows.length` when none is). */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def lowerBound(rows: IArray[Int], row: Int): Int =
    var lo = 0
    var hi = rows.length
    while lo < hi do
      val mid = (lo + hi) >>> 1
      if rows(mid) < row then lo = mid + 1 else hi = mid
    lo

  /**
   * Build from the slot of every formula the pass evaluates and the stripped cells it does not
   * evaluate (a stripped cell that is also evaluated counts as evaluated).
   */
  def build(slots: Map[QualifiedRef, Int], stripped: Set[QualifiedRef]): Readiness =
    val entries =
      slots.iterator.map((q, slot) => (q, slot, false)) ++
        stripped.iterator.filterNot(slots.contains).map(q => (q, -1, true))
    val columns = entries.toVector
      .groupBy((q, _, _) => (q.sheet, q.ref.col.index0))
      .view
      .mapValues(members => column(members.map((q, slot, s) => (q.ref.row.index0, slot, s))))
      .toMap
    new Readiness(columns)

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def column(members: Vector[(Int, Int, Boolean)]): Column =
    val sorted = members.sortBy(_._1)
    val n = sorted.length
    val rows = Array.tabulate(n)(i => sorted(i)._1)
    val strippedBefore = new Array[Int](n + 1)
    var i = 0
    while i < n do
      strippedBefore(i + 1) = strippedBefore(i) + (if sorted(i)._3 then 1 else 0)
      i += 1
    val levelCount = 32 - Integer.numberOfLeadingZeros(math.max(n, 1))
    val levels = new Array[IArray[Int]](levelCount)
    var current = Array.tabulate(n)(i => sorted(i)._2)
    levels(0) = IArray.unsafeFromArray(current)
    var k = 1
    while k < levelCount do
      val half = 1 << (k - 1)
      val width = n - (1 << k) + 1
      val next = new Array[Int](math.max(width, 0))
      var j = 0
      while j < width do
        next(j) = math.max(current(j), current(j + half))
        j += 1
      levels(k) = IArray.unsafeFromArray(next)
      current = next
      k += 1
    new Column(
      IArray.unsafeFromArray(rows),
      IArray.unsafeFromArray(levels),
      IArray.unsafeFromArray(strippedBefore)
    )

/**
 * #767: one column of one sheet, rows `from..to`, as the pass's final values: a formula cell is its
 * cached value, a constant its value, an absent cell `Empty`. `r0..r1` (inside `from..to`, empty
 * when the sheet's used range misses it) are the rows held; every other row of the domain lies
 * outside the sheet's used range and is blank. Immutable once built.
 */
private[formula] final class ColumnSlice(
  val sheet: SheetName,
  val col: Int,
  val from: Int,
  val to: Int,
  r0: Int,
  values: IArray[CellValue],
  val generation: Long
):
  /** The value at `row`; `Empty` outside the held rows. Total. */
  def valueAt(row: Int): CellValue =
    val i = row - r0
    if i >= 0 && i < values.length then values(i) else CellValue.Empty

  /** Whether rows `a..b` of `col` on `sheet` lie inside this slice's domain. */
  def covers(sheet: SheetName, col: Int, a: Int, b: Int): Boolean =
    this.col == col && this.sheet == sheet && from <= a && b <= to

  /** Rows held. */
  def heldRows: Int = values.length

  /** The first row held; rows `firstHeldRow until firstHeldRow + heldRows` are held. */
  def firstHeldRow: Int = r0

  /** This slice's identity within its pass: one slice per domain and structural generation. */
  def id: ColumnSlice.Id = ColumnSlice.Id(sheet, col, from, to, generation)

private[formula] object ColumnSlice:
  final case class Id(sheet: SheetName, col: Int, from: Int, to: Int, generation: Long)

/**
 * #767 per-pass caches, owned by one [[Evaluator.AggregateMemo]] and living as long as it: the used
 * range of each sheet, static defined-name resolutions, parsed name bodies, [[ColumnSlice]]s (one
 * per column when the column is ready, else per requested box), and, when `criteriaIndex` is on,
 * each slice's equality index ([[EqIndex]]) and the candidate rows of criteria calls
 * ([[CriteriaCandidates]]). When `parseCache` is on, it also holds the formula parses of the pass's
 * graph build ([[FormulaParses]], #790).
 *
 * Name caches and formula parses are valid whenever the memo caches at all: a resolution depends
 * only on the name table and the sheet list, which every entry records and checks, and a parse only
 * on its text, which every formula parse is looked up with. Used ranges and slices depend on cell
 * contents, so they wait for [[begin]], which the passes that thread their own sheets call with the
 * [[Readiness]] of the order they are about to run. Until then (a memo a test drives by hand) every
 * request scans.
 *
 * Every structure here is safe under parallel waves: a used range or a slice is computed once
 * (single flight), so the counters it adds are the same in sequential and parallel passes.
 */
private[formula] final class PassCache(
  val limits: PassCache.Limits,
  counters: RecalcCounters,
  criteriaIndex: Boolean = true,
  parseCache: Boolean = true
):
  import PassCache.*

  private val readiness = new AtomicReference[Option[Deferred]](scala.None)

  /**
   * Arm the cell-content caches for a pass that runs in the order `ready` describes. The
   * [[Readiness]] is built on the first slice request, so a pass that never asks for a slice never
   * pays for it. Used ranges, generations and slices recorded before are dropped: they describe
   * another pass's sheets. The pass calls this before it evaluates anything.
   */
  def begin(ready: => Readiness): Unit =
    generations.clear()
    usedRanges.clear()
    slices.clear()
    covering.clear()
    demands.clear()
    cachedCells.set(0L)
    indexes.clear()
    candidateMemo.clear()
    indexBytes.set(0L)
    memoBytes.set(0L)
    readiness.set(Some(Deferred(() => ready)))

  /** Whether [[begin]] armed the cell-content caches. */
  def active: Boolean = readiness.get().isDefined

  // ----- structural generations and used ranges -----

  private val generations = new ConcurrentHashMap[SheetName, AtomicLong]()

  private def generationCell(sheet: SheetName): AtomicLong =
    generations.computeIfAbsent(sheet, _ => new AtomicLong(0L))

  /** The pass-local structural generation of `sheet`. */
  def generation(sheet: SheetName): Long = generationCell(sheet).get()

  /**
   * The pass wrote a cell of `sheet` other than a formula record's cache: its populated cells may
   * have changed, so used ranges and slices recorded before are stale.
   */
  def structureChanged(sheet: SheetName): Unit =
    generationCell(sheet).incrementAndGet()
    ()

  private val usedRanges = new ConcurrentHashMap[(SheetName, Long), UsedRange]()

  /**
   * `sheet.usedRange`, scanned once per sheet and structural generation while the pass is active
   * (every scan counts as a `UsedRangeScans`), and on every call otherwise.
   */
  def usedRange(sheet: Sheet): Option[CellRange] =
    if !active then scanUsedRange(sheet)
    else
      usedRanges
        .computeIfAbsent((sheet.name, generation(sheet.name)), _ => UsedRange(scanUsedRange(sheet)))
        .range

  private def scanUsedRange(sheet: Sheet): Option[CellRange] =
    counters.increment(RecalcCounters.Counter.UsedRangeScans)
    sheet.usedRange

  // ----- formula parses -----

  private val formulaParses = new AtomicReference[FormulaParses](FormulaParses.none)

  /**
   * #790: the parses of the graph build of the pass about to run, which evaluation looks up before
   * it parses a formula cell; ignored when the parse cache is off. [[begin]] keeps them: a parse
   * depends on its text alone.
   */
  def useParses(parses: FormulaParses): Unit =
    if parseCache then formulaParses.set(parses)

  /** The formula parses [[useParses]] supplied, or none. */
  def parses: FormulaParses = formulaParses.get()

  // ----- defined names -----

  private val nameBodies = new ConcurrentHashMap[String, Either[ParseError, TExpr[?]]]()

  /**
   * The parse of a defined name's body; parsed once per pass (each parse counts as a `NameParses`).
   */
  def parsedNameBody(formula: String)(
    parse: String => Either[ParseError, TExpr[?]]
  ): Either[ParseError, TExpr[?]] =
    nameBodies.computeIfAbsent(
      formula,
      text =>
        counters.increment(RecalcCounters.Counter.NameParses)
        parse(text)
    )

  private val staticNames = new ConcurrentHashMap[NameKey, NameEntry]()

  /**
   * A cached static range resolution of the name `name` (qualified by `scope`) read from `reader`,
   * valid for the workbook's name table and sheet list. The entry holds the target's sheet by name
   * (or as "the reader's sheet") and its range; the `Sheet` is bound again from `workbook` on every
   * hit, so no snapshot outlives the call that read it.
   */
  def cachedName(
    name: String,
    scope: Option[SheetName],
    reader: Sheet,
    workbook: Workbook
  ): Option[(Sheet, CellRange)] =
    Option(staticNames.get(NameKey(name, scope, reader.name)))
      .filter(_.validFor(workbook.metadata.definedNames, workbook))
      .flatMap(entry => bind(entry.origin, reader, workbook).map((_, entry.range)))

  /** Record a static resolution for [[cachedName]]. */
  def recordName(
    name: String,
    scope: Option[SheetName],
    reader: Sheet,
    workbook: Workbook,
    origin: NameOrigin,
    range: CellRange
  ): Unit =
    staticNames.put(
      NameKey(name, scope, reader.name),
      NameEntry(workbook.metadata.definedNames, workbook.sheets.map(_.name), origin, range)
    )
    ()

  private def bind(origin: NameOrigin, reader: Sheet, workbook: Workbook): Option[Sheet] =
    origin match
      case NameOrigin.Reader => Some(reader)
      case NameOrigin.Book(sheet) => workbook(sheet).toOption

  // ----- slices -----

  private val slices = new ConcurrentHashMap[SliceKey, SliceEntry]()
  private val covering = new ConcurrentHashMap[(SheetName, Int), Vector[ColumnSlice]]()
  private val demands = new ConcurrentHashMap[SliceKey, Int]()
  private val cachedCells = new AtomicLong(0L)

  /**
   * The slice that serves rows `a..b` of column `col` on `sheet` (the reader's threaded snapshot)
   * for a reader at `at`.
   *
   *   - An already-built slice whose domain covers `a..b` serves any request, whatever its height.
   *   - A box shorter than `max(minSliceRows, usedRows / sliceRowsDivisor)` builds nothing
   *     (`Small`).
   *   - Otherwise the whole column is tried first: its domain is every row, its held rows the
   *     sheet's used rows, and it serves every later window of the column. It needs [[Readiness]]
   *     over the used rows.
   *   - When the whole column is not ready, the box `a..b` itself is built, but only on its
   *     `boxDemand`-th request in the generation (the second, by default): a window asked for once
   *     (a running `COUNTIF($A$2:A9,…)`, a sliding window) would copy as many cells as its scan
   *     reads and hold them for the pass, so the requests before scan.
   *
   * `Refused` (readiness, an uncached formula in the box, the cell budget, or a box's request below
   * `boxDemand` after the whole column was refused) is a fallback the caller counts.
   */
  def slice(sheet: Sheet, col: Int, a: Int, b: Int, at: EvalPosition): SliceAnswer =
    readiness.get() match
      case scala.None => SliceAnswer.Ineligible
      case Some(_) if !at.isDefined || a > b => SliceAnswer.Ineligible
      case Some(deferred) =>
        val name = sheet.name
        val gen = generation(name)
        val existing = Option(covering.get((name, col))).flatMap(
          _.find(s => s.generation == gen && s.covers(name, col, a, b))
        )
        existing match
          case Some(found) => SliceAnswer.Use(found)
          case scala.None =>
            val used = usedRange(sheet)
            val usedRows = used.fold(0)(_.height)
            val threshold = math.max(limits.minSliceRows, usedRows / limits.sliceRowsDivisor)
            if b - a + 1 < threshold then SliceAnswer.Small
            else
              val wholeColumn = buildBox(sheet, col, 0, Row.MaxIndex0, used, gen, deferred, at)
              wholeColumn match
                case Some(built) => SliceAnswer.Use(built)
                case scala.None =>
                  val box = SliceKey(name, col, a, b, gen)
                  if demands.merge(box, 1, _ + _) < limits.boxDemand then SliceAnswer.Refused
                  else
                    buildBox(sheet, col, a, b, used, gen, deferred, at) match
                      case Some(built) => SliceAnswer.Use(built)
                      case scala.None => SliceAnswer.Refused

  /**
   * The slice of rows `a..b` of `col`, held over the used rows inside them, built once per
   * generation. `None` when [[Readiness]] refuses it at `at` (not remembered: a later slot may
   * pass), or when the copy meets an uncached formula or the cell budget is spent (remembered for
   * the generation: once readiness has passed, every formula of the box the pass evaluates is done
   * and every other cell keeps its value, so a later copy would meet the same cells).
   */
  private def buildBox(
    sheet: Sheet,
    col: Int,
    a: Int,
    b: Int,
    used: Option[CellRange],
    gen: Long,
    deferred: Deferred,
    at: EvalPosition
  ): Option[ColumnSlice] =
    val name = sheet.name
    val r0 = used.fold(a)(u => math.max(a, u.rowStart.index0))
    val r1 = used.fold(a - 1)(u => math.min(b, u.rowEnd.index0))
    slices
      .computeIfAbsent(SliceKey(name, col, a, b, gen), _ => SliceEntry())
      .getOrBuild(deferred.value.allows(name, col, r0, r1, at))(
        build(sheet, col, a, b, r0, r1, gen)
      )

  /**
   * Copy rows `r0..r1` of `col`, or nothing when one holds an uncached formula (a value the pass
   * has not computed; the reader scans, which evaluates it) or the cell budget is spent.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def build(
    sheet: Sheet,
    col: Int,
    a: Int,
    b: Int,
    r0: Int,
    r1: Int,
    gen: Long
  ): Option[ColumnSlice] =
    val n = math.max(0, r1 - r0 + 1)
    if cachedCells.addAndGet(n.toLong) > limits.maxCachedCells then
      cachedCells.addAndGet(-n.toLong)
      scala.None
    else
      val values = new Array[CellValue](n)
      var i = 0
      var complete = true
      while complete && i < n do
        sheet.cells.get(ARef.from0(col, r0 + i)).map(_.value) match
          case scala.None => values(i) = CellValue.Empty
          case Some(CellValue.Formula(_, scala.None, _)) => complete = false
          case Some(CellValue.Formula(_, Some(cached), _)) => values(i) = cached
          case Some(value) => values(i) = value
        i += 1
      if !complete then
        cachedCells.addAndGet(-n.toLong)
        scala.None
      else
        val slice =
          new ColumnSlice(sheet.name, col, a, b, r0, IArray.unsafeFromArray(values), gen)
        covering.merge((sheet.name, col), Vector(slice), (known, more) => known ++ more)
        counters.increment(RecalcCounters.Counter.SlicesBuilt)
        Some(slice)

  // ----- equality indexes and candidates -----

  private val indexes = new ConcurrentHashMap[ColumnSlice.Id, IndexEntry]()
  private val candidateMemo =
    new ConcurrentHashMap[CriteriaCandidates.MemoKey, CandidatesEntry]()
  private val indexBytes = new AtomicLong(0L)
  private val memoBytes = new AtomicLong(0L)

  /** Whether criteria calls may use equality indexes. */
  def indexing: Boolean = criteriaIndex && active

  /**
   * The equality index of `slice`, built once per slice (single flight). `None` when indexing is
   * off or the index would take the pass past `maxIndexBytes`; the call then tests every row.
   */
  def eqIndex(slice: ColumnSlice): Option[EqIndex] =
    if !indexing then scala.None
    else indexes.computeIfAbsent(slice.id, _ => new IndexEntry(slice)).value

  /**
   * The candidate offsets of a criteria call whose indexed pairs are `pairs` (each: its slice, its
   * lookup keys and its range's first row) over `n` rows, computed once per pass for equal pairs
   * and `n` while the budget lasts; `None` when an index is missing or every offset must be tested
   * ([[CriteriaCandidates.compute]]). Memoised lists hold at most half of `maxIndexBytes`, so they
   * never use the budget up before the indexes they are computed from.
   */
  def candidates(
    pairs: List[(ColumnSlice, List[MatchKey], Int)],
    n: Int
  ): Option[ArraySeq.ofInt] =
    val indexed = pairs.map((slice, keys, offset) => eqIndex(slice).map((slice, keys, offset, _)))
    if indexed.exists(_.isEmpty) then scala.None
    else
      val resolved = indexed.flatten
      val key = CriteriaCandidates.MemoKey(
        n,
        resolved.map((slice, keys, offset, _) => CriteriaCandidates.PairKey(slice.id, keys, offset))
      )
      def compute: Option[ArraySeq.ofInt] =
        CriteriaCandidates.compute(
          resolved.map((_, keys, offset, index) => CriteriaCandidates.Pair(index, keys, offset)),
          n
        )
      val entry = candidateMemo.computeIfAbsent(key, _ => new CandidatesEntry(() => compute))
      val result = entry.value
      if !entry.charged(indexBytes, memoBytes, limits.maxIndexBytes) then
        candidateMemo.remove(key, entry)
      result

  /** A slice's index, built on first use and charged against the pass's index budget. */
  private final class IndexEntry(slice: ColumnSlice):
    lazy val value: Option[EqIndex] =
      val index = EqIndex.build(slice)
      if indexBytes.addAndGet(index.bytes) > limits.maxIndexBytes then
        indexBytes.addAndGet(-index.bytes)
        scala.None
      else Some(index)

private[formula] object PassCache:

  /**
   * Bounds on what one pass holds. `maxCachedCells` caps the cells all slices hold together; a box
   * is sliced only when it spans at least `max(minSliceRows, usedRows / sliceRowsDivisor)` rows
   * (smaller windows scan, or read a slice that covers them); a box whose column is not ready is
   * sliced on its `boxDemand`-th request. `maxIndexBytes` caps what equality indexes and memoised
   * candidate lists hold together, the lists at most half of it. Past a cap a call scans; results
   * are unaffected.
   */
  final case class Limits(
    maxCachedCells: Long = 64L * 1024 * 1024,
    minSliceRows: Int = 256,
    sliceRowsDivisor: Int = 16,
    boxDemand: Int = 2,
    maxIndexBytes: Long = 256L * 1024 * 1024
  )

  val DefaultLimits: Limits = Limits()

  /** What [[PassCache.slice]] answers. */
  enum SliceAnswer:
    case Use(slice: ColumnSlice)

    /**
     * Readiness, an uncached formula, the cell budget or the repeat-demand rule refused the build:
     * the call scans and counts a fallback.
     */
    case Refused

    /** The box is below the size threshold: the call scans. */
    case Small

    /** The pass has no readiness, or the reader is not at its own slot: the call scans. */
    case Ineligible

  /** Where a static name's target sheet comes from when it is bound again. */
  enum NameOrigin derives CanEqual:
    /** The sheet the name is read from (a workbook-scoped name whose body names no sheet). */
    case Reader

    /** The workbook's sheet of this name. */
    case Book(sheet: SheetName)

  private final case class UsedRange(range: Option[CellRange])

  /** A [[Readiness]] built once, on first use; a lazy val is safe to share between waves. */
  private final class Deferred(build: () => Readiness):
    lazy val value: Readiness = build()

  private final case class NameKey(name: String, scope: Option[SheetName], reader: SheetName)

  private final case class NameEntry(
    names: Vector[DefinedName],
    sheets: Vector[SheetName],
    origin: NameOrigin,
    range: CellRange
  ):
    def validFor(otherNames: Vector[DefinedName], workbook: Workbook): Boolean =
      (names eq otherNames) && sheets.corresponds(workbook.sheets)(_ == _.name)

  /**
   * One memoised candidate list, computed once (single flight) and charged against the index budget
   * by the first caller that sees it computed.
   */
  private final class CandidatesEntry(compute: () => Option[ArraySeq.ofInt]):
    lazy val value: Option[ArraySeq.ofInt] = compute()
    private val charge = new java.util.concurrent.atomic.AtomicInteger(0)

    /**
     * Whether the entry fits the budget (charging it once): `used` (indexes and lists together) at
     * most `max`, and `memo` (lists alone) at most half of it. False evicts it.
     */
    def charged(used: AtomicLong, memo: AtomicLong, max: Long): Boolean =
      if charge.compareAndSet(0, 1) then
        val bytes = CriteriaCandidates.bytes(value)
        if memo.addAndGet(bytes) > max / 2 then
          memo.addAndGet(-bytes)
          charge.set(2)
          false
        else if used.addAndGet(bytes) > max then
          used.addAndGet(-bytes)
          memo.addAndGet(-bytes)
          charge.set(2)
          false
        else true
      else charge.get() == 1

  private final case class SliceKey(sheet: SheetName, col: Int, a: Int, b: Int, generation: Long)

  /**
   * Single-flight build of one slice. A readiness refusal is not remembered (the next request
   * retries); a build that failed after readiness passed is (see `buildBox`).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private final class SliceEntry:
    private var built: Option[ColumnSlice] = scala.None
    private var failed: Boolean = false

    def getOrBuild(ready: => Boolean)(build: => Option[ColumnSlice]): Option[ColumnSlice] =
      synchronized {
        built match
          case found @ Some(_) => found
          case scala.None if failed || !ready => scala.None
          case scala.None =>
            val result = build
            built = result
            failed = result.isEmpty
            result
      }
