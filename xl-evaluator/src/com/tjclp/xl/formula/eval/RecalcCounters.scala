package com.tjclp.xl.formula.eval

import java.util.concurrent.atomic.LongAdder

/**
 * Work counters and per-phase wall time for one recalculation pass (#767).
 *
 * A pass reads its counters from its [[Evaluator.AggregateMemo]]; every public entry point builds
 * that memo with [[RecalcCounters.disabled]], whose methods do nothing and whose `time` runs the
 * body without reading the clock. Only the internal seams `WorkbookEvaluator.recalculateCounted`
 * and `recalculateUncachedCounted` take an enabled instance, so results never depend on them.
 *
 * Counts are exact functions of the workbook and the algorithm. A hot loop accumulates into a local
 * [[RecalcCounters.Tally]] and adds it once per call, so parallel waves touch the shared
 * `LongAdder`s once per function call, not once per row.
 *
 * Phase times are inclusive and may nest: `Parse` runs inside `GraphBuild` and `Eval`, `Criteria`
 * and `SheetPut` inside `Eval`. Under parallel waves `Criteria` and `Parse` sum the time of every
 * worker thread, so they can exceed the wall time of `Eval`.
 */
private[formula] final class RecalcCounters private (val enabled: Boolean):
  private val counts: IArray[LongAdder] =
    IArray.fill(RecalcCounters.Counter.values.length)(new LongAdder)
  private val nanos: IArray[LongAdder] =
    IArray.fill(RecalcCounters.Phase.values.length)(new LongAdder)

  /** Add `n` to `counter`; a no-op when disabled. */
  def add(counter: RecalcCounters.Counter, n: Long): Unit =
    if enabled && n != 0L then counts(counter.ordinal).add(n)

  /** Add one to `counter`; a no-op when disabled. */
  def increment(counter: RecalcCounters.Counter): Unit =
    if enabled then counts(counter.ordinal).increment()

  /** Add elapsed nanoseconds to `phase`; a no-op when disabled. */
  def addNanos(phase: RecalcCounters.Phase, elapsed: Long): Unit =
    if enabled then nanos(phase.ordinal).add(elapsed)

  /**
   * Run `body` and add its wall time to `phase`. Disabled counters run the body without reading the
   * clock.
   *
   * `body` appears once in the expansion: an inline parameter is copied at every use, and a second
   * use would compile each timed block (some are whole recalculation phases) twice.
   */
  inline def time[A](phase: RecalcCounters.Phase)(inline body: A): A =
    val start = if enabled then System.nanoTime() else 0L
    try body
    finally if enabled then addNanos(phase, System.nanoTime() - start)

  /** A fresh per-call tally for `counter`, or the shared no-op tally when disabled. */
  def tally(counter: RecalcCounters.Counter): RecalcCounters.Tally =
    if enabled then new RecalcCounters.LiveTally(this, counter) else RecalcCounters.NoTally

  /** The current totals; an immutable copy. */
  def snapshot: RecalcCounters.Snapshot =
    RecalcCounters.Snapshot(
      RecalcCounters.Counter.values.iterator.map(c => c -> counts(c.ordinal).sum()).toMap,
      RecalcCounters.Phase.values.iterator.map(p => p -> nanos(p.ordinal).sum()).toMap
    )

private[formula] object RecalcCounters:

  /**
   * What a pass counts. The legacy graph (`RecalcEngine.Legacy`) has no virtual or range nodes and
   * does not count `GraphWork`; the per-pass slice cache and criteria index do not exist yet, so
   * `CriteriaIndexed`, `CriteriaWindowOnly`, `CriteriaNotIndexable`, `SlicesBuilt` and
   * `ScanFallbacks` stay 0 until the code that owns them lands; every criteria call is a
   * `CriteriaScan` today.
   */
  enum Counter derives CanEqual:
    /** Formula nodes of the recalculation graph. */
    case GraphNodes

    /** Formula-to-formula edges of the recalculation graph. */
    case GraphEdges

    /** Segment nodes of the compact graph. */
    case VirtualNodes

    /** Distinct (sheet, range) nodes of the compact graph. */
    case RangeNodes

    /**
     * Edges walked by the compact graph's passes: Tarjan, Kahn's in-degree count and order, and the
     * reachability walks (blocked cells, the dynamic bucket, invalidation). 0 for the legacy graph.
     */
    case GraphWork

    /** Criteria calls answered through the equality index. */
    case CriteriaIndexed

    /** Criteria calls whose criteria are all numeric windows (no equality pair). */
    case CriteriaWindowOnly

    /** Criteria calls with no indexable criterion. */
    case CriteriaNotIndexable

    /** Criteria calls answered by walking every row of their ranges. */
    case CriteriaScan

    /** `CriteriaMatcher.matches` calls made by the criteria functions' row loops. */
    case MatcherCalls

    /** Per-pass column slices built. */
    case SlicesBuilt

    /** Criteria calls that wanted a slice and fell back to the scan. */
    case ScanFallbacks

    /** `Sheet.usedRange` scans made to bound whole-column or whole-row ranges. */
    case UsedRangeScans

    /** Defined-name bodies parsed during evaluation. */
    case NameParses

  /** Where a pass spends its time. `Write` is measured by the caller that writes the file. */
  enum Phase derives CanEqual:
    case Parse, GraphBuild, Tarjan, Kahn, Dynamic, Eval, Criteria, SheetPut, Finalize, Write

  /** The instance every public entry point uses: counts nothing, never reads the clock. */
  val disabled: RecalcCounters = new RecalcCounters(enabled = false)

  /** A fresh enabled instance for one measured pass. */
  def create(): RecalcCounters = new RecalcCounters(enabled = true)

  /** A per-call accumulator, added to its counter once with [[Tally.flush]]. */
  sealed trait Tally:
    def increment(): Unit
    def flush(): Unit

  private object NoTally extends Tally:
    def increment(): Unit = ()
    def flush(): Unit = ()

  /** Confined to the thread running one function call; never shared. */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private final class LiveTally(owner: RecalcCounters, counter: Counter) extends Tally:
    private var n = 0L
    def increment(): Unit = n += 1L
    def flush(): Unit =
      owner.add(counter, n)
      n = 0L

  /** Immutable totals of one pass. */
  final case class Snapshot(counters: Map[Counter, Long], phaseNanos: Map[Phase, Long]):
    def apply(counter: Counter): Long = counters.getOrElse(counter, 0L)
    def nanos(phase: Phase): Long = phaseNanos.getOrElse(phase, 0L)
    def seconds(phase: Phase): Double = nanos(phase).toDouble / 1e9

    /** One `key=value` line per counter, then one per phase in seconds, in declaration order. */
    def render: String =
      val countLines = Counter.values.iterator.map(c => s"${c.toString}=${apply(c)}")
      val phaseLines = Phase.values.iterator.map(p => f"${p.toString}Seconds=${seconds(p)}%.3f")
      (countLines ++ phaseLines).mkString("\n")
