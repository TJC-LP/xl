package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.Clock

import com.tjclp.xl.cells.CellError

/**
 * Typeclass for aggregate functions over cell ranges.
 *
 * An Aggregator encapsulates:
 *   - Function metadata (name)
 *   - Accumulation logic (fold over cell values)
 *   - Type transformation (input cell value type -> output result type)
 *
 * Laws:
 *   - Associativity: combine(combine(a, b), c) == combine(a, combine(b, c))
 *   - Identity: combine(empty, a) == a == combine(a, empty)
 *
 * @tparam A
 *   The accumulator type (e.g., BigDecimal for SUM, Long for COUNT, (BigDecimal, Long) for AVERAGE)
 */
trait Aggregator[A]:
  /** Function name as used in Excel (uppercase) */
  def name: String

  /** Initial accumulator value */
  def empty: A

  /** Combine accumulator with a numeric cell value */
  def combine(acc: A, value: BigDecimal): A

  /**
   * #750: combine `value` into the accumulator `n` times, equal to `n` successive [[combine]]
   * calls. Every built-in aggregator overrides it in closed form, so a sparse array's fill (a
   * million blank or equal cells) folds in one step; this default repeats [[combine]].
   */
  def combineN(acc: A, value: BigDecimal, n: Long): A =
    @scala.annotation.tailrec
    def loop(a: A, left: Long): A = if left <= 0L then a else loop(combine(a, value), left - 1L)
    loop(acc, n)

  /** Finalize the result after processing all cells */
  def finalize(acc: A): BigDecimal

  /**
   * Finalize with error handling for functions that can fail (like AVERAGE on empty range).
   *
   * By default, wraps `finalize` in Right. Override for error-returning functions.
   */
  def finalizeWithError(acc: A): Either[EvalError, BigDecimal] =
    Right(finalize(acc))

  /** Whether to skip non-numeric cells (true) or include them in count (false) */
  def skipNonNumeric: Boolean = true

  /**
   * Whether this aggregator counts non-empty cells (like COUNTA).
   *
   * When true, the evaluator will count any non-empty cell rather than calling `combine` with
   * numeric values. This is used for COUNTA which counts all non-empty cells regardless of type.
   */
  def countsNonEmpty: Boolean = false

  /**
   * Whether this aggregator counts empty cells (like COUNTBLANK).
   *
   * When true, the evaluator will count cells that are empty rather than calling `combine` with
   * values. This is used for COUNTBLANK which counts truly empty cells.
   */
  def countsEmpty: Boolean = false

  /**
   * GH-337/GH-344: whether a carried error ELEMENT consumed by this aggregate — from an evaluated
   * expression array (`SUM((A1:A2)*1)`) or a raw-range cell (`SUM(A1:A2)`, GH-344 item 6's
   * resolve-then-police) — propagates as the element's Excel error VALUE.
   *
   * Excel propagates errors through nearly every aggregate; COUNT alone genuinely ignores them (it
   * counts numbers, and an error is not a number). COUNTA/COUNTBLANK are governed by their counts*
   * modes instead (an error element is non-empty).
   */
  def propagatesErrors: Boolean = true

object Aggregator:
  /**
   * GH-630: what an aggregate does with an ARGUMENT that is itself an Excel error value — the
   * literal `#REF!`, `1/0`, a range slot that resolved to an error, a defined name bound to one.
   * COUNTA counts it (`Some(acc + 1)`: an error is not empty), COUNT skips it (`Some(acc)`: an
   * error is not a number), and every other aggregate — COUNTBLANK included, whose argument must be
   * a range — propagates it (`None`). Shared by the `Call` form (`FunctionSpecsAggregate`) and the
   * typed `TExpr.Aggregate` node (`Evaluator`), so `Aggregate ≡ Call` holds for the error form too.
   */
  def onErrorArgument[A](agg: Aggregator[A], acc: A): Option[A] =
    if agg.countsNonEmpty then Some(agg.combine(acc, BigDecimal(1)))
    else if !agg.propagatesErrors && !agg.countsEmpty then Some(acc)
    else None

  /** Registry of all aggregators by name (uppercase) */
  private lazy val registry: Map[String, Aggregator[?]] = Map(
    "SUM" -> sumAggregator,
    "COUNT" -> countAggregator,
    "COUNTA" -> countaAggregator,
    "COUNTBLANK" -> countblankAggregator,
    "MIN" -> minAggregator,
    "MAX" -> maxAggregator,
    "AVERAGE" -> averageAggregator,
    "MEDIAN" -> medianAggregator,
    "STDEV" -> stdevAggregator,
    "STDEVP" -> stdevpAggregator,
    "VAR" -> varAggregator,
    "VARP" -> varpAggregator
  )

  /** Look up an aggregator by name (case-insensitive) */
  def lookup(name: String): Option[Aggregator[?]] =
    registry.get(name.toUpperCase)

  /** All registered aggregator names */
  def all: Seq[String] = registry.keys.toSeq.sorted

  /** Check if a name is a registered aggregator */
  def isAggregator(name: String): Boolean =
    registry.contains(name.toUpperCase)

  // ==================== Given Instances ====================

  /** SUM: Add all numeric values in a range */
  given sumAggregator: Aggregator[BigDecimal] with
    def name = "SUM"
    def empty = BigDecimal(0)
    def combine(acc: BigDecimal, value: BigDecimal) = acc + value
    override def combineN(acc: BigDecimal, value: BigDecimal, n: Long) =
      if n <= 0L then acc else acc + value * BigDecimal(n)
    def finalize(acc: BigDecimal) = acc

  /** COUNT: Count numeric values in a range */
  given countAggregator: Aggregator[Long] with
    def name = "COUNT"
    def empty = 0L
    def combine(acc: Long, value: BigDecimal) = acc + 1L
    override def combineN(acc: Long, value: BigDecimal, n: Long) = acc + math.max(n, 0L)
    def finalize(acc: Long) = BigDecimal(acc)
    // GH-337: Excel's COUNT counts numbers only — error elements are skipped, not propagated
    override def propagatesErrors: Boolean = false

  /** COUNTA: Count non-empty cells in a range (not just numeric) */
  given countaAggregator: Aggregator[Long] with
    def name = "COUNTA"
    def empty = 0L
    def combine(acc: Long, value: BigDecimal) = acc + 1L
    override def combineN(acc: Long, value: BigDecimal, n: Long) = acc + math.max(n, 0L)
    def finalize(acc: Long) = BigDecimal(acc)
    override def countsNonEmpty: Boolean = true

  /** COUNTBLANK: Count empty cells in a range */
  given countblankAggregator: Aggregator[Long] with
    def name = "COUNTBLANK"
    def empty = 0L
    def combine(acc: Long, value: BigDecimal) = acc + 1L
    override def combineN(acc: Long, value: BigDecimal, n: Long) = acc + math.max(n, 0L)
    def finalize(acc: Long) = BigDecimal(acc)
    override def countsEmpty: Boolean = true

  /** MIN: Find the minimum numeric value in a range */
  given minAggregator: Aggregator[Option[BigDecimal]] with
    def name = "MIN"
    def empty = None
    def combine(acc: Option[BigDecimal], value: BigDecimal) =
      Some(acc.fold(value)(_.min(value)))
    override def combineN(acc: Option[BigDecimal], value: BigDecimal, n: Long) =
      if n <= 0L then acc else combine(acc, value)
    def finalize(acc: Option[BigDecimal]) =
      acc.getOrElse(BigDecimal(0)) // Excel returns 0 for empty range

  /** MAX: Find the maximum numeric value in a range */
  given maxAggregator: Aggregator[Option[BigDecimal]] with
    def name = "MAX"
    def empty = None
    def combine(acc: Option[BigDecimal], value: BigDecimal) =
      Some(acc.fold(value)(_.max(value)))
    override def combineN(acc: Option[BigDecimal], value: BigDecimal, n: Long) =
      if n <= 0L then acc else combine(acc, value)
    def finalize(acc: Option[BigDecimal]) =
      acc.getOrElse(BigDecimal(0)) // Excel returns 0 for empty range

  /** AVERAGE: Calculate the arithmetic mean of numeric values in a range */
  given averageAggregator: Aggregator[(BigDecimal, Long)] with
    def name = "AVERAGE"
    def empty = (BigDecimal(0), 0L)
    def combine(acc: (BigDecimal, Long), value: BigDecimal) =
      (acc._1 + value, acc._2 + 1L)
    override def combineN(acc: (BigDecimal, Long), value: BigDecimal, n: Long) =
      if n <= 0L then acc else (acc._1 + value * BigDecimal(n), acc._2 + n)
    def finalize(acc: (BigDecimal, Long)) =
      if acc._2 == 0L then BigDecimal(0) // Fallback; use finalizeWithError for proper error
      else acc._1 / BigDecimal(acc._2)
    override def finalizeWithError(acc: (BigDecimal, Long)) =
      if acc._2 == 0L then Left(EvalError.DivByZero("AVERAGE(empty range)", "count=0"))
      else Right(acc._1 / BigDecimal(acc._2))

  /**
   * MEDIAN: Find the middle value(s) in a range. The accumulator keeps each value with how many
   * times it was combined (#750: a sparse array's fill is one entry, however many cells it covers).
   */
  given medianAggregator: Aggregator[Vector[(BigDecimal, Long)]] with
    def name = "MEDIAN"
    def empty = Vector.empty
    def combine(acc: Vector[(BigDecimal, Long)], value: BigDecimal) = acc :+ ((value, 1L))
    override def combineN(acc: Vector[(BigDecimal, Long)], value: BigDecimal, n: Long) =
      if n <= 0L then acc else acc :+ ((value, n))
    def finalize(acc: Vector[(BigDecimal, Long)]) =
      val total = acc.iterator.map(_._2).sum
      if total == 0L then BigDecimal(0) // Fallback; use finalizeWithError for proper error
      else
        val sorted = acc.sortBy(_._1)
        // the value at 0-based position `i` of the expanded, sorted values
        def at(i: Long): BigDecimal =
          sorted.iterator
            .scanLeft((BigDecimal(0), 0L)) { case ((_, seen), (v, k)) => (v, seen + k) }
            .drop(1)
            .find(_._2 > i)
            .map(_._1)
            .getOrElse(BigDecimal(0))
        val mid = total / 2
        if total % 2 == 0 then (at(mid - 1) + at(mid)) / 2
        else at(mid)
    override def finalizeWithError(acc: Vector[(BigDecimal, Long)]) =
      if acc.isEmpty then Left(EvalError.EvalFailed("MEDIAN requires at least 1 value", None))
      else Right(finalize(acc))

  /**
   * Welford accumulator for numerically stable variance computation. Stores (count, mean, m2) where
   * m2 = sum of squared deviations from mean. Reference: Welford, B. P. (1962) and Knuth TAOCP Vol
   * 2, Section 4.2.2
   */
  private type WelfordAcc = (Long, BigDecimal, BigDecimal)

  /**
   * A variance statistic over fewer values than its divisor needs — STDEV and VAR with fewer than
   * two, STDEVP and VARP with none — is Excel's #DIV/0! error VALUE (IFERROR catches it, recalc
   * caches it), never a host failure. A plain cell's `STDEV(A1:A10*1)` reaches this with its one
   * intersected value.
   */
  private def tooFewValues(reason: String, count: Long): Either[EvalError, BigDecimal] =
    Left(EvalError.ErrorValue(CellError.Div0, Some(s"$reason (count=$count)")))

  /** Welford's online algorithm update step */
  private def welfordCombine(acc: WelfordAcc, value: BigDecimal): WelfordAcc =
    val (count, mean, m2) = acc
    val newCount = count + 1L
    val delta = value - mean
    val newMean = mean + delta / BigDecimal(newCount)
    val delta2 = value - newMean
    val newM2 = m2 + delta * delta2
    (newCount, newMean, newM2)

  /**
   * #750: `n` copies of `value` merged into the accumulator in one step (Chan et al.'s pairwise
   * update with a block of `n` equal values, whose own m2 is 0).
   */
  private def welfordCombineN(acc: WelfordAcc, value: BigDecimal, n: Long): WelfordAcc =
    if n <= 0L then acc
    else if n == 1L then welfordCombine(acc, value)
    else
      val (count, mean, m2) = acc
      val newCount = count + n
      val delta = value - mean
      val weight = BigDecimal(n) / BigDecimal(newCount)
      (newCount, mean + delta * weight, m2 + delta * delta * BigDecimal(count) * weight)

  /** STDEV: Sample standard deviation (divides by n-1) */
  given stdevAggregator: Aggregator[WelfordAcc] with
    def name = "STDEV"
    def empty = (0L, BigDecimal(0), BigDecimal(0))
    def combine(acc: WelfordAcc, value: BigDecimal) = welfordCombine(acc, value)
    override def combineN(acc: WelfordAcc, value: BigDecimal, n: Long) =
      welfordCombineN(acc, value, n)
    def finalize(acc: WelfordAcc) =
      val (n, _, m2) = acc
      if n < 2 then BigDecimal(0) // Fallback; use finalizeWithError for proper error
      else BigDecimal(math.sqrt((m2 / BigDecimal(n - 1L)).toDouble))
    override def finalizeWithError(acc: WelfordAcc) =
      if acc._1 < 2 then tooFewValues("STDEV requires at least 2 values", acc._1)
      else Right(finalize(acc))

  /** STDEVP: Population standard deviation (divides by n) */
  given stdevpAggregator: Aggregator[WelfordAcc] with
    def name = "STDEVP"
    def empty = (0L, BigDecimal(0), BigDecimal(0))
    def combine(acc: WelfordAcc, value: BigDecimal) = welfordCombine(acc, value)
    override def combineN(acc: WelfordAcc, value: BigDecimal, n: Long) =
      welfordCombineN(acc, value, n)
    def finalize(acc: WelfordAcc) =
      val (n, _, m2) = acc
      if n < 1 then BigDecimal(0) // Fallback; use finalizeWithError for proper error
      else BigDecimal(math.sqrt((m2 / BigDecimal(n)).toDouble))
    override def finalizeWithError(acc: WelfordAcc) =
      if acc._1 < 1 then tooFewValues("STDEVP requires at least 1 value", acc._1)
      else Right(finalize(acc))

  /** VAR: Sample variance (divides by n-1) */
  given varAggregator: Aggregator[WelfordAcc] with
    def name = "VAR"
    def empty = (0L, BigDecimal(0), BigDecimal(0))
    def combine(acc: WelfordAcc, value: BigDecimal) = welfordCombine(acc, value)
    override def combineN(acc: WelfordAcc, value: BigDecimal, n: Long) =
      welfordCombineN(acc, value, n)
    def finalize(acc: WelfordAcc) =
      val (n, _, m2) = acc
      if n < 2 then BigDecimal(0) // Fallback; use finalizeWithError for proper error
      else m2 / BigDecimal(n - 1L)
    override def finalizeWithError(acc: WelfordAcc) =
      if acc._1 < 2 then tooFewValues("VAR requires at least 2 values", acc._1)
      else Right(finalize(acc))

  /** VARP: Population variance (divides by n) */
  given varpAggregator: Aggregator[WelfordAcc] with
    def name = "VARP"
    def empty = (0L, BigDecimal(0), BigDecimal(0))
    def combine(acc: WelfordAcc, value: BigDecimal) = welfordCombine(acc, value)
    override def combineN(acc: WelfordAcc, value: BigDecimal, n: Long) =
      welfordCombineN(acc, value, n)
    def finalize(acc: WelfordAcc) =
      val (n, _, m2) = acc
      if n < 1 then BigDecimal(0) // Fallback; use finalizeWithError for proper error
      else m2 / BigDecimal(n)
    override def finalizeWithError(acc: WelfordAcc) =
      if acc._1 < 1 then tooFewValues("VARP requires at least 1 value", acc._1)
      else Right(finalize(acc))
