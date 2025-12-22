package com.tjclp.xl.formula

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
 *   The accumulator type (e.g., BigDecimal for SUM, Int for COUNT, (BigDecimal, Int) for AVERAGE)
 */
trait Aggregator[A]:
  /** Function name as used in Excel (uppercase) */
  def name: String

  /** Initial accumulator value */
  def empty: A

  /** Combine accumulator with a numeric cell value */
  def combine(acc: A, value: BigDecimal): A

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

object Aggregator:
  /** Registry of all aggregators by name (uppercase) */
  private lazy val registry: Map[String, Aggregator[?]] = Map(
    "SUM" -> sumAggregator,
    "COUNT" -> countAggregator,
    "COUNTA" -> countaAggregator,
    "COUNTBLANK" -> countblankAggregator,
    "MIN" -> minAggregator,
    "MAX" -> maxAggregator,
    "AVERAGE" -> averageAggregator
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
    def finalize(acc: BigDecimal) = acc

  /** COUNT: Count numeric values in a range */
  given countAggregator: Aggregator[Int] with
    def name = "COUNT"
    def empty = 0
    def combine(acc: Int, value: BigDecimal) = acc + 1
    def finalize(acc: Int) = BigDecimal(acc)

  /** COUNTA: Count non-empty cells in a range (not just numeric) */
  given countaAggregator: Aggregator[Int] with
    def name = "COUNTA"
    def empty = 0
    def combine(acc: Int, value: BigDecimal) = acc + 1
    def finalize(acc: Int) = BigDecimal(acc)
    override def countsNonEmpty: Boolean = true

  /** COUNTBLANK: Count empty cells in a range */
  given countblankAggregator: Aggregator[Int] with
    def name = "COUNTBLANK"
    def empty = 0
    def combine(acc: Int, value: BigDecimal) = acc + 1
    def finalize(acc: Int) = BigDecimal(acc)
    override def countsEmpty: Boolean = true

  /** MIN: Find the minimum numeric value in a range */
  given minAggregator: Aggregator[Option[BigDecimal]] with
    def name = "MIN"
    def empty = None
    def combine(acc: Option[BigDecimal], value: BigDecimal) =
      Some(acc.fold(value)(_.min(value)))
    def finalize(acc: Option[BigDecimal]) =
      acc.getOrElse(BigDecimal(0)) // Excel returns 0 for empty range

  /** MAX: Find the maximum numeric value in a range */
  given maxAggregator: Aggregator[Option[BigDecimal]] with
    def name = "MAX"
    def empty = None
    def combine(acc: Option[BigDecimal], value: BigDecimal) =
      Some(acc.fold(value)(_.max(value)))
    def finalize(acc: Option[BigDecimal]) =
      acc.getOrElse(BigDecimal(0)) // Excel returns 0 for empty range

  /** AVERAGE: Calculate the arithmetic mean of numeric values in a range */
  given averageAggregator: Aggregator[(BigDecimal, Int)] with
    def name = "AVERAGE"
    def empty = (BigDecimal(0), 0)
    def combine(acc: (BigDecimal, Int), value: BigDecimal) =
      (acc._1 + value, acc._2 + 1)
    def finalize(acc: (BigDecimal, Int)) =
      if acc._2 == 0 then BigDecimal(0) // Fallback; use finalizeWithError for proper error
      else acc._1 / acc._2
    override def finalizeWithError(acc: (BigDecimal, Int)) =
      if acc._2 == 0 then Left(EvalError.DivByZero("AVERAGE(empty range)", "count=0"))
      else Right(acc._1 / acc._2)
