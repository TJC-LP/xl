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

  /** Whether to skip non-numeric cells (true) or include them in count (false) */
  def skipNonNumeric: Boolean = true

object Aggregator:
  /** Registry of all aggregators by name (uppercase) */
  private lazy val registry: Map[String, Aggregator[?]] = Map(
    "SUM" -> sumAggregator,
    "COUNT" -> countAggregator,
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
      if acc._2 == 0 then BigDecimal(0) // Could return #DIV/0! error
      else acc._1 / acc._2
