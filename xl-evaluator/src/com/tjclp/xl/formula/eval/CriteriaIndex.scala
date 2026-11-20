package com.tjclp.xl.formula.eval

import scala.collection.immutable.ArraySeq

import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.ExprValue

/**
 * #767: an equality key of a criteria-index bucket. A cell value has the keys under which an
 * `Exact` criterion can match it ([[CriteriaKeys.cellKeys]]); a criterion looks up the keys of the
 * values it can match ([[CriteriaKeys.lookup]]).
 */
private[formula] enum MatchKey derives CanEqual:
  /** Text, folded per UTF-16 unit ([[CriteriaKeys.fold]]). */
  case Text(folded: String)

  /** A number without trailing zeros, so value-equal numbers share a key. */
  case Num(stripped: java.math.BigDecimal)

  case Bool(value: Boolean)

  case Err(error: CellError)

/**
 * #767: the keys of the criteria index, derived from `CriteriaMatcher.matchesExact`.
 *
 * The index is a filter, never a judge: it returns a superset of the rows that can match, and the
 * caller tests every candidate with the unchanged `CriteriaMatcher.matches`. What makes the
 * superset complete:
 *
 * for every cell value `v` and every criterion `q` that [[lookup]] indexes, `matches(v, q)` implies
 * that `v` is residual or that `cellKeys(v)` and `lookup(q)` share a key.
 *
 * `CriteriaIndexSpec` checks this over edge-biased values and criteria. No indexed criterion
 * matches `Empty` (only `Exact(Text(""))` does, and it is not indexed), so a row outside a slice's
 * held rows, which reads as `Empty`, never needs to be a candidate.
 */
private[formula] object CriteriaKeys:

  /** Which rows of the index a cell value belongs to. */
  enum CellKeys:
    /** No indexed criterion matches the value. */
    case NoKeys

    /** Every indexed criterion may match the value: the row is always a candidate. */
    case Residual

    case One(key: MatchKey)

    case Two(first: MatchKey, second: MatchKey)

  private val one = MatchKey.Num(java.math.BigDecimal.ONE)
  private val zero = MatchKey.Num(java.math.BigDecimal.ZERO)
  private val trueKey = MatchKey.Text("true")
  private val falseKey = MatchKey.Text("false")

  /**
   * `Character.toLowerCase(Character.toUpperCase(c))` per UTF-16 unit; the same instance when the
   * text is already folded. Two strings of BMP characters (no surrogate) that `equalsIgnoreCase`
   * accepts have equal folds: it compares unit by unit and accepts a pair when the units, their
   * upper cases, or the lower cases of their upper cases are equal, and each of those implies equal
   * folds. `CriteriaIndexSpec` checks this for every BMP character against its case orbit. Text
   * with a surrogate is compared by code point, so it is not keyed ([[hasSurrogate]]).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def fold(s: String): String =
    var i = 0
    while i < s.length && foldChar(s.charAt(i)) == s.charAt(i) do i += 1
    if i == s.length then s
    else
      val out = s.toCharArray
      while i < out.length do
        out(i) = foldChar(out(i))
        i += 1
      new String(out)

  private def foldChar(c: Char): Char = Character.toLowerCase(Character.toUpperCase(c))

  /** Whether `s` holds a UTF-16 surrogate unit. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def hasSurrogate(s: String): Boolean =
    var i = 0
    while i < s.length && !Character.isSurrogate(s.charAt(i)) do i += 1
    i < s.length

  /**
   * The key of `n` without trailing zeros: value-equal numbers (`4.20`, `4.2`) give equal keys.
   * `None` when the stripped form's scale does not fit an `Int` (`100E+2147483647`), where
   * `BigDecimal.stripTrailingZeros` throws; every value-equal number has the same stripped form, so
   * all of them get `None`. Stripping removes fewer zeros than the precision, so a scale at least
   * that far above `Int.MinValue` never overflows and skips the `Try`.
   */
  def strip(n: BigDecimal): Option[MatchKey] =
    stripped(n).map(MatchKey.Num(_))

  /** `n` without trailing zeros, or `None` where `stripTrailingZeros` throws ([[strip]]). */
  private def stripped(n: BigDecimal): Option[java.math.BigDecimal] =
    val bd = n.bigDecimal
    if bd.scale.toLong - bd.precision > Int.MinValue.toLong then Some(bd.stripTrailingZeros)
    else scala.util.Try(bd.stripTrailingZeros).toOption

  /** The largest scale magnitude [[numberText]] writes in plain notation. */
  private val MaxPlainScale = 400L

  /**
   * #780: a number as text, total and canonical: value-equal numbers (`4.20`, `4.2`) give the same
   * text. A number whose stripped scale is within ±400 (every double, and more) is written plainly
   * without trailing zeros (`4.2`, `1000`), as `toPlainString` writes it; any other in scientific
   * notation (`1E+999999999`), where `toPlainString` would build a string of up to two billion
   * characters. Where `stripTrailingZeros` throws (`100E+2147483647`), the zeros are counted in a
   * `Long` exponent instead.
   */
  def numberText(n: BigDecimal): String =
    stripped(n) match
      case Some(s) if math.abs(s.scale.toLong) <= MaxPlainScale => s.toPlainString
      case Some(s) => s.toString
      case None => scientific(n.bigDecimal)

  /** `bd` as `<d>[.<digits>]E<sign><exponent>`, its trailing zeros moved into a `Long` exponent. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def scientific(bd: java.math.BigDecimal): String =
    var unscaled = bd.unscaledValue
    var exponent = -bd.scale.toLong
    if unscaled.signum == 0 then "0"
    else
      var divided = unscaled.divideAndRemainder(java.math.BigInteger.TEN)
      while divided(1).signum == 0 do
        unscaled = divided(0)
        exponent += 1
        divided = unscaled.divideAndRemainder(java.math.BigInteger.TEN)
      val digits = unscaled.abs.toString
      val sign = if unscaled.signum < 0 then "-" else ""
      val mantissa =
        if digits.length == 1 then digits else s"${digits.charAt(0)}.${digits.substring(1)}"
      val adjusted = exponent + digits.length - 1
      s"$sign${mantissa}E${if adjusted < 0 then "-" else "+"}${math.abs(adjusted)}"

  private def textKeys(s: String): CellKeys =
    if hasSurrogate(s) then CellKeys.Residual
    else
      val text = MatchKey.Text(fold(s))
      CriteriaMatcher.parseNumeric(s) match
        case Some(n) => strip(n).fold(CellKeys.Residual)(CellKeys.Two(text, _))
        case None => CellKeys.One(text)

  /**
   * The keys of a cell value, as `matchesExact` compares it: text by its fold (and its number when
   * it parses as one), a number by its stripped value, a boolean as itself and as 1 or 0, an error
   * by its code. A date-time matches text, dates and serial numbers by rules no single key
   * captures, so it is residual, as is text with a surrogate, a number (or numeric text) whose
   * stripped form has no `Int` scale ([[strip]]), and any value this match does not list.
   */
  def cellKeys(value: CellValue): CellKeys =
    value match
      case CellValue.Empty => CellKeys.NoKeys
      case CellValue.Text(s) => textKeys(s)
      case CellValue.RichText(rt) =>
        // matchesExact compares rich text only with text
        val plain = rt.toPlainText
        if hasSurrogate(plain) then CellKeys.Residual
        else CellKeys.One(MatchKey.Text(fold(plain)))
      case CellValue.Number(n) => strip(n).fold(CellKeys.Residual)(CellKeys.One(_))
      case CellValue.Bool(b) => CellKeys.Two(MatchKey.Bool(b), if b then one else zero)
      case CellValue.Error(e) => CellKeys.One(MatchKey.Err(e))
      case CellValue.Formula(_, Some(cached), _) => cellKeys(cached)
      // matchesExact never matches an uncached formula
      case CellValue.Formula(_, None, _) => CellKeys.NoKeys
      case _ => CellKeys.Residual

  /**
   * The keys an `Exact` criterion looks up, plus the residual rows, which every lookup includes;
   * `None` when the criterion is not indexed and the call tests every row. Only `Exact` with a
   * non-empty operand is indexed: `Exact(Text(""))` matches blanks, which the index does not hold,
   * and comparisons, wildcards and negations (`<>x` is a complement) are never buckets. A date or
   * date-time operand matches only date-time cells, all residual, so it looks up no key.
   */
  def lookup(criterion: CriteriaMatcher.Criterion): Option[List[MatchKey]] =
    criterion match
      case CriteriaMatcher.Exact(expected) =>
        expected match
          case ExprValue.Text(s) if s.isEmpty || hasSurrogate(s) => None
          case ExprValue.Text(s) =>
            val folded = fold(s)
            val text = MatchKey.Text(folded)
            val bool =
              if folded == "true" then Some(MatchKey.Bool(true))
              else if folded == "false" then Some(MatchKey.Bool(false))
              else None
            // numeric text whose number has no key is not indexed
            CriteriaMatcher.parseNumeric(s) match
              case Some(n) => strip(n).map(number => text :: number :: bool.toList)
              case None => Some(text :: bool.toList)
          case ExprValue.Number(n) => strip(n).map(List(_))
          case ExprValue.Bool(b) => Some(List(MatchKey.Bool(b), if b then trueKey else falseKey))
          case ExprValue.Date(_) | ExprValue.DateTime(_) => Some(Nil)
          case ExprValue.Cell(CellValue.Error(e)) => Some(List(MatchKey.Err(e)))
          case _ => None
      case _ => None

  /** Whether the criterion is a numeric window (`>`, `>=`, `<`, `<=` against a number). */
  def isNumericWindow(criterion: CriteriaMatcher.Criterion): Boolean =
    criterion match
      case CriteriaMatcher.Compare(op, _) => op != CriteriaMatcher.CompareOp.Neq
      case _ => false

  /**
   * #920: the order key of a number, as [[OrderIndex]] sorts it: the nearest `Double`, `-0.0` read
   * as `0.0`. `BigDecimal.doubleValue` rounds correctly, so it never reverses an order: `a < b`
   * implies `orderKey(a) <= orderKey(b)` (values too large for a double round to an infinity, which
   * keeps the order). Equal keys may hold unequal numbers: the index settles them by the exact
   * values it keeps per key.
   */
  def orderKey(n: BigDecimal): Double = n.bigDecimal.doubleValue + 0.0

  /**
   * #920: the window a numeric criterion (`>`, `>=`, `<`, `<=` against a number) looks up; `None`
   * for every other criterion. `matchesCompare` matches it only on a cell whose orderable number (a
   * number or a date, #798) lies past the bound, and `Empty` has none, so no window matches a row
   * outside a slice's held rows.
   */
  def window(criterion: CriteriaMatcher.Criterion): Option[IndexLookup.Window] =
    import CriteriaMatcher.CompareOp
    criterion match
      case CriteriaMatcher.Compare(op, t) =>
        op match
          case CompareOp.Gt => Some(IndexLookup.Window(Some(Bound(t, inclusive = false)), None))
          case CompareOp.Gte => Some(IndexLookup.Window(Some(Bound(t, inclusive = true)), None))
          case CompareOp.Lt => Some(IndexLookup.Window(None, Some(Bound(t, inclusive = false))))
          case CompareOp.Lte => Some(IndexLookup.Window(None, Some(Bound(t, inclusive = true))))
          case CompareOp.Neq => None
      case _ => None

/**
 * #920: one end of a numeric window: `value`, inclusive or not. Equal by value (`4.2` and `4.20`
 * are one bound) and hashed by its order key, so a number whose Scala `BigDecimal.hashCode` throws
 * (`100E+2147483647`) still makes a memo key.
 */
private[formula] final class Bound(val value: BigDecimal, val inclusive: Boolean):
  val key: Double = CriteriaKeys.orderKey(value)

  /** Whether `v` lies on the window's side of this bound when it is a lower bound. */
  def admitsAbove(v: BigDecimal): Boolean =
    val c = v.bigDecimal.compareTo(value.bigDecimal)
    c > 0 || (c == 0 && inclusive)

  /** Whether `v` lies on the window's side of this bound when it is an upper bound. */
  def admitsBelow(v: BigDecimal): Boolean =
    val c = v.bigDecimal.compareTo(value.bigDecimal)
    c < 0 || (c == 0 && inclusive)

  /** The tighter of two lower bounds: the larger value, exclusive on a tie. */
  def maxLower(other: Bound): Bound =
    val c = value.bigDecimal.compareTo(other.value.bigDecimal)
    if c > 0 || (c == 0 && !inclusive) then this else other

  /** The tighter of two upper bounds: the smaller value, exclusive on a tie. */
  def minUpper(other: Bound): Bound =
    val c = value.bigDecimal.compareTo(other.value.bigDecimal)
    if c < 0 || (c == 0 && !inclusive) then this else other

  override def equals(other: Any): Boolean = other match
    case b: Bound =>
      inclusive == b.inclusive && value.bigDecimal.compareTo(b.value.bigDecimal) == 0
    case _ => false

  override def hashCode: Int = java.lang.Double.hashCode(key) * 31 + (if inclusive then 1 else 0)

  override def toString: String = s"Bound(${value.bigDecimal}, inclusive = $inclusive)"

/**
 * What one indexed criteria pair looks up: equality keys in its slice's [[EqIndex]] (#767), or a
 * numeric window in its slice's [[OrderIndex]] (#920).
 */
private[formula] enum IndexLookup derives CanEqual:
  case Keys(keys: List[MatchKey])

  /** Rows whose orderable number lies past `lower` and before `upper` (absent: unbounded). */
  case Window(lower: Option[Bound], upper: Option[Bound])

private[formula] object IndexLookup:

  /**
   * #920: the lookups of the criteria a call reads through slices, each `(slice, firstRow,
   * criterion)`: an equality lookup per keyed criterion, in order, then one window per range (slice
   * and first row) holding numeric windows, the intersection of its windows (`>=start` and `<end`
   * on one column become one interval). A row that matches every window of its range lies past
   * every bound, so inside the intersection.
   */
  def of(
    sliced: List[(ColumnSlice, Int, CriteriaMatcher.Criterion)]
  ): List[(ColumnSlice, IndexLookup, Int)] =
    val keyed = sliced.flatMap((slice, firstRow, criterion) =>
      CriteriaKeys.lookup(criterion).map(keys => (slice, Keys(keys), firstRow))
    )
    val windows = sliced.flatMap((slice, firstRow, criterion) =>
      CriteriaKeys.window(criterion).map(w => ((slice, firstRow), w))
    )
    val ranges = windows.map(_._1).distinctBy((slice, firstRow) => (slice.id, firstRow))
    val merged = ranges.map { (slice, firstRow) =>
      val mine = windows.collect {
        case ((s, r), w: Window) if s.id == slice.id && r == firstRow => w
      }
      val lower = mine.flatMap(_.lower).reduceOption(_.maxLower(_))
      val upper = mine.flatMap(_.upper).reduceOption(_.minUpper(_))
      (slice, Window(lower, upper), firstRow)
    }
    keyed ++ merged

/**
 * #767: the equality index of one [[ColumnSlice]], as compressed sparse rows: `keyOff(k) until
 * keyOff(k + 1)` of `rows` are the sheet rows (0-based) holding key `k`, ascending; `residual` the
 * rows that are candidates for every lookup, ascending. Built once per slice and never changed.
 */
private[formula] final class EqIndex private (
  keyIds: Map[MatchKey, Int],
  keyOff: IArray[Int],
  rows: IArray[Int],
  residual: IArray[Int],
  /**
   * An estimate of the memory the index holds, charged against `PassCache.Limits.maxIndexBytes`.
   */
  val bytes: Long
):
  /** The posting lists `keys` select (keys absent from the slice select nothing), then residual. */
  private[eval] def postings(keys: List[MatchKey]): List[EqIndex.Postings] =
    val keyed = keys.distinct.flatMap(keyIds.get).map { k =>
      EqIndex.Postings(rows, keyOff(k), keyOff(k + 1))
    }
    if residual.isEmpty then keyed else keyed :+ EqIndex.Postings(residual, 0, residual.length)

  /** Distinct keys. */
  def keyCount: Int = keyOff.length - 1

  /** Rows that hold a key. */
  def keyedRows: Int = rows.length

  /** Rows that are candidates for every lookup. */
  def residualRows: Int = residual.length

private[formula] object EqIndex:

  /** `rows(from until until)`: one ascending posting list. */
  private[eval] final case class Postings(rows: IArray[Int], from: Int, until: Int):
    /** The first index in `from until until` whose row is at least `row` (`until` when none). */
    @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
    def lowerBound(row: Int, start: Int): Int =
      var lo = start
      var hi = until
      while lo < hi do
        val mid = (lo + hi) >>> 1
        if rows(mid) < row then lo = mid + 1 else hi = mid
      lo

    /** How many rows lie in `a..b`. */
    def countIn(a: Int, b: Int): Int = lowerBound(b + 1, from) - lowerBound(a, from)

  /**
   * Index the rows `slice` holds. Each held row's value gives its keys ([[CriteriaKeys.cellKeys]]);
   * the rows are grouped by key with a counting sort, so each list stays ascending.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def build(slice: ColumnSlice): EqIndex =
    val first = slice.firstHeldRow
    val held = slice.heldRows
    val ids = scala.collection.mutable.HashMap.empty[MatchKey, Int]
    val order = scala.collection.mutable.ArrayBuffer.empty[MatchKey]
    // (key id, row) pairs in row order, and the residual rows
    var pairKeys = new Array[Int](math.max(16, held))
    var pairRows = new Array[Int](math.max(16, held))
    var pairs = 0
    val residual = scala.collection.mutable.ArrayBuilder.make[Int]
    def add(key: MatchKey, row: Int): Unit =
      val id = ids.getOrElseUpdate(key, { order += key; order.length - 1 })
      if pairs == pairKeys.length then
        pairKeys = java.util.Arrays.copyOf(pairKeys, pairs * 2)
        pairRows = java.util.Arrays.copyOf(pairRows, pairs * 2)
      pairKeys(pairs) = id
      pairRows(pairs) = row
      pairs += 1
    var i = 0
    while i < held do
      val row = first + i
      CriteriaKeys.cellKeys(slice.valueAt(row)) match
        case CriteriaKeys.CellKeys.NoKeys => ()
        case CriteriaKeys.CellKeys.Residual => residual += row
        case CriteriaKeys.CellKeys.One(key) => add(key, row)
        case CriteriaKeys.CellKeys.Two(a, b) =>
          add(a, row)
          add(b, row)
      i += 1
    val keyCount = order.length
    val keyOff = new Array[Int](keyCount + 1)
    i = 0
    while i < pairs do
      keyOff(pairKeys(i) + 1) += 1
      i += 1
    var k = 0
    while k < keyCount do
      keyOff(k + 1) += keyOff(k)
      k += 1
    val next = java.util.Arrays.copyOf(keyOff, keyCount)
    val rows = new Array[Int](pairs)
    i = 0
    while i < pairs do
      val id = pairKeys(i)
      rows(next(id)) = pairRows(i)
      next(id) += 1
      i += 1
    val residualRows = residual.result()
    val keyBytes = order.iterator.map {
      case MatchKey.Text(s) => 64L + 2L * s.length
      case _ => 64L
    }.sum
    new EqIndex(
      order.iterator.zipWithIndex.toMap,
      IArray.unsafeFromArray(keyOff),
      IArray.unsafeFromArray(rows),
      IArray.unsafeFromArray(residualRows),
      keyBytes + 4L * (keyCount + 1) + 4L * pairs + 4L * residualRows.length
    )

/**
 * #920: the order index of one [[ColumnSlice]]: its number and date rows grouped by order key
 * ([[CriteriaKeys.orderKey]]), keys ascending, in compressed sparse rows: `keyOff(k) until keyOff(k
 * + 1)` of `rows` are the sheet rows (0-based, ascending) whose key is `keys(k)`, and
 * `least(k)`/`greatest(k)` the smallest and largest exact number among them. `byRow(i)` is the key
 * of held row `first + i`, `NaN` for a row no numeric window matches (text, booleans, blanks,
 * errors). Built once per slice and never changed.
 */
private[formula] final class OrderIndex private (
  first: Int,
  byRow: IArray[Double],
  keys: IArray[Double],
  keyOff: IArray[Int],
  rows: IArray[Int],
  least: IArray[BigDecimal],
  greatest: IArray[BigDecimal],
  /**
   * An estimate of the memory the index holds, charged against `PassCache.Limits.maxIndexBytes`.
   */
  val bytes: Long
):
  /**
   * The rows a window selects: one contiguous run of the sorted order. Order keys never reverse an
   * order, so every key group strictly inside the bounds' keys matches and every group outside
   * cannot; a group whose key equals a bound's key is kept when its greatest (lower bound) or least
   * (upper bound) number lies past the bound. The run is a superset of the matching rows, exact but
   * for a boundary group that mixes matching and non-matching numbers.
   */
  def window(w: IndexLookup.Window): CriteriaCandidates.Selection =
    val a = w.lower.fold(0) { b =>
      val i = OrderIndex.lowerBound(keys, b.key)
      if i < keys.length && keys(i) == b.key && !b.admitsAbove(greatest(i)) then i + 1 else i
    }
    val z = w.upper.fold(keys.length) { b =>
      val j = OrderIndex.lowerBoundAbove(keys, b.key)
      if j > 0 && keys(j - 1) == b.key && !b.admitsBelow(least(j - 1)) then j - 1 else j
    }
    CriteriaCandidates.Run(
      this,
      keyOff(a),
      keyOff(math.max(a, z)),
      w.lower.fold(Double.NegativeInfinity)(_.key),
      w.upper.fold(Double.PositiveInfinity)(_.key)
    )

  /** Row `i` of the sorted order. */
  private[eval] def rowAt(i: Int): Int = rows(i)

  /** Whether `row`'s key lies in `lo..hi` (never for a row without a key). */
  private[eval] def within(row: Int, lo: Double, hi: Double): Boolean =
    val i = row - first
    if i < 0 || i >= byRow.length then false
    else
      val k = byRow(i)
      lo <= k && k <= hi

  /** Whether rows `a..b` include every held row. */
  private[eval] def spans(a: Int, b: Int): Boolean = a <= first && first + byRow.length - 1 <= b

  /** Distinct keys. */
  def keyCount: Int = keys.length

  /** Rows that hold a key. */
  def keyedRows: Int = rows.length

private[formula] object OrderIndex:

  /** The first index whose key is at least `k` (`keys.length` when none). */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def lowerBound(keys: IArray[Double], k: Double): Int =
    var lo = 0
    var hi = keys.length
    while lo < hi do
      val mid = (lo + hi) >>> 1
      if keys(mid) < k then lo = mid + 1 else hi = mid
    lo

  /** The first index whose key is above `k` (`keys.length` when none). */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def lowerBoundAbove(keys: IArray[Double], k: Double): Int =
    var lo = 0
    var hi = keys.length
    while lo < hi do
      val mid = (lo + hi) >>> 1
      if keys(mid) <= k then lo = mid + 1 else hi = mid
    lo

  /**
   * Index the rows `slice` holds: each held row with an orderable number (a number or a date, as
   * `CriteriaMatcher.orderableNumber` reads it) gets its order key. The keys are sorted and
   * deduplicated, then the rows are grouped by key with a counting sort, so each group stays
   * ascending: `O(n log n)` once per slice.
   */
  @SuppressWarnings(
    Array("org.wartremover.warts.Var", "org.wartremover.warts.While", "org.wartremover.warts.Null")
  )
  def build(slice: ColumnSlice): OrderIndex =
    val first = slice.firstHeldRow
    val held = slice.heldRows
    val byRow = new Array[Double](held)
    val exact = new Array[BigDecimal](held)
    var keyed = 0
    var i = 0
    while i < held do
      CriteriaMatcher.orderableNumber(slice.valueAt(first + i)) match
        case Some(n) =>
          byRow(i) = CriteriaKeys.orderKey(n)
          exact(i) = n
          keyed += 1
        case None => byRow(i) = Double.NaN
      i += 1
    val sorted = new Array[Double](keyed)
    var w = 0
    i = 0
    while i < held do
      if !byRow(i).isNaN then
        sorted(w) = byRow(i)
        w += 1
      i += 1
    java.util.Arrays.sort(sorted)
    // distinct keys, in place
    var distinct = 0
    i = 0
    while i < keyed do
      if distinct == 0 || sorted(distinct - 1) != sorted(i) then
        sorted(distinct) = sorted(i)
        distinct += 1
      i += 1
    val keys = java.util.Arrays.copyOf(sorted, distinct)
    val keyOff = new Array[Int](distinct + 1)
    val least = new Array[BigDecimal](distinct)
    val greatest = new Array[BigDecimal](distinct)
    val rank = new Array[Int](held)
    i = 0
    while i < held do
      if !byRow(i).isNaN then
        val r = java.util.Arrays.binarySearch(keys, byRow(i))
        rank(i) = r
        keyOff(r + 1) += 1
        val n = exact(i)
        if least(r) == null || n.bigDecimal.compareTo(least(r).bigDecimal) < 0 then least(r) = n
        if greatest(r) == null || n.bigDecimal.compareTo(greatest(r).bigDecimal) > 0 then
          greatest(r) = n
      i += 1
    var k = 0
    while k < distinct do
      keyOff(k + 1) += keyOff(k)
      k += 1
    val next = java.util.Arrays.copyOf(keyOff, distinct)
    val rows = new Array[Int](keyed)
    i = 0
    while i < held do
      if !byRow(i).isNaN then
        rows(next(rank(i))) = first + i
        next(rank(i)) += 1
      i += 1
    new OrderIndex(
      first,
      IArray.unsafeFromArray(byRow),
      IArray.unsafeFromArray(keys),
      IArray.unsafeFromArray(keyOff),
      IArray.unsafeFromArray(rows),
      IArray.unsafeFromArray(least),
      IArray.unsafeFromArray(greatest),
      // per held row its key; per key its value, offset and two numbers (a date's are its own)
      64L + 8L * held + (8L + 4L + 2L * 64L) * distinct + 4L + 4L * keyed
    )

/**
 * #767: candidate rows of a criteria call from its indexed criteria. Every criteria range of the
 * call has `n` rows; range `k` starts at sheet row `offset_k`, so offset `i` of the call is row
 * `offset_k + i` of range `k`.
 *
 * Each indexed pair `k` selects the offsets whose row holds one of its lookup keys or is residual
 * (an equality pair), or whose number lies in its window's run of the sorted order (a window pair,
 * #920). The candidates are the offsets every indexed pair selects: the pair selecting the fewest
 * rows is listed, and each further pair keeps the listed offsets it selects, found by galloping
 * search forward through its posting lists, or by reading each candidate's order key. A pair's
 * selection is the DISTINCT rows its posting lists hold in the window: a row holding both keys of a
 * two-key lookup (text `"1"` has the Text and Num keys a lookup for `"1"` reads) sits in two lists
 * and counts once (#842). When even the smallest selection is more than half the rows, listing it
 * costs about as much as testing every row, so the call tests every row (`None`); the count stops
 * as soon as it passes that bound.
 */
private[formula] object CriteriaCandidates:

  /**
   * The sheet rows one indexed pair selects: posting lists of an [[EqIndex]] ([[Postings]]) or a
   * run of an [[OrderIndex]] ([[Run]]). Every method takes the pair's window, `n` rows from sheet
   * row `offset`.
   */
  sealed trait Selection:
    /** The distinct selected rows in the window, exact up to `limit` ([[selectionSize]]). */
    def size(offset: Int, n: Int, limit: Int): Int

    /** The ascending, distinct offsets `row - offset` of the selected rows in the window. */
    def list(offset: Int, n: Int): Array[Int]

    /** The offsets of `cand` (ascending) whose row `offset + i` is selected. */
    def keep(cand: Array[Int], offset: Int): Array[Int]

  /** #767: the rows of an equality lookup's posting lists. */
  final case class Postings(lists: List[EqIndex.Postings]) extends Selection:
    def size(offset: Int, n: Int, limit: Int): Int = selectionSize(lists, offset, n, limit)
    def list(offset: Int, n: Int): Array[Int] = union(lists, offset, n)
    def keep(cand: Array[Int], offset: Int): Array[Int] =
      CriteriaCandidates.keep(cand, lists, offset)

  /**
   * #920: rows `index.rowAt(from until until)`, the run of an [[OrderIndex]] whose keys lie in
   * `lo..hi`. The run is in key order, so listing it sorts the rows in the window (`O(k log k)` for
   * `k` of them); keeping candidates reads each candidate's key (`O(1)` a candidate).
   */
  final case class Run(index: OrderIndex, from: Int, until: Int, lo: Double, hi: Double)
      extends Selection:
    @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
    def size(offset: Int, n: Int, limit: Int): Int =
      val last = offset + n - 1
      if index.spans(offset, last) then until - from
      else
        var count = 0
        var i = from
        while i < until && count <= limit do
          val row = index.rowAt(i)
          if row >= offset && row <= last then count += 1
          i += 1
        count

    @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
    def list(offset: Int, n: Int): Array[Int] =
      val last = offset + n - 1
      val out = new Array[Int](until - from)
      var w = 0
      var i = from
      while i < until do
        val row = index.rowAt(i)
        if row >= offset && row <= last then
          out(w) = row - offset
          w += 1
        i += 1
      val listed = java.util.Arrays.copyOf(out, w)
      java.util.Arrays.sort(listed)
      listed

    def keep(cand: Array[Int], offset: Int): Array[Int] =
      cand.filter(i => index.within(offset + i, lo, hi))

  /** One indexed pair of a call: the rows it selects, and its range's first row. */
  final case class Pair(selection: Selection, offset: Int)

  object Pair:
    /** An equality pair: the posting lists `keys` select in `index`. */
    def apply(index: EqIndex, keys: List[MatchKey], offset: Int): Pair =
      Pair(Postings(index.postings(keys)), offset)

  /**
   * One indexed pair as a memo key: the slice by its identity in the pass, the lookup, and the
   * first row. Two calls whose indexed pairs have equal lookups and `n` share their candidates.
   */
  final case class PairKey(slice: ColumnSlice.Id, lookup: IndexLookup, offset: Int)

  final case class MemoKey(n: Int, pairs: List[PairKey])

  /** The ascending candidate offsets in `0 until n`, or `None` to test every offset. */
  def compute(pairs: List[Pair], n: Int): Option[ArraySeq.ofInt] =
    if pairs.isEmpty || n <= 0 then None
    else
      val limit = n / 2
      val selections = pairs.map(p => (p.selection, p.offset, p.selection.size(p.offset, n, limit)))
      val sorted = selections.sortBy(_._3)
      sorted match
        case (selection, offset, size) :: rest =>
          if size > limit then None
          else
            val listed = selection.list(offset, n)
            Some(
              ArraySeq.ofInt(rest.foldLeft(listed) { case (cand, (more, off, _)) =>
                if cand.isEmpty then cand else more.keep(cand, off)
              })
            )
        case Nil => None

  /**
   * The distinct rows of `lists` in `offset .. offset + n - 1`, exact up to `limit`: a larger count
   * is reported as some count past `limit`. One list is counted whole by binary search (`countIn`);
   * several are merged in row order, a row in more than one list counting once, and the merge stops
   * one past `limit`, so a refusal costs at most about `limit` steps per list (#842). A summed
   * count overstated a window of two-key rows by up to twice and refused it.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While")) // hot path
  private[eval] def selectionSize(
    lists: List[EqIndex.Postings],
    offset: Int,
    n: Int,
    limit: Int
  ): Int =
    val last = offset + n - 1
    lists match
      case Nil => 0
      case List(p) => p.countIn(offset, last)
      case _ =>
        // each list's window, found once: its bounds give its count and start the merge
        val parts = lists.map { p =>
          val from = p.lowerBound(offset, p.from)
          (p.rows, from, p.lowerBound(last + 1, from))
        }
        // every row of one list is distinct: past the limit, no merge is needed
        if parts.exists((_, from, until) => until - from > limit) then limit + 1
        else
          val rows = parts.map(_._1).toArray
          val cursors = parts.map(_._2).toArray
          val untils = parts.map(_._3).toArray
          // the smallest row at any cursor, Int.MaxValue when every list is spent
          def smallest: Int =
            var min = Int.MaxValue
            var i = 0
            while i < rows.length do
              if cursors(i) < untils(i) then min = math.min(min, rows(i)(cursors(i)))
              i += 1
            min
          var count = 0
          var row = smallest
          while count <= limit && row != Int.MaxValue do
            count += 1
            var i = 0
            while i < rows.length do
              if cursors(i) < untils(i) && rows(i)(cursors(i)) == row then cursors(i) += 1
              i += 1
            row = smallest
          count

  /** The ascending, distinct offsets `row - offset` of the rows of `lists` in the window. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def union(lists: List[EqIndex.Postings], offset: Int, n: Int): Array[Int] =
    val last = offset + n - 1
    val parts = lists.map { p =>
      val from = p.lowerBound(offset, p.from)
      val until = p.lowerBound(last + 1, from)
      (p.rows, from, until)
    }
    parts match
      case List((rows, from, until)) =>
        Array.tabulate(until - from)(i => rows(from + i) - offset)
      case _ =>
        val all = new Array[Int](parts.map((_, f, u) => u - f).sum)
        var at = 0
        parts.foreach { (rows, from, until) =>
          var i = from
          while i < until do
            all(at) = rows(i) - offset
            at += 1
            i += 1
        }
        java.util.Arrays.sort(all)
        // drop repeats: a row may sit in more than one list (text "1" holds Text and Num keys)
        var w = 0
        var r = 0
        while r < all.length do
          if w == 0 || all(w - 1) != all(r) then
            all(w) = all(r)
            w += 1
          r += 1
        java.util.Arrays.copyOf(all, w)

  /** The offsets of `cand` (ascending) whose row `offset + i` lies in one of `lists`. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def keep(cand: Array[Int], lists: List[EqIndex.Postings], offset: Int): Array[Int] =
    val selected = new Array[Boolean](cand.length)
    lists.foreach { p =>
      var cursor = p.from
      var i = 0
      while i < cand.length && cursor < p.until do
        val row = cand(i) + offset
        cursor = gallop(p, row, cursor)
        if cursor < p.until && p.rows(cursor) == row then selected(i) = true
        i += 1
    }
    var count = 0
    var i = 0
    while i < cand.length do
      if selected(i) then count += 1
      i += 1
    val out = new Array[Int](count)
    var w = 0
    i = 0
    while i < cand.length do
      if selected(i) then
        out(w) = cand(i)
        w += 1
      i += 1
    out

  /**
   * The first index at or after `cursor` whose row is at least `row`: doubling steps from the
   * cursor, then a binary search inside the last step, so a pass over ascending rows costs
   * `O(log gap)` per row.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def gallop(p: EqIndex.Postings, row: Int, cursor: Int): Int =
    if cursor >= p.until || p.rows(cursor) >= row then cursor
    else
      var lo = cursor
      var step = 1
      while lo + step < p.until && p.rows(lo + step) < row do
        lo += step
        step *= 2
      val hi = math.min(p.until, lo + step + 1)
      EqIndex.Postings(p.rows, lo + 1, hi).lowerBound(row, lo + 1)

  /** An estimate of the memory a memoised candidate list holds. */
  def bytes(candidates: Option[ArraySeq.ofInt]): Long =
    64L + 4L * candidates.fold(0)(_.length)
