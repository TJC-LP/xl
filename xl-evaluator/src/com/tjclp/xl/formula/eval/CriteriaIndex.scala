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
    val bd = n.bigDecimal
    if bd.scale.toLong - bd.precision > Int.MinValue.toLong then
      Some(MatchKey.Num(bd.stripTrailingZeros))
    else scala.util.Try(bd.stripTrailingZeros).toOption.map(MatchKey.Num(_))

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
 * #767: candidate rows of a criteria call from its indexed criteria. Every criteria range of the
 * call has `n` rows; range `k` starts at sheet row `offset_k`, so offset `i` of the call is row
 * `offset_k + i` of range `k`.
 *
 * Each indexed pair `k` selects the offsets whose row holds one of its lookup keys or is residual.
 * The candidates are the offsets every indexed pair selects: the pair selecting the fewest rows is
 * listed, and each further pair keeps the listed offsets it selects, found by galloping search
 * forward through its posting lists. When even the smallest selection is more than half the rows,
 * listing it costs about as much as testing every row, so the call tests every row (`None`).
 */
private[formula] object CriteriaCandidates:

  /** One indexed pair of a call: its slice's index, its lookup keys, and its range's first row. */
  final case class Pair(index: EqIndex, keys: List[MatchKey], offset: Int)

  /**
   * One indexed pair as a memo key: the slice by its identity in the pass, the keys, and the first
   * row. Two calls whose indexed pairs have equal keys and `n` share their candidates.
   */
  final case class PairKey(slice: ColumnSlice.Id, keys: List[MatchKey], offset: Int)

  final case class MemoKey(n: Int, pairs: List[PairKey])

  /** The ascending candidate offsets in `0 until n`, or `None` to test every offset. */
  def compute(pairs: List[Pair], n: Int): Option[ArraySeq.ofInt] =
    if pairs.isEmpty || n <= 0 then None
    else
      val selections = pairs.map { p =>
        val lists = p.index.postings(p.keys)
        // an upper bound of the pair's selection: its lists may share rows
        val size = lists.map(_.countIn(p.offset, p.offset + n - 1)).sum
        (lists, p.offset, size)
      }
      val sorted = selections.sortBy(_._3)
      sorted match
        case (lists, offset, size) :: rest =>
          if size > n / 2 then None
          else
            val listed = union(lists, offset, n)
            Some(
              ArraySeq.ofInt(rest.foldLeft(listed) { case (cand, (more, off, _)) =>
                if cand.isEmpty then cand else keep(cand, more, off)
              })
            )
        case Nil => None

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
