package com.tjclp.xl.formula.eval

import java.time.LocalDateTime

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.ExprValue
import com.tjclp.xl.richtext.RichText
import org.scalacheck.{Gen, Prop}

/**
 * #920: the order index of numeric windows is a filter, never a judge.
 *
 *   - Completeness: whenever `CriteriaMatcher.matches(v, q)` holds for a numeric window `q` (`>`,
 *     `>=`, `<`, `<=` against a number), the value has an order key inside the window's interval.
 *     Checked over the cross product of edge-biased values and thresholds (numbers a `Double`
 *     cannot hold, infinities, signed zeros, dates and date-times), and on random ones.
 *   - Candidate offsets of mixed equality and window pairs contain every offset whose criteria all
 *     match, ascending and distinct (merged windows included).
 *   - A run is sized exactly up to the bound it is compared with.
 *
 * Seeds follow the law-fuzz convention (GH-308): fixed by default, random under
 * `XL_ROUNDTRIP_RANDOM_SEED=true`, replayable with `-Dxl.roundtrip.seed=<printed seed>`.
 */
class OrderIndexSpec extends LawFuzzSuite("#920 order index", 20261120L):

  private val data = SheetName.unsafe("Data")

  // ---------- edge-biased pools ----------

  /** Numbers around the edges of `Double`: too precise, too large, too small, signed zeros. */
  private val numbers: Vector[BigDecimal] = Vector(
    "0",
    "0.0",
    "-0",
    "1",
    "1.0",
    "-1",
    "0.1",
    "0.1000000000000000000000001",
    "0.0999999999999999999999999",
    "0.30000000000000004",
    "0.3",
    "2.5",
    "45659",
    "45659.5",
    "45658.99999999999999999999",
    "9007199254740993", // 2^53 + 1: no double holds it
    "9007199254740992",
    "9007199254740994",
    "1E+308",
    "1.7976931348623157E+308",
    "1E+309",
    "-1E+309",
    "1E+999999999",
    "-1E+999999999",
    "1E-400",
    "-1E-400",
    "4.9E-324",
    "1E-999999999"
  ).map(BigDecimal(_)) ++ Vector(
    BigDecimal(new java.math.BigDecimal(java.math.BigInteger.valueOf(100), Int.MinValue + 1)),
    BigDecimal(new java.math.BigDecimal(java.math.BigInteger.valueOf(-7), Int.MaxValue))
  )

  private val dateTimes: Vector[LocalDateTime] = Vector(
    LocalDateTime.of(1900, 1, 1, 0, 0),
    LocalDateTime.of(2024, 1, 31, 0, 0),
    LocalDateTime.of(2024, 2, 1, 0, 0),
    LocalDateTime.of(2024, 2, 1, 12, 0),
    LocalDateTime.of(2024, 2, 1, 23, 59, 59, 999_999_999)
  )

  private val plainValues: Vector[CellValue] =
    Vector[CellValue](
      CellValue.Empty,
      CellValue.Text("1"),
      CellValue.Text("45659"),
      CellValue.Text("x"),
      CellValue.RichText(RichText.plain("2")),
      CellValue.Bool(true),
      CellValue.Bool(false),
      CellValue.Error(CellError.NA)
    ) ++ numbers.map(CellValue.Number(_)) ++ dateTimes.map(CellValue.DateTime(_))

  private val values: Vector[CellValue] =
    plainValues ++
      plainValues.filter(_ != CellValue.Empty).map(v => CellValue.Formula("A1", Some(v))) ++
      Vector(CellValue.Formula("A1", None))

  private val ops: Vector[CriteriaMatcher.CompareOp] = Vector(
    CriteriaMatcher.CompareOp.Gt,
    CriteriaMatcher.CompareOp.Gte,
    CriteriaMatcher.CompareOp.Lt,
    CriteriaMatcher.CompareOp.Lte
  )

  private val thresholds: Vector[BigDecimal] =
    numbers ++ dateTimes.map(dt => BigDecimal(CellValue.dateTimeToExcelSerial(dt)))

  private val windows: Vector[CriteriaMatcher.Criterion] =
    for
      op <- ops
      t <- thresholds
    yield CriteriaMatcher.Compare(op, t)

  /** The closed interval of order keys a window's bounds span. */
  private def keyInterval(w: IndexLookup.Window): (Double, Double) =
    (
      w.lower.fold(Double.NegativeInfinity)(_.key),
      w.upper.fold(Double.PositiveInfinity)(_.key)
    )

  private def bound(value: String, inclusive: Boolean): Bound =
    Bound(BigDecimal(value), inclusive)

  /** The completeness lemma for one pair; `None` when it holds. */
  private def violation(v: CellValue, q: CriteriaMatcher.Criterion): Option[String] =
    CriteriaKeys.window(q) match
      case None => None
      case Some(w) =>
        if !CriteriaMatcher.matches(v, q) then None
        else
          CriteriaMatcher.orderableNumber(v) match
            case None => Some(s"$v matches $q without an orderable number")
            case Some(n) =>
              val k = CriteriaKeys.orderKey(n)
              val (lo, hi) = keyInterval(w)
              val exact = w.lower.forall(_.admitsAbove(n)) && w.upper.forall(_.admitsBelow(n))
              if !exact then Some(s"$v matches $q outside its bounds $w")
              else if lo <= k && k <= hi then None
              else Some(s"$v (key $k) matches $q, keys $lo..$hi")

  // ---------- completeness ----------

  test("completeness: every edge-biased value and numeric window (full cross product)") {
    val bad = for
      v <- values
      q <- windows
      problem <- violation(v, q)
    yield problem
    assertEquals(bad.take(10), Vector.empty)
    // not vacuous
    val matches = for
      v <- values
      q <- windows if CriteriaMatcher.matches(v, q)
    yield ()
    assert(matches.size > 1000, s"${matches.size} matches")
  }

  test("what has a window: >, >=, <, <= against a number; never <>, text or equality") {
    def window(s: String): Option[IndexLookup] =
      CriteriaKeys.window(CriteriaMatcher.parse(ExprValue.Text(s)))
    assertEquals(window(">=4"), Some(IndexLookup.Window(Some(bound("4", true)), None)))
    assertEquals(window(">4.0"), Some(IndexLookup.Window(Some(bound("4", false)), None)))
    assertEquals(window("<4"), Some(IndexLookup.Window(None, Some(bound("4", false)))))
    assertEquals(window("<=-0"), Some(IndexLookup.Window(None, Some(bound("0", true)))))
    // value-equal bounds are one memo key, even where Scala's BigDecimal.hashCode throws
    assertEquals(bound("4.20", true).hashCode, bound("4.2", true).hashCode)
    val huge =
      BigDecimal(new java.math.BigDecimal(java.math.BigInteger.valueOf(100), Int.MinValue + 1))
    assertEquals(Bound(huge, false), Bound(huge, false))
    assertEquals(Bound(huge, false).hashCode, Bound(huge, false).hashCode)
    Vector("<>4", ">x", "<=", "4", "=4", "a*", "", "<>").foreach(s =>
      assertEquals(window(s), None, s)
    )
    assertEquals(CriteriaKeys.window(CriteriaMatcher.parse(ExprValue.Number(4))), None)
  }

  test("no window matches Empty, so rows outside a slice are never candidates") {
    windows.foreach(q => assert(!CriteriaMatcher.matches(CellValue.Empty, q), q.toString))
  }

  test("order keys never reverse an order and read -0 as 0") {
    val sorted = numbers.sorted
    sorted.zip(sorted.drop(1)).foreach { (a, b) =>
      assert(CriteriaKeys.orderKey(a) <= CriteriaKeys.orderKey(b), s"$a, $b")
    }
    assert(CriteriaKeys.orderKey(BigDecimal("-1E-400")).equals(0.0))
    assertEquals(CriteriaKeys.orderKey(BigDecimal("1E+999999999")), Double.PositiveInfinity)
  }

  private val genNumber: Gen[BigDecimal] = Gen.oneOf(
    Gen.oneOf(numbers),
    Gen.choose(-5000L, 5000L).map(n => BigDecimal(n) / 100),
    Gen.choose(-1e12, 1e12).map(BigDecimal(_)),
    Gen.zip(Gen.choose(-99999L, 99999L), Gen.choose(-30, 30)).map((u, s) => BigDecimal(u, s))
  )

  private val genValue: Gen[CellValue] = Gen.frequency(
    2 -> Gen.oneOf(values),
    4 -> genNumber.map(CellValue.Number(_)),
    1 -> Gen
      .choose(0L, 3000000L)
      .map(m => CellValue.DateTime(LocalDateTime.of(2020, 1, 1, 0, 0).plusMinutes(m)))
  )

  private val genWindow: Gen[CriteriaMatcher.Criterion] =
    Gen.zip(Gen.oneOf(ops), genNumber).map((op, t) => CriteriaMatcher.Compare(op, t))

  property("completeness: random values and numeric windows") {
    Prop.forAll(genValue, genWindow) { (v, q) =>
      val problem = violation(v, q)
      Prop(problem.isEmpty) :| problem.getOrElse("")
    }
  }

  // ---------- the index as a filter, end to end ----------

  private def slice(col: Int, r0: Int, values: Vector[CellValue]): ColumnSlice =
    new ColumnSlice(data, col, 0, 1048575, r0, IArray.from(values), 0L)

  private val genColumn: Gen[(Int, Vector[CellValue])] =
    for
      r0 <- Gen.choose(0, 10)
      n <- Gen.choose(0, 60)
      vs <- Gen.listOfN(n, genValue.filter(_ != CellValue.Formula("A1", None)))
    yield (r0, vs.toVector)

  private val genCriterion: Gen[CriteriaMatcher.Criterion] = Gen.frequency(
    3 -> genWindow,
    1 -> genNumber.map(n => CriteriaMatcher.Exact(ExprValue.Number(n))),
    1 -> Gen.oneOf("1", "x", "<>1", ">x").map(s => CriteriaMatcher.parse(ExprValue.Text(s)))
  )

  property("filter law: every offset whose criteria all match is a candidate (windows merged)") {
    // up to three columns, each read by up to three criteria through one range
    Prop.forAll(
      Gen
        .choose(1, 3)
        .flatMap(k =>
          Gen.listOfN(
            k,
            Gen.zip(
              genColumn,
              Gen.choose(0, 15),
              Gen.choose(1, 3).flatMap(Gen.listOfN(_, genCriterion))
            )
          )
        ),
      Gen.choose(1, 60)
    ) { (columns, n) =>
      val ranges = columns.zipWithIndex.map { case (((r0, vs), offset, qs), i) =>
        (slice(i, r0, vs), offset, qs)
      }
      val lookups = IndexLookup.of(ranges.flatMap((s, offset, qs) => qs.map(q => (s, offset, q))))
      val pairs = lookups.map { (s, lookup, offset) =>
        lookup match
          case IndexLookup.Keys(keys) => CriteriaCandidates.Pair(EqIndex.build(s), keys, offset)
          case w: IndexLookup.Window =>
            CriteriaCandidates.Pair(OrderIndex.build(s).window(w), offset)
      }
      val matching = (0 until n).filter(i =>
        ranges.forall((s, offset, qs) =>
          qs.forall(CriteriaMatcher.matches(s.valueAt(offset + i), _))
        )
      )
      val windowsMerged = lookups.count {
        case (_, _: IndexLookup.Window, _) => true
        case _ => false
      } <= ranges.size
      CriteriaCandidates.compute(pairs, n) match
        case Some(offsets) =>
          val listed = offsets.toVector
          Prop(windowsMerged) :| "more than one window per range" &&
          Prop(matching.forall(listed.contains)) :|
            s"missing ${matching.filterNot(listed.contains)}" &&
            Prop(listed == listed.distinct.sorted) :| s"not ascending and distinct: $listed" &&
            Prop(listed.forall(i => i >= 0 && i < n)) :| s"outside 0 until $n: $listed"
        case None => Prop(windowsMerged) :| "more than one window per range"
    }
  }

  property("a run holds every match and nothing outside its keys; sized exactly up to its bound") {
    Prop.forAll(genColumn, Gen.choose(0, 15), Gen.choose(1, 70), Gen.choose(0, 40), genWindow) {
      case ((r0, vs), offset, n, limit, q) =>
        val s = slice(0, r0, vs)
        val w = CriteriaKeys.window(q).getOrElse(fail(q.toString))
        val (lo, hi) = keyInterval(w)
        val run = OrderIndex.build(s).window(w)
        val inKeys = (0 until n).filter { i =>
          CriteriaMatcher
            .orderableNumber(s.valueAt(offset + i))
            .map(CriteriaKeys.orderKey)
            .exists(k => lo <= k && k <= hi)
        }
        val matching = (0 until n).filter(i => CriteriaMatcher.matches(s.valueAt(offset + i), q))
        val listed = run.list(offset, n).toVector
        val size = run.size(offset, n, limit)
        Prop(
          matching.forall(listed.contains)
        ) :| s"missing ${matching.filterNot(listed.contains)}" &&
        Prop(listed.forall(inKeys.contains)) :| s"outside the keys: $listed" &&
        Prop(listed == listed.distinct.sorted) :| s"not ascending and distinct: $listed" &&
        Prop(math.min(size, limit + 1) == math.min(listed.size, limit + 1)) :|
          s"size $size, listed ${listed.size}, limit $limit" &&
          Prop(
            run.keep((0 until n).toArray, offset).toVector == inKeys.toVector
          ) :| "kept rows differ"
    }
  }

  // ---------- structure ----------

  test("an order index holds number and date rows by key, ascending rows within a key") {
    val s = slice(
      0,
      5,
      Vector(
        CellValue.Number(BigDecimal(3)),
        CellValue.Text("1"),
        CellValue.Number(BigDecimal("1.0")),
        CellValue.Bool(true),
        CellValue.Empty,
        CellValue.DateTime(LocalDateTime.of(1900, 1, 1, 0, 0)), // serial 1
        CellValue.Number(BigDecimal(1)),
        CellValue.Error(CellError.NA)
      )
    )
    val index = OrderIndex.build(s)
    // keys 1 and 3; rows 7, 10, 11 hold 1, row 5 holds 3
    assertEquals(index.keyCount, 2)
    assertEquals(index.keyedRows, 4)
    def listed(lower: Option[Bound], upper: Option[Bound]) =
      index.window(IndexLookup.Window(lower, upper)).list(0, 100).toVector
    assertEquals(listed(Some(bound("1", true)), Some(bound("1", true))), Vector(7, 10, 11))
    assertEquals(listed(None, None), Vector(5, 7, 10, 11))
    assertEquals(listed(Some(bound("2", true)), None), Vector(5))
    // a strict bound on a key's exact numbers skips the whole group
    assertEquals(listed(Some(bound("1", false)), None), Vector(5))
    assertEquals(listed(None, Some(bound("3", false))), Vector(7, 10, 11))
    // an empty interval (">=3" and "<1" merged) selects nothing
    assertEquals(listed(Some(bound("3", true)), Some(bound("1", false))), Vector.empty)
  }

  test("a key group holding several numbers is kept while any of them lies past the bound") {
    // 1 and 1 + 1E-22 share the order key 1.0
    val s = slice(
      0,
      0,
      Vector(
        CellValue.Number(BigDecimal(1)),
        CellValue.Number(BigDecimal("1.0000000000000000000001"))
      )
    )
    val index = OrderIndex.build(s)
    assertEquals(index.keyCount, 1)
    def listed(w: IndexLookup.Window) = index.window(w).list(0, 2).toVector
    assertEquals(listed(IndexLookup.Window(Some(bound("1", false)), None)), Vector(0, 1))
    assertEquals(
      listed(IndexLookup.Window(None, Some(bound("1.0000000000000000000001", false)))),
      Vector(0, 1)
    )
    assertEquals(listed(IndexLookup.Window(None, Some(bound("1", false)))), Vector.empty)
    assertEquals(
      listed(IndexLookup.Window(Some(bound("1.0000000000000000000001", false)), None)),
      Vector.empty
    )
  }

  test("windows on one range merge; windows on different ranges or first rows stay apart") {
    val a = slice(0, 0, Vector(CellValue.Number(BigDecimal(1))))
    val b = slice(1, 0, Vector(CellValue.Number(BigDecimal(1))))
    def q(s: String) = CriteriaMatcher.parse(ExprValue.Text(s))
    val merged = IndexLookup.of(List((a, 0, q(">=2")), (a, 0, q("<5")), (a, 0, q("x"))))
    assertEquals(
      merged.map((s, l, o) => (s.id, l, o)),
      List(
        (a.id, IndexLookup.Keys(List(MatchKey.Text("x"))), 0),
        (a.id, IndexLookup.Window(Some(bound("2", true)), Some(bound("5", false))), 0)
      )
    )
    val apart = IndexLookup.of(List((a, 0, q(">=2")), (b, 0, q("<5")), (a, 3, q("<1"))))
    assertEquals(apart.map(_._2).length, 3)
    // of two bounds at one value, the exclusive one is tighter
    assertEquals(
      IndexLookup
        .of(List((a, 0, q(">=2")), (a, 0, q(">2.0")), (a, 0, q("<=7")), (a, 0, q("<9"))))
        .map(_._2),
      List(IndexLookup.Window(Some(bound("2", false)), Some(bound("7", true))))
    )
  }

  // ---------- the pass's order index ----------

  private def datedSheet: com.tjclp.xl.sheets.Sheet =
    (0 until 600).foldLeft(com.tjclp.xl.sheets.Sheet(data)) { (s, r) =>
      s.put(ARef.from0(0, r), CellValue.Number(BigDecimal(r % 60)))
    }

  private def passCache(limits: PassCache.Limits, index: Boolean = true): PassCache =
    val cache = new PassCache(limits, RecalcCounters.disabled, index)
    cache.begin(Readiness.build(Map.empty, Set.empty))
    cache

  private def sliceOf(cache: PassCache): ColumnSlice =
    cache.slice(datedSheet, 0, 0, 599, EvalPosition.at(0)) match
      case PassCache.SliceAnswer.Use(s) => s
      case other => fail(other.toString)

  test("a window call is answered from one order index per slice, memoised per lookup") {
    val cache = passCache(PassCache.DefaultLimits)
    val s = sliceOf(cache)
    // values 0..59 ten times each: 10 <= v <= 14 is 50 rows
    val w = IndexLookup.Window(Some(bound("10", true)), Some(bound("14", true)))
    val first = cache.candidates(List((s, w, 0)), 600)
    assertEquals(first.map(_.length), Some(50))
    assert(cache.candidates(List((s, w, 0)), 600).zip(first).exists(_ eq _))
    assert(cache.orderIndex(s).zip(cache.orderIndex(s)).exists(_ eq _), "built once")
    // more than half the rows: every row is tested
    assertEquals(
      cache.candidates(
        List((s, IndexLookup.Window(Some(bound("0", true)), Some(bound("40", true))), 0)),
        600
      ),
      None
    )
    // with an equality pair on the same slice: the intersection
    val ten = CriteriaKeys.strip(BigDecimal(10)).toList
    val both = cache.candidates(List((s, IndexLookup.Keys(ten), 0), (s, w, 0)), 600)
    assertEquals(both.map(_.length), Some(10))
  }

  test("past maxIndexBytes, or with the index off, a window call gets no candidates") {
    val w = IndexLookup.Window(Some(bound("10", true)), Some(bound("14", true)))
    val tight = passCache(PassCache.DefaultLimits.copy(maxIndexBytes = 64L))
    assertEquals(tight.orderIndex(sliceOf(tight)), None)
    assertEquals(tight.candidates(List((sliceOf(tight), w, 0)), 600), None)
    val off = passCache(PassCache.DefaultLimits, index = false)
    assertEquals(off.candidates(List((sliceOf(off), w, 0)), 600), None)
    // the order index is charged against the same budget as the equality index
    val bytes = OrderIndex.build(sliceOf(passCache(PassCache.DefaultLimits))).bytes
    assert(bytes >= 8L * 600, bytes.toString)
  }

  test("a mixed lookup keeps the equality candidates when its order index exceeds the budget") {
    val full = passCache(PassCache.DefaultLimits)
    val eqBytes = full.eqIndex(sliceOf(full)).fold(fail("no equality index"))(_.bytes)
    val cache = passCache(PassCache.DefaultLimits.copy(maxIndexBytes = eqBytes + 1024L))
    val s = sliceOf(cache)
    val q = CriteriaMatcher.parse(ExprValue.Number(BigDecimal(1)))
    val eq = IndexLookup.Keys(CriteriaKeys.lookup(q).getOrElse(fail("no key")))
    val w = IndexLookup.Window(Some(bound("0", false)), None)
    val expected = cache.candidates(List((s, eq, 0)), 600)
    assertEquals(expected.map(_.length), Some(10))
    assertEquals(cache.orderIndex(s), None)
    assertEquals(cache.candidates(List((s, eq, 0), (s, w, 0)), 600), expected)
  }
