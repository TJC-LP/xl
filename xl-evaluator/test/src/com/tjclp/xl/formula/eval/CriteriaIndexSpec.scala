package com.tjclp.xl.formula.eval

import java.time.{LocalDate, LocalDateTime}

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.ExprValue
import com.tjclp.xl.richtext.RichText
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * #767 PR 2: the criteria index is a filter, never a judge.
 *
 *   - Completeness: whenever `CriteriaMatcher.matches(v, q)` holds for a criterion the index looks
 *     up, the value's keys meet the criterion's lookup keys, or the value is residual. Checked over
 *     the full cross product of edge-biased values and criteria, and on random ones.
 *   - `fold` agrees with `String.equalsIgnoreCase` on every BMP character and its case orbit, in
 *     the Latin-1, UTF-16 and mixed string representations.
 *   - `parseNumeric`'s leading-character check refuses only what the parse refuses.
 *   - Candidate offsets equal a brute-force model of the index's selection.
 *
 * Seeds follow the law-fuzz convention (GH-308): fixed by default, random under
 * `XL_ROUNDTRIP_RANDOM_SEED=true`, replayable with `-Dxl.roundtrip.seed=<printed seed>`.
 */
class CriteriaIndexSpec extends ScalaCheckSuite:

  private val minSuccess: Int =
    sys.props
      .get("xl.roundtrip.minSuccess")
      .orElse(sys.env.get("XL_ROUNDTRIP_MIN_SUCCESS"))
      .flatMap(_.toIntOption)
      .getOrElse(500)

  private val explicitSeed: Option[org.scalacheck.rng.Seed] =
    sys.props.get("xl.roundtrip.seed").flatMap { raw =>
      raw.toLongOption
        .map(org.scalacheck.rng.Seed.apply)
        .orElse(org.scalacheck.rng.Seed.fromBase64(raw).toOption)
    }

  private val randomSeed: Boolean =
    explicitSeed.isEmpty &&
      sys.env.get("XL_ROUNDTRIP_RANDOM_SEED").exists(_.equalsIgnoreCase("true"))

  override def scalaCheckTestParameters: org.scalacheck.Test.Parameters =
    val base = super.scalaCheckTestParameters.withMinSuccessfulTests(minSuccess)
    if randomSeed then
      val seed = org.scalacheck.rng.Seed.random()
      println(s"[law-fuzz] random seed for this run: ${seed.toBase64} (#767 criteria index)")
      base.withInitialSeed(seed)
    else base.withInitialSeed(explicitSeed.getOrElse(org.scalacheck.rng.Seed(20261014L)))

  // ---------- edge-biased pools ----------

  /** Case pairs, numeric text, booleans, wildcards, tildes, error spellings and surrogates. */
  private val texts: Vector[String] = Vector(
    "ABC",
    "abc",
    "aBc",
    "İ",
    "i",
    "I",
    "ı",
    "ß",
    "SS",
    "ss",
    "ẞ",
    "ς",
    "σ",
    "Σ",
    "K", // Kelvin sign
    "k",
    "K",
    "ſ", // long s
    "s",
    "S",
    "µ", // micro sign
    "μ",
    "Μ",
    "ÿ",
    "Ÿ",
    "Ꭰ", // Cherokee
    "ꭰ",
    "ǅ",
    "Ǆ",
    "ǆ",
    "𐐀", // Deseret, a surrogate pair
    "𐐨",
    "a𐐀",
    "\uD800", // a lone surrogate
    "042",
    "42",
    " 42 ",
    "4.20",
    "4.2",
    "1e2",
    "100",
    "-0",
    "0",
    "0.0",
    "1",
    "1.0",
    "+5",
    ".5",
    "5.",
    "٣", // Arabic-Indic three
    "１２", // full-width 12
    "NaN",
    "Infinity",
    "TRUE",
    "true",
    "True",
    "FALSE",
    "false",
    "x",
    "X",
    "a*b",
    "a?c",
    "~*",
    "a~*b",
    "~~",
    "*",
    "?",
    "#N/A",
    "#n/a",
    "#DIV/0!",
    "=x",
    "=42",
    "=",
    "<>x",
    ">5",
    "",
    " ",
    "2025-01-02",
    "2025-01-02T00:00",
    // numbers whose stripped form has no Int scale: no number key, residual
    "100E+2147483647",
    "1E+2147483647",
    "10E-2147483647"
  )

  private val numbers: Vector[BigDecimal] =
    Vector("0", "0.0", "1", "1.0", "42", "42.00", "4.2", "4.20", "100", "1E+2", "-1", "0.5")
      .map(BigDecimal(_)) ++ Vector(BigDecimal(45659), BigDecimal("45659.5"))

  private val dateTimes: Vector[LocalDateTime] =
    Vector(LocalDateTime.of(2025, 1, 2, 0, 0), LocalDateTime.of(2025, 1, 2, 12, 0))

  private val errors: Vector[CellError] = CellError.values.toVector

  private val plainValues: Vector[CellValue] =
    Vector[CellValue](CellValue.Empty) ++
      texts.map(CellValue.Text(_)) ++
      texts.map(s => CellValue.RichText(RichText.plain(s))) ++
      numbers.map(CellValue.Number(_)) ++
      Vector(true, false).map(CellValue.Bool(_)) ++
      dateTimes.map(CellValue.DateTime(_)) ++
      errors.map(CellValue.Error(_))

  /**
   * A number whose stripped form has no `Int` scale. Kept out of the pools that are deduplicated:
   * Scala's `BigDecimal.hashCode` throws for it.
   */
  private val hugeNumber: CellValue = CellValue.Number(BigDecimal("100E+2147483647"))

  private val values: Vector[CellValue] =
    plainValues ++
      plainValues.filter(_ != CellValue.Empty).map(v => CellValue.Formula("A1", Some(v))) ++
      Vector(hugeNumber, CellValue.Formula("A1", Some(hugeNumber)), CellValue.Formula("A1", None))

  private val criterionValues: Vector[ExprValue] =
    texts.map(ExprValue.Text(_)) ++
      numbers.map(ExprValue.Number(_)) ++
      Vector(true, false).map(ExprValue.Bool(_)) ++
      Vector(ExprValue.Date(LocalDate.of(2025, 1, 2))) ++
      dateTimes.map(ExprValue.DateTime(_)) ++
      plainValues.map(ExprValue.Cell(_)) ++
      Vector(
        ExprValue.Cell(CellValue.Formula("A1", None)),
        ExprValue.Cell(CellValue.Formula("A1", Some(CellValue.Text("x")))),
        ExprValue.Opaque(List(1))
      )

  private val criteria: Vector[CriteriaMatcher.Criterion] =
    criterionValues.map(CriteriaMatcher.parse).distinct

  private def keysOf(keys: CriteriaKeys.CellKeys): Option[Set[MatchKey]] =
    keys match
      case CriteriaKeys.CellKeys.NoKeys => Some(Set.empty)
      case CriteriaKeys.CellKeys.Residual => None
      case CriteriaKeys.CellKeys.One(k) => Some(Set(k))
      case CriteriaKeys.CellKeys.Two(a, b) => Some(Set(a, b))

  /** The completeness lemma for one pair; `None` when it holds. */
  private def violation(v: CellValue, q: CriteriaMatcher.Criterion): Option[String] =
    CriteriaKeys.lookup(q) match
      case None => None
      case Some(lookupKeys) =>
        if !CriteriaMatcher.matches(v, q) then None
        else
          keysOf(CriteriaKeys.cellKeys(v)) match
            case None => None
            case Some(cellKeys) if cellKeys.exists(lookupKeys.contains) => None
            case Some(cellKeys) => Some(s"$v matches $q, keys $cellKeys miss $lookupKeys")

  // ---------- completeness ----------

  test("completeness: every edge-biased value and criterion (full cross product)") {
    val bad = for
      v <- values
      q <- criteria
      problem <- violation(v, q)
    yield problem
    assertEquals(bad.take(10), Vector.empty)
    // not vacuous: most criteria are indexed, and the index sees many matches
    val indexed = criteria.count(CriteriaKeys.lookup(_).isDefined)
    assert(indexed > criteria.size / 2, s"$indexed of ${criteria.size} indexed")
    val matches = for
      v <- values
      q <- criteria if CriteriaKeys.lookup(q).isDefined && CriteriaMatcher.matches(v, q)
    yield ()
    assert(matches.size > 500, s"${matches.size} matches")
  }

  test("no indexed criterion matches Empty, so rows outside a slice are never candidates") {
    criteria.filter(CriteriaKeys.lookup(_).isDefined).foreach { q =>
      assert(!CriteriaMatcher.matches(CellValue.Empty, q), q.toString)
    }
  }

  test("what is not indexed: empty text, comparisons, wildcards, negations, other cells") {
    def lookup(v: ExprValue) = CriteriaKeys.lookup(CriteriaMatcher.parse(v))
    Vector("", "=", ">5", "<=x", "<>x", "<>", "<>4", "a*", "=a*", "?", "<>#N/A", "𐐀").foreach {
      s =>
        assertEquals(lookup(ExprValue.Text(s)), None, s)
    }
    assertEquals(lookup(ExprValue.Cell(CellValue.Empty)), None)
    assertEquals(lookup(ExprValue.Cell(CellValue.Formula("A1", None))), None)
    assertEquals(lookup(ExprValue.Opaque(1)), None)
    // a date matches only date-time cells, every one of which is residual
    assertEquals(lookup(ExprValue.Date(LocalDate.of(2025, 1, 2))), Some(Nil))
    // tilde escapes are settled by the parse: "~*" is the text "*"
    assertEquals(lookup(ExprValue.Text("~*")), Some(List(MatchKey.Text("*"))))
  }

  private val genText: Gen[String] = Gen.oneOf(
    Gen.oneOf(texts),
    Gen.listOfN(3, Gen.oneOf(texts)).map(_.mkString),
    Gen.asciiPrintableStr,
    Gen.stringOf(Gen.oneOf(Gen.choose('\u0000', '˿'), Gen.choose('Ͱ', 'Ͽ')))
  )

  private val genValue: Gen[CellValue] = Gen.oneOf(
    Gen.oneOf(values),
    genText.map(CellValue.Text(_)),
    genText.map(s => CellValue.RichText(RichText.plain(s))),
    Gen.choose(-5000L, 5000L).map(n => CellValue.Number(BigDecimal(n) / 100))
  )

  private val genCriterion: Gen[CriteriaMatcher.Criterion] = Gen
    .oneOf(
      Gen.oneOf(criterionValues),
      genText.map(ExprValue.Text(_)),
      Gen.choose(-5000L, 5000L).map(n => ExprValue.Number(BigDecimal(n) / 100))
    )
    .map(CriteriaMatcher.parse)

  property("completeness: random values and criteria") {
    Prop.forAll(genValue, genCriterion) { (v, q) =>
      val problem = violation(v, q)
      Prop(problem.isEmpty) :| problem.getOrElse("")
    }
  }

  test("a number whose stripped form has no Int scale gets no key and never throws") {
    val huge = BigDecimal("100E+2147483647")
    assertEquals(CriteriaKeys.strip(huge), None)
    assertEquals(CriteriaKeys.cellKeys(CellValue.Number(huge)), CriteriaKeys.CellKeys.Residual)
    assertEquals(
      CriteriaKeys.cellKeys(CellValue.Text("100E+2147483647")),
      CriteriaKeys.CellKeys.Residual
    )
    assertEquals(CriteriaKeys.lookup(CriteriaMatcher.Exact(ExprValue.Number(huge))), None)
    assertEquals(
      CriteriaKeys.lookup(CriteriaMatcher.parse(ExprValue.Text("100E+2147483647"))),
      None
    )
    // the other end of the scale strips normally
    assertEquals(
      CriteriaKeys.strip(BigDecimal("10E-2147483647")),
      Some(MatchKey.Num(new java.math.BigDecimal("1E-2147483646")))
    )
    assertEquals(CriteriaKeys.strip(BigDecimal("4.20")), CriteriaKeys.strip(BigDecimal("4.2")))
    assertEquals(
      CriteriaKeys.strip(BigDecimal("0.000")),
      Some(MatchKey.Num(java.math.BigDecimal.ZERO))
    )
  }

  test("#780: numberText is total, canonical, and plain for ordinary numbers") {
    def text(s: String) = CriteriaKeys.numberText(BigDecimal(s))
    assertEquals(text("4.20"), "4.2")
    assertEquals(text("4.2"), "4.2")
    assertEquals(text("1E+3"), "1000")
    assertEquals(text("1000.00"), "1000")
    assertEquals(text("-0.50"), "-0.5")
    assertEquals(text("0.000"), "0")
    assertEquals(text("1E+300"), "1" + "0" * 300)
    assertEquals(text("1E-300"), "0." + "0" * 299 + "1")
    // past ±400 places: scientific, never a string of up to two billion characters
    assertEquals(text("1E+999999999"), "1E+999999999")
    assertEquals(text("12.5E+999999999"), "1.25E+1000000000")
    assertEquals(text("1E-999999999"), "1E-999999999")
    // stripTrailingZeros throws here; value-equal forms still give one text
    val huge =
      BigDecimal(new java.math.BigDecimal(java.math.BigInteger.valueOf(100), Int.MinValue + 1))
    val sameValue =
      BigDecimal(new java.math.BigDecimal(java.math.BigInteger.valueOf(1000), Int.MinValue + 2))
    assertEquals(CriteriaKeys.numberText(huge), "1E+2147483649")
    assertEquals(CriteriaKeys.numberText(sameValue), "1E+2147483649")
    assertEquals(CriteriaKeys.numberText(-huge), "-1E+2147483649")
    assertEquals(text("100E+2147483647"), "1E+2147483649")
  }

  // ---------- the index as a filter, end to end ----------

  private val genAnyColumn: Gen[Vector[CellValue]] =
    Gen
      .choose(0, 60)
      .flatMap(n => Gen.listOfN(n, genValue.filter(_ != CellValue.Formula("A1", None))))
      .map(_.toVector)

  property("filter law: every offset whose indexed criteria all match is a candidate") {
    Prop.forAll(
      Gen
        .choose(1, 3)
        .flatMap(k =>
          Gen.listOfN(k, Gen.zip(Gen.choose(0, 10), genAnyColumn, Gen.choose(0, 15), genCriterion))
        ),
      Gen.choose(1, 60)
    ) { (columns, n) =>
      val pairs = columns.zipWithIndex.flatMap { case ((r0, values, offset, q), i) =>
        CriteriaKeys
          .lookup(q)
          .map(keys =>
            (new ColumnSlice(data, i, 0, 1048575, r0, IArray.from(values), 0L), keys, offset, q)
          )
      }
      val result = CriteriaCandidates.compute(
        pairs.map((s, keys, offset, _) => CriteriaCandidates.Pair(EqIndex.build(s), keys, offset)),
        n
      )
      val matching = (0 until n).filter(i =>
        pairs.forall((s, _, offset, q) => CriteriaMatcher.matches(s.valueAt(offset + i), q))
      )
      result match
        case Some(offsets) =>
          val listed = offsets.toVector
          Prop(
            matching.forall(listed.contains)
          ) :| s"missing ${matching.filterNot(listed.contains)}" &&
          Prop(listed == listed.distinct.sorted) :| s"not ascending and distinct: $listed" &&
          Prop(listed.forall(i => i >= 0 && i < n)) :| s"outside 0 until $n: $listed"
        case None => Prop.passed
    }
  }

  // ---------- fold ----------

  test("fold agrees with equalsIgnoreCase on every BMP character and its case orbit") {
    def orbit(c: Char): Vector[Char] =
      val up = Character.toUpperCase(c)
      val low = Character.toLowerCase(c)
      Vector(c, up, low, Character.toLowerCase(up), Character.toUpperCase(low))
        .map(_.toChar)
        .distinct
        .filterNot(Character.isSurrogate)
    // the suffix moves a Latin-1 string into the UTF-16 representation
    val utf16 = "Ā"
    val bad = (0 to 0xffff).iterator
      .map(_.toChar)
      .filterNot(Character.isSurrogate)
      .flatMap(c => orbit(c).map(d => (c, d)))
      .filter { (c, d) =>
        val same = CriteriaKeys.fold(c.toString) == CriteriaKeys.fold(d.toString)
        // Latin-1 or mixed (one-character strings) and UTF-16 (both suffixed)
        c.toString.equalsIgnoreCase(d.toString) != same ||
        (c.toString + utf16).equalsIgnoreCase(d.toString + utf16) != same ||
        ("a" + c).equalsIgnoreCase("A" + d) != same
      }
      .take(10)
      .toVector
    assertEquals(bad.map((c, d) => (c.toInt.toHexString, d.toInt.toHexString)), Vector.empty)
  }

  property("fold: equalsIgnoreCase implies equal folds on surrogate-free text") {
    Prop.forAll(genText, genText) { (a, b) =>
      val clean = !CriteriaKeys.hasSurrogate(a) && !CriteriaKeys.hasSurrogate(b)
      Prop(!clean || !a.equalsIgnoreCase(b) || CriteriaKeys.fold(a) == CriteriaKeys.fold(b))
    }
  }

  test("fold returns the same instance for folded text") {
    val s = "already folded 42"
    assert(CriteriaKeys.fold(s) eq s)
    assertEquals(CriteriaKeys.fold("MiXeD"), "mixed")
  }

  // ---------- parseNumeric ----------

  private val genNumericish: Gen[String] =
    val pieces = Vector(
      "",
      " ",
      "\t",
      " ",
      "+",
      "-",
      ".",
      "e",
      "E",
      "1",
      "0",
      "9",
      "42",
      "٣",
      "１",
      "x",
      "NaN",
      "Infinity",
      "_",
      "e5",
      "$"
    )
    Gen.oneOf(
      Gen.listOf(Gen.oneOf(pieces)).map(_.mkString),
      Gen.asciiPrintableStr,
      Gen.oneOf(texts),
      Gen.stringOf(Gen.choose('\u0000', '￿'))
    )

  property("parseNumeric's leading-character check equals the parse alone") {
    Prop.forAll(genNumericish) { s =>
      Prop(CriteriaMatcher.parseNumeric(s) == CriteriaMatcher.parseNumericByParse(s)) :| s
    }
  }

  test("parseNumeric's check keeps every accepted form") {
    Vector("42", " 42 ", "+5", "-0", ".5", "5.", "1e2", "٣", "１２").foreach { s =>
      assertEquals(CriteriaMatcher.parseNumeric(s), CriteriaMatcher.parseNumericByParse(s), s)
      assert(CriteriaMatcher.parseNumeric(s).isDefined, s)
    }
  }

  // ---------- candidates ----------

  private val data = SheetName.unsafe("Data")

  private val genSliceValue: Gen[CellValue] = Gen.frequency(
    3 -> Gen.choose(0, 5).map(n => CellValue.Number(BigDecimal(n))),
    3 -> Gen.oneOf("a", "A", "b", "1", "TRUE", "x").map(CellValue.Text(_)),
    1 -> Gen.oneOf(true, false).map(CellValue.Bool(_)),
    1 -> Gen.const(CellValue.Empty),
    1 -> Gen.const(CellValue.DateTime(LocalDateTime.of(1900, 1, 1, 0, 0))),
    1 -> Gen.const(CellValue.Error(CellError.NA))
  )

  /** A slice of rows `r0 until r0 + values.length`; its domain is the whole column. */
  private def slice(col: Int, r0: Int, values: Vector[CellValue]): ColumnSlice =
    new ColumnSlice(data, col, 0, 1048575, r0, IArray.from(values), 0L)

  private val genCriterionForSlice: Gen[CriteriaMatcher.Criterion] =
    Gen
      .oneOf[ExprValue](
        ExprValue.Number(BigDecimal(1)),
        ExprValue.Number(BigDecimal(3)),
        ExprValue.Text("a"),
        ExprValue.Text("1"),
        ExprValue.Text("true"),
        ExprValue.Bool(true),
        ExprValue.Text("#N/A"),
        ExprValue.Text("zzz"),
        ExprValue.Date(LocalDate.of(1900, 1, 1))
      )
      .map(CriteriaMatcher.parse)

  private final case class Column(r0: Int, values: Vector[CellValue], offset: Int)

  private val genColumn: Gen[Column] =
    for
      r0 <- Gen.choose(0, 20)
      len <- Gen.choose(0, 80)
      values <- Gen.listOfN(len, genSliceValue)
      offset <- Gen.choose(0, 40)
    yield Column(r0, values.toVector, offset)

  property("candidates equal the brute-force selection of every indexed pair") {
    Prop.forAll(
      Gen.choose(1, 3).flatMap(k => Gen.listOfN(k, Gen.zip(genColumn, genCriterionForSlice))),
      Gen.choose(1, 70)
    ) { (pairs, n) =>
      val built = pairs.zipWithIndex.map { case ((c, q), i) =>
        val s = slice(i, c.r0, c.values)
        (s, c, q, CriteriaKeys.lookup(q).getOrElse(Nil))
      }
      val result = CriteriaCandidates.compute(
        built.map((s, c, _, keys) => CriteriaCandidates.Pair(EqIndex.build(s), keys, c.offset)),
        n
      )
      def selected(s: ColumnSlice, keys: List[MatchKey], offset: Int, i: Int): Boolean =
        CriteriaKeys.cellKeys(s.valueAt(offset + i)) match
          case CriteriaKeys.CellKeys.Residual => true
          case other => keysOf(other).exists(_.exists(keys.contains))
      val model =
        (0 until n).filter(i => built.forall((s, c, _, keys) => selected(s, keys, c.offset, i)))
      val matching = (0 until n).filter(i =>
        built.forall((s, c, q, _) => CriteriaMatcher.matches(s.valueAt(c.offset + i), q))
      )
      val smallest = built.map((s, c, _, keys) => (0 until n).count(selected(s, keys, c.offset, _)))
      result match
        case Some(offsets) =>
          Prop(offsets.toVector == model.toVector) :| s"$offsets != $model" &&
          Prop(matching.forall(offsets.contains)) :| "a matching offset is missing"
        // the call tests every offset: allowed only when listing would not pay
        case None =>
          Prop(smallest.forall(_ > n / 2)) :| s"no candidates although $smallest <= ${n / 2}"
    }
  }

  test("an index holds its keys in CSR form, residual rows apart") {
    val s = slice(
      0,
      5,
      Vector(
        CellValue.Text("A"),
        CellValue.Number(BigDecimal("1.0")),
        CellValue.Text("1"),
        CellValue.Bool(true),
        CellValue.Empty,
        CellValue.DateTime(LocalDateTime.of(2000, 1, 1, 0, 0)),
        CellValue.Text("a")
      )
    )
    val index = EqIndex.build(s)
    // Text(a), Num(1), Text(1), Bool(true)
    assertEquals(index.keyCount, 4)
    // a: rows 5 and 11; 1: rows 6, 7 (text "1" also as a number) and 8 (TRUE as 1); "1": row 7;
    // TRUE: row 8
    assertEquals(index.keyedRows, 2 + 3 + 1 + 1)
    assertEquals(index.residualRows, 1)
    val offsets = CriteriaCandidates.compute(
      List(CriteriaCandidates.Pair(index, List(MatchKey.Text("a")), 0)),
      100
    )
    // rows 5 and 11 hold "a", row 10 is residual
    assertEquals(offsets.map(_.toVector), Some(Vector(5, 10, 11)))
  }

  // ---------- the pass's index and candidate memo ----------

  private def keyedSheet: com.tjclp.xl.sheets.Sheet =
    (0 until 600).foldLeft(com.tjclp.xl.sheets.Sheet(data)) { (s, r) =>
      s.put(com.tjclp.xl.addressing.ARef.from0(0, r), CellValue.Text(s"k${r % 30}"))
    }

  private def passCache(limits: PassCache.Limits, index: Boolean = true): PassCache =
    val cache = new PassCache(limits, RecalcCounters.disabled, index)
    cache.begin(Readiness.build(Map.empty, Set.empty))
    cache

  private def sliceOf(cache: PassCache): ColumnSlice =
    cache.slice(keyedSheet, 0, 0, 599, EvalPosition.at(0)) match
      case PassCache.SliceAnswer.Use(s) => s
      case other => fail(other.toString)

  test("equal calls share one memoised candidate list; other windows get their own") {
    val cache = passCache(PassCache.DefaultLimits)
    val s = sliceOf(cache)
    val keys = List(MatchKey.Text("k7"))
    val first = cache.candidates(List((s, keys, 0)), 600)
    assertEquals(first.map(_.length), Some(20))
    assert(cache.candidates(List((s, keys, 0)), 600).zip(first).exists(_ eq _))
    // the same bucket under another window is another list
    assertEquals(cache.candidates(List((s, keys, 100)), 300).map(_.length), Some(10))
    assert(cache.eqIndex(s).zip(cache.eqIndex(s)).exists(_ eq _), "the index is built once")
  }

  test("past maxIndexBytes, or with the index off, a call gets no candidates") {
    val keys = List(MatchKey.Text("k7"))
    val tight = passCache(PassCache.DefaultLimits.copy(maxIndexBytes = 64L))
    assertEquals(tight.eqIndex(sliceOf(tight)), None)
    assertEquals(tight.candidates(List((sliceOf(tight), keys, 0)), 600), None)
    val off = passCache(PassCache.DefaultLimits, index = false)
    assert(!off.indexing)
    assertEquals(off.candidates(List((sliceOf(off), keys, 0)), 600), None)
    // a budget that holds the index but not the memo: candidates are computed every time
    val indexBytes = EqIndex.build(sliceOf(passCache(PassCache.DefaultLimits))).bytes
    // a candidate list of 20 rows is charged 144 bytes
    val noMemo = passCache(PassCache.DefaultLimits.copy(maxIndexBytes = indexBytes + 100L))
    val s = sliceOf(noMemo)
    assert(noMemo.eqIndex(s).isDefined, "the index fits")
    val again = noMemo.candidates(List((s, keys, 0)), 600)
    assertEquals(again.map(_.length), Some(20))
    assert(!noMemo.candidates(List((s, keys, 0)), 600).zip(again).exists(_ eq _))
  }

  test("memoised candidate lists never use up the budget an index needs") {
    val sheet = (0 until 600).foldLeft(com.tjclp.xl.sheets.Sheet(data)) { (s, r) =>
      s.put(com.tjclp.xl.addressing.ARef.from0(0, r), CellValue.Text(s"k${r % 30}"))
        .put(com.tjclp.xl.addressing.ARef.from0(1, r), CellValue.Text(s"j${r % 30}"))
    }
    def sliceAt(cache: PassCache, col: Int): ColumnSlice =
      cache.slice(sheet, col, 0, 599, EvalPosition.at(0)) match
        case PassCache.SliceAnswer.Use(s) => s
        case other => fail(other.toString)
    val probe = passCache(PassCache.DefaultLimits)
    val bytesA = probe.eqIndex(sliceAt(probe, 0)).fold(0L)(_.bytes)
    val bytesB = probe.eqIndex(sliceAt(probe, 1)).fold(0L)(_.bytes)
    // room for both indexes twice over: lists may take half of it, never the second index's share
    val cache = passCache(PassCache.DefaultLimits.copy(maxIndexBytes = 2 * (bytesA + bytesB)))
    val a = sliceAt(cache, 0)
    val keys = List(MatchKey.Text("k7"))
    // distinct windows of column A, each its own memoised list, well past the budget
    (0 until 300).foreach(offset =>
      assert(cache.candidates(List((a, keys, offset)), 300).isDefined)
    )
    assert(cache.eqIndex(sliceAt(cache, 1)).isDefined, "the second column's index was refused")
  }
