package com.tjclp.xl.formula.eval

import java.time.LocalDateTime

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.eval.RecalcCounters.Counter
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * #767 PR 2: answering criteria calls from the equality index changes no result.
 *
 * Each law runs one book three ways: slices with the index (every range sliced, thresholds at one
 * row), the same slices without the index, and no per-pass caches at all; it compares results cell
 * by cell (numbers by their exact `BigDecimal`, scale included) and error by error. The books key
 * criteria columns with case variants, numeric text, booleans, errors, date-times, rich text and
 * blanks, carry a helper formula column with wrong caches, and read them through windows with
 * different first rows, whole columns (blank tails), criteria arrays and empty-text criteria.
 * Criteria columns also hold numbers with extreme exponents (#780), and a law checks that no
 * criteria call fails on them.
 */
class CriteriaIndexLawSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(40)

  private val data = SheetName.unsafe("Data")
  private val other = SheetName.unsafe("Other")
  private val out = SheetName.unsafe("Out")
  private val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 2))

  private val Rows = 40
  private val Last = Rows + 1

  private val eager =
    PassCache.Limits(minSliceRows = 1, sliceRowsDivisor = Int.MaxValue, boxDemand = 1)

  private def indexOn(counters: RecalcCounters = RecalcCounters.disabled) =
    new Evaluator.AggregateMemo(counters = counters, limits = eager)

  private def indexOff(counters: RecalcCounters = RecalcCounters.disabled) =
    new Evaluator.AggregateMemo(counters = counters, limits = eager, criteriaIndex = false)

  private def cachesOff() = new Evaluator.AggregateMemo(columnCache = false)

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def at(a1: String): ARef = ARef.parse(a1).fold(e => fail(e.toString), identity)

  private def render(value: CellValue): String = value match
    case CellValue.Number(n) => s"N(${n.bigDecimal.toString})"
    case CellValue.Formula(text, cached, kind) => s"F($text,${cached.map(render)},$kind)"
    case other => other.toString

  private def cells(wb: Workbook): Vector[String] =
    wb.sheets.flatMap { sheet =>
      sheet.cells.toVector
        .sortBy((ref, _) => (ref.row.index0, ref.col.index0))
        .map((ref, cell) => s"${sheet.name.value}!${ref.toA1}=${render(cell.value)}")
    }

  private def exactlySame(a: RecalcResult, b: RecalcResult, label: String): Prop =
    val diff = cells(a.workbook).zip(cells(b.workbook)).filter(_ != _).take(3)
    Prop.all(
      Prop(diff.isEmpty && cells(a.workbook).size == cells(b.workbook).size) :|
        s"$label: cells differ $diff",
      Prop(a.errors == b.errors) :| s"$label: errors differ\n${a.errors}\n${b.errors}",
      Prop(a == b) :| s"$label: results differ"
    )

  // ---------- the keyed book ----------

  private val genKey: Gen[Option[CellValue]] = Gen.frequency(
    1 -> Gen.const(None),
    4 -> Gen
      .oneOf("a", "A", "b", "B", "1", " 1 ", "true", "İ", "i", "ß")
      .map(s => Some(CellValue.Text(s))),
    1 -> Gen.oneOf("a", "1").map(s => Some(CellValue.RichText(RichText.plain(s)))),
    3 -> Gen.oneOf("0", "1", "1.0", "2", "2.50").map(s => Some(CellValue.Number(BigDecimal(s)))),
    1 -> Gen.oneOf(true, false).map(b => Some(CellValue.Bool(b))),
    1 -> Gen.oneOf(CellError.NA, CellError.Div0).map(e => Some(CellValue.Error(e))),
    1 -> Gen.const(Some(CellValue.DateTime(LocalDateTime.of(1900, 1, 1, 0, 0)))),
    1 -> Gen.const(Some(CellValue.Text("")))
  )

  /**
   * #780: numbers with extreme exponents. `1E+999999999` parses as a `BigDecimal`; on
   * `100E+2147483647` `BigDecimal.stripTrailingZeros` throws (its stripped scale has no `Int`).
   */
  private val extremeNumbers: Vector[BigDecimal] = Vector(
    BigDecimal("1E+999999999"),
    BigDecimal(new java.math.BigDecimal(java.math.BigInteger.valueOf(100), Int.MinValue + 1))
  )

  /** A criteria-column key: [[genKey]], or a number with an extreme exponent. */
  private val genCriteriaKey: Gen[Option[CellValue]] = Gen.frequency(
    12 -> genKey,
    1 -> Gen.oneOf(extremeNumbers).map(n => Some(CellValue.Number(n)))
  )

  private val genCriterion: Gen[String] = Gen.oneOf(
    "\"a\"",
    "\"A\"",
    "\"i\"",
    "\"1\"",
    "1",
    "2.5",
    "TRUE",
    "\"true\"",
    "#N/A",
    "\"=1\"",
    "\"=a\"",
    "\"\"",
    "\"<>a\"",
    "\">0\"",
    "\"a*\"",
    "\"=a*\"",
    "\"=\"",
    "\"<>\"",
    "\"<>1\"",
    "\"<>*\"",
    "\"1E+999999999\"",
    "\"<>1E+999999999\"",
    "\"1E*\"",
    "DATE(1900,1,1)",
    "1+0*0",
    "Out!$Z$1",
    "Out!$Z$2"
  )

  /** A window of `rows` rows of `col` starting at `first`, sheet-qualified. */
  private def window(sheet: String, col: Char, first: Int, rows: Int): String =
    s"$sheet!$col$first:$col${first + rows - 1}"

  private val genReader: Gen[String] =
    for
      rows <- Gen.choose(5, Rows)
      fa <- Gen.choose(2, Last - rows + 1)
      fb <- Gen.choose(2, Last - rows + 1)
      fo <- Gen.choose(1, 10)
      c1 <- genCriterion
      c2 <- genCriterion
      c3 <- genCriterion
      a = window("Data", 'A', fa, rows)
      b = window("Data", 'B', fb, rows)
      d = window("Data", 'D', fa, rows)
      v = window("Data", 'C', fb, rows)
      o = window("Other", 'A', fo, rows)
      f <- Gen.oneOf(
        s"COUNTIF($a,$c1)",
        s"COUNTIF($d,$c1)",
        s"SUMIF($a,$c1,$v)",
        s"AVERAGEIF($b,$c1,$v)",
        s"COUNTIFS($a,$c1,$b,$c2)",
        s"COUNTIFS($a,$c1,$b,$c2,$o,$c3)",
        s"SUMIFS($v,$a,$c1,$b,$c2)",
        s"SUMIFS($v,$d,$c1,$o,$c2)",
        s"AVERAGEIFS($v,$a,$c1,$o,$c2)",
        s"MINIFS($v,$b,$c1,$a,$c2)",
        s"MAXIFS($v,$a,$c1,$d,$c3)",
        s"SUM(COUNTIFS($a,{\"a\",\"b\",1},$b,$c2))",
        s"SUM(SUMIFS($v,$b,{1,2,TRUE},$a,$c1))",
        // whole columns: the blank tail counts when every criterion matches a blank
        s"COUNTIFS(Data!A:A,$c1,Data!B:B,$c2)",
        s"COUNTIF(Data!A:A,$c1)",
        s"SUMIFS(Data!C:C,Data!A:A,$c1,Data!D:D,$c2)",
        // a whole-column defined name, and a dynamic criteria range beside a sliced one (the call
        // scans: an unsliced range may read an uncached formula on any row)
        s"COUNTIFS(acol,$c1,Data!B:B,$c2)",
        s"SUMIFS($v,OFFSET(Data!A$fa,0,0,$rows,1),$c1,$b,$c2)"
      )
    yield f

  private def genHelper(r: Int): Gen[CellValue] =
    for
      text <- Gen.oneOf(
        s"""IF(B$r=1,"a","b")""",
        s"""A$r&"" """.trim,
        s"B$r*1",
        s"""IF(ISNUMBER(B$r),B$r,"A")"""
      )
      cache <- Gen.oneOf(None, Some(CellValue.Text("a")), Some(num(1)), Some(num(-1)))
    yield CellValue.Formula(text, cache)

  private val genBook: Gen[Workbook] =
    for
      as <- Gen.listOfN(Rows, genCriteriaKey)
      bs <- Gen.listOfN(Rows, genCriteriaKey)
      cs <- Gen.listOfN(
        Rows,
        Gen.frequency(4 -> genKey, 3 -> Gen.choose(-3, 9).map(n => Some(num(n))))
      )
      helpers <- Gen.sequence[List[CellValue], CellValue]((2 to Last).map(genHelper))
      os <- Gen.listOfN(Rows + 10, genCriteriaKey)
      readers <- Gen.listOfN(24, genReader)
      z1 <- genCriteriaKey
      z2 <- genCriteriaKey
      readerCache <- Gen.oneOf(None, Some(num(-7)))
    yield
      def column(sheet: Sheet, col: Int, first: Int, values: List[Option[CellValue]]): Sheet =
        values.zipWithIndex.foldLeft(sheet) {
          case (s, (Some(v), i)) => s.put(ARef.from0(col, first - 1 + i), v)
          case (s, (None, _)) => s
        }
      val withData = column(column(column(Sheet(data), 0, 2, as), 1, 2, bs), 2, 2, cs)
      val dataSheet = helpers.zipWithIndex.foldLeft(withData) { case (s, (f, i)) =>
        s.put(ARef.from0(3, i + 1), f)
      }
      val otherSheet = column(Sheet(other), 0, 1, os)
      val withReaders = readers.zipWithIndex.foldLeft(Sheet(out)) { case (s, (f, i)) =>
        s.put(ARef.from0(0, i), CellValue.Formula(f, readerCache))
      }
      // criteria read from cells (Out!$Z$1, Out!$Z$2): any value, blank included
      val outSheet = Vector(z1, z2).zipWithIndex.foldLeft(withReaders) {
        case (s, (Some(v), i)) => s.put(ARef.from0(25, i), v)
        case (s, (None, _)) => s
      }
      Workbook(Vector(dataSheet, otherSheet, outSheet)).withDefinedName("acol", "Data!$A:$A")

  private val sequential =
    () => RecalcOptions(clock = clock, rng = Rng.seeded(5L), iterative = IterativeMode.Off)
  private val parallel =
    () => RecalcOptions(clock = clock, iterative = IterativeMode.Off, parallelism = 4)

  property("index law: recalculate with the index equals without it and without any cache") {
    Prop.forAll(genBook) { book =>
      Prop.all(
        exactlySame(
          recalculateWithMemo(book, sequential(), indexOn()),
          recalculateWithMemo(book, sequential(), indexOff()),
          "sequential on/off"
        ),
        exactlySame(
          recalculateWithMemo(book, sequential(), indexOn()),
          recalculateWithMemo(book, sequential(), cachesOff()),
          "sequential on/no caches"
        ),
        exactlySame(
          recalculateWithMemo(book, parallel(), indexOn()),
          recalculateWithMemo(book, parallel(), indexOff()),
          "parallel on/off"
        )
      )
    }
  }

  property("#780: no criteria call fails on a number with an extreme exponent") {
    Prop.forAll(genBook) { book =>
      val results = Vector(
        recalculateWithMemo(book, sequential(), indexOn()),
        recalculateWithMemo(book, sequential(), indexOff())
      )
      val defects = results
        .flatMap(r => r.errors.map(_.toString) ++ r.errorReasons.values)
        .filter(_.contains("internal evaluator defect"))
      Prop(defects.isEmpty) :| s"defects: ${defects.take(2)}"
    }
  }

  property("index law: recalculateUncached and after-edit with the index equal without it") {
    Prop.forAll(genBook, Gen.choose(2, Last), genKey) { (book, row, value) =>
      val ref = ARef.from0(0, row - 1)
      val edited = book
        .update(data, s => value.fold(s.remove(ref))(s.put(ref, _)))
        .fold(e => fail(e.toString), identity)
      Prop.all(
        exactlySame(
          recalculateUncachedWithMemo(book, sequential(), indexOn()),
          recalculateUncachedWithMemo(book, sequential(), indexOff()),
          "uncached"
        ),
        exactlySame(
          DependentRecalculation.recalculateAfterEdit(edited, data, Set(ref), clock, indexOn()),
          DependentRecalculation.recalculateAfterEdit(edited, data, Set(ref), clock, indexOff()),
          "after-edit"
        )
      )
    }
  }

  property("the index law is not vacuous: calls are indexed and test fewer rows") {
    Prop.forAll(genBook) { book =>
      val on = RecalcCounters.create()
      val off = RecalcCounters.create()
      recalculateWithMemo(book, sequential(), indexOn(on))
      recalculateWithMemo(book, sequential(), indexOff(off))
      val (a, b) = (on.snapshot, off.snapshot)
      Prop(a(Counter.CriteriaIndexed) > 0L) :| s"nothing indexed\n${a.render}" &&
      Prop(b(Counter.CriteriaIndexed) == 0L) :| "indexed with the index off" &&
      Prop(a(Counter.MatcherCalls) <= b(Counter.MatcherCalls)) :| "more matcher calls" &&
      Prop(
        a(Counter.CriteriaIndexed) + a(Counter.CriteriaScan) == b(Counter.CriteriaScan)
      ) :| "a call counted twice or not at all"
    }
  }

  // ---------- fixed cases ----------

  /** Data!A2:A1001 cycles ten lanes, B2:B1001 years 2020..2024, C2:C1001 = row number. */
  private val keyed: Workbook =
    val sheet = (2 to 1001).foldLeft(Sheet(data)) { (s, r) =>
      s.put(ARef.from0(0, r - 1), CellValue.Text(s"LN${r % 10}"))
        .put(ARef.from0(1, r - 1), num(2020 + r % 5))
        .put(ARef.from0(2, r - 1), num(r))
        .put(ARef.from0(3, r - 1), CellValue.Formula(s"""IF(C$r>500,"hi","lo")""", Some(num(-1))))
    }
    val formulas = Vector(
      """COUNTIFS(Data!A2:A1001,"ln3",Data!B2:B1001,2023)""",
      """SUMIFS(Data!C2:C1001,Data!A2:A1001,"LN3",Data!B2:B1001,2023)""",
      """SUMIFS(Data!C2:C1001,Data!B2:B1001,2023,Data!A2:A1001,"LN3")""",
      """MAXIFS(Data!C:C,Data!A:A,"LN3",Data!D:D,"lo")""",
      """COUNTIF(Data!A2:A1001,"LN3")""",
      """COUNTIFS(Data!A2:A1001,"LN3",Data!B2:B1001,">2022")""",
      """COUNTIFS(Data!B2:B1001,">2022")"""
    )
    val outSheet = formulas.zipWithIndex.foldLeft(Sheet(out)) { case (s, (f, i)) =>
      s.put(ARef.from0(0, i), CellValue.Formula(f))
    }
    Workbook(Vector(sheet, outSheet))

  test("the index tests only the candidate rows, and shares candidates between equal calls") {
    val counters = RecalcCounters.create()
    val result =
      recalculateWithMemo(keyed, sequential(), new Evaluator.AggregateMemo(counters = counters))
    assertEquals(result, recalculateWithMemo(keyed, sequential(), cachesOff()))
    val values = result.evaluated.getOrElse(out, Map.empty)
    // rows r = 3 (mod 10) and r = 3 (mod 5): r = 3 (mod 10), 100 rows; C = r
    val lane3 = (2 to 1001).filter(_ % 10 == 3)
    assertEquals(values.get(at("A1")), Some(num(lane3.size)))
    assertEquals(values.get(at("A2")), Some(num(lane3.sum)))
    assertEquals(values.get(at("A3")), Some(num(lane3.sum)))
    assertEquals(values.get(at("A4")), Some(num(lane3.filter(_ <= 500).foldLeft(0)(math.max))))
    assertEquals(values.get(at("A5")), Some(num(lane3.size)))
    val snap = counters.snapshot
    // six calls have an equality criterion; the last is window-only and scans
    assertEquals(snap(Counter.CriteriaIndexed), 6L)
    assertEquals(snap(Counter.CriteriaScan), 1L)
    assertEquals(snap(Counter.CriteriaWindowOnly), 1L)
    assertEquals(snap(Counter.ScanFallbacks), 0L)
    // indexed calls test their 100 candidates: one pair each for A1, A5 (lane bucket = the year
    // bucket's intersection), two for A2, A3, A4 (MAXIFS: D only on lane matches) and A6; the
    // window-only call tests all 1,000 rows
    assert(snap(Counter.MatcherCalls) <= 6L * 2 * 100 + 1000L, snap.render)
  }

  test("criteria that select more than half the rows test every row, through the slice") {
    val counters = RecalcCounters.create()
    val wb = keyed.update(
      out,
      _.put(at("B1"), CellValue.Formula("""COUNTIFS(Data!D2:D1001,"hi")"""))
    )
    val book = wb.fold(e => fail(e.toString), identity)
    val result =
      recalculateWithMemo(book, sequential(), new Evaluator.AggregateMemo(counters = counters))
    assertEquals(result.evaluated.get(out).flatMap(_.get(at("B1"))), Some(num(501)))
    // "hi" holds rows 501..1001: more than half, so B1 scans its 1,000 rows
    assertEquals(counters.snapshot(Counter.CriteriaIndexed), 6L)
    assertEquals(counters.snapshot(Counter.CriteriaScan), 2L)
  }

  test("an empty-text criterion stays on the matcher, blank tail included") {
    val sheet = (2 to 401).foldLeft(Sheet(data)) { (s, r) =>
      val withA = s.put(ARef.from0(0, r - 1), CellValue.Text(if r % 2 == 0 then "x" else "y"))
      if r % 3 == 0 then withA else withA.put(ARef.from0(1, r - 1), num(r))
    }
    val outSheet = Sheet(out)
      .put(at("A1"), CellValue.Formula("""COUNTIFS(Data!A:A,"x",Data!B:B,"")"""))
      .put(at("A2"), CellValue.Formula("""COUNTIFS(Data!B:B,"")"""))
      .put(at("A3"), CellValue.Formula("""COUNTIF(Data!B2:B401,"")"""))
    val wb = Workbook(Vector(sheet, outSheet))
    val counters = RecalcCounters.create()
    val result =
      recalculateWithMemo(wb, sequential(), new Evaluator.AggregateMemo(counters = counters))
    assertEquals(result, recalculateWithMemo(wb, sequential(), cachesOff()))
    val values = result.evaluated.getOrElse(out, Map.empty)
    val blankB = (2 to 401).count(_ % 3 == 0)
    assertEquals(values.get(at("A1")), Some(num((2 to 401).count(r => r % 6 == 0))))
    // A:A and B:B bound to rows 2..401; B's blank tail is row 1 and rows 402..1048576
    assertEquals(values.get(at("A2")), Some(num(blankB + 1 + (1048576 - 401))))
    assertEquals(values.get(at("A3")), Some(num(blankB)))
    // only A1 has an indexed criterion
    assertEquals(counters.snapshot(Counter.CriteriaIndexed), 1L)
    assertEquals(counters.snapshot(Counter.CriteriaNotIndexable), 2L)
  }

  test("a tiny index budget scans instead, with the same results") {
    val counters = RecalcCounters.create()
    val memo = new Evaluator.AggregateMemo(
      counters = counters,
      limits = PassCache.DefaultLimits.copy(maxIndexBytes = 16L)
    )
    val result = recalculateWithMemo(keyed, sequential(), memo)
    assertEquals(result, recalculateWithMemo(keyed, sequential(), cachesOff()))
    assertEquals(counters.snapshot(Counter.CriteriaIndexed), 0L)
    assertEquals(counters.snapshot(Counter.CriteriaScan), 7L)
  }
