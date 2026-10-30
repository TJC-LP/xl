package com.tjclp.xl.formula.eval

import java.time.LocalDateTime

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.eval.RecalcCounters.Counter
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * #767 PR 1: the per-pass caches (used ranges, static names, column slices) never change a result.
 *
 * Every law runs one book twice, with the caches on (slices at every height, so the small books
 * here use them) and off, and compares the two results cell by cell — numbers by their exact
 * `BigDecimal`, scale included — and error by error. The books carry wrong caches on every formula
 * (a slice built before its column is recomputed would read them), windowed and whole-column
 * readers of the same formula columns, a dynamic-bucket member and a cycle (stripped cells a slice
 * must refuse), and run sequentially, in parallel waves, iteratively, uncached and after an edit.
 */
class PassCacheLawSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(40)

  private val data = SheetName.unsafe("Data")
  private val other = SheetName.unsafe("Other")
  private val out = SheetName.unsafe("Out")
  private val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 2))

  /** Data rows 2..Rows+1; 40 helper formulas make a wave wide enough for the parallel path. */
  private val Rows = 40
  private val Last = Rows + 1

  /** Slices at every height (the threshold is 1 row), and every box on its first request. */
  private val eager =
    PassCache.Limits(minSliceRows = 1, sliceRowsDivisor = Int.MaxValue, boxDemand = 1)

  private def cachesOn(counters: RecalcCounters = RecalcCounters.disabled) =
    new Evaluator.AggregateMemo(counters = counters, limits = eager)

  private def cachesOff() = new Evaluator.AggregateMemo(columnCache = false)

  /** Every per-pass cache on except the reuse of the graph build's formula parses (#790). */
  private def parsesOff(counters: RecalcCounters = RecalcCounters.disabled) =
    new Evaluator.AggregateMemo(counters = counters, limits = eager, parseCache = false)

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def at(a1: String): ARef = ARef.parse(a1).fold(e => fail(e.toString), identity)

  // ---------- exact comparison ----------

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

  private def exactlySame(on: RecalcResult, off: RecalcResult, label: String): Prop =
    val diff = cells(on.workbook).zip(cells(off.workbook)).filter(_ != _).take(3)
    Prop.all(
      Prop(diff.isEmpty && cells(on.workbook).size == cells(off.workbook).size) :|
        s"$label: cells differ $diff",
      Prop(on.errors == off.errors) :| s"$label: errors differ\n${on.errors}\n${off.errors}",
      Prop(on.evaluated == off.evaluated) :| s"$label: evaluated maps differ",
      Prop(on == off) :| s"$label: results differ"
    )

  // ---------- the stale-cache book ----------

  private val genConstant: Gen[Option[CellValue]] = Gen.frequency(
    1 -> Gen.const(None),
    4 -> Gen.choose(-2, 9).map(n => Some(num(n))),
    1 -> Gen.oneOf("x", "hi", "lo", "7", " 4 ").map(s => Some(CellValue.Text(s))),
    1 -> Gen.oneOf(true, false).map(b => Some(CellValue.Bool(b))),
    1 -> Gen.const(Some(CellValue.Error(CellError.NA)))
  )

  /** A helper formula for Data!F{r}, as written, with a wrong cache or none. */
  private def genHelper(r: Int): Gen[CellValue] =
    for
      text <- Gen.oneOf(
        s"A$r*2",
        s"""IF(ISNUMBER(A$r),A$r+1,"hi")""",
        s"""A$r&"" """.trim,
        s"1/(A$r-3)",
        s"""IF(A$r>4,"hi","lo")"""
      )
      cache <- Gen.oneOf(
        None,
        Some(num(-1)),
        Some(CellValue.Text("zz")),
        Some(CellValue.Error(CellError.Ref))
      )
    yield CellValue.Formula(text, cache)

  private val genCriterion: Gen[String] = Gen.oneOf(
    "\">2\"",
    "\"hi\"",
    "\"<>lo\"",
    "#DIV/0!",
    "4",
    "\"*i\"",
    "-1",
    "\"zz\"",
    "\"\"",
    "\"<5\""
  )

  /** A criteria reader of the helper column F (windowed, whole-column or named) and A. */
  private val genReader: Gen[String] =
    for
      k <- Gen.choose(3, Last)
      window <- Gen.oneOf(s"Data!F2:F$k", "Data!F:F", "fcol", s"Data!F$k:F$Last")
      sumWindow = window match
        case "Data!F:F" => "Data!A:A"
        case "fcol" => s"Data!A2:A$Last"
        case w => w.replace('F', 'A')
      c1 <- genCriterion
      c2 <- genCriterion
      f <- Gen.oneOf(
        s"COUNTIF($window,$c1)",
        s"SUMIF($window,$c1,$sumWindow)",
        s"AVERAGEIF($window,$c1,$sumWindow)",
        s"COUNTIFS($window,$c1,$sumWindow,$c2)",
        s"SUMIFS($sumWindow,$window,$c1)",
        s"AVERAGEIFS($sumWindow,$window,$c1,$window,$c2)",
        s"MINIFS($sumWindow,$window,$c1)",
        s"MAXIFS($sumWindow,$window,$c1,$sumWindow,$c2)"
      )
    yield f

  private final case class StaleBook(
    workbook: Workbook,
    edit: ARef,
    edited: CellValue,
    cyclic: Boolean
  )

  private val genStaleBook: Gen[StaleBook] =
    for
      as <- Gen.listOfN(Rows, genConstant)
      helpers <- Gen.sequence[List[CellValue], CellValue]((2 to Last).map(genHelper))
      readers <- Gen.listOfN(24, genReader)
      sameSheet <- Gen.listOfN(4, genReader)
      dynamicRow <- Gen.choose(2, Last)
      cycle <- Gen.frequency(3 -> false, 1 -> true)
      editRow <- Gen.choose(2, Last)
      edited <- Gen.choose(-2, 9).map(num)
      readerCache <- Gen.oneOf(None, Some(num(-7)))
    yield
      val withA = as.zipWithIndex.foldLeft(Sheet(data)) {
        case (s, (Some(v), i)) => s.put(ARef.from0(0, i + 1), v)
        case (s, (None, _)) => s
      }
      val withF = helpers.zipWithIndex.foldLeft(withA) { case (s, (f, i)) =>
        s.put(ARef.from0(5, i + 1), f)
      }
      // a dynamic member of the helper column: stripped, evaluated in the deferred bucket
      val withDynamic = withF.put(
        ARef.from0(5, dynamicRow - 1),
        CellValue.Formula(s"""INDIRECT("A$dynamicRow")*1""", Some(num(-1)))
      )
      // same-sheet readers of F in column H
      val dataSheet = sameSheet.zipWithIndex.foldLeft(withDynamic) { case (s, (f, i)) =>
        s.put(ARef.from0(7, i + 1), CellValue.Formula(f, readerCache))
      }
      val readerSheet = readers.zipWithIndex.foldLeft(Sheet(out)) { case (s, (f, i)) =>
        s.put(ARef.from0(1, i), CellValue.Formula(f, readerCache))
      }
      // a second-order reader over the readers' column, and, optionally, a cycle through F
      val outSheet =
        val chained =
          readerSheet.put(at("D1"), CellValue.Formula("""COUNTIF(Out!B:B,">0")""", Some(num(-3))))
        if cycle then chained.put(at("D2"), CellValue.Formula("""COUNTIF(Data!F:F,">1")""", None))
        else chained
      val finalData =
        if cycle then dataSheet.put(at("G2"), CellValue.Formula("Out!D2+1", Some(num(-1))))
        else dataSheet
      val withCycle =
        if cycle then finalData.put(at(s"F$Last"), CellValue.Formula("G2*1", Some(num(-1))))
        else finalData
      StaleBook(
        Workbook(Vector(withCycle, outSheet))
          .withDefinedName("fcol", s"Data!$$F$$2:$$F$$$Last"),
        ARef.from0(0, editRow - 1),
        edited,
        cycle
      )

  private val sequential =
    () => RecalcOptions(clock = clock, rng = Rng.seeded(5L), iterative = IterativeMode.Off)
  private val parallel =
    () => RecalcOptions(clock = clock, iterative = IterativeMode.Off, parallelism = 4)
  private val iterative = () =>
    RecalcOptions(
      clock = clock,
      rng = Rng.seeded(5L),
      iterative = IterativeMode.Force(IterativeCalc(10, BigDecimal("0.01")))
    )

  property("stale-cache law: recalculate with the caches on equals off, sequential and parallel") {
    Prop.forAll(genStaleBook) { book =>
      Prop.all(
        exactlySame(
          recalculateWithMemo(book.workbook, sequential(), cachesOn()),
          recalculateWithMemo(book.workbook, sequential(), cachesOff()),
          "sequential"
        ),
        exactlySame(
          recalculateWithMemo(book.workbook, parallel(), cachesOn()),
          recalculateWithMemo(book.workbook, parallel(), cachesOff()),
          "parallel"
        ),
        exactlySame(
          recalculateWithMemo(book.workbook, iterative(), cachesOn()),
          recalculateWithMemo(book.workbook, iterative(), cachesOff()),
          "iterative"
        )
      )
    }
  }

  property("stale-cache law: recalculateUncached and after-edit with the caches on equal off") {
    Prop.forAll(genStaleBook) { book =>
      val edited = book.workbook
        .update(data, _.put(book.edit, book.edited))
        .fold(e => fail(e.toString), identity)
      Prop.all(
        exactlySame(
          recalculateUncachedWithMemo(book.workbook, sequential(), cachesOn()),
          recalculateUncachedWithMemo(book.workbook, sequential(), cachesOff()),
          "uncached"
        ),
        exactlySame(
          DependentRecalculation
            .recalculateAfterEdit(edited, data, Set(book.edit), clock, cachesOn()),
          DependentRecalculation
            .recalculateAfterEdit(edited, data, Set(book.edit), clock, cachesOff()),
          "after-edit"
        )
      )
    }
  }

  /** An edit of the stale book: a constant, a cleared cell, or a helper replaced. */
  private def genEdit: Gen[(ARef, Option[CellValue])] =
    for
      r <- Gen.choose(2, Last)
      edit <- Gen.oneOf(
        Gen.choose(-2, 9).map(n => (ARef.from0(0, r - 1), Some(num(n)))),
        Gen.const((ARef.from0(0, r - 1), None)),
        Gen
          .oneOf(s"A$r*3", s"""IF(A$r>2,"hi",A$r)""", s"F${math.max(2, r - 1)}+1")
          .map(f => (ARef.from0(5, r - 1), Some(CellValue.Formula(f, Some(num(-5)))))),
        Gen.oneOf(num(4), CellValue.Text("hi")).map(v => (ARef.from0(5, r - 1), Some(v)))
      )
    yield edit

  property("stale-cache law: after-edit with the caches on equals off, for every kind of edit") {
    Prop.forAll(genStaleBook, Gen.choose(1, 3).flatMap(n => Gen.listOfN(n, genEdit))) {
      (book, edits) =>
        val edited = book.workbook
          .update(
            data,
            s =>
              edits.foldLeft(s) {
                case (acc, (ref, Some(v))) => acc.put(ref, v)
                case (acc, (ref, None)) => acc.remove(ref)
              }
          )
          .fold(e => fail(e.toString), identity)
        val refs = edits.map(_._1).toSet
        exactlySame(
          DependentRecalculation.recalculateAfterEdit(edited, data, refs, clock, cachesOn()),
          DependentRecalculation.recalculateAfterEdit(edited, data, refs, clock, cachesOff()),
          s"after-edit $edits"
        )
    }
  }

  property("the stale-cache law is not vacuous: its acyclic books build slices") {
    // a cycle through F blocks every reader of F, so only acyclic books must slice
    Prop.forAll(genStaleBook) { book =>
      val counters = RecalcCounters.create()
      recalculateWithMemo(book.workbook, sequential(), cachesOn(counters))
      val snap = counters.snapshot
      Prop(book.cyclic || snap(Counter.SlicesBuilt) > 0L) :| s"no slice built: ${snap.render}"
    }
  }

  // ---------- #790: formula parses reused from the graph build ----------

  property("parse-cache law: reusing the build's parses equals parsing again, every pass") {
    Prop.forAll(genStaleBook) { book =>
      Prop.all(
        Vector(sequential, parallel, iterative).map { options =>
          exactlySame(
            recalculateWithMemo(book.workbook, options(), cachesOn()),
            recalculateWithMemo(book.workbook, options(), parsesOff()),
            s"parses ${options().parallelism} ${options().iterative}"
          )
        } :+ exactlySame(
          recalculateUncachedWithMemo(book.workbook, sequential(), cachesOn()),
          recalculateUncachedWithMemo(book.workbook, sequential(), parsesOff()),
          "parses uncached"
        )*
      )
    }
  }

  property("the parse-cache law is not vacuous: on, a pass parses each formula once; off, more") {
    Prop.forAll(genStaleBook) { book =>
      Prop.all(Vector(sequential, parallel, iterative).map { options =>
        val on = RecalcCounters.create()
        val off = RecalcCounters.create()
        recalculateWithMemo(book.workbook, options(), cachesOn(on))
        recalculateWithMemo(book.workbook, options(), parsesOff(off))
        val formulas = on.snapshot(Counter.GraphNodes)
        Prop(on.snapshot(Counter.FormulaParses) == formulas) :| s"on: ${on.snapshot.render}" &&
        Prop(off.snapshot(Counter.FormulaParses) > formulas) :| s"off: ${off.snapshot.render}"
      }*)
    }
  }

  test("a formula whose text changed after the build is parsed again") {
    // the memo holds the parses of `wrongCaches`; the after-edit pass runs on a book whose F2
    // reads A2*3: a reused parse of A2*2 would compute F2 = 2 and its readers from it
    val memo = cachesOn()
    memo.passCache.foreach(
      _.useParses(com.tjclp.xl.formula.graph.CompactGraph.build(wrongCaches).parses)
    )
    val edited = wrongCaches
      .update(data, _.put(at("F2"), CellValue.Formula("A2*3", Some(num(-1)))))
      .fold(e => fail(e.toString), identity)
    val reused =
      DependentRecalculation.recalculateAfterEdit(edited, data, Set(at("F2")), clock, memo)
    assertEquals(
      reused,
      DependentRecalculation.recalculateAfterEdit(edited, data, Set(at("F2")), clock, cachesOff())
    )
    assertEquals(reused.evaluated.get(data).flatMap(_.get(at("F2"))), Some(num(3)))
  }

  test("build parses answer only for their own cell and text, unparseable formulas never") {
    // an extreme-exponent literal: a hash of its BigDecimal throws (#780), and no lookup takes one
    val text = "A1+100E+2147483646"
    val sheet = Sheet(data)
      .put(at("A1"), num(1))
      .put(at("B1"), CellValue.Formula(text))
      .put(at("B2"), CellValue.Formula("SUM(("))
    val parses =
      com.tjclp.xl.formula.graph.CompactGraph.build(Workbook(Vector(sheet))).parses
    assertEquals(
      parses.get(data, at("B1"), text).map(_.toString),
      com.tjclp.xl.formula.parser.FormulaParser.parse(text).toOption.map(_.toString)
    )
    assert(parses.get(data, at("B1"), text).isDefined)
    assertEquals(parses.get(data, at("B1"), "A1+1"), None)
    assertEquals(parses.get(data, at("C1"), text), None)
    assertEquals(parses.get(other, at("B1"), text), None)
    assertEquals(parses.get(data, at("B2"), "SUM(("), None)
  }

  // ---------- fixed cases ----------

  /** Data!F2:F301 = A*2 over A = 1..300, every cache -1; readers windowed and whole-column. */
  private val wrongCaches: Workbook =
    val sheet = (2 to 301).foldLeft(Sheet(data)) { (s, r) =>
      s.put(ARef.from0(0, r - 1), num(r - 1))
        .put(ARef.from0(5, r - 1), CellValue.Formula(s"A$r*2", Some(num(-1))))
    }
    val readers = Vector(
      "COUNTIF(Data!F2:F301,-1)",
      "COUNTIF(Data!F:F,-1)",
      """COUNTIF(Data!F2:F100,">100")""",
      """SUMIFS(Data!A:A,Data!F:F,">500")""",
      "COUNTIF(fcol,600)"
    )
    val outSheet = readers.zipWithIndex.foldLeft(Sheet(out)) { case (s, (f, i)) =>
      s.put(ARef.from0(0, i), CellValue.Formula(f, Some(num(-9))))
    }
    Workbook(Vector(sheet, outSheet)).withDefinedName("fcol", "Data!$F$2:$F$301")

  private def outValues(result: RecalcResult): Vector[Option[CellValue]] =
    (0 until 5).map(i => result.evaluated.get(out).flatMap(_.get(ARef.from0(0, i)))).toVector

  test("a slice never reads a cache its pass has not recomputed (default thresholds)") {
    val expected = Vector(0, 0, 49, 300 * 301 / 2 - 250 * 251 / 2, 1).map(n => Some(num(n)))
    Vector(sequential(), parallel()).foreach { options =>
      val counters = RecalcCounters.create()
      val result =
        recalculateWithMemo(
          wrongCaches,
          options,
          new Evaluator.AggregateMemo(counters = counters)
        )
      assertEquals(outValues(result), expected)
      val snap = counters.snapshot
      // F2:F301 (the whole column bounds to the same rows); F2:F100 is below the threshold
      assertEquals(snap(Counter.SlicesBuilt), 1L)
      assertEquals(snap(Counter.ScanFallbacks), 0L)
      assertEquals(snap(Counter.UsedRangeScans), 1L)
    }
  }

  test("recalculateUncached reads every cache as it is, through slices too (GH-468)") {
    // the readers lose their caches; the helper column keeps its wrong ones, which this pass trusts
    val uncachedReaders = wrongCaches
      .update(
        out,
        s =>
          s.cells.foldLeft(s) { case (acc, (ref, cell)) =>
            cell.value match
              case f: CellValue.Formula => acc.put(ref, f.copy(cachedValue = None))
              case _ => acc
          }
      )
      .fold(e => fail(e.toString), identity)
    val counters = RecalcCounters.create()
    val on = recalculateUncachedWithMemo(
      uncachedReaders,
      sequential(),
      new Evaluator.AggregateMemo(counters = counters)
    )
    assertEquals(on, recalculateUncachedWithMemo(uncachedReaders, sequential(), cachesOff()))
    assertEquals(outValues(on), Vector(300, 300, 0, 0, 0).map(n => Some(num(n))))
    assertEquals(counters.snapshot(Counter.SlicesBuilt), 1L)
    assertEquals(counters.snapshot(Counter.ScanFallbacks), 0L)
  }

  test("a box holding a formula its pass left uncached is refused, and the call scans") {
    // F150 does not parse: its evaluation fails and the pass strips its cache, so the slice
    // build meets an uncached formula, gives up, and every reader scans (and fails as before)
    val failing = wrongCaches
      .update(data, _.put(ARef.from0(5, 149), CellValue.Formula("SUM((", Some(num(-1)))))
      .fold(e => fail(e.toString), identity)
    Vector(sequential(), parallel()).foreach { options =>
      val counters = RecalcCounters.create()
      // default thresholds: F2:F100 is too short to slice, so every slice request spans F150
      val on =
        recalculateWithMemo(failing, options, new Evaluator.AggregateMemo(counters = counters))
      assertEquals(on, recalculateWithMemo(failing, options, cachesOff()))
      // the four readers spanning F150 wanted the one slice and fell back
      assertEquals(counters.snapshot(Counter.ScanFallbacks), 4L)
      assertEquals(counters.snapshot(Counter.SlicesBuilt), 0L)
    }
  }

  test("used-range scans: at most one per sheet and structural generation") {
    val counters = RecalcCounters.create()
    recalculateWithMemo(wrongCaches, sequential(), cachesOn(counters))
    assert(counters.snapshot(Counter.UsedRangeScans) <= wrongCaches.sheets.size.toLong)
  }

  test("a structural change re-scans the used range; a formula cache write does not") {
    val cache = new PassCache(eager, RecalcCounters.disabled)
    val sheet = Sheet(data).put(at("A1"), num(1)).put(at("A3"), CellValue.Formula("A1"))
    // inactive: every call scans
    assertEquals(cache.usedRange(sheet).map(_.toA1), Some("A1:A3"))
    cache.begin(Readiness.build(Map.empty, Set.empty))
    assertEquals(cache.usedRange(sheet).map(_.toA1), Some("A1:A3"))
    val grown = sheet.put(at("B9"), num(2))
    // same generation: the recorded range stands (a pass never grows a sheet without saying so)
    assertEquals(cache.usedRange(grown).map(_.toA1), Some("A1:A3"))
    assert(SheetEvaluator.threadChangesStructure(sheet, at("B9")))
    assert(!SheetEvaluator.threadChangesStructure(sheet, at("A3")))
    cache.structureChanged(data)
    assertEquals(cache.usedRange(grown).map(_.toA1), Some("A1:B9"))
  }

  test("running windows build one slice per column, never one per window") {
    // B{n} = COUNTIF($A$2:A{n},A{n}) over constants, D{n} = COUNTIF($C$2:C{n},">10") over the
    // formula column C{n} = A{n}*2: every window is a different box. With default thresholds each
    // window past the size threshold would otherwise copy and keep its own slice (quadratic cells).
    val n = 1200
    val sheet = (2 to n + 1).foldLeft(Sheet(data)) { (s, r) =>
      s.put(ARef.from0(0, r - 1), num(r % 17))
        .put(ARef.from0(1, r - 1), CellValue.Formula(s"COUNTIF($$A$$2:A$r,A$r)", Some(num(-1))))
        .put(ARef.from0(2, r - 1), CellValue.Formula(s"A$r*2", Some(num(-1))))
        .put(ARef.from0(3, r - 1), CellValue.Formula(s"""COUNTIF($$C$$2:C$r,">10")""", None))
    }
    val wb = Workbook(Vector(sheet))
    Vector(sequential(), parallel()).foreach { options =>
      val counters = RecalcCounters.create()
      val on = recalculateWithMemo(wb, options, new Evaluator.AggregateMemo(counters = counters))
      assertEquals(on, recalculateWithMemo(wb, options, cachesOff()))
      assertEquals(on.evaluated.get(data).flatMap(_.get(at(s"B${n + 1}"))), Some(num(71)))
      assert(counters.snapshot(Counter.SlicesBuilt) <= 2L, counters.snapshot.render)
    }
  }

  test("a memo used for a second pass forgets the first pass's slices and used ranges") {
    val memo = cachesOn()
    val first = recalculateWithMemo(wrongCaches, sequential(), memo)
    // the second book has the same shape and different constants: a slice kept from the first
    // pass would answer with the first book's helper values
    val changed = wrongCaches
      .update(data, s => (2 to 301).foldLeft(s)((acc, r) => acc.put(ARef.from0(0, r - 1), num(1))))
      .fold(e => fail(e.toString), identity)
    val second = recalculateWithMemo(changed, sequential(), memo)
    assertEquals(second, recalculateWithMemo(changed, sequential(), cachesOff()))
    assertNotEquals(outValues(first), outValues(second))
  }

  test("readiness decides a build at each slot; a failed copy is remembered, a refusal is not") {
    // F2..F400 are formulas; F300 evaluates at slot 10, every other one at slot 0
    val sheet = (2 to 400).foldLeft(Sheet(data)) { (s, r) =>
      s.put(ARef.from0(5, r - 1), CellValue.Formula(s"A$r", Some(num(r))))
    }
    val slots = (2 to 400).map { r =>
      com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef(data, ARef.from0(5, r - 1)) ->
        (if r == 300 then 10 else 0)
    }.toMap
    val counters = RecalcCounters.create()
    val cache = new PassCache(PassCache.DefaultLimits, counters)
    cache.begin(Readiness.build(slots, Set.empty))
    def ask(a: Int, b: Int, slot: Int) = cache.slice(sheet, 5, a, b, EvalPosition.at(slot))
    def served(answer: PassCache.SliceAnswer): Boolean = answer match
      case PassCache.SliceAnswer.Use(_) => true
      case _ => false
    // before F300 is done: the column and any box over F300 are refused, every time
    assertEquals(ask(1, 399, 5), PassCache.SliceAnswer.Refused)
    assertEquals(ask(1, 399, 5), PassCache.SliceAnswer.Refused)
    // a box clear of F300 is built on its second request and then serves contained boxes
    assertEquals(ask(1, 290, 5), PassCache.SliceAnswer.Refused)
    ask(1, 290, 5) match
      case PassCache.SliceAnswer.Use(slice) => assertEquals(slice.valueAt(9), num(10))
      case other => fail(other.toString)
    assert(served(ask(3, 280, 5)))
    // after F300 is done: the whole column is built and serves every window
    ask(1, 399, 11) match
      case PassCache.SliceAnswer.Use(slice) =>
        assertEquals(slice.valueAt(299), num(300))
        assertEquals(slice.valueAt(5000), CellValue.Empty)
      case other => fail(other.toString)
    assert(served(ask(350, 360, 12)))
    assertEquals(counters.snapshot(Counter.SlicesBuilt), 2L)

    // an uncached formula: the column copy fails once and is not copied again
    val holed = sheet.put(ARef.from0(5, 199), CellValue.Formula("A200", None))
    val again = new PassCache(PassCache.DefaultLimits, RecalcCounters.disabled)
    again.begin(Readiness.build(slots, Set.empty))
    // first request: the column copy fails; second: the box's own copy fails
    assertEquals(again.slice(holed, 5, 1, 399, EvalPosition.at(11)), PassCache.SliceAnswer.Refused)
    assertEquals(again.slice(holed, 5, 1, 399, EvalPosition.at(11)), PassCache.SliceAnswer.Refused)
    // a copy would now succeed, but both failures stand for the generation
    assertEquals(again.slice(sheet, 5, 1, 399, EvalPosition.at(11)), PassCache.SliceAnswer.Refused)
    // a new generation copies again
    again.structureChanged(data)
    assert(served(again.slice(sheet, 5, 1, 399, EvalPosition.at(11))))
  }

  // ---------- names ----------

  test("stale-sheet name law: a workbook-scoped unqualified name reads each reader's own sheet") {
    def sheet(name: SheetName, values: Vector[Int]): Sheet =
      values.zipWithIndex.foldLeft(Sheet(name)) { case (s, (v, i)) =>
        s.put(ARef.from0(1, i), num(v))
      }
    val dataSheet = sheet(data, Vector(1, 2, 3, 4, 5))
      .put(at("D1"), CellValue.Formula("""COUNTIF(loc,">2")""", Some(num(-1))))
      .put(at("D2"), CellValue.Formula("SUM(loc)", Some(num(-1))))
    val otherSheet = sheet(other, Vector(10, 20, 30, 40, 50))
      .put(at("D1"), CellValue.Formula("""COUNTIF(loc,">2")""", Some(num(-1))))
      .put(at("D2"), CellValue.Formula("SUM(loc)", Some(num(-1))))
      .put(at("D3"), CellValue.Formula("""SUMIF(loc,">25")""", Some(num(-1))))
    val wb = Workbook(Vector(dataSheet, otherSheet)).withDefinedName("loc", "$B$1:$B$5")
    val counters = RecalcCounters.create()
    val on = recalculateWithMemo(wb, sequential(), cachesOn(counters))
    assertEquals(on, recalculateWithMemo(wb, sequential(), cachesOff()))
    def value(sheet: SheetName, a1: String) = on.evaluated.get(sheet).flatMap(_.get(at(a1)))
    assertEquals(value(data, "D1"), Some(num(3)))
    assertEquals(value(data, "D2"), Some(num(15)))
    assertEquals(value(other, "D1"), Some(num(5)))
    assertEquals(value(other, "D2"), Some(num(150)))
    assertEquals(value(other, "D3"), Some(num(120)))
    // one parse of the shared body, however many readers and sheets resolve it
    assertEquals(counters.snapshot(Counter.NameParses), 1L)
  }

  test("stale-sheet name law: a name read before and after its cells compute binds the new sheet") {
    // Out!A1 resolves `h` while F2 still holds its stale cache and is evaluated on demand; the
    // dynamic reader Out!A3 resolves it again after F2 and F3 are threaded. A cached Sheet would
    // hand A3 the snapshot from A1's time; the cache keeps the range and rebinds the sheet.
    val dataSheet = Sheet(data)
      .put(at("A2"), num(4))
      .put(at("A3"), num(6))
      .put(at("F2"), CellValue.Formula("""INDIRECT("A2")*10""", Some(num(-1))))
      .put(at("F3"), CellValue.Formula("A3*10", Some(num(-1))))
    val outSheet = Sheet(out)
      .put(at("A1"), CellValue.Formula("""SUMIF(h,">0")""", Some(num(-1))))
      .put(at("A2"), CellValue.Formula("SUM(h)", Some(num(-1))))
      .put(at("A3"), CellValue.Formula("""SUMIF(h,">0")+0*INDIRECT("A1")""", Some(num(-1))))
    val wb = Workbook(Vector(dataSheet, outSheet)).withDefinedName("h", "Data!$F$2:$F$3")
    Vector(sequential(), parallel()).foreach { options =>
      val on = recalculateWithMemo(wb, options, cachesOn())
      assertEquals(on, recalculateWithMemo(wb, options, cachesOff()))
      val outValues = on.evaluated.getOrElse(out, Map.empty)
      assertEquals(outValues.get(at("A1")), Some(num(100)))
      assertEquals(outValues.get(at("A2")), Some(num(100)))
      assertEquals(outValues.get(at("A3")), Some(num(100)))
    }
  }

  test("a sheet-scoped name shadows the workbook-scoped one through the cache") {
    val dataSheet = Sheet(data)
      .put(at("B1"), num(1))
      .put(at("B2"), num(2))
      .put(at("D1"), CellValue.Formula("SUM(loc)"))
    val otherSheet = Sheet(other)
      .put(at("B1"), num(10))
      .put(at("C1"), num(100))
      .put(at("D1"), CellValue.Formula("SUM(loc)"))
      .put(at("D2"), CellValue.Formula("SUM(Data!D1)+SUM(loc)"))
    val wb = Workbook(Vector(dataSheet, otherSheet))
      .withDefinedName("loc", "$B$1:$B$2")
      .withDefinedName("loc", "Other!$C$1", other)
      .fold(e => fail(e.toString), identity)
    val on = recalculateWithMemo(wb, sequential(), cachesOn())
    assertEquals(on, recalculateWithMemo(wb, sequential(), cachesOff()))
    assertEquals(on.evaluated.get(data).flatMap(_.get(at("D1"))), Some(num(3)))
    assertEquals(on.evaluated.get(other).flatMap(_.get(at("D1"))), Some(num(100)))
    assertEquals(on.evaluated.get(other).flatMap(_.get(at("D2"))), Some(num(103)))
  }

  // ---------- function parity over every value class ----------

  private val genAnyValue: Gen[Option[CellValue]] = Gen.frequency(
    2 -> Gen.const(None),
    3 -> Gen.choose(-3, 9).map(n => Some(num(n))),
    1 -> Gen.oneOf("4.20", "0.50").map(s => Some(CellValue.Number(BigDecimal(s)))),
    3 -> Gen
      .oneOf("x", "X", "xy", "", " 42 ", "042", "4.2", "TRUE", "a*b", "~*", "1e2", "-0")
      .map(s => Some(CellValue.Text(s))),
    1 -> Gen.oneOf("x", "42").map(s => Some(CellValue.RichText(RichText.plain(s)))),
    1 -> Gen.oneOf(true, false).map(b => Some(CellValue.Bool(b))),
    1 -> Gen.const(Some(CellValue.DateTime(LocalDateTime.of(1900, 1, 3, 0, 0)))),
    1 -> Gen
      .oneOf(CellError.NA, CellError.Div0, CellError.Value)
      .map(e => Some(CellValue.Error(e))),
    2 -> Gen
      .oneOf("1+1", "\"x\"", "TRUE", "1/0", "\"042\"")
      .flatMap(f =>
        Gen
          .oneOf(Some(num(7)), Some(CellValue.Text("x")), None)
          .map(c => Some(CellValue.Formula(f, c)))
      )
  )

  private val genParityCriterion: Gen[String] = Gen.oneOf(
    "\">2\"",
    "\"x\"",
    "\"X*\"",
    "\"<>x\"",
    "\"\"",
    "\"=\"",
    "\"=042\"",
    "42",
    "TRUE",
    "#N/A",
    "\"<=5\"",
    "\"~*\"",
    "\"<>\"",
    "\"4.2\"",
    "7"
  )

  /** Data has values in A3..A12 and B3..B12; Other in A5..A20 and B5..B20 (different used rows). */
  private val genParityBook: Gen[Workbook] =
    for
      da <- Gen.listOfN(10, genAnyValue)
      db <- Gen.listOfN(10, genAnyValue)
      oa <- Gen.listOfN(16, genAnyValue)
      ob <- Gen.listOfN(16, genAnyValue)
      criteria <- Gen.listOfN(6, genParityCriterion)
      window <- Gen.oneOf("A1:A20", "A3:A12", "A2:A9", "A:A", "A6:A30")
    yield
      def column(sheet: Sheet, col: Int, first: Int, values: List[Option[CellValue]]): Sheet =
        values.zipWithIndex.foldLeft(sheet) {
          case (s, (Some(v), i)) => s.put(ARef.from0(col, first - 1 + i), v)
          case (s, (None, _)) => s
        }
      val dataSheet = column(column(Sheet(data), 0, 3, da), 1, 3, db)
      val otherSheet = column(column(Sheet(other), 0, 5, oa), 1, 5, ob)
      val sumWindow = window.replace('A', 'B')
      val List(c1, c2, c3, c4, c5, c6) = criteria: @unchecked
      val formulas = Vector(
        s"COUNTIF(Data!$window,$c1)",
        s"COUNTIFS(Data!$window,$c1,Data!$sumWindow,$c2)",
        s"SUMIF(Data!$window,$c3,Data!$sumWindow)",
        s"SUMIFS(Data!$sumWindow,Data!$window,$c1,Data!$window,$c4)",
        s"AVERAGEIF(Data!$window,$c5,Data!$sumWindow)",
        s"AVERAGEIFS(Data!$sumWindow,Data!$window,$c6)",
        s"MINIFS(Data!$sumWindow,Data!$window,$c2)",
        s"MAXIFS(Data!$sumWindow,Data!$window,$c3)",
        // whole columns across sheets with different used ranges share one bound
        s"COUNTIFS(Data!A:A,$c1,Other!A:A,$c2)",
        s"SUMIFS(Other!B:B,Data!A:A,$c4,Other!A:A,$c5)",
        s"MAXIFS(Data!B:B,Other!A:A,$c6)",
        s"COUNTIF(Other!$window,$c3)",
        s"SUM(COUNTIF(Data!$window,{1,2,\"x\"}))",
        // slots a slice never serves: a LET name, a computed name, two columns; and SUMIF's
        // sum range sized from its corner
        s"LET(r,Data!$window,COUNTIF(r,$c1))",
        s"COUNTIF(dyn,$c2)",
        s"COUNTIF(Data!A3:B12,$c4)",
        s"SUMIF(Data!$window,$c5,Data!B3)",
        s"COUNTIFS(stat,$c6,Data!$sumWindow,$c1)"
      )
      val outSheet = formulas.zipWithIndex.foldLeft(Sheet(out)) { case (s, (f, i)) =>
        s.put(ARef.from0(0, i), CellValue.Formula(f))
      }
      Workbook(Vector(dataSheet, otherSheet, outSheet))
        .withDefinedName("dyn", "INDEX(Data!$A:$A,1):INDEX(Data!$A:$A,20)")
        .withDefinedName(
          "stat",
          "Data!" + window.replaceAll("([A-Z]+)", "\\$$1").replaceAll("([0-9]+)", "\\$$1")
        )

  property("function parity: the slice path equals the scan path for every criteria function") {
    Prop.forAll(genParityBook) { book =>
      Prop.all(
        exactlySame(
          recalculateWithMemo(book, sequential(), cachesOn()),
          recalculateWithMemo(book, sequential(), cachesOff()),
          "recalculate"
        ),
        exactlySame(
          recalculateUncachedWithMemo(book, sequential(), cachesOn()),
          recalculateUncachedWithMemo(book, sequential(), cachesOff()),
          "uncached"
        )
      )
    }
  }

  property("function parity is not vacuous: uncached books with no formula in range use slices") {
    Prop.forAll(genParityBook) { book =>
      val counters = RecalcCounters.create()
      recalculateUncachedWithMemo(book, sequential(), cachesOn(counters))
      val hasUncachedInData = book.sheets.exists(s =>
        s.name != out && s.cells.valuesIterator.exists(_.value match
          case CellValue.Formula(_, None, _) => true
          case _ => false)
      )
      Prop(hasUncachedInData || counters.snapshot(Counter.SlicesBuilt) > 0L) :|
        counters.snapshot.render
    }
  }

  test("spill-extent criteria ranges read as the scan reads them (GH-695)") {
    val anchor = CellValue.Formula(
      "SEQUENCE(5)",
      Some(num(1)),
      FormulaKind.ArrayFormula(CellRange(at("E1"), at("E5")))
    )
    // E2:E5 hold the previous generation's spill, which the threaded sheet keeps
    val dataSheet = (2 to 5)
      .foldLeft(Sheet(data).put(at("E1"), anchor)) { (s, r) =>
        s.put(ARef.from0(4, r - 1), num(10 * r))
      }
      .put(at("A1"), num(1))
    val outSheet = Sheet(out)
      .put(at("A1"), CellValue.Formula("""COUNTIF(Data!E1:E5,">2")"""))
      .put(at("A2"), CellValue.Formula("""SUMIFS(Data!E1:E5,Data!E1:E5,">0")"""))
      .put(at("A3"), CellValue.Formula("""COUNTIF(Data!E:E,">20")"""))
      .put(at("A4"), CellValue.Formula("""SUM(Data!E1#)"""))
    val wb = Workbook(Vector(dataSheet, outSheet))
    Vector(sequential(), parallel()).foreach { options =>
      val counters = RecalcCounters.create()
      val on = recalculateWithMemo(wb, options, cachesOn(counters))
      val off = recalculateWithMemo(wb, options, cachesOff())
      assertEquals(cells(on.workbook), cells(off.workbook))
      assertEquals(on, off)
      assert(counters.snapshot(Counter.SlicesBuilt) > 0L, counters.snapshot.render)
    }
  }

  // ---------- readiness ----------

  property("Readiness.allows equals its definition on random columns") {
    val genColumn =
      for
        n <- Gen.choose(0, 40)
        rows <- Gen.pick(n, 0 until 60).map(_.toVector.sorted)
        slots <- Gen.listOfN(n, Gen.choose(-1, 30))
        stripped <- Gen.listOfN(n, Gen.frequency(9 -> false, 1 -> true))
        r0 <- Gen.choose(-2, 62)
        r1 <- Gen.choose(-2, 62)
        slot <- Gen.choose(-1, 32)
      yield (rows.zip(slots).zip(stripped).map { case ((r, p), s) => (r, p, s) }, r0, r1, slot)
    Prop.forAll(genColumn) { (members, r0, r1, slot) =>
      val slots = members.collect {
        case (row, p, false) if p >= 0 =>
          com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef(data, ARef.from0(2, row)) -> p
      }.toMap
      val stripped = members.collect { case (row, _, true) =>
        com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef(data, ARef.from0(2, row))
      }.toSet
      val readiness = Readiness.build(slots, stripped)
      val inBox = members.filter((row, _, _) => row >= r0 && row <= r1)
      val expected =
        slot >= 0 && (r0 > r1 || inBox.forall((_, p, s) => !s && p < slot))
      val position = EvalPosition.at(slot)
      Prop(readiness.allows(data, 2, r0, r1, position) == expected) &&
      Prop(readiness.allows(other, 2, r0, r1, position) == (slot >= 0))
    }
  }

  test("a column slice is total: Empty outside the rows it holds") {
    val slice =
      new ColumnSlice(data, 0, 0, 9, 3, IArray[CellValue](num(1), CellValue.Text("x")), 0L)
    assertEquals(slice.valueAt(3), num(1))
    assertEquals(slice.valueAt(4), CellValue.Text("x"))
    Vector(-1, 0, 2, 5, 9, 100, Int.MaxValue, Int.MinValue).foreach(r =>
      assertEquals(slice.valueAt(r), CellValue.Empty, r.toString)
    )
    assert(slice.covers(data, 0, 0, 9))
    assert(slice.covers(data, 0, 4, 4))
    assert(!slice.covers(data, 0, 0, 10))
    assert(!slice.covers(data, 1, 0, 9))
    assert(!slice.covers(other, 0, 0, 9))
  }
