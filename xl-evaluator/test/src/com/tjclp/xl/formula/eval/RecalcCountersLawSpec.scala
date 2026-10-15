package com.tjclp.xl.formula.eval

import java.time.LocalDateTime

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.eval.RecalcCounters.Counter
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{CalcPr, Workbook}
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * #767: the measured seams return exactly what the public entry points return, on random books.
 *
 * Books mix every `CriteriaMatcher` input class (blanks, text, numeric text, wildcard text,
 * booleans, dates, errors, cached and uncached formulas) under the seven criteria functions,
 * bounded, whole-column and named ranges, cross-sheet reads, OFFSET/INDIRECT, spills and cycles.
 * Each book runs through every `recalculate(options)` dispatch (sequential, parallel, iterative
 * from calcPr and forced, table seeding) and `recalculateUncached`, with counters enabled and
 * disabled.
 */
class RecalcCountersLawSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(60)

  private val data = SheetName.unsafe("Data")
  private val out = SheetName.unsafe("Out")
  private val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 2))
  private val Rows = 8

  private val genValue: Gen[Option[CellValue]] = Gen.frequency(
    2 -> Gen.const(None),
    3 -> Gen.choose(-3, 9).map(n => Some(CellValue.Number(BigDecimal(n)))),
    1 -> Gen.const(Some(CellValue.Number(BigDecimal("4.20")))),
    3 -> Gen
      .oneOf("x", "X", "xy", "", " 42 ", "042", "4.2", "TRUE", "a*b", "~*", "1e2", "-0")
      .map(s => Some(CellValue.Text(s))),
    1 -> Gen.oneOf(true, false).map(b => Some(CellValue.Bool(b))),
    1 -> Gen.const(Some(CellValue.DateTime(LocalDateTime.of(1900, 1, 3, 0, 0)))),
    1 -> Gen
      .oneOf(CellError.NA, CellError.Div0, CellError.Value)
      .map(e => Some(CellValue.Error(e))),
    1 -> Gen.const(Some(CellValue.Formula("B1*2", Some(CellValue.Number(BigDecimal(99)))))),
    2 -> Gen
      .oneOf("B1*2", "A1&\"\"", "1/0", "IF(B2>1,\"x\",TRUE)")
      .map(f => Some(CellValue.Formula(f)))
  )

  private val genRange: Gen[String] = Gen.oneOf(
    s"Data!A1:A$Rows",
    "Data!A:A",
    "Data!A2:A5",
    "colA",
    "boxA",
    s"Data!C1:C$Rows",
    "Out!F1:F4"
  )

  private val genSumRange: Gen[String] =
    Gen.oneOf(s"Data!B1:B$Rows", "Data!B:B", "colB", s"Data!C1:C$Rows")

  private val genCriterion: Gen[String] = Gen.oneOf(
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
    "Data!B1",
    "\"~*\"",
    "\"<>\""
  )

  private val genFormula: Gen[String] =
    for
      r1 <- genRange
      r2 <- genRange
      s <- genSumRange
      c1 <- genCriterion
      c2 <- genCriterion
      f <- Gen.oneOf(
        s"COUNTIF($r1,$c1)",
        s"COUNTIFS($r1,$c1,$r2,$c2)",
        s"SUMIF($r1,$c1)",
        s"SUMIF($r1,$c1,$s)",
        s"SUMIFS($s,$r1,$c1,$r2,$c2)",
        s"AVERAGEIF($r1,$c1,$s)",
        s"AVERAGEIFS($s,$r1,$c1)",
        s"MINIFS($s,$r1,$c1)",
        s"MAXIFS($s,$r1,$c1,$r2,$c2)",
        s"SUM(OFFSET(Data!A1,0,0,3,1))+COUNTIF($r1,$c1)",
        s"COUNTIF(INDIRECT(\"Data!A1:A5\"),$c1)",
        "SUM(Out!A1:A3)",
        "SUM(Out!E1#)",
        s"COUNTIF($r1,Out!E2)",
        "Out!B1+1",
        "Out!F1+1"
      )
    yield f

  private val genBook: Gen[Workbook] =
    for
      as <- Gen.listOfN(Rows, genValue)
      bs <- Gen.listOfN(Rows, genValue)
      cs <- Gen.listOfN(Rows, genValue)
      formulas <- Gen.listOfN(6, genFormula)
      cyclic <- Gen.oneOf(true, false)
      iterate <- Gen.oneOf(true, false)
    yield
      def column(sheet: Sheet, col: Int, values: List[Option[CellValue]]): Sheet =
        values.zipWithIndex.foldLeft(sheet) {
          case (s, (Some(v), row)) => s.put(ARef.from0(col, row), v)
          case (s, (None, _)) => s
        }
      val dataSheet = column(column(column(Sheet(data), 0, as), 1, bs), 2, cs)
      // D1:D6 the drawn formulas; E1 a spill anchor (`Out!E1#`); F1:F4 plain formulas over Data
      val drawn = formulas.zipWithIndex.foldLeft(Sheet(out)) { case (s, (f, i)) =>
        s.put(ARef.from0(3, i), CellValue.Formula(f))
      }
      val outBase =
        (0 until 4).foldLeft(drawn.put(ARef.from0(4, 0), CellValue.Formula("SEQUENCE(3)"))) {
          (s, i) => s.put(ARef.from0(5, i), CellValue.Formula(s"Data!B${i + 1}*2"))
        }
      val outSheet =
        if cyclic then
          outBase
            .put(ARef.from0(1, 0), CellValue.Formula("B2*0.5+1"))
            .put(ARef.from0(1, 1), CellValue.Formula("B1*0.5"))
        else outBase.put(ARef.from0(1, 0), CellValue.Number(BigDecimal(3)))
      val book = Workbook(Vector(dataSheet, outSheet))
        .withDefinedName("colA", "Data!$A:$A")
        .withDefinedName("colB", "Data!$B:$B")
        .withDefinedName("boxA", s"Data!$$A$$1:$$A$$$Rows")
      if iterate then
        book.withCalcPr(
          CalcPr(
            iterativeCalculation = true,
            maxIterations = Some(20),
            maxChange = Some(BigDecimal("0.001"))
          )
        )
      else book

  /** Every dispatch `recalculate(options)` has, each with a fresh seeded generator. */
  private val optionVariants: Vector[() => RecalcOptions] = Vector(
    () => RecalcOptions(clock = clock, rng = Rng.seeded(7L), iterative = IterativeMode.Off),
    () => RecalcOptions(clock = clock, rng = Rng.seeded(7L)),
    () =>
      RecalcOptions(
        clock = clock,
        rng = Rng.seeded(7L),
        iterative = IterativeMode.Force(IterativeCalc(10, BigDecimal("0.01")))
      ),
    () => RecalcOptions(clock = clock, iterative = IterativeMode.Off, parallelism = 2),
    () => RecalcOptions(clock = clock, rng = Rng.seeded(7L), seedTables = true)
  )

  private def same(actual: RecalcResult, expected: RecalcResult, label: String): Prop =
    Prop(actual == expected) :| s"$label: counted ≠ public\n${actual.errors}\n${expected.errors}"

  property("recalculateCounted ≡ recalculate(options) for every dispatch, counters on or off") {
    Prop.forAll(genBook) { book =>
      Prop.all(optionVariants.zipWithIndex.flatMap { case (options, i) =>
        val expected = book.recalculate(options())
        Vector(
          same(recalculateCounted(book, options(), RecalcCounters.create()), expected, s"on #$i"),
          same(recalculateCounted(book, options(), RecalcCounters.disabled), expected, s"off #$i")
        )
      }*)
    }
  }

  property("recalculateUncachedCounted ≡ recalculateUncached(options), counters on or off") {
    Prop.forAll(genBook) { book =>
      val options = RecalcOptions(clock = clock, rng = Rng.seeded(3L))
      val expected = book.recalculateUncached(options)
      val seeded = () => RecalcOptions(clock = clock, rng = Rng.seeded(3L))
      Prop.all(
        same(recalculateUncachedCounted(book, seeded(), RecalcCounters.create()), expected, "on"),
        same(recalculateUncachedCounted(book, seeded(), RecalcCounters.disabled), expected, "off")
      )
    }
  }

  property("graph counters depend only on the book; disabled counters never record") {
    Prop.forAll(genBook) { book =>
      def graph(run: RecalcCounters => RecalcResult): (Long, Long) =
        val counters = RecalcCounters.create()
        run(counters)
        val snap = counters.snapshot
        (snap(Counter.GraphNodes), snap(Counter.GraphEdges))
      val off = RecalcOptions(clock = clock, iterative = IterativeMode.Off)
      val sequential = graph(recalculateCounted(book, off, _))
      val formulas = book.sheets.iterator
        .map(_.cells.valuesIterator.count(_.value match
          case _: CellValue.Formula => true
          case _ => false))
        .sum
        .toLong
      recalculateCounted(book, off, RecalcCounters.disabled)
      Prop.all(
        Prop(sequential._1 == formulas) :| s"GraphNodes ${sequential._1} ≠ $formulas formulas",
        Prop(graph(recalculateCounted(book, off.copy(parallelism = 2), _)) == sequential),
        Prop(graph(recalculateUncachedCounted(book, off, _)) == sequential),
        Prop(RecalcCounters.disabled.snapshot.counters.values.forall(_ == 0L)),
        Prop(RecalcCounters.disabled.snapshot.phaseNanos.values.forall(_ == 0L))
      )
    }
  }
