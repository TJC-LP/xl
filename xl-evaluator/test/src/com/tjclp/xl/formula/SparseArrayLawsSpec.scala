package com.tjclp.xl.formula

import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.forAll

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import com.tjclp.xl.formula.eval.{Aggregator, ArrayArithmetic, ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.units.StyleId

/**
 * #750: the laws that let a sparse [[ArrayResult]] (a populated box plus a fill: one value, or one
 * per row or per column) stand in for the dense array it describes. Its run walk is its cells, its
 * tally their counts with each value first met in row-major order; every operation agrees with the
 * same operation on the dense array; a range read through its populated box evaluates every formula
 * as the position-by-position read does; and an aggregate's closed-form `combineN` is `n` combines.
 */
class SparseArrayLawsSpec extends ScalaCheckSuite:

  // The regression sheets (a header over 200,000 rows, 600,000 rows plus B1) take ~6 s locally but
  // over 30 s on a loaded CI runner, mostly building the sheets
  override val munitTimeout = scala.concurrent.duration.Duration(120, "s")

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(300)

  private val genCell: Gen[CellValue] = Gen.frequency(
    4 -> Gen.choose(-5, 5).map(i => CellValue.Number(BigDecimal(i))),
    1 -> Gen.choose(-50, 50).map(i => CellValue.Number(BigDecimal(i) / 4)),
    2 -> Gen.oneOf("a", "B", "", "3").map(CellValue.Text(_)),
    1 -> Gen.oneOf(true, false).map(CellValue.Bool(_)),
    3 -> Gen.const(CellValue.Empty),
    1 -> Gen.oneOf(CellError.NA, CellError.Div0).map(CellValue.Error(_))
  )

  /** A fill for a `rows` x `cols` array: one value, one per row or one per column. */
  private def genFill(rows: Int, cols: Int): Gen[ArrayResult.Fill] =
    Gen.frequency(
      3 -> genCell.map(ArrayResult.Fill.Const(_)),
      1 -> Gen.listOfN(rows, genCell).map(vs => ArrayResult.Fill.ByRow(vs.toVector)),
      1 -> Gen.listOfN(cols, genCell).map(vs => ArrayResult.Fill.ByCol(vs.toVector))
    )

  /**
   * A sparse array: a box anywhere inside a 1..7 x 1..7 shape (possibly empty), and a fill that is
   * one value, or one value per row or per column.
   */
  private val genSparse: Gen[ArrayResult] =
    for
      rows <- Gen.choose(1, 7)
      cols <- Gen.choose(1, 7)
      top <- Gen.choose(0, rows - 1)
      left <- Gen.choose(0, cols - 1)
      height <- Gen.choose(0, rows - top)
      width <- Gen.choose(1, cols - left)
      box <- Gen.listOfN(height, Gen.listOfN(width, genCell).map(_.toVector)).map(_.toVector)
      fill <- genFill(rows, cols)
    yield ArrayResult.sparse(rows, cols, top, left, box, fill)

  private def dense(ar: ArrayResult): ArrayResult = ArrayResult(ar.dense)

  property("runs expanded are the dense cells, in row-major order") {
    forAll(genSparse) { ar =>
      val expanded = ar.runs.flatMap((v, n) => Iterator.fill(n.toInt)(v)).toVector
      assertEquals(expanded, ar.dense.flatten)
      assertEquals(ar.runs.map(_._2).sum, ar.cellCount)
      assertEquals(ar.cells.toVector, ar.dense.flatten)
    }
  }

  property("tally holds the cells' counts and meets each value first in row-major order") {
    def firsts(values: Iterator[CellValue]): Vector[CellValue] =
      values.foldLeft(Vector.empty[CellValue])((seen, v) =>
        if seen.contains(v) then seen else seen :+ v
      )
    def counts(groups: Iterator[(CellValue, Long)]): Map[CellValue, Long] =
      groups.foldLeft(Map.empty[CellValue, Long]) { case (m, (v, n)) =>
        m.updated(v, m.getOrElse(v, 0L) + n)
      }
    forAll(genSparse) { ar =>
      assertEquals(counts(ar.tally), counts(ar.dense.flatten.iterator.map(_ -> 1L)))
      assertEquals(firsts(ar.tally.map(_._1)), firsts(ar.dense.flatten.iterator))
      assert(ar.tally.size <= ar.runs.size)
    }
  }

  test("a fill by column folds in a handful of groups, whatever the height") {
    val n = (i: Int) => CellValue.Number(BigDecimal(i))
    val tall = ArrayResult.sparse(
      1048576,
      3,
      2,
      0,
      Vector(Vector(n(4), n(8), n(12)), Vector(n(5), n(10), n(15))),
      ArrayResult.Fill.ByCol(Vector(n(0), n(1), n(2)))
    )
    // two rows above the box (a row of three, then three groups), its six cells, and the rows
    // below it (a row of three, then three groups)
    assertEquals(tally(tall).size, 18)
    assertEquals(tally(tall).map(_._2).sum, 3L * 1048576)
    assertEquals(tall(1048575, 2), n(2))
    assertEquals(tall.transpose(2, 1048575), n(2))
    assertEquals(tall.transpose.fill, ArrayResult.Fill.ByRow(Vector(n(0), n(1), n(2))))
  }

  private def tally(ar: ArrayResult): Vector[(CellValue, Long)] = ar.tally.toVector

  property("a sparse array equals its dense form, cell by cell, with the same hash") {
    forAll(genSparse) { ar =>
      val d = dense(ar)
      assertEquals(ar, d)
      assertEquals(d, ar)
      assertEquals(ar.hashCode, d.hashCode)
      (0 until ar.rows).foreach(r =>
        (0 until ar.cols).foreach(c => assertEquals(ar(r, c), d(r, c)))
      )
    }
  }

  property("map, transpose and slice agree with the dense array; map keeps the first error") {
    val failOn: CellValue => Either[EvalError, CellValue] = {
      case CellValue.Error(e) => Left(EvalError.EvalFailed(s"element ${e.toExcel}", None))
      case CellValue.Text(s) => Right(CellValue.Text(s + "!"))
      case other => Right(other)
    }
    forAll(genSparse, Gen.choose(-1, 7), Gen.choose(0, 8), Gen.choose(-1, 7), Gen.choose(0, 8)) {
      (ar, r0, nr, c0, nc) =>
        val d = dense(ar)
        assertEquals(ar.map(failOn), d.map(failOn))
        assertEquals(
          ar.mapPure(cv => CellValue.Bool(cv == CellValue.Empty)),
          d.mapPure(cv => CellValue.Bool(cv == CellValue.Empty))
        )
        assertEquals(ar.transpose, d.transpose)
        assertEquals(ar.slice(r0, nr, c0, nc), d.slice(r0, nr, c0, nc))
    }
  }

  property("an elementwise broadcast over sparse operands equals the dense broadcast") {
    val elem: Vector[CellValue] => Either[EvalError, CellValue] = elements =>
      elements.collectFirst { case CellValue.Error(CellError.Div0) => () } match
        case Some(_) => Left(EvalError.EvalFailed("div0 element", None))
        case None => Right(CellValue.Text(elements.map(_.toString.take(6)).mkString("|")))
    // vector constants ({1,2,3}, {1;2;3}) broadcast along the other axis as a fill by column or row
    val genVector = for
      across <- Gen.oneOf(true, false)
      n <- Gen.choose(1, 7)
      cells <- Gen.listOfN(n, genCell)
    yield
      if across then ArrayResult(Vector(cells.toVector))
      else ArrayResult(cells.toVector.map(Vector(_)))
    val genOperand = Gen.frequency(
      4 -> genSparse,
      1 -> genCell.map(ArrayResult.single),
      2 -> genVector,
      1 -> genSparse.map(_.slice(0, 1, 0, 7).fold(_ => ArrayResult.empty, identity)),
      1 -> genSparse.map(_.slice(0, 7, 0, 1).fold(_ => ArrayResult.empty, identity))
    )
    forAll(Gen.listOfN(3, genOperand)) { ops =>
      val sparse = ArrayArithmetic.broadcastN(ops.toVector)(elem)
      val reference = ArrayArithmetic.broadcastN(ops.toVector.map(dense))(elem)
      assertEquals(sparse, reference)
    }
  }

  // ===== combineN =====

  private val genValue: Gen[BigDecimal] =
    Gen.oneOf(
      Gen.choose(-1000, 1000).map(BigDecimal(_)),
      Gen.choose(-100000, 100000).map(i => BigDecimal(i) / 7)
    )

  private def close(a: BigDecimal, b: BigDecimal): Boolean =
    (a - b).abs <= BigDecimal("1e-25") * (a.abs.max(b.abs).max(BigDecimal(1)))

  private def lawFor[A](agg: Aggregator[A])(same: (A, A) => Boolean): Prop =
    forAll(Gen.listOf(genValue), genValue, Gen.choose(0, 60)) { (prefix, v, n) =>
      val start = prefix.foldLeft(agg.empty)(agg.combine)
      val repeated = (0 until n).foldLeft(start)((acc, _) => agg.combine(acc, v))
      val closed = agg.combineN(start, v, n.toLong)
      assert(same(closed, repeated), s"${agg.name}: $closed vs $repeated")
      (agg.finalizeWithError(closed), agg.finalizeWithError(repeated)) match
        case (Right(x), Right(y)) => assert(close(x, y), s"${agg.name}: $x vs $y")
        case (x, y) => assertEquals(x.isLeft, y.isLeft)
    }

  property("combineN(acc, v, n) is n combines, for every built-in aggregator") {
    import Aggregator.given
    def welford(a: (Long, BigDecimal, BigDecimal), b: (Long, BigDecimal, BigDecimal)) =
      a._1 == b._1 && close(a._2, b._2) && close(a._3, b._3)
    lawFor(sumAggregator)(close) &&
    lawFor(countAggregator)(_ == _) &&
    lawFor(countaAggregator)(_ == _) &&
    lawFor(countblankAggregator)(_ == _) &&
    lawFor(minAggregator)(_ == _) &&
    lawFor(maxAggregator)(_ == _) &&
    lawFor(averageAggregator)((a, b) => a._2 == b._2 && close(a._1, b._1)) &&
    lawFor(medianAggregator)((a, b) =>
      a.flatMap((v, k) => Vector.fill(k.toInt)(v)).sorted ==
        b.flatMap((v, k) => Vector.fill(k.toInt)(v)).sorted
    ) &&
    lawFor(stdevAggregator)(welford) &&
    lawFor(stdevpAggregator)(welford) &&
    lawFor(varAggregator)(welford) &&
    lawFor(varpAggregator)(welford)
  }

  // ===== Sparse evaluation is dense evaluation =====

  /** A sheet with some of A1:F8 populated (numbers, text, blanks-as-values, errors). */
  private val genSheet: Gen[Sheet] =
    Gen
      .listOf(
        for
          col <- Gen.choose(0, 5)
          row <- Gen.choose(0, 7)
          value <- genCell
        yield (ARef.from0(col, row), value)
      )
      .map(_.foldLeft(Sheet(SheetName.unsafe("S"))) { case (s, (at, v)) => s.put(at, v) })

  private val genRange: Gen[String] =
    for
      c0 <- Gen.choose(0, 7)
      r0 <- Gen.choose(0, 9)
      w <- Gen.choose(1, 4)
      h <- Gen.choose(1, 6)
    yield s"${ARef.from0(c0, r0).toA1}:${ARef.from0(c0 + w - 1, r0 + h - 1).toA1}"

  private val genFormula: Gen[String] =
    for
      r <- genRange
      s <- genRange
      template <- Gen.oneOf(
        "=$R*2",
        "=$R=\"\"",
        "=$R>0",
        "=-$R",
        "=$R&\"x\"",
        "=$R+$S",
        "=$R<>$S",
        "=IF($R>1,$R,$S)",
        "=IF($R=\"\",1,0)",
        "=NOT($R)",
        "=SUM($R*1)",
        "=SUM(($R=\"\")*1)",
        "=COUNT($R*1)",
        "=COUNTA($R&\"\")",
        "=AVERAGE($R*1)",
        "=MEDIAN($R*1)",
        "=STDEV($R*1)",
        "=MAX($R+$S)",
        "=SUMPRODUCT(($R>0)*1)",
        "=SUMPRODUCT(($R=\"\")*($S=\"\"))",
        "=AND($R>-10)",
        "=OR($R=\"a\")",
        "=TRANSPOSE($R)",
        "=INDEX($R*1,0,1)",
        "=ROWS($R*1)",
        "=$R*{1,2}",
        "=$R*{1;2;3}",
        "=SUM($R*{1,2,3})",
        "=SUM(($R=\"\")*{1;2})",
        "=SUMPRODUCT(($R=\"\")*{1,2,3})",
        "=COUNT($R*{1,2,3})",
        // a fill by row times a fill by column: SUMPRODUCT's outside sum by inclusion-exclusion
        "=SUMPRODUCT((A1:D6=\"\")*{1;2;3;4;5;6},(A1:D6>0)*{1,2,3,4})",
        "=SUMPRODUCT((B2:E7=\"\")*{1;2;3;4;5;6},(B2:E7=\"\")*{2,3,4,5}+1)",
        "=AND($R*{1,2}>=0)",
        "=MAX(($R=\"\")*{1;2;3})",
        "=INDEX($R*{1,2,3},0,2)",
        "=ABS($R)",
        "=ISBLANK($R)",
        "=IFERROR($R/$S,\"e\")",
        "=$R"
      )
    yield template.replace("$R", r).replace("$S", s)

  private def evalWith(mode: ArrayArithmetic.RangeRead, sheet: Sheet, formula: String): Any =
    FormulaParser.parse(formula) match
      case Left(err) => s"parse: $err"
      case Right(expr) =>
        Evaluator
          .arrayInstanceReading(Rng.seeded(0), mode)
          .eval(expr, sheet, Clock.system, None, None)

  property("a range read through its populated box evaluates as the positional read does") {
    forAll(genSheet, genFormula) { (sheet, formula) =>
      val sparse = evalWith(ArrayArithmetic.RangeRead.Sparse, sheet, formula)
      val dense = evalWith(ArrayArithmetic.RangeRead.Dense, sheet, formula)
      assertEquals(sparse, dense, formula)
    }
  }

  property("SUMPRODUCT over boxes in disjoint corners evaluates as the positional read does") {
    // each argument reads one half of A1:F8, so its box holds only that half's populated cells:
    // the boxes sit apart, and every tile of the grid their edges cut is summed exactly once
    val halves = List(
      ("A1:F4", "A5:F8"),
      ("A1:C8", "D1:F8"),
      ("A1:C4", "D5:F8"),
      ("D1:F4", "A5:C8")
    )
    val templates = List(
      "=SUMPRODUCT(($R>0)*1,($S>0)*1)",
      "=SUMPRODUCT(($R=\"\")*1,($S=\"\")*1)",
      "=SUMPRODUCT(($R>0)*1+1,($S=\"\")*2+1)",
      "=SUMPRODUCT(($R=\"\")*{1,2,3},($S>0)*{4;5;6;7}+1)",
      "=SUMPRODUCT(($R=\"\")*{1;2;3;4},($S=\"\")*{1,2,3})",
      "=SUMPRODUCT(($R=\"\")*1,($S=\"\")*1,($R>0)+($S>0)+1)"
    )
    forAll(genSheet, Gen.oneOf(halves), Gen.oneOf(templates)) { (sheet, half, template) =>
      val (r, s) = half
      val formula = template.replace("$R", r).replace("$S", s)
      val sparse = evalWith(ArrayArithmetic.RangeRead.Sparse, sheet, formula)
      val dense = evalWith(ArrayArithmetic.RangeRead.Dense, sheet, formula)
      assertEquals(sparse, dense, formula)
    }
  }

  test("vector operands give fills by row and by column (the vector laws are not vacuous)") {
    val sheet = Sheet(SheetName.unsafe("S")).put(ARef.from0(1, 1), CellValue.Number(BigDecimal(3)))
    def fillOf(formula: String) = evalWith(ArrayArithmetic.RangeRead.Sparse, sheet, formula) match
      case Right(ar: ArrayResult) => ar.fill
      case other => fail(s"expected an array, got $other")
    val n = (i: Int) => CellValue.Number(BigDecimal(i))
    assertEquals(
      fillOf("=(A1:D6=\"\")*{1;2;3;4;5;6}"),
      ArrayResult.Fill.ByRow(Vector(n(1), n(2), n(3), n(4), n(5), n(6)))
    )
    assertEquals(
      fillOf("=(A1:D6=\"\")*{1,2,3,4}"),
      ArrayResult.Fill.ByCol(Vector(n(1), n(2), n(3), n(4)))
    )
    // every blank cell's (row + 1) * (col + 1): 21 * 10 less B2's 2 * 2
    val mixed = "=SUMPRODUCT((A1:D6=\"\")*{1;2;3;4;5;6},(A1:D6=\"\")*{1,2,3,4})"
    assertEquals(evalWith(ArrayArithmetic.RangeRead.Sparse, sheet, mixed), Right(BigDecimal(206)))
    assertEquals(evalWith(ArrayArithmetic.RangeRead.Dense, sheet, mixed), Right(BigDecimal(206)))
  }

  test("the sparse read is genuinely sparse (the law above is not vacuous)") {
    val sheet = Sheet(SheetName.unsafe("S")).put(ARef.from0(1, 1), CellValue.Number(BigDecimal(3)))
    evalWith(ArrayArithmetic.RangeRead.Sparse, sheet, "=A1:C5*2") match
      case Right(ar: ArrayResult) =>
        assert(!ar.isDense, ar.toString)
        assertEquals((ar.boxTop, ar.boxLeft, ar.boxRows, ar.boxCols), (1, 1, 1, 1))
        assertEquals(ar.fill, ArrayResult.Fill.Const(CellValue.Number(BigDecimal(0))))
        assertEquals(Right(ar), evalWith(ArrayArithmetic.RangeRead.Dense, sheet, "=A1:C5*2"))
      case other => fail(s"expected an array, got $other")
  }

  // ===== The box bound never refuses an ordinary table (#750 review) =====

  private def evalAuto(sheet: Sheet, formula: String): Any =
    FormulaParser.parse(formula) match
      case Left(err) => s"parse: $err"
      case Right(expr) =>
        Evaluator.arrayInstance(Rng.seeded(0)).eval(expr, sheet, Clock.system, None, None)

  private def num(n: Long): CellValue = CellValue.Number(BigDecimal(n))

  private def timed[A](body: => A): (A, Long) =
    val start = System.nanoTime()
    val result = body
    (result, (System.nanoTime() - start) / 1000000L)

  test("a header row over a tall column evaluates, as the positional read does") {
    // A1:J1 numbered headers over A2:A200000: `A:J` is read through its box (200,000 × 10), which
    // the first bound (the larger of 1,048,576 cells and the sheet's 200,009) refused
    val header = (0 until 10).foldLeft(Sheet(SheetName.unsafe("S")))((s, c) =>
      s.put(ARef.from0(c, 0), num(c + 1L))
    )
    val sheet = (1 until 200000).foldLeft(header)((s, r) => s.put(ARef.from0(0, r), num(1L)))
    val cases = List(
      "=SUM((A:J=\"\")*1)" -> BigDecimal(10285751L),
      "=SUM(A:J*1)" -> BigDecimal(55L + 199999L),
      "=ROWS(A:J*1)" -> BigDecimal(1048576L),
      "=SUMPRODUCT((A:J<>\"\")*1)" -> BigDecimal(200009L)
    )
    cases.foreach { (formula, expected) =>
      val (auto, ms) = timed(evalAuto(sheet, formula))
      assertEquals(auto, Right(expected), s"$formula (${ms} ms)")
      assertEquals(evalWith(ArrayArithmetic.RangeRead.Dense, sheet, formula), auto, formula)
    }
  }

  test("a full column beside one cell evaluates, as the positional read does") {
    // A1:A600000 and B1: `A:B`'s box is 600,000 × 2 (1,200,000 cells), past the first bound
    val column = (0 until 600000).foldLeft(Sheet(SheetName.unsafe("S")))((s, r) =>
      s.put(ARef.from0(0, r), num(1L))
    )
    val sheet = column.put(ARef.from0(1, 0), num(1L))
    List("=SUM((A:B=\"\")*1)" -> BigDecimal(1497151L), "=SUM(A:B*2)" -> BigDecimal(1200002L))
      .foreach { (formula, expected) =>
        val (auto, ms) = timed(evalAuto(sheet, formula))
        assertEquals(auto, Right(expected), s"$formula (${ms} ms)")
        assertEquals(evalWith(ArrayArithmetic.RangeRead.Dense, sheet, formula), auto, formula)
      }
  }

  test("a style-only record does not widen the box: it reads as the Empty fill does") {
    val styled = (at: ARef) => Cell(at, CellValue.Empty, Some(StyleId(1)))
    val data =
      Sheet(SheetName.unsafe("S")).put(ARef.from0(0, 0), num(1L)).put(ARef.from0(0, 4), num(2L))
    // formatted blanks in XFD1 and XFD1048576: counted as cells, the box would be the whole sheet
    val sheet = data.put(styled(ARef.from0(16383, 0))).put(styled(ARef.from0(16383, 1048575)))
    assertEquals(evalAuto(sheet, "=SUM((A:XFD=\"\")*1)"), Right(BigDecimal(17179869182L)))
    assertEquals(evalAuto(sheet, "=SUM(A:XFD*1)"), Right(BigDecimal(3L)))
    evalWith(ArrayArithmetic.RangeRead.Sparse, sheet, "=A:XFD*1") match
      case Right(ar: ArrayResult) =>
        assertEquals((ar.boxTop, ar.boxLeft, ar.boxRows, ar.boxCols), (0, 0, 5, 1))
      case other => fail(s"expected an array, got $other")
    // on a small range the dense read agrees (the generated sheets above hold Empty records too)
    val small = data.put(styled(ARef.from0(4, 4))).put(styled(ARef.from0(3, 0)))
    List("=SUM((A1:E5=\"\")*1)", "=A1:E5*1", "=ISBLANK(A1:E5)").foreach { formula =>
      assertEquals(
        evalWith(ArrayArithmetic.RangeRead.Sparse, small, formula),
        evalWith(ArrayArithmetic.RangeRead.Dense, small, formula),
        formula
      )
    }
  }

  test("a box no read could build is refused, naming the bound") {
    // A1 and Q1048576: a box of 1,048,576 × 17 cells, past 16,777,216, on a two-cell sheet
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ARef.from0(0, 0), num(1L))
      .put(ARef.from0(16, 1048575), num(1L))
    evalAuto(sheet, "=SUM(A:Q*1)") match
      case Left(EvalError.EvalFailed(msg, _)) =>
        assert(msg.contains("16,777,216-cell bound on a range's populated box"), msg)
      case other => fail(s"expected the box bound, got $other")
  }
