package com.tjclp.xl.formula

import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.forAll

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{Aggregator, ArrayArithmetic, ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet

/**
 * #750: the laws that let a sparse [[ArrayResult]] (a populated box plus one fill value) stand in
 * for the dense array it describes. Its run walk is its cells; every operation agrees with the same
 * operation on the dense array; a range read through its populated box evaluates every formula as
 * the position-by-position read does; and an aggregate's closed-form `combineN` is `n` combines.
 */
class SparseArrayLawsSpec extends ScalaCheckSuite:

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

  /** A sparse array: a box anywhere inside a 1..7 x 1..7 shape (possibly empty), and a fill. */
  private val genSparse: Gen[ArrayResult] =
    for
      rows <- Gen.choose(1, 7)
      cols <- Gen.choose(1, 7)
      top <- Gen.choose(0, rows - 1)
      left <- Gen.choose(0, cols - 1)
      height <- Gen.choose(0, rows - top)
      width <- Gen.choose(1, cols - left)
      box <- Gen.listOfN(height, Gen.listOfN(width, genCell).map(_.toVector)).map(_.toVector)
      fill <- genCell
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
    val genOperand = Gen.frequency(
      4 -> genSparse,
      1 -> genCell.map(ArrayResult.single),
      1 -> genSparse.map(_.slice(0, 1, 0, 7)),
      1 -> genSparse.map(_.slice(0, 7, 0, 1))
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
        ArrayArithmetic.rangeRead.withValue(mode) {
          Evaluator.arrayInstance(Rng.seeded(0)).eval(expr, sheet, Clock.system, None, None)
        }

  property("a range read through its populated box evaluates as the positional read does") {
    forAll(genSheet, genFormula) { (sheet, formula) =>
      val sparse = evalWith(ArrayArithmetic.RangeRead.Sparse, sheet, formula)
      val dense = evalWith(ArrayArithmetic.RangeRead.Dense, sheet, formula)
      assertEquals(sparse, dense, formula)
    }
  }

  test("the sparse read is genuinely sparse (the law above is not vacuous)") {
    val sheet = Sheet(SheetName.unsafe("S")).put(ARef.from0(1, 1), CellValue.Number(BigDecimal(3)))
    evalWith(ArrayArithmetic.RangeRead.Sparse, sheet, "=A1:C5*2") match
      case Right(ar: ArrayResult) =>
        assert(!ar.isDense, ar.toString)
        assertEquals((ar.boxTop, ar.boxLeft, ar.boxRows, ar.boxCols), (1, 1, 1, 1))
        assertEquals(ar.fill, CellValue.Number(BigDecimal(0)))
        assertEquals(Right(ar), evalWith(ArrayArithmetic.RangeRead.Dense, sheet, "=A1:C5*2"))
      case other => fail(s"expected an array, got $other")
  }
