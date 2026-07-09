package com.tjclp.xl.formula

import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.forAll

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.codec.CodecError
import com.tjclp.xl.formula.eval.{ArrayArithmetic, ArrayResult, EvalError}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

/**
 * GH-337: errors propagate ELEMENTWISE through array operations instead of failing the whole
 * formula.
 *
 * Excel semantics: `(A1:A3<5)*1` with a #DIV/0! cell at A2 yields the array {1, #DIV/0!, 1} — the
 * error only surfaces when an aggregation consumes that element. Pre-GH-337 the comparison refused
 * with a whole-formula Left ("cannot compare #REF! error value"), which was total and deterministic
 * but not Excel.
 *
 * The demotion table lives in [[EvalError.toCellError]]; the carriage lives in the ArrayArithmetic
 * broadcasts (compare, arithmetic, IF). CircularRef stays a fatal Left.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class ArrayErrorPropagationSpec extends ScalaCheckSuite:

  // ===== Commit 1: the EvalError -> CellError demotion table =====

  test("GH-337: toCellError maps each EvalError to its carried Excel error") {
    val a1 = ARef.from0(0, 0)
    val table: List[(EvalError, Option[CellError])] = List(
      EvalError.DivByZero("10", "0") -> Some(CellError.Div0),
      EvalError.RefError(a1, "cell not found") -> Some(CellError.Ref),
      EvalError.TypeMismatch("arithmetic", "number", "text: x") -> Some(CellError.Value),
      EvalError.CodecFailed(a1, CodecError.TypeMismatch("Numeric", CellValue.Text("x"))) -> Some(
        CellError.Value
      ),
      EvalError.EvalFailed("anything", None) -> Some(CellError.Value),
      // CircularRef is a structural failure of the sheet, never an element-local one: it must
      // stay a fatal Left rather than demote to an error element.
      EvalError.CircularRef(List(a1)) -> None
    )
    table.foreach { case (evalError, expected) =>
      assertEquals(EvalError.toCellError(evalError), expected, s"toCellError($evalError)")
    }
  }

  // ===== Helpers =====

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private val div0 = CellValue.Error(CellError.Div0)
  private val refErr = CellValue.Error(CellError.Ref)
  private val na = CellValue.Error(CellError.NA)
  private val valueErr = CellValue.Error(CellError.Value)

  /** 1xN row array. */
  private def rowOf(values: CellValue*): ArrayResult = ArrayResult(Vector(values.toVector))

  /** Nx1 column array. */
  private def colOf(values: CellValue*): ArrayResult =
    ArrayResult(values.toVector.map(Vector(_)))

  private def elementsOf(result: Either[EvalError, ArrayResult])(implicit
    loc: munit.Location
  ): Vector[CellValue] =
    result match
      case Right(ar) => ar.values.flatten
      case Left(err) => fail(s"expected Right(ArrayResult), got Left($err)")

  // ===== Commit 2: compare carriage =====

  test("GH-337: ordered compare carries error elements, left operand's error wins") {
    val left = rowOf(num(1), refErr, div0)
    val right = rowOf(div0, num(2), na)
    val result = ArrayArithmetic.broadcastOrderedCompare(left, right, _ < 0)
    assertEquals(elementsOf(result), Vector(div0, refErr, div0))
  }

  test("GH-337: equality compare carries error elements (array-array and array-scalar)") {
    val arrArr =
      ArrayArithmetic.broadcastEqualityCompare(
        rowOf(num(1), div0),
        Left(rowOf(num(1), num(2))),
        false
      )
    assertEquals(elementsOf(arrArr), Vector(CellValue.Bool(true), div0))

    val arrScalar =
      ArrayArithmetic.broadcastEqualityCompare(rowOf(num(1), num(2)), Right(na), false)
    assertEquals(elementsOf(arrScalar), Vector(na, na))

    // Left operand's error wins over the scalar's, and negate doesn't flip errors
    val leftWins = ArrayArithmetic.broadcastEqualityCompare(rowOf(div0, num(2)), Right(na), true)
    assertEquals(elementsOf(leftWins), Vector(div0, na))
  }

  test("GH-337: carried errors recurse through cached formula values") {
    val cachedErr = CellValue.Formula("=1/0", Some(div0))
    val compared =
      ArrayArithmetic.broadcastOrderedCompare(rowOf(cachedErr), rowOf(num(1)), _ < 0)
    assertEquals(elementsOf(compared), Vector(div0))

    val summed = ArrayArithmetic.broadcast(
      ArrayArithmetic.ArrayOperand.Array(rowOf(cachedErr)),
      ArrayArithmetic.ArrayOperand.Scalar(BigDecimal(1)),
      ArrayArithmetic.add
    )
    summed match
      case Right(ArrayArithmetic.ArrayOperand.Array(out)) => assertEquals(out(0, 0), div0)
      case other => fail(s"expected Array result, got $other")
  }

  // ===== Commit 2: arithmetic carriage =====

  test("GH-337: arithmetic carries error elements and maps local failures per element") {
    // {1, #REF!, "abc", 0} entering 10/x: carried #REF! stays, text is #VALUE!, 0 is #DIV/0!
    val operand = rowOf(num(1), refErr, CellValue.Text("abc"), num(0))
    val result = ArrayArithmetic.broadcast(
      ArrayArithmetic.ArrayOperand.Scalar(BigDecimal(10)),
      ArrayArithmetic.ArrayOperand.Array(operand),
      ArrayArithmetic.div
    )
    result match
      case Right(ArrayArithmetic.ArrayOperand.Array(out)) =>
        assertEquals(out(0, 0), num(10))
        assertEquals(out(0, 1), refErr)
        assertEquals(out(0, 2), valueErr)
        assertEquals(out(0, 3), div0)
      case other => fail(s"expected Array result, got $other")
  }

  test("GH-337: array-array arithmetic: left operand's error wins") {
    val result = ArrayArithmetic.broadcast(
      ArrayArithmetic.ArrayOperand.Array(rowOf(refErr, num(2))),
      ArrayArithmetic.ArrayOperand.Array(rowOf(div0, na)),
      ArrayArithmetic.add
    )
    result match
      case Right(ArrayArithmetic.ArrayOperand.Array(out)) =>
        assertEquals(out(0, 0), refErr)
        assertEquals(out(0, 1), na)
      case other => fail(s"expected Array result, got $other")
  }

  test("GH-337: a 1x1 scalar error array broadcasts across the other operand (NA()+range shape)") {
    // The =NA()+A1:A3 shape: a scalar error operand enters as a 1x1 error array and broadcasts.
    val result = ArrayArithmetic.broadcast(
      ArrayArithmetic.ArrayOperand.Array(ArrayResult.single(na)),
      ArrayArithmetic.ArrayOperand.Array(colOf(num(1), num(2), num(3))),
      ArrayArithmetic.add
    )
    result match
      case Right(ArrayArithmetic.ArrayOperand.Array(out)) =>
        assertEquals(out.rows, 3)
        assertEquals(out.values.flatten, Vector(na, na, na))
      case other => fail(s"expected Array result, got $other")
  }

  test("GH-337: dimension mismatch stays a whole-broadcast Left even with error elements") {
    val left = rowOf(num(1), div0, num(3)) // 1x3
    val right = rowOf(num(1), num(2)) // 1x2
    assert(ArrayArithmetic.broadcastOrderedCompare(left, right, _ < 0).isLeft)
    assert(
      ArrayArithmetic
        .broadcast(
          ArrayArithmetic.ArrayOperand.Array(left),
          ArrayArithmetic.ArrayOperand.Array(right),
          ArrayArithmetic.mul
        )
        .isLeft
    )
    assert(ArrayArithmetic.broadcastEqualityCompare(left, Left(right), false).isLeft)
  }

  // ===== Commit 2: end-to-end shapes =====

  test("GH-337: =A1:A3<5 with an error cell spills {TRUE; #DIV/0!; TRUE}") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(1))
      .put(ref"A2", div0)
      .put(ref"A3", num(2))
    val result = sheet.evaluateArrayFormula("=A1:A3<5", ref"C1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, spill) = result.toOption.get
    assertEquals(spill.height, 3)
    assertEquals(updated(ref"C1").value, CellValue.Bool(true))
    assertEquals(updated(ref"C2").value, div0)
    assertEquals(updated(ref"C3").value, CellValue.Bool(true))
  }

  test("GH-337: =10/A1:A3 with a zero yields an elementwise #DIV/0!") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(1))
      .put(ref"A2", num(0))
      .put(ref"A3", num(2))
    val result = sheet.evaluateArrayFormula("=10/A1:A3", ref"C1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, _) = result.toOption.get
    assertEquals(updated(ref"C1").value, num(10))
    assertEquals(updated(ref"C2").value, div0)
    assertEquals(updated(ref"C3").value, num(5))
  }

  test("GH-337: text elements in array arithmetic become #VALUE! elements") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("abc"))
      .put(ref"A2", num(3))
    val result = sheet.evaluateArrayFormula("=A1:A2*2", ref"C1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, _) = result.toOption.get
    assertEquals(updated(ref"C1").value, valueErr)
    assertEquals(updated(ref"C2").value, num(6))
  }

  test("GH-337: comparing a range against a scalar error cell carries the error elementwise") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(1))
      .put(ref"A2", num(2))
      .put(ref"B1", na)
    val ordered = sheet.evaluateArrayFormula("=A1:A2<B1", ref"C1")
    assert(ordered.isRight, s"expected Right, got $ordered")
    val (updatedOrd, _) = ordered.toOption.get
    assertEquals(updatedOrd(ref"C1").value, na)
    assertEquals(updatedOrd(ref"C2").value, na)

    val equality = sheet.evaluateArrayFormula("=A1:A2=B1", ref"D1")
    assert(equality.isRight, s"expected Right, got $equality")
    val (updatedEq, _) = equality.toOption.get
    assertEquals(updatedEq(ref"D1").value, na)
    assertEquals(updatedEq(ref"D2").value, na)
  }

  test("GH-337: SCALAR comparison against an error cell still refuses (unchanged boundary)") {
    // Carriage is an array-path behavior: the scalar compare path keeps its clean Left.
    val sheet = Sheet("Test").put(ref"A1", num(1)).put(ref"B1", na)
    val result = sheet.evaluateFormula("=A1<B1")
    assert(result.isLeft, s"expected Left, got $result")
    assert(
      result.left.toOption.get.message.contains("cannot compare"),
      s"unexpected error: $result"
    )
  }

  test("GH-337: scalar entry collapses — error at top-left refuses, error elsewhere intersects") {
    // Implicit intersection takes the top-left element; a carried error there refuses cleanly,
    // while an error elsewhere in the array no longer poisons the collapsed scalar (improvement
    // over the pre-GH-337 whole-formula Left).
    val errTopLeft = Sheet("Test").put(ref"A1", div0).put(ref"A2", num(5))
    val collapsed = errTopLeft.evaluateFormula("=A1:A2*2")
    assert(collapsed.isLeft, s"expected Left, got $collapsed")
    assert(
      collapsed.left.toOption.get.message.contains("cannot coerce"),
      s"unexpected error: $collapsed"
    )

    val errBelow = Sheet("Test").put(ref"A1", num(5)).put(ref"A2", div0)
    assertEquals(errBelow.evaluateFormula("=A1:A2*2"), Right(num(10)))
  }

  // ===== Commit 3: broadcastIf condition carriage =====

  test("GH-337: broadcastIf carries a condition element's error to the output position") {
    val cond = colOf(CellValue.Bool(true), div0, CellValue.Bool(false))
    val ifTrue = colOf(num(1), num(2), num(3))
    val ifFalse = ArrayResult.single(num(99))
    val result = ArrayArithmetic.broadcastIf(cond, ifTrue, ifFalse)
    assertEquals(elementsOf(result), Vector(num(1), div0, num(99)))
  }

  test("GH-337: broadcastIf demotes a text condition element to #VALUE! (was whole-formula Left)") {
    val cond = colOf(CellValue.Bool(true), CellValue.Text("x"))
    val result =
      ArrayArithmetic.broadcastIf(cond, ArrayResult.single(num(1)), ArrayResult.single(num(2)))
    assertEquals(elementsOf(result), Vector(num(1), valueErr))
  }

  test("GH-337: broadcastIf condition errors mask branch errors positionally") {
    // The condition's own error wins at its position; other positions still select normally,
    // including selecting error elements from a branch.
    val cond = colOf(na, CellValue.Bool(false))
    val result = ArrayArithmetic.broadcastIf(
      cond,
      ArrayResult.single(num(1)),
      ArrayResult.single(div0)
    )
    assertEquals(elementsOf(result), Vector(na, div0))
  }

  test("GH-337: =IF over a condition range with an error cell selects errors positionally") {
    val sheet = Sheet("Test")
      .put(ref"A1", num(1))
      .put(ref"A2", div0)
      .put(ref"A3", num(-1))
    val result = sheet.evaluateArrayFormula("=IF(A1:A3>0,1,2)", ref"C1")
    assert(result.isRight, s"expected Right, got $result")
    val (updated, _) = result.toOption.get
    assertEquals(updated(ref"C1").value, num(1))
    assertEquals(updated(ref"C2").value, div0)
    assertEquals(updated(ref"C3").value, num(2))
  }

  property("GH-337: broadcastIf never Lefts for element-local condition failures") {
    val genCondElement: Gen[CellValue] = Gen.frequency(
      3 -> Gen.oneOf(
        Gen.oneOf(true, false).map(CellValue.Bool(_)),
        Gen.chooseNum(-5, 5).map(n => CellValue.Number(BigDecimal(n))),
        Gen.const(CellValue.Empty)
      ),
      1 -> genErrorValue,
      1 -> Gen.oneOf("abc", "").map(CellValue.Text(_))
    )
    forAll(
      Gen.choose(1, 3).flatMap(n => Gen.listOfN(n, genCondElement)),
      Gen.chooseNum(-50, 50),
      Gen.chooseNum(-50, 50)
    ) { (condElems, t, f) =>
      val cond = ArrayResult(condElems.toVector.map(Vector(_)))
      val result = ArrayArithmetic.broadcastIf(
        cond,
        ArrayResult.single(CellValue.Number(BigDecimal(t))),
        ArrayResult.single(CellValue.Number(BigDecimal(f)))
      )
      result match
        case Right(out) =>
          Prop.all(
            condElems.zipWithIndex.map { case (condCV, idx) =>
              val expected = condCV match
                case CellValue.Error(err) => CellValue.Error(err)
                case CellValue.Text(_) => valueErr
                case CellValue.Bool(b) =>
                  if b then CellValue.Number(BigDecimal(t)) else CellValue.Number(BigDecimal(f))
                case CellValue.Number(n) =>
                  if n.signum != 0 then CellValue.Number(BigDecimal(t))
                  else CellValue.Number(BigDecimal(f))
                case _ => CellValue.Number(BigDecimal(f)) // Empty is FALSE
              Prop(out(idx, 0) == expected) :| s"idx=$idx cond=$condCV out=${out(idx, 0)}"
            }*
          )
        case Left(err) => Prop.falsified :| s"expected Right, got Left($err)"
    }
  }

  // ===== Commit 2: property tests =====

  private val genPlainValue: Gen[CellValue] = Gen.oneOf(
    Gen.chooseNum(-50, 50).map(n => CellValue.Number(BigDecimal(n))),
    Gen.oneOf("abc", "42", "").map(CellValue.Text(_)),
    Gen.oneOf(true, false).map(CellValue.Bool(_)),
    Gen.const(CellValue.Empty)
  )

  private val genErrorValue: Gen[CellValue] =
    Gen.oneOf(CellError.values.toIndexedSeq).map(CellValue.Error(_))

  private val genElement: Gen[CellValue] =
    Gen.frequency(4 -> genPlainValue, 1 -> genErrorValue)

  private def genArrayOf(gen: Gen[CellValue], rows: Int, cols: Int): Gen[ArrayResult] =
    Gen
      .listOfN(rows * cols, gen)
      .map(values => ArrayResult.fromFlat(values.toVector, rows, cols))

  /** A pair of broadcast-compatible arrays (each axis matches the output or is 1). */
  private def genBroadcastablePair(gen: Gen[CellValue]): Gen[(ArrayResult, ArrayResult)] =
    for
      outRows <- Gen.choose(1, 3)
      outCols <- Gen.choose(1, 3)
      lRows <- Gen.oneOf(1, outRows)
      lCols <- Gen.oneOf(1, outCols)
      rRows <- Gen.oneOf(1, outRows)
      rCols <- Gen.oneOf(1, outCols)
      left <- genArrayOf(gen, lRows, lCols)
      right <- genArrayOf(gen, rRows, rCols)
    yield (left, right)

  private val genOp: Gen[ArrayArithmetic.BinaryOp] =
    Gen.oneOf(ArrayArithmetic.add, ArrayArithmetic.sub, ArrayArithmetic.mul, ArrayArithmetic.div)

  /** Broadcast indexing: a 1-sized axis repeats. */
  private def broadcastAt(a: ArrayResult, row: Int, col: Int): CellValue =
    a.values(if a.rows == 1 then 0 else row)(if a.cols == 1 then 0 else col)

  private def isErrorElement(cv: CellValue): Boolean = cv match
    case CellValue.Error(_) => true
    case _ => false

  property("GH-337: broadcast compare/arith never Left for element-local failures") {
    forAll(genBroadcastablePair(genElement), genOp) { case ((left, right), op) =>
      val compare = ArrayArithmetic.broadcastOrderedCompare(left, right, _ < 0)
      val equality = ArrayArithmetic.broadcastEqualityCompare(left, Left(right), false)
      val arith = ArrayArithmetic.broadcast(
        ArrayArithmetic.ArrayOperand.Array(left),
        ArrayArithmetic.ArrayOperand.Array(right),
        op
      )
      Prop(compare.isRight && equality.isRight && arith.isRight) :| s"$compare / $equality / $arith"
    }
  }

  property("GH-337: arithmetic output element is Error iff an input carries one or the op fails") {
    forAll(genBroadcastablePair(genElement), genOp) { case ((left, right), op) =>
      ArrayArithmetic.broadcast(
        ArrayArithmetic.ArrayOperand.Array(left),
        ArrayArithmetic.ArrayOperand.Array(right),
        op
      ) match
        case Right(ArrayArithmetic.ArrayOperand.Array(out)) =>
          Prop.all(
            (for
              row <- 0 until out.rows
              col <- 0 until out.cols
            yield
              val l = broadcastAt(left, row, col)
              val r = broadcastAt(right, row, col)
              val localFailure =
                isErrorElement(l) || isErrorElement(r) ||
                  (for
                    ln <- ArrayArithmetic.cellValueToNumeric(l)
                    rn <- ArrayArithmetic.cellValueToNumeric(r)
                    v <- op(ln, rn)
                  yield v).isLeft
              Prop(isErrorElement(out(row, col)) == localFailure) :|
                s"($row,$col): l=$l r=$r out=${out(row, col)}"
            )*
          )
        case other => Prop.falsified :| s"expected Array result, got $other"
    }
  }

  property("GH-337: compare output element is Error iff an input element carries an error") {
    forAll(genBroadcastablePair(genElement)) { case (left, right) =>
      ArrayArithmetic.broadcastOrderedCompare(left, right, _ < 0) match
        case Right(out) =>
          Prop.all(
            (for
              row <- 0 until out.rows
              col <- 0 until out.cols
            yield
              val carried =
                isErrorElement(broadcastAt(left, row, col)) ||
                  isErrorElement(broadcastAt(right, row, col))
              Prop(isErrorElement(out(row, col)) == carried) :| s"($row,$col)"
            )*
          )
        case Left(err) => Prop.falsified :| s"expected Right, got Left($err)"
    }
  }

  property("GH-337: error-free arithmetic agrees with the numeric-matrix reference") {
    forAll(genBroadcastablePair(genPlainValue), genOp) { case ((left, right), op) =>
      val reference: Either[EvalError, Vector[Vector[CellValue]]] =
        for
          lNums <- ArrayArithmetic.arrayToNumeric(left)
          rNums <- ArrayArithmetic.arrayToNumeric(right)
          rows = math.max(left.rows, right.rows)
          cols = math.max(left.cols, right.cols)
          out <- (0 until rows).foldLeft[Either[EvalError, Vector[Vector[CellValue]]]](
            Right(Vector.empty)
          ) { (accRows, row) =>
            accRows.flatMap { rowsSoFar =>
              (0 until cols)
                .foldLeft[Either[EvalError, Vector[CellValue]]](Right(Vector.empty)) {
                  (accCols, col) =>
                    accCols.flatMap { colsSoFar =>
                      val l = lNums(if left.rows == 1 then 0 else row)(
                        if left.cols == 1 then 0 else col
                      )
                      val r = rNums(if right.rows == 1 then 0 else row)(
                        if right.cols == 1 then 0 else col
                      )
                      op(l, r).map(v => colsSoFar :+ CellValue.Number(v))
                    }
                }
                .map(rowsSoFar :+ _)
            }
          }
        yield out
      reference match
        case Right(expected) =>
          ArrayArithmetic.broadcast(
            ArrayArithmetic.ArrayOperand.Array(left),
            ArrayArithmetic.ArrayOperand.Array(right),
            op
          ) match
            case Right(ArrayArithmetic.ArrayOperand.Array(out)) =>
              Prop(out.values == expected) :| s"out=${out.values} expected=$expected"
            case other => Prop.falsified :| s"expected Array result, got $other"
        // Reference itself fails (text/div-zero): covered by the local-failure properties above
        case Left(_) => Prop.passed
    }
  }

  property("GH-337: error-free compares agree with elementwise compareCellValues") {
    forAll(genBroadcastablePair(genPlainValue)) { case (left, right) =>
      ArrayArithmetic.broadcastOrderedCompare(left, right, _ < 0) match
        case Right(out) =>
          Prop.all(
            (for
              row <- 0 until out.rows
              col <- 0 until out.cols
            yield
              val expected = ArrayArithmetic
                .compareCellValues(broadcastAt(left, row, col), broadcastAt(right, row, col))
                .map(sign => CellValue.Bool(sign < 0))
              Prop(expected == Right(out(row, col))) :| s"($row,$col)"
            )*
          )
        case Left(err) => Prop.falsified :| s"expected Right, got Left($err)"
    }
  }
