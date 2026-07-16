package com.tjclp.xl.formula

import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.forAll

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{ArrayArithmetic, ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.SheetName

/**
 * GH-344: laws for Excel error VALUES as first-class evaluation results.
 *
 * The two-table law: [[EvalError.toCellError]] (permissive element demotion, GH-337) and
 * [[EvalError.toErrorValue]] (strict boundary promotion) are distinct judgments. An
 * `EvalError.ErrorValue` travels the Left channel (Either short-circuit = Excel strict-position
 * absorption) and promotes to `Right(CellValue.Error(code))` only at the CellValue boundaries
 * (evaluateFormula / evaluateCell / evaluateArrayFormula / cross-sheet reads). Infrastructure
 * failures (parse, missing workbook/sheet, unknown function, cycles, unbound names) must never
 * launder into error cells — they stay loud Lefts.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class ErrorValueLawsSpec extends ScalaCheckSuite:

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  private val genError: Gen[CellError] = Gen.oneOf(CellError.values.toIndexedSeq)

  /** A1 bears the generated error; B1/B2 are plain numbers for mixed shapes. */
  private def sheetWithError(err: CellError): Sheet =
    Sheet("Test")
      .put(ref"A1", CellValue.Error(err))
      .put(ref"B1", num(3))
      .put(ref"B2", num(4))

  /**
   * Error-VALUE producers: formula shapes that consume A1's error through a strict position. Every
   * one must surface the error as a catchable value, never a loud Left.
   */
  private val errorProducers: List[String] = List(
    "=A1", // bare reference
    "=A1+1", // scalar arithmetic over a typed ref decode
    "=1+A1",
    "=A1*B1",
    "=ABS(A1)", // typed numeric argument position
    "=1<A1", // scalar ordered comparison
    "=A1>=B1",
    "=1=A1", // scalar equality
    "=A1<>1",
    "=\"x\"&A1", // text concatenation
    "=IF(A1,1,2)", // scalar condition position
    "=SWITCH(A1,1,\"one\")", // SWITCH target
    "=LET(x,A1,x)" // LET binding used in the body
  )

  // ===== The two tables =====

  test("GH-344: toErrorValue promotes only ErrorValue and DivByZero") {
    val a1 = ARef.from0(0, 0)
    assertEquals(
      EvalError.toErrorValue(EvalError.ErrorValue(CellError.Num, None)),
      Some(CellError.Num)
    )
    assertEquals(
      EvalError.toErrorValue(EvalError.DivByZero("1", "0")),
      Some(CellError.Div0)
    )
    val loud: List[EvalError] = List(
      EvalError.RefError(a1, "cell not found"),
      EvalError.TypeMismatch("op", "number", "text"),
      EvalError.CodecFailed(
        a1,
        com.tjclp.xl.codec.CodecError.TypeMismatch("Numeric", CellValue.Text("x"))
      ),
      EvalError.EvalFailed("anything", None),
      EvalError.CircularRef(List(a1))
    )
    loud.foreach(e => assertEquals(EvalError.toErrorValue(e), None, s"toErrorValue($e)"))
  }

  test("GH-344: toCellError gains the ErrorValue identity arm (code fidelity)") {
    CellError.values.foreach { err =>
      assertEquals(EvalError.toCellError(EvalError.ErrorValue(err, None)), Some(err))
    }
  }

  // ===== L1: boundary law =====

  property("L1: every error producer yields Right(CellValue.Error(_)) at evaluateFormula") {
    forAll(genError) { err =>
      val sheet = sheetWithError(err)
      Prop.all(
        errorProducers.map { f =>
          sheet.evaluateFormula(f) match
            case Right(CellValue.Error(_)) => Prop.passed :| f
            case other => Prop.falsified :| s"$f with $err: expected Right(Error), got $other"
        }*
      )
    }
  }

  property("L1: code fidelity — pass-through shapes preserve the exact error code") {
    val passThrough = List("=A1", "=A1+1", "=ABS(A1)", "=1<A1", "=1=A1", "=\"x\"&A1")
    forAll(genError) { err =>
      val sheet = sheetWithError(err)
      Prop.all(
        passThrough.map { f =>
          Prop(sheet.evaluateFormula(f) == Right(CellValue.Error(err))) :|
            s"$f with $err: got ${sheet.evaluateFormula(f)}"
        }*
      )
    }
  }

  test("L1: =1/0 promotes DivByZero at the boundary (headline contract)") {
    val sheet = Sheet("Test")
    assertEquals(sheet.evaluateFormula("=1/0"), Right(CellValue.Error(CellError.Div0)))
    assertEquals(sheet.evaluateFormula("=10/(5-5)"), Right(CellValue.Error(CellError.Div0)))
  }

  test("L1 canaries: infrastructure failures stay loud Lefts") {
    val sheet = Sheet("Test").put(ref"A1", num(1))
    // parse failure
    assert(sheet.evaluateFormula("=SUM(").isLeft, "parse failure must stay Left")
    // missing workbook context for a cross-sheet ref
    assert(sheet.evaluateFormula("=Other!A1").isLeft, "missing workbook must stay Left")
    // missing sheet with workbook context
    val wb = Workbook(sheet)
    assert(
      wb.evaluateFormula("=Nope!A1", SheetName.unsafe("Test")).isLeft,
      "missing sheet must stay Left"
    )
    // unknown function
    assert(sheet.evaluateFormula("=NOSUCHFN(1)").isLeft, "unknown function must stay Left")
    // unbound defined name
    assert(sheet.evaluateFormula("=LET(x, 1, y)").isLeft, "unbound name must stay Left")
    // cycle: structural, fatal
    val cyclic = Sheet("Test")
      .put(ref"A1", CellValue.Formula("=B1", None))
      .put(ref"B1", CellValue.Formula("=A1", None))
    assert(cyclic.evaluateWithDependencyCheck().isLeft, "cycle must stay Left")
  }

  // ===== L2: catchability =====

  property("L2: IFERROR catches and ISERROR detects every producer") {
    forAll(genError) { err =>
      val sheet = sheetWithError(err)
      Prop.all(
        errorProducers.flatMap { f =>
          val body = f.stripPrefix("=")
          List(
            Prop(sheet.evaluateFormula(s"=IFERROR($body, 42)") == Right(num(42))) :|
              s"IFERROR($body) with $err: got ${sheet.evaluateFormula(s"=IFERROR($body, 42)")}",
            Prop(sheet.evaluateFormula(s"=ISERROR($body)") == Right(CellValue.Bool(true))) :|
              s"ISERROR($body) with $err"
          )
        }*
      )
    }
  }

  test("L2: ISERR excludes #N/A on every channel") {
    val naSheet = sheetWithError(CellError.NA)
    assertEquals(naSheet.evaluateFormula("=ISERR(A1)"), Right(CellValue.Bool(false)))
    assertEquals(naSheet.evaluateFormula("=ISERR(1<A1)"), Right(CellValue.Bool(false)))
    val divSheet = sheetWithError(CellError.Div0)
    assertEquals(divSheet.evaluateFormula("=ISERR(A1)"), Right(CellValue.Bool(true)))
    assertEquals(divSheet.evaluateFormula("=ISERR(1<A1)"), Right(CellValue.Bool(true)))
  }

  // ===== L3: absorption, precedence, argument-order determinism =====

  test("L3: left operand's error wins in scalar arithmetic, comparison, and equality") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Error(CellError.Ref))
      .put(ref"B1", CellValue.Error(CellError.Div0))
    assertEquals(sheet.evaluateFormula("=A1+B1"), Right(CellValue.Error(CellError.Ref)))
    assertEquals(sheet.evaluateFormula("=B1+A1"), Right(CellValue.Error(CellError.Div0)))
    assertEquals(sheet.evaluateFormula("=A1<B1"), Right(CellValue.Error(CellError.Ref)))
    assertEquals(sheet.evaluateFormula("=A1=B1"), Right(CellValue.Error(CellError.Ref)))
    assertEquals(sheet.evaluateFormula("=A1&B1"), Right(CellValue.Error(CellError.Ref)))
  }

  test("L3: errors absorb through nested strict positions") {
    val sheet = sheetWithError(CellError.NA)
    assertEquals(sheet.evaluateFormula("=ABS(A1+1)*2"), Right(CellValue.Error(CellError.NA)))
    assertEquals(sheet.evaluateFormula("=(1<A1)=TRUE"), Right(CellValue.Error(CellError.NA)))
  }

  // ===== L5: broadcasts are TOTAL in Right — mismatched dims pad with #N/A =====

  property("L5: broadcast compare/arith/equality never Left; beyond-extent positions are #N/A") {
    val genPlain: Gen[CellValue] =
      Gen.chooseNum(-9, 9).map(n => CellValue.Number(BigDecimal(n)))
    def genArray(rows: Int, cols: Int): Gen[ArrayResult] =
      Gen.listOfN(rows * cols, genPlain).map(vs => ArrayResult.fromFlat(vs.toVector, rows, cols))
    val genDims = Gen.choose(1, 4)
    forAll(genDims, genDims, genDims, genDims) { (lr, lc, rr, rc) =>
      forAll(genArray(lr, lc), genArray(rr, rc)) { (left, right) =>
        def outDim(l: Int, r: Int): Int = if l == 1 then r else if r == 1 then l else math.max(l, r)
        val outRows = outDim(lr, rr)
        val outCols = outDim(lc, rc)
        def padded(aRows: Int, aCols: Int, row: Int, col: Int): Boolean =
          (aRows != 1 && row >= aRows) || (aCols != 1 && col >= aCols)
        val compare = ArrayArithmetic.broadcastOrderedCompare(left, right, _ < 0)
        val arith = ArrayArithmetic.broadcast(
          ArrayArithmetic.ArrayOperand.Array(left),
          ArrayArithmetic.ArrayOperand.Array(right),
          ArrayArithmetic.add
        )
        val equality = ArrayArithmetic.broadcastEqualityCompare(left, Left(right), false)
        (compare, arith, equality) match
          case (Right(cmp), Right(ArrayArithmetic.ArrayOperand.Array(sum)), Right(eq)) =>
            Prop.all(
              (for
                row <- 0 until outRows
                col <- 0 until outCols
              yield
                val isPadded = padded(lr, lc, row, col) || padded(rr, rc, row, col)
                val expectNA =
                  if isPadded then
                    cmp(row, col) == CellValue.Error(CellError.NA) &&
                    sum(row, col) == CellValue.Error(CellError.NA) &&
                    eq(row, col) == CellValue.Error(CellError.NA)
                  else
                    cmp(row, col) != CellValue.Error(CellError.NA) &&
                    sum(row, col) != CellValue.Error(CellError.NA) &&
                    eq(row, col) != CellValue.Error(CellError.NA)
                Prop(expectNA) :| s"($row,$col) padded=$isPadded dims=($lr,$lc)x($rr,$rc)"
              )*
            ) && Prop(cmp.rows == outRows && cmp.cols == outCols) :| "output dims"
          case other => Prop.falsified :| s"expected all Right, got $other"
      }
    }
  }

  // ===== L4: Aggregate(id, range) ≡ Call(spec, range), including error policy =====

  property("L4: the TExpr.Aggregate node and the registry Call agree over error-bearing ranges") {
    import com.tjclp.xl.formula.parser.FormulaParser
    val aggregates = List("SUM", "COUNT", "COUNTA", "COUNTBLANK", "MIN", "MAX", "AVERAGE")
    forAll(genError) { err =>
      val sheet = Sheet("Test")
        .put(ref"A1", num(1))
        .put(ref"A2", CellValue.Error(err))
        .put(ref"A3", num(2))
      val range = com.tjclp.xl.addressing.CellRange(ref"A1", ref"A3")
      Prop.all(
        aggregates.map { name =>
          val node = TExpr.Aggregate(name, TExpr.RangeLocation.Local(range))
          val nodeResult = Evaluator.eval(node, sheet)
          val callResult = FormulaParser
            .parse(s"=$name(A1:A3)")
            .left
            .map(pe => EvalError.EvalFailed(pe.toString, None): EvalError)
            .flatMap(expr => Evaluator.eval(expr, sheet))
          val agree = (nodeResult, callResult) match
            case (Right(a), Right(b)) => a == b
            case (Left(a), Left(b)) =>
              EvalError.toErrorValue(a) == EvalError.toErrorValue(b) &&
              EvalError.toErrorValue(a).isDefined
            case _ => false
          Prop(agree) :| s"$name with $err: node=$nodeResult call=$callResult"
        }*
      )
    }
  }
