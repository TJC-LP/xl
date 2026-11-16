package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * #758: INDEX quotes its call (`INDEX(PROBE(), 5)`) in its error messages only, so it prints the
 * call on the error path alone. A selection that succeeds never prints its operand: printing a
 * large array constant would cost as much as the array.
 */
class IndexCallTextSpec extends FunSuite:

  private def N(n: String): CellValue = CellValue.Number(BigDecimal(n))

  private val dataName = SheetName.unsafe("Data")

  private val data: Sheet = Sheet(dataName)
    .put(ref"A1", N("1"))
    .put(ref"A2", N("2"))
    .put(ref"A3", N("3"))

  private val book: Workbook = Workbook(Vector(data))

  /**
   * INDEX(PROBE(), row) over a test function PROBE() that returns the column {1;2;3} and counts how
   * often the printer renders it.
   */
  private final class PrintCounter:
    private val renders = new java.util.concurrent.atomic.AtomicInteger(0)
    def count: Int = renders.get
    private val probe: FunctionSpec[ArrayResult] { type Args = EmptyTuple } =
      FunctionSpec.simple[ArrayResult, EmptyTuple](
        "PROBE",
        Arity.none,
        renderFn = Some { (_, _) =>
          renders.incrementAndGet()
          "PROBE()"
        }
      )((_, _) => Right(ArrayResult(Vector(Vector(N("1")), Vector(N("2")), Vector(N("3"))))))
    @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
    def index(row: Int): Either[EvalError, ArrayResult] =
      val operand = TExpr.Call(probe, EmptyTuple).asInstanceOf[TExpr[Any]]
      Evaluator.instance.eval(
        TExpr.Call(FunctionSpecs.index, (Right(operand), TExpr.Lit(BigDecimal(row)), None, None)),
        data
      )

  test("#758: an INDEX selection that succeeds never prints its operand") {
    val counter = PrintCounter()
    assertEquals(counter.index(2), Right(ArrayResult(Vector(Vector(N("2"))))))
    assertEquals(counter.count, 0)
  }

  test("#758: an INDEX position past the array prints its operand once, in the message") {
    val counter = PrintCounter()
    assertEquals(
      counter.index(5),
      Left(
        EvalError.ErrorValue(
          CellError.Ref,
          Some(
            "INDEX: row_num 5 is out of bounds (array has 3 rows, valid range: 1-3): INDEX(PROBE(), 5)"
          )
        )
      )
    )
    assertEquals(counter.count, 1)
  }

  test("#758: INDEX's error messages quote the call as written") {
    List(
      "INDEX({1;2;3},5)" -> "INDEX: row_num 5 is out of bounds (array has 3 rows, valid range: 1-3): INDEX({1;2;3}, 5)",
      "INDEX({1;2;3},-1)" -> "INDEX: row_num -1 is negative: INDEX({1;2;3}, -1)",
      "INDEX({1,2;3,4},1,3)" -> "INDEX: col_num 3 is out of bounds (array has 2 columns, valid range: 1-2): INDEX({1,2;3,4}, 1, 3)",
      "INDEX(Data!A1:A3,1,1,0)" -> "INDEX: area_num 0 is below 1 (INDEX(Data!A1:A3, 1, 1, 0))",
      "INDEX(Data!A1:A3,1,1,2)" -> "INDEX: area_num 2 is past the reference's 1 area(s) (INDEX(Data!A1:A3, 1, 1, 2))"
    ).foreach { (formula, message) =>
      val expr = FormulaParser.parse(s"=$formula").fold(e => fail(s"$formula: $e"), identity)
      Evaluator.instance.eval(expr, data, workbook = Some(book)) match
        case Left(EvalError.ErrorValue(_, Some(got))) => assertEquals(got, message, formula)
        case other => fail(s"$formula: expected an INDEX error message, got $other")
    }
  }
