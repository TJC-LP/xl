package com.tjclp.xl.formula

import java.time.{LocalDate, LocalDateTime}

import scala.util.boundary

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * Totality at the evaluation boundary (#681 item 5): a throwable escaping a function body is an
 * internal defect, not an Excel error value. The outermost evaluator frame turns NonFatal throws
 * and StackOverflowError into `EvalError.EvalFailed` naming the defect; OutOfMemoryError and
 * InterruptedException still propagate (MemoryGuard's RESOURCE_LIMIT, recalc cancellation), and no
 * formula construct (IFERROR, ISERROR) can swallow a defect because only the outermost frame
 * catches.
 */
class EvaluationTotalitySpec extends FunSuite:

  private val name = SheetName.unsafe("Sheet1")

  private def throwingClock(t: () => Throwable): Clock = new Clock:
    def today(): LocalDate = throw t()
    def now(): LocalDateTime = throw t()

  private def throwingRng(t: () => Throwable): Rng = new Rng:
    def nextDouble(): Double = throw t()

  private val boom = throwingClock(() => new IllegalStateException("clock down"))
  private val deep = throwingClock(() => new StackOverflowError())

  private val sheet: Sheet = Sheet(name)
    .put(ref"A1", CellValue.Number(BigDecimal(1)))
    .put(ref"B1", CellValue.Formula("TODAY()+1"))
    .put(ref"C1", CellValue.Formula("A1+1"))

  private def parse(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(e => fail(s"$formula does not parse: $e"), identity)

  private val defectAtB1 =
    "Evaluation threw java.lang.IllegalStateException: clock down at B1 — an internal evaluator " +
      "defect, not an Excel error value; please report it with the formula"

  private val stackAtB1 =
    "Evaluation exhausted the stack at B1: the formula's nesting or reference chain is too deep"

  private def factories: List[(String, Evaluator)] = List(
    "instance" -> Evaluator.instance,
    "instance(rng)" -> Evaluator.instance(Rng.system),
    "instance(workbookPath)" -> Evaluator.instance(Option.empty[String]),
    "recalculationInstance" ->
      Evaluator.recalculationInstance(Rng.system, new Evaluator.AggregateMemo),
    "arrayInstance" -> Evaluator.arrayInstance,
    "arrayInstance(rng)" -> Evaluator.arrayInstance(Rng.system)
  )

  test("a NonFatal throw in a function body is EvalFailed with the defect wording, every factory") {
    factories.foreach { (label, evaluator) =>
      assertEquals(
        evaluator.eval(parse("=TODAY()+1"), sheet, boom, None, Some(ref"B1")),
        Left(EvalError.EvalFailed(defectAtB1)),
        label
      )
    }
  }

  test("GH-713: a defect inside a computed range or XLOOKUP's reference is contained too") {
    factories.foreach { (label, evaluator) =>
      List(
        "=SUM(A1:INDEX(A1:A3,TODAY()))",
        "=ROWS(XLOOKUP(TODAY(),A1:A3,A1:A3):A1)",
        "=A1:INDEX(A1:A3,TODAY())"
      ).foreach { formula =>
        assertEquals(
          evaluator.eval(parse(formula), sheet, boom, None, Some(ref"B1")),
          Left(EvalError.EvalFailed(defectAtB1)),
          s"$label $formula"
        )
      }
    }
  }

  test("a StackOverflowError is EvalFailed naming the exhausted stack, every factory") {
    factories.foreach { (label, evaluator) =>
      assertEquals(
        evaluator.eval(parse("=TODAY()+1"), sheet, deep, None, Some(ref"B1")),
        Left(EvalError.EvalFailed(stackAtB1)),
        label
      )
    }
  }

  test("a throwing randomness source is contained the same way") {
    val rng = throwingRng(() => new ClassCastException("not a number"))
    Evaluator.instance(rng).eval(parse("=RAND()*2"), sheet) match
      case Left(EvalError.EvalFailed(reason, _)) =>
        assert(reason.startsWith("Evaluation threw java.lang.ClassCastException: not a number"))
        assert(reason.contains("internal evaluator defect"), reason)
        assert(!reason.contains(" at "), s"no cell position was given: $reason")
      case other => fail(s"expected a contained defect, got $other")
  }

  test("the defect message caps the throwable's message and tolerates a null one") {
    val long = throwingClock(() => new IllegalStateException("x" * 500))
    Evaluator.instance.eval(parse("=TODAY()"), sheet, long, None, None) match
      case Left(EvalError.EvalFailed(reason, _)) =>
        assert(reason.contains("x" * 200 + "…"), reason)
        assert(!reason.contains("x" * 201), reason)
      case other => fail(s"expected a contained defect, got $other")
    val silent = throwingClock(() => new IllegalStateException())
    assertEquals(
      Evaluator.instance.eval(parse("=TODAY()"), sheet, silent, None, Some(ref"B1")),
      Left(
        EvalError.EvalFailed(
          "Evaluation threw java.lang.IllegalStateException at B1 — an internal evaluator " +
            "defect, not an Excel error value; please report it with the formula"
        )
      )
    )
  }

  test("IFERROR and ISERROR never swallow a defect") {
    List("=IFERROR(TODAY(),0)", "=ISERROR(TODAY())", "=IFERROR(1/0,TODAY())").foreach { f =>
      Evaluator.instance.eval(parse(f), sheet, boom, None, Some(ref"B1")) match
        case Left(EvalError.EvalFailed(reason, _)) => assertEquals(reason, defectAtB1, f)
        case other => fail(s"$f: a defect must stay a defect, got $other")
    }
    assert(sheet.evaluateFormula("=IFERROR(TODAY(),0)", boom).isLeft)
  }

  test("the companion Evaluator.eval is guarded too") {
    Evaluator.eval(parse("=TODAY()"), sheet, boom) match
      case Left(EvalError.EvalFailed(reason, _)) =>
        assert(reason.contains("internal evaluator defect"))
      case other => fail(s"expected a contained defect, got $other")
  }

  test("every SheetEvaluator entry point reports the defect as a Left naming the formula") {
    def assertDefect(label: String, result: Either[com.tjclp.xl.error.XLError, Any]): Unit =
      result match
        case Left(err) =>
          assert(err.message.contains("internal evaluator defect"), s"$label: ${err.message}")
          assert(err.message.contains("TODAY()"), s"$label names the formula: ${err.message}")
        case Right(v) => fail(s"$label: expected a Left, got $v")
    assertDefect("evaluateFormula", sheet.evaluateFormula("=TODAY()+1", boom))
    assertDefect("evaluateCell", sheet.evaluateCell(ref"B1", boom))
    assertDefect("evaluateWithDependencyCheck", sheet.evaluateWithDependencyCheck(boom))
    assertDefect("evaluateForRange", sheet.evaluateForRange(CellRange(ref"A1", ref"C1"), boom))
    assertDefect("evaluateArrayFormula", sheet.evaluateArrayFormula("=TODAY()+1", ref"H1", boom))
    assertDefect(
      "wb.evaluateFormula",
      Workbook(sheet).evaluateFormula("=TODAY()+1", name, boom)
    )
  }

  test("recalculate contains a defect per cell and still caches the rest (sequential, parallel)") {
    List(boom, deep).foreach { clock =>
      List(Workbook(sheet).recalculate(clock), Workbook(sheet).recalculateParallel(clock, 4))
        .foreach { result =>
          assertEquals(result.errors.map(_.ref), Vector(ref"B1"))
          val message = result.errors.headOption.map(_.error.message).getOrElse("")
          assert(
            message.contains("internal evaluator defect") || message.contains(
              "exhausted the stack"
            ),
            message
          )
          assertEquals(result.evaluated(name).get(ref"C1"), Some(CellValue.Number(BigDecimal(2))))
        }
    }
  }

  test("OutOfMemoryError and InterruptedException are not contained") {
    // munit's intercept rethrows fatal errors, so the escape is observed directly
    def escaped(body: => Any): Option[Throwable] =
      try
        body
        None
      catch case t: Throwable => Some(t)
    def today(clock: Clock): Either[EvalError, Any] =
      Evaluator.instance.eval(parse("=TODAY()"), sheet, clock)
    def isOom(t: Option[Throwable]): Boolean = t match
      case Some(_: OutOfMemoryError) => true
      case _ => false
    val oom = throwingClock(() => new OutOfMemoryError("test heap"))
    assert(isOom(escaped(today(oom))))
    assert(isOom(escaped(sheet.evaluateFormula("=TODAY()", oom))))
    escaped(today(throwingClock(() => new InterruptedException("stop")))) match
      case Some(_: InterruptedException) => ()
      case other => fail(s"InterruptedException must propagate, got $other")
  }

  test("a boundary.Break crossing the guard reaches its boundary instead of becoming a defect") {
    // Break is a RuntimeException in Scala 3, so NonFatal alone would contain it
    def escapes(run: Clock => Any): Either[String, Any] =
      boundary[Either[String, Any]] {
        val breaking = new Clock:
          def today(): LocalDate = boundary.break(Left("escaped"))
          def now(): LocalDateTime = boundary.break(Left("escaped"))
        Right(run(breaking))
      }
    factories.foreach { (label, evaluator) =>
      assertEquals(escapes(evaluator.eval(parse("=TODAY()+1"), sheet, _)), Left("escaped"), label)
    }
    assertEquals(escapes(Workbook(sheet).recalculate(_)), Left("escaped"), "recalculate")
  }

  test("a defect is never promoted to an Excel error value") {
    assert(EvalError.toErrorValue(EvalError.EvalFailed(defectAtB1)).isEmpty)
    sheet.evaluateCell(ref"B1", boom) match
      case Right(value) => fail(s"a defect must not become a cell value: $value")
      case Left(_) => ()
  }

  test("#750: whole-axis ranges in any position evaluate, never materialised, every factory") {
    // array mode holds a whole-axis value as its populated box plus one fill: every one of these
    // finishes in milliseconds whatever its declared size (R:KHI is 8 billion positions)
    val ranges = List("A:A", "B:C", "R:KHI", "kHi:r", "A:XFD", "2:2", "1:1048576")
    val shapes = List(
      "=$R",
      "=$R*2",
      "=-$R",
      "=$R>0",
      "=$R=\"\"",
      "=$R&\"\"",
      "=$R+$R",
      "=SUM($R)",
      "=SUM(($R=\"\")*1)",
      "=COUNT($R*1)",
      "=MEDIAN($R*1)",
      "=ROWS($R*1)",
      "=AND($R>0)",
      "=NOT($R)",
      "=IF($R>0,1,0)",
      "=SUMPRODUCT(($R>0)*1)",
      "=TRANSPOSE($R)",
      "=INDEX($R,1,1)",
      "=($R $R)",
      "=ABS($R)",
      "=IFERROR($R/0,0)"
    )
    val started = System.nanoTime()
    for
      (label, evaluator) <- factories
      r <- ranges
      shape <- shapes
      at <- List(None, Some(ref"B2"))
    do
      val formula = shape.replace("$R", r)
      // any value or Left is total; only a throw (or an OOM) would fail here
      evaluator.eval(parse(formula), sheet, Clock.system, Some(Workbook(sheet)), at): Unit
    val seconds = (System.nanoTime() - started) / 1e9
    assert(seconds < 60.0, s"took $seconds s")
  }
