package com.tjclp.xl.formula

import java.time.{LocalDate, LocalDateTime}

import scala.util.{Failure, Success, Try}

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.{BindingCoercion, TExpr}
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.functions.{FunctionRegistry, FunctionSpecs}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * The array-mode typing invariant: any `TExpr[A]` may evaluate to an ArrayResult in array mode, and
 * every consumer either handles the array or turns it into a value of its position's type (or a
 * Left) — never a mistyped value that throws inside a function body. The fixture is the Weaver
 * repro (B2:D4 = 5,-3,1 / -2,4,-6 / 7,1,2).
 *
 * A lifted function (ArrayLiftingSpec) answers Excel's element-wise array instead; the typed
 * top-left collapse remains the rule for an array value in a scalar slot of a function that does
 * not lift. A plain cell intersects its references with its own row or column first
 * (PlainCellIntersectionSpec).
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
class ArrayTypingLawsSpec extends FunSuite:

  private val name = SheetName.unsafe("Sheet1")
  private def num(n: BigDecimal): CellValue = CellValue.Number(n)
  private def text(s: String): CellValue = CellValue.Text(s)

  private val sheet: Sheet = List[(ARef, CellValue)](
    ref"A1" -> text("Label"),
    ref"B1" -> text("x"),
    ref"C1" -> text("y"),
    ref"D1" -> text("z"),
    ref"B2" -> num(5),
    ref"C2" -> num(-3),
    ref"D2" -> num(1),
    ref"B3" -> num(-2),
    ref"C3" -> num(4),
    ref"D3" -> num(-6),
    ref"B4" -> num(7),
    ref"C4" -> num(1),
    ref"D4" -> num(2),
    ref"A5" -> text("a"),
    ref"A6" -> CellValue.RichText(RichText.plain("rich")),
    ref"E2" -> num(1),
    ref"E3" -> CellValue.Error(CellError.NA)
  ).foldLeft(Sheet(name)) { case (s, (at, v)) => s.put(at, v) }

  private val workbook = Workbook(sheet)
  private val here = Some(ref"H1")

  private def parse(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(e => fail(s"$formula does not parse: $e"), identity)

  private def arrayEval(formula: String): Either[EvalError, Any] =
    Evaluator.arrayInstance.eval(parse(formula), sheet, Clock.system, Some(workbook), here)

  private def scalarEval(formula: String): Either[EvalError, Any] =
    Evaluator.instance.eval(parse(formula), sheet, Clock.system, Some(workbook), here)

  /** The plain-cell evaluation at `at`, inside the data rows. */
  private def scalarEvalAt(formula: String, at: ARef): Either[EvalError, Any] =
    Evaluator.instance.eval(parse(formula), sheet, Clock.system, Some(workbook), Some(at))

  private def isValueError(result: Either[EvalError, Any]): Boolean = result match
    case Left(EvalError.ErrorValue(CellError.Value, _)) => true
    case _ => false

  private def isDefect(message: String): Boolean =
    message.contains("internal evaluator defect") || message.contains("exhausted the stack")

  private def isDefect(e: EvalError): Boolean = e match
    case EvalError.EvalFailed(reason, _) => isDefect(reason)
    case _ => false

  private def column(values: CellValue*): ArrayResult = ArrayResult(values.toVector.map(Vector(_)))

  /** A result is sound when nothing escaped and no defect was contained on the way. */
  private def assertSound(label: String, result: => Either[EvalError, Any]): Unit =
    Try(result) match
      case Failure(t) => fail(s"$label threw ${t.getClass.getName}: ${t.getMessage}")
      case Success(Left(e)) if isDefect(e) => fail(s"$label: contained defect $e")
      case Success(_) => ()

  // ===== The ClassCastException family (Weaver 9/17 view --eval crash) =====

  private val typedArrayArguments = List(
    "=ABS(C2:C4+D2:D4)",
    "=SUM(ABS(C2:C4+D2:D4))",
    "=SUMPRODUCT(ROUND(C2:C4/3,1))",
    "=SUMPRODUCT(--(B2:B4>0),ABS(C2:C4+D2:D4))",
    "=IF(B2:B4>0,ABS(C2:C4+0),0)",
    "=MOD(C2:C4+0,2)",
    "=RANK(B2:B4+0,B2:B4)",
    "=SQRT(ABS(B2:B4*2))",
    "=LEFT(\"abcdef\",C2:C4+D2:D4+5)",
    "=DATE(2020,B2:B4+0,1)",
    "=DATE(2020,B2:B4+0,1)+1",
    "=YEAR(DATE(2020,1,1)+B2:B4)",
    "=UPPER(B2:B4&\"\")",
    "=LEN(TRANSPOSE(A1:D1)&\"\")",
    "=NOT(B2:B4>0)",
    "=EDATE(DATE(2020,1,31),B2:B4*1)"
  )

  test("an array-valued expression in a typed argument never reaches a function body mistyped") {
    typedArrayArguments.foreach { f =>
      assertSound(s"$f (array mode)", arrayEval(f))
      assertSound(s"$f (scalar mode)", scalarEval(f))
      assertSound(
        s"$f (evaluateFormula)",
        sheet.evaluateFormula(f).left.map(e => EvalError.EvalFailed(e.message))
      )
    }
  }

  test("a numeric slot lifts in array mode; a plain cell intersects the reference first") {
    assertEquals(arrayEval("=ABS(C2:C4+D2:D4)"), Right(column(num(2), num(2), num(3))))
    assertEquals(arrayEval("=SUM(ABS(C2:C4+D2:D4))"), Right(BigDecimal(7)))
    assertEquals(arrayEval("=SUMPRODUCT(ROUND(C2:C4/3,1))"), Right(BigDecimal("0.6")))
    assertEquals(scalarEvalAt("=MOD(C2:C4+0,2)", ref"H3"), Right(BigDecimal(0)))
    assert(isValueError(scalarEval("=MOD(C2:C4+0,2)")), "row 1 does not cross C2:C4")
  }

  test("Weaver's F2 is Excel's 5 (ABS lifts, SUMPRODUCT sees two 3x1 arrays)") {
    val withF2 = sheet.put(ref"F2", CellValue.Formula("SUMPRODUCT(--(B2:B4>0),ABS(C2:C4+D2:D4))"))
    assertEquals(withF2.evaluateCell(ref"F2"), Right(num(5)))
  }

  test(
    "an integer slot lifts in array mode; a plain cell intersects first (ToInt), then truncates"
  ) {
    assertEquals(
      arrayEval("=LEFT(\"abcdef\",C2:C4+D2:D4+5)"),
      Right(column(text("abc"), text("abc"), text("abcdef")))
    )
    assertEquals(
      arrayEval("=DATE(2020,B2:B4+0,1)"),
      Right(
        column(
          CellValue.DateTime(LocalDate.of(2020, 5, 1).atStartOfDay()),
          CellValue.DateTime(LocalDate.of(2019, 10, 1).atStartOfDay()),
          CellValue.DateTime(LocalDate.of(2020, 7, 1).atStartOfDay())
        )
      )
    )
    assertEquals(scalarEvalAt("=DATE(2020,B2:B4+0,1)", ref"H3"), Right(LocalDate.of(2019, 10, 1)))
  }

  test("the typed collapse coerces by the node's static kind, so a text element is a clean Left") {
    // A1:A2 top-left is "Label": a numeric slot that does not lift refuses it as a type
    // mismatch, and so does a plain cell's ABS; a lifting ABS answers #VALUE! per element
    arrayEval("=SEQUENCE(A1:A2&\"\")") match
      case Left(EvalError.TypeMismatch(_, _, _)) => ()
      case other => fail(s"expected a TypeMismatch, got $other")
    scalarEval("=ABS(A1:A2&\"\")") match
      case Left(EvalError.TypeMismatch(_, _, _)) => ()
      case other => fail(s"expected a TypeMismatch, got $other")
    assertEquals(
      arrayEval("=ABS(A1:A2&\"\")"),
      Right(column(CellValue.Error(CellError.Value), CellValue.Error(CellError.Value)))
    )
  }

  // ===== Concat broadcasts element-wise =====

  test("text concatenation broadcasts over an array operand in array mode") {
    assertEquals(
      arrayEval("=(C2:C4+D2:D4)&\"x\""),
      Right(column(text("-2x"), text("-2x"), text("3x")))
    )
    assertEquals(arrayEval("=\"<\"&B2:B4"), Right(column(text("<5"), text("<-2"), text("<7"))))
    assertEquals(
      arrayEval("=TRANSPOSE(B2:B4)&\"x\""),
      Right(ArrayResult(Vector(Vector(text("5x"), text("-2x"), text("7x")))))
    )
  }

  test("text concatenation never renders an ArrayResult as text") {
    def texts(value: Any): Vector[String] = value match
      case s: String => Vector(s)
      case CellValue.Text(s) => Vector(s)
      case ar: ArrayResult => ar.values.flatten.flatMap(texts)
      case _ => Vector.empty
    List("=(C2:C4+D2:D4)&\"x\"", "=TRANSPOSE(B2:B4)&\"x\"", "=\"x\"&TRANSPOSE(B2:B4)").foreach {
      f =>
        List(arrayEval(f), scalarEvalAt(f, ref"H3")).foreach { r =>
          val rendered = r.fold(_ => Vector.empty, texts)
          assert(rendered.nonEmpty, s"$f produced no text: $r")
          assert(!rendered.exists(_.contains("ArrayResult")), s"$f rendered an array as text: $r")
        }
    }
  }

  test("text concatenation: a carried error element wins, mismatched shapes pad #N/A") {
    assertEquals(
      arrayEval("=E2:E3&\"x\""),
      Right(column(text("1x"), CellValue.Error(CellError.NA)))
    )
    assertEquals(
      arrayEval("=\"x\"&E2:E3"),
      Right(column(text("x1"), CellValue.Error(CellError.NA)))
    )
    assertEquals(
      arrayEval("=B2:B4&C2:C3"),
      Right(column(text("5-3"), text("-24"), CellValue.Error(CellError.NA)))
    )
  }

  test("text concatenation renders rich-text elements as their plain text") {
    assertEquals(arrayEval("=A5:A6&\"!\""), Right(column(text("a!"), text("rich!"))))
  }

  test("a plain cell's concatenation intersects a reference, and reads an array value's top-left") {
    assertEquals(scalarEvalAt("=(C2:C4+D2:D4)&\"x\"", ref"H3"), Right("-2x"))
    assert(isValueError(scalarEval("=(C2:C4+D2:D4)&\"x\"")), "row 1 does not cross C2:C4")
    assertEquals(scalarEval("=TRANSPOSE(B2:B4)&\"x\""), Right("5x"))
    assertEquals(sheet.evaluateFormula("=B2:B4&\"x\""), Right(text("5x")))
    // two scalars are unchanged, and an error operand still propagates as that error
    assertEquals(scalarEval("=B2&\"x\""), Right("5x"))
    assertEquals(sheet.evaluateFormula("=E3&\"x\""), Right(CellValue.Error(CellError.NA)))
  }

  // ===== Date serial nodes are total =====

  private val datesColumn = column(
    CellValue.DateTime(LocalDate.of(2020, 1, 1).atStartOfDay()),
    num(5),
    CellValue.Error(CellError.NA)
  )

  test("DateToSerial maps an array element-wise in array mode and collapses in scalar mode") {
    val node = TExpr.DateToSerial(TExpr.Lit(datesColumn).asInstanceOf[TExpr[LocalDate]])
    assertEquals(
      Evaluator.arrayInstance.eval(node, sheet): Either[EvalError, Any],
      Right(column(num(43831), num(5), CellValue.Error(CellError.NA)))
    )
    assertEquals(Evaluator.instance.eval(node, sheet), Right(BigDecimal(43831)))
  }

  test("DateTimeToSerial maps an array element-wise and converts a scalar date-time") {
    val node = TExpr.DateTimeToSerial(TExpr.Lit(datesColumn).asInstanceOf[TExpr[LocalDateTime]])
    assertEquals(
      Evaluator.arrayInstance.eval(node, sheet): Either[EvalError, Any],
      Right(column(num(43831), num(5), CellValue.Error(CellError.NA)))
    )
    val noon = TExpr.DateTimeToSerial(TExpr.Lit(LocalDateTime.of(2020, 1, 1, 12, 0)))
    assertEquals(Evaluator.instance.eval(noon, sheet), Right(BigDecimal("43831.5")))
  }

  test("DateToSerial coerces a non-date operand totally instead of throwing") {
    val serial = TExpr.DateToSerial(TExpr.Lit(BigDecimal(5)).asInstanceOf[TExpr[LocalDate]])
    assertEquals(Evaluator.instance.eval(serial, sheet), Right(BigDecimal(5)))
    val textual = TExpr.DateToSerial(TExpr.Lit("abc").asInstanceOf[TExpr[LocalDate]])
    Evaluator.instance.eval(textual, sheet) match
      case Left(EvalError.TypeMismatch(_, _, _)) => ()
      case other => fail(s"expected a TypeMismatch, got $other")
  }

  // ===== Programmatic ASTs: literal arms and TExpr.cond =====

  test("a literal of the wrong runtime type in a typed slot coerces instead of being cast") {
    val date = LocalDate.of(2020, 1, 1)
    assertEquals(
      Evaluator.instance.eval(TExpr.asNumericExpr(TExpr.Lit(date)), sheet),
      Right(BigDecimal(43831))
    )
    assertEquals(Evaluator.instance.eval(TExpr.asStringExpr(TExpr.Lit(num(2))), sheet), Right("2"))
    assertEquals(
      Evaluator.instance.eval(TExpr.asBooleanExpr(TExpr.Lit(CellValue.Bool(true))), sheet),
      Right(true)
    )
    // parser shapes are unchanged: a number literal stays a bare literal in a numeric slot
    assertEquals(TExpr.asNumericExpr(TExpr.Lit(BigDecimal(1))), TExpr.Lit(BigDecimal(1)))
    assertEquals(TExpr.asBooleanExpr(TExpr.Lit(true)), TExpr.Lit(true))
    assertEquals(TExpr.asStringExpr(TExpr.Lit("a")), TExpr.Lit("a"))
  }

  test("TExpr.cond with a mistyped branch is a Left, never a mistyped value") {
    val mistyped = TExpr.Lit("abc").asInstanceOf[TExpr[BigDecimal]]
    val expr = TExpr.cond(TExpr.Lit(false), TExpr.Lit(BigDecimal(1)), mistyped)
    Evaluator.instance.eval(expr, sheet) match
      case Left(EvalError.TypeMismatch(_, _, _)) => ()
      case other => fail(s"expected a TypeMismatch, got $other")
    assertEquals(
      Evaluator.instance
        .eval(TExpr.cond(TExpr.Lit(true), TExpr.Lit(BigDecimal(1)), mistyped), sheet),
      Right(BigDecimal(1))
    )
  }

  // ===== The static scalar kind table =====

  test("scalarKind names the static scalar kind of every typed node, None for raw values") {
    val n = TExpr.Lit(BigDecimal(1))
    val today = parse("=TODAY()")
    val abs = parse("=ABS(1)")
    val iff = parse("=IF(TRUE,1,2)")
    assertEquals(TExpr.scalarKind(TExpr.Add(n, n)), Some(BindingCoercion.Numeric))
    assertEquals(TExpr.scalarKind(TExpr.Percent(n)), Some(BindingCoercion.Numeric))
    assertEquals(TExpr.scalarKind(TExpr.ToInt(n)), Some(BindingCoercion.Integer))
    assertEquals(
      TExpr.scalarKind(TExpr.Concat(TExpr.Lit("a"), TExpr.Lit("b"))),
      Some(BindingCoercion.Text)
    )
    assertEquals(TExpr.scalarKind(TExpr.Lt(n, n)), Some(BindingCoercion.Bool))
    assertEquals(TExpr.scalarKind(TExpr.UnaryPlus(TExpr.Mul(n, n))), Some(BindingCoercion.Numeric))
    assertEquals(
      TExpr.scalarKind(TExpr.Coerced[Any](n.asInstanceOf[TExpr[Any]], BindingCoercion.Date)),
      Some(BindingCoercion.Date)
    )
    assertEquals(TExpr.scalarKind(abs), Some(BindingCoercion.Numeric))
    assertEquals(TExpr.scalarKind(today), Some(BindingCoercion.Date))
    assertEquals(TExpr.scalarKind(iff), None)
    assertEquals(TExpr.scalarKind(n), None)
    assertEquals(TExpr.scalarKind(TExpr.Missing), None)
  }

  // ===== Registry-driven law: no registered function throws on array-valued arguments =====

  /** Array-producing argument text for each slot kind the parser describes. */
  private def arrayShapes(kind: String): List[String] = kind match
    case "number" | "integer" | "cell" | "value" | "number or range" =>
      List("B2:B4+0", "(B2:B4)*1", "TRANSPOSE(B2:B4)", "IF(B2:B4>0,B2:B4,0)", "-C2:C4")
    case "text" => List("A1:D1&\"\"", "TRANSPOSE(A1:D1)")
    case "boolean" => List("B2:B4>0", "NOT(B2:B4>0)")
    case "date" => List("B2:B4+45000", "DATE(2020,1,1)+B2:B4")
    case _ => Nil

  /** A plain scalar (or range) argument for each slot kind. */
  private def scalarArgument(kind: String): String = kind match
    case "text" => "\"a\""
    case "boolean" => "TRUE"
    case "date" => "DATE(2026,1,1)"
    case "range" | "array or range" => "B2:D4"
    case _ => "1"

  private def slotKinds(parts: List[String], arity: Int): List[String] =
    val cleaned = parts.map(_.stripPrefix("optional ").stripSuffix("..."))
    val last = cleaned.lastOption.getOrElse("value")
    (0 until arity).toList.map(i => cleaned.lift(i).getOrElse(last))

  private def arities(arity: Arity): List[Int] = arity match
    case Arity.Exact(n) => List(n)
    case Arity.Range(min, max) => (min to math.min(max, min + 2)).toList
    case Arity.AtLeast(min) => (min to min + 2).toList

  test("registry law: every function answers array-valued typed arguments with a value or a Left") {
    val probes =
      for
        spec <- FunctionRegistry.all
        arity <- arities(spec.arity)
        kinds = slotKinds(spec.argSpec.describeParts, arity)
        (kind, position) <- kinds.zipWithIndex
        shape <- arrayShapes(kind)
      yield
        val args = kinds.zipWithIndex.map { (k, i) =>
          if i == position then shape else scalarArgument(k)
        }
        s"=${spec.name}(${args.mkString(",")})"
    assert(probes.sizeIs > 1000, s"only ${probes.size} probes generated")
    val parsed = probes.flatMap(f => FormulaParser.parse(f).toOption.map(f -> _))
    assert(parsed.sizeIs > 1000, s"only ${parsed.size} of ${probes.size} probes parse")
    parsed.foreach { (f, expr) =>
      assertSound(
        s"$f (array mode)",
        Evaluator.arrayInstance.eval(expr, sheet, Clock.system, Some(workbook), here)
      )
      assertSound(
        s"$f (scalar mode)",
        Evaluator.instance.eval(expr, sheet, Clock.system, Some(workbook), here)
      )
    }
  }
