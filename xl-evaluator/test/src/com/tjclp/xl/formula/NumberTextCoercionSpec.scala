package com.tjclp.xl.formula

import munit.ScalaCheckSuite
import org.scalacheck.Prop.forAll

import java.time.LocalDate

import com.tjclp.xl.Generators.genWideBigDecimal
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-665: every number → text conversion in the evaluator (`&`, CONCATENATE, text-typed arguments,
 * numeric literals in text positions, TEXT(x,"General")) renders Excel's General text — 15
 * significant digits, trailing zeros stripped, plain up to 20 characters then E notation — never
 * `BigDecimal.toString`, which carries the stored scale (`<v>2.0</v>` read as "2.0", `SUM` of
 * scaled cells as "1070.0", `1/3` as 34 digits).
 *
 * The rule itself is pinned in xl-core's GeneralTextSpec; this suite pins that each evaluator site
 * routes through it (decodeAsString, ScalarCoercion.coerceText, Evaluator.concatText, the
 * asStringExpr literal arm, TEXT) and that the GH-561 date-as-serial behaviour survives.
 */
class NumberTextCoercionSpec extends ScalaCheckSuite:

  // A1..A3 carry the scale a batch-JSON `2` lands with (`<v>2.0</v>`); B3:B7 sum to 1070 with
  // scale so `SUM` yields 1070.0; T1 is genuine text that must never be renumbered.
  private val sheet = Sheet("Test")
    .put(ref"A1", CellValue.Number(BigDecimal("2.0")))
    .put(ref"A2", CellValue.Number(BigDecimal("3.0")))
    .put(ref"A3", CellValue.Number(BigDecimal("5.00")))
    .put(ref"B3", CellValue.Number(BigDecimal("200.0")))
    .put(ref"B4", CellValue.Number(BigDecimal("300.0")))
    .put(ref"B5", CellValue.Number(BigDecimal("250.0")))
    .put(ref"B6", CellValue.Number(BigDecimal("170.0")))
    .put(ref"B7", CellValue.Number(BigDecimal("150.0")))
    .put(ref"B8", CellValue.Formula("SUM(B3:B7)", None))
    .put(ref"T1", CellValue.Text("2.50"))
    .put(ref"D1", CellValue.DateTime(LocalDate.of(2026, 1, 1).atStartOfDay()))

  private def assertScalar(formula: String, expected: CellValue)(implicit
    loc: munit.Location
  ): Unit =
    assertEquals(sheet.evaluateFormula(formula), Right(expected), formula)

  private def text(s: String): CellValue = CellValue.Text(s)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  // ===== The issue's table =====

  test("GH-665: `&` on a scaled stored number drops the scale (2.0 → \"2\")") {
    assertScalar("=A1&\"\"", text("2"))
  }

  test("GH-665: `&` on scaled arithmetic drops the scale (16.00 → \"16 x\")") {
    assertScalar("=A1*(A2+A3)&\" x\"", text("16 x"))
  }

  test("GH-665: `&` on a SUM of scaled cells (\"Total 1070\", not \"Total 1070.0\")") {
    assertScalar("=\"Total \"&B8", text("Total 1070"))
  }

  test("GH-665: `&` on 1/3 renders 15 significant digits, not 34") {
    assertScalar("=1/3&\"\"", text("0.333333333333333"))
    assertScalar("=2/3&\"\"", text("0.666666666666667"))
  }

  test("GH-665: concatenated scaled numbers re-enter arithmetic (=(A1&A2)+1 is 24)") {
    // The issue's row says 25 with its inputs; on this sheet A1=2, A2=3 → "23" + 1
    assertScalar("=(A1&A2)+1", num(24))
    assertScalar("=\"24\"+1", num(25))
  }

  test("GH-665: literals still concatenate (=2&3 is \"23\")") {
    assertScalar("=2&3", text("23"))
  }

  // ===== Each conversion site once =====

  test("GH-665: numeric LITERAL in a text position renders General text (asStringExpr arm)") {
    assertScalar("=2.50&\"\"", text("2.5"))
    assertScalar("=LEN(2.50)", num(3))
    assertScalar("=1E3&\"\"", text("1000"))
  }

  test("GH-665: text-typed cell reads render General text (decodeAsString)") {
    assertScalar("=CONCATENATE(A1,\"-\",A2)", text("2-3"))
    assertScalar("=LEFT(A1,5)", text("2"))
  }

  test("GH-665: runtime-polymorphic text positions render General text (coerceText)") {
    assertScalar("=UPPER(1/3)", text("0.333333333333333"))
    assertScalar("=LET(x,A1,\"v: \"&x)", text("v: 2"))
    assertScalar("=\"s\"&SUM(A1:A2)", text("s5"))
  }

  test("GH-665: `&` between two scaled cells (concatText)") {
    assertScalar("=A1&A2", text("23"))
  }

  test("GH-665: genuine text is never renumbered") {
    assertScalar("=T1&\"\"", text("2.50"))
    assertScalar("=LEN(T1)", num(4))
  }

  // ===== Thresholds end-to-end =====

  test("GH-665: the 20-character plain/E switch is reachable from formulas") {
    assertScalar("=10^19&\"\"", text("10000000000000000000"))
    assertScalar("=10^20&\"\"", text("1E+20"))
    assertScalar("=1E21&\"\"", text("1E+21"))
    assertScalar("=2^60&\"\"", text("1152921504606850000"))
    assertScalar("=0.00001&\"\"", text("0.00001"))
    assertScalar("=1/1024/1024&\"\"", text("9.5367431640625E-07"))
    assertScalar("=-A1&\"\"", text("-2"))
    assertScalar("=0&\"\"", text("0"))
    assertScalar("=(-1*0)&\"\"", text("0"))
    assertScalar("=PI()&\"\"", text("3.14159265358979"))
  }

  // ===== TEXT(x, "General") =====

  test("GH-665: TEXT(x,\"General\") is the width-independent text rule") {
    assertScalar("=TEXT(1/3,\"General\")", text("0.333333333333333"))
    assertScalar("=TEXT(A1,\"general\")", text("2"))
    assertScalar("=TEXT(DATE(2026,1,1),\"General\")", text("46023"))
    assertScalar("=TEXT(1/3,\"General\")=1/3&\"\"", CellValue.Bool(true))
  }

  test("#672: the General keyword inside a TEXT code follows the text rule, not cell display") {
    // a second section cannot change a positive value, so the codes must agree
    assertScalar("=TEXT(123456789012,\"General\")", text("123456789012"))
    assertScalar("=TEXT(123456789012,\"General;-General\")", text("123456789012"))
    assertScalar("=TEXT(123456789012,\"General\"\" u\"\"\")", text("123456789012 u"))
    assertScalar("=TEXT(0.000012345,\"General;-General\")", text("0.000012345"))
    assertScalar("=TEXT(-5,\"General;-General\")", text("-5"))
    assertScalar("=TEXT(1/3,\"General;-General\")", text("0.333333333333333"))
    // a text-only code renders a number as General too
    assertScalar("=TEXT(123456789012,\"@\")", text("123456789012"))
  }

  test("GH-665: TEXT with an explicit format is unchanged") {
    assertScalar("=TEXT(A1,\"0.00\")", text("2.00"))
    assertScalar("=TEXT(A1,\"0\")", text("2"))
  }

  // ===== Round trips and GH-561 =====

  test("GH-665: text → number round trip through VALUE") {
    assertScalar("=VALUE(A1&\"\")=A1", CellValue.Bool(true))
    assertScalar("=ABS(VALUE(1/3&\"\")-1/3)<1E-14", CellValue.Bool(true))
  }

  test("GH-561 preserved: dates in text positions are still their serial") {
    assertScalar("=\">=\"&DATE(2026,1,1)", text(">=46023"))
    assertScalar("=\">=\"&D1", text(">=46023"))
    assertScalar("=LEN(D1)", num(5))
  }

  // ===== Printer: the literal survives in the AST =====

  test("GH-665: a numeric literal in a text position re-prints byte-identically") {
    List("=2.50&\"\"", "=2&3", "=A1&2.50").foreach { f =>
      FormulaParser.parse(f) match
        case Right(expr) => assertEquals(FormulaPrinter.print(expr), f)
        case Left(err) => fail(s"parse failed for $f: $err")
    }
  }

  // ===== Three-arm parity law =====

  // Arms: decodeAsString (=A1&""), coerceText via LET binding and via the INDEX call result,
  // CONCATENATE's TextList (asStringExpr → decodeAsString), and (#671) a cached formula cell in a
  // text position, which reads its cached value like any other reference.
  property("GH-665: every text-position arm agrees with NumFmtFormatter.generalText") {
    forAll(genWideBigDecimal) { (n: BigDecimal) =>
      val s = Sheet("P")
        .put(ref"A1", CellValue.Number(n))
        .put(ref"B1", CellValue.Formula("A1", Some(CellValue.Number(n))))
      val expected = Right(CellValue.Text(NumFmtFormatter.generalText(n)))
      val direct = s.evaluateFormula("=A1&\"\"")
      val viaBinding = s.evaluateFormula("=LET(x,A1,x&\"\")")
      val viaConcatenate = s.evaluateFormula("=CONCATENATE(\"\",A1)")
      val viaCall = s.evaluateFormula("=\"x\"&INDEX(A1:A1,1)").map {
        case CellValue.Text(t) => CellValue.Text(t.stripPrefix("x"))
        case other => other
      }
      assertEquals(direct, expected, s"=A1&\"\" for $n")
      assertEquals(viaBinding, expected, s"LET arm for $n")
      assertEquals(viaConcatenate, expected, s"CONCATENATE arm for $n")
      assertEquals(viaCall, expected, s"INDEX arm for $n")
      assertEquals(s.evaluateFormula("=B1&\"\""), expected, s"cached formula arm for $n")
    }
  }

  // ===== #671: a cached formula cell in a text position yields its value =====

  // LibreOffice: ="" & C2 with C2 = =A1*2 over A1 = 10 reads "20"; Excel likewise reads the
  // formula's value. These cells are never recalculated first: the cached value is what is read.
  private val cached = Sheet("C")
    .put(ref"B1", CellValue.Formula("A1", Some(CellValue.Number(BigDecimal(0)))))
    .put(ref"B2", CellValue.Formula("\"x\"&\"y\"", Some(CellValue.Text("xy"))))
    .put(ref"B3", CellValue.Formula("1=1", Some(CellValue.Bool(true))))
    .put(
      ref"B4",
      CellValue
        .Formula("DATE(2026,1,1)", Some(CellValue.DateTime(LocalDate.of(2026, 1, 1).atStartOfDay)))
    )
    .put(ref"B5", CellValue.Formula("1/0", Some(CellValue.Error(CellError.Div0))))
    .put(ref"B6", CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal("20.0")))))
    .put(
      ref"B7",
      CellValue.Formula(
        "TABLE(,A1)",
        None,
        FormulaKind.DataTable(CellRange(ref"B7", ref"B7"), false, false, Some(ref"A1"), None)
      )
    )

  private def cachedCase(formula: String, expected: CellValue)(implicit loc: munit.Location): Unit =
    assertEquals(cached.evaluateFormula(formula), Right(expected), formula)

  test("#671: a cached formula cell in a text position reads its cached value, not its text") {
    cachedCase("=\"\"&B1", text("0"))
    cachedCase("=B2&\"!\"", text("xy!"))
    cachedCase("=B3&\"\"", text("TRUE"))
    cachedCase("=B4&\"\"", text("46023"))
    cachedCase("=\"\"&B6", text("20"))
  }

  test("#671: text-typed function arguments read the cached value too") {
    cachedCase("=LEN(B1)", num(1))
    cachedCase("=UPPER(B2)", text("XY"))
    cachedCase("=CONCATENATE(B1,B6)", text("020"))
    cachedCase("=LEFT(B6,1)", text("2"))
  }

  test("#671: a cached error in a text position propagates the error") {
    cachedCase("=B5&\"\"", CellValue.Error(CellError.Div0))
    cachedCase("=LEN(B5)", CellValue.Error(CellError.Div0))
  }

  test("#671: an uncached data-table record reads as the empty text") {
    cachedCase("=B7&\"|\"", text("|"))
  }

  test("#671: a cross-sheet cached formula cell reads its cached value") {
    val main = Sheet("Main")
    val wb = Workbook(Vector(main, cached))
    assertEquals(
      main.evaluateFormula("=C!B1&\"\"", workbook = Some(wb)),
      Right(text("0"))
    )
    assertEquals(main.evaluateFormula("=LEN(C!B2)", workbook = Some(wb)), Right(num(2)))
  }

  // ===== #671: Boolean and date literals in text positions survive in the AST =====

  test("#671: TRUE/FALSE in a text position render at evaluation time and re-print as written") {
    List("=TRUE&\"\"", "=FALSE&1", "=LEN(TRUE)", "=UPPER(FALSE)").foreach { f =>
      FormulaParser.parse(f) match
        case Right(expr) =>
          assertEquals(FormulaPrinter.print(expr), f)
          assertEquals(FormulaParser.parse(FormulaPrinter.print(expr)), Right(expr))
        case Left(err) => fail(s"parse failed for $f: $err")
    }
    assertScalar("=TRUE&\"\"", text("TRUE"))
    assertScalar("=FALSE&1", text("FALSE1"))
    assertScalar("=LEN(TRUE)", num(4))
  }

  test("#671: a programmatic date literal in a text position renders its serial at evaluation") {
    val date = TExpr.asStringExpr(TExpr.Lit(LocalDate.of(2026, 1, 1)))
    val dateTime = TExpr.asStringExpr(TExpr.Lit(LocalDate.of(2026, 1, 1).atTime(12, 0)))
    assert(!date.isInstanceOf[TExpr.Lit[?]], s"folded to $date")
    assertEquals(Evaluator.instance.eval(date, sheet), Right("46023"))
    assertEquals(Evaluator.instance.eval(dateTime, sheet), Right("46023.5"))
  }
