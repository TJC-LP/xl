package com.tjclp.xl.formula

import java.time.LocalDate

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * Excel array lifting: a scalar parameter handed an array evaluates once per element
 * (`ABS({-2;-2;3})` is `{2;2;3}`), implemented once at the Call node for the flagged functions. The
 * fixture is the Weaver repro r1b.xlsx (A1:D1 = Label,x,y,z; B2:D4 = 5,-3,1 / -2,4,-6 / 7,1,2),
 * plus a date column (E2:E4), a text column with duplicates (G1:G5) and a #N/A cell (H3).
 */
class ArrayLiftingSpec extends FunSuite:

  private val name = SheetName.unsafe("Sheet1")
  private def num(n: BigDecimal): CellValue = CellValue.Number(n)
  private def dec(s: String): CellValue = CellValue.Number(BigDecimal(s))
  private def text(s: String): CellValue = CellValue.Text(s)
  private def bool(b: Boolean): CellValue = CellValue.Bool(b)
  private def date(y: Int, m: Int, d: Int): CellValue =
    CellValue.DateTime(LocalDate.of(y, m, d).atStartOfDay())
  private def err(e: CellError): CellValue = CellValue.Error(e)

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
    ref"E2" -> date(2020, 1, 15),
    ref"E3" -> date(2020, 2, 10),
    ref"E4" -> date(2020, 3, 31),
    ref"G1" -> text("a"),
    ref"G2" -> text("b"),
    ref"G3" -> text("a"),
    ref"G4" -> text("c"),
    ref"G5" -> text("b"),
    ref"H2" -> num(1),
    ref"H3" -> err(CellError.NA)
  ).foldLeft(Sheet(name)) { case (s, (at, v)) => s.put(at, v) }

  private val workbook = Workbook(sheet)
  private val weaver = "SUMPRODUCT(--(B2:B4>0),ABS(C2:C4+D2:D4))"

  private def parse(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(e => fail(s"$formula does not parse: $e"), identity)

  /** Array mode (evala / SUMPRODUCT arguments / spill anchors), anchored at Z1. */
  private def arrayEval(formula: String): Either[EvalError, Any] =
    Evaluator.arrayInstance.eval(parse(formula), sheet, Clock.system, Some(workbook), Some(ref"Z1"))

  /** A plain (Normal-kind) formula cell at `at`. */
  private def cellAt(formula: String, at: ARef): XLResult[CellValue] =
    sheet.evaluateFormula(formula, Clock.system, Some(workbook), Some(at))

  private def column(values: CellValue*): ArrayResult = ArrayResult(values.toVector.map(Vector(_)))
  private def row(values: CellValue*): ArrayResult = ArrayResult(Vector(values.toVector))

  private def number(result: Either[EvalError, Any]): BigDecimal = result match
    case Right(n: BigDecimal) => n
    case Right(CellValue.Number(n)) => n
    case other => fail(s"expected a number, got $other")

  // ===== The Weaver finding =====

  test("Weaver F2 is Excel's 5 through every evaluation route") {
    val withF2 = sheet.put(ref"F2", CellValue.Formula(weaver))
    assertEquals(cellAt(s"=$weaver", ref"F2"), Right(num(5)))
    assertEquals(withF2.evaluateCell(ref"F2"), Right(num(5)))
    assertEquals(withF2.evaluateWithDependencyCheck().map(_.get(ref"F2")), Right(Some(num(5))))
    assertEquals(
      withF2.evaluateForRange(CellRange(ref"F2", ref"F2")).map(_.get(ref"F2")),
      Right(Some(num(5)))
    )
    assertEquals(number(arrayEval(s"=$weaver")), BigDecimal(5))
    val recalc = Workbook(withF2).recalculate()
    assert(recalc.isClean, recalc.errors.toString)
    assertEquals(
      recalc.workbook.sheets.headOption.map(_(ref"F2").value),
      Some(CellValue.Formula(weaver, Some(num(5))))
    )
  }

  test("the brief's repro values") {
    assertEquals(number(arrayEval("=SUMPRODUCT(ABS(C2:C4+D2:D4))")), BigDecimal(7))
    assertEquals(number(arrayEval("=SUM(ABS(C2:C4+D2:D4))")), BigDecimal(7))
    assertEquals(cellAt("=SUM(ABS(C2:C4+D2:D4))", ref"F9"), Right(num(7)))
    assertEquals(number(arrayEval("=SUMPRODUCT(ROUND(C2:C4/3,1))")), BigDecimal("0.6"))
    assertEquals(number(arrayEval("=SUMPRODUCT(LEN(A1:D1))")), BigDecimal(8))
    assertEquals(number(arrayEval("=SUMPRODUCT(--ISNUMBER(B2:B4))")), BigDecimal(3))
    assertEqualsDouble(
      number(arrayEval("=SUMPRODUCT(SQRT(ABS(C2:C4)))")).toDouble,
      math.sqrt(3) + 2 + 1,
      1e-12
    )
    assertEquals(arrayEval("=ABS(C2:C4)"), Right(column(num(3), num(4), num(1))))
    assertEquals(
      arrayEval("=IFERROR(1/(B2:B4-5),0)"),
      Right(column(num(0), num(BigDecimal(1) / BigDecimal(-7)), num(BigDecimal("0.5"))))
    )
  }

  // ===== Math =====

  test("math: element-wise over ranges and computed arrays") {
    assertEquals(arrayEval("=ABS(C2:C4+D2:D4)"), Right(column(num(2), num(2), num(3))))
    assertEquals(arrayEval("=SQRT(C2:C4)"), Right(column(err(CellError.Num), num(2), num(1))))
    assertEquals(arrayEval("=SIGN(B2:B4)"), Right(column(num(1), num(-1), num(1))))
    assertEquals(arrayEval("=INT(C2:C4/3)"), Right(column(num(-1), num(1), num(0))))
    assertEquals(arrayEval("=ROUND(C2:C4/3,1)"), Right(column(num(-1), dec("1.3"), dec("0.3"))))
    assertEquals(arrayEval("=MOD(B2:B4,3)"), Right(column(num(2), num(1), num(1))))
  }

  test("math: the digits slot lifts too, and so do both slots together") {
    assertEquals(arrayEval("=ROUND(1.23456,SEQUENCE(2))"), Right(column(dec("1.2"), dec("1.23"))))
    assertEquals(arrayEval("=POWER(B2:B4,H2)"), Right(column(num(5), num(-2), num(7))))
  }

  // ===== Broadcasting =====

  test("broadcasting: a column against a row fills the grid; a shorter operand pads #N/A") {
    // MOD(B_i, {C2, D2}) = MOD(B_i, -3), MOD(B_i, 1)
    assertEquals(
      arrayEval("=MOD(B2:B4,C2:D2)"),
      Right(
        ArrayResult(
          Vector(
            Vector(num(-1), num(0)),
            Vector(num(-2), num(0)),
            Vector(num(-2), num(0))
          )
        )
      )
    )
    assertEquals(
      arrayEval("=MOD(B2:B4,C2:C3)"),
      Right(column(num(-1), num(2), err(CellError.NA)))
    )
  }

  test("a 1x1 array argument behaves as its scalar") {
    assertEquals(arrayEval("=ABS(INDEX(C2:C4,1))"), Right(BigDecimal(3)))
    assertEquals(arrayEval("=ABS(C2:C2)"), Right(BigDecimal(3)))
  }

  test("nested arrays collapse to their top-left element") {
    assertEquals(
      arrayEval("=INDEX(B2:D4,SEQUENCE(3),1)"),
      Right(column(num(5), num(-2), num(7)))
    )
    // each row selection is a 1x3 array; the element keeps its first cell
    assertEquals(arrayEval("=INDEX(B2:D4,SEQUENCE(2),0)"), Right(column(num(5), num(-2))))
  }

  // ===== Text =====

  test("text: element-wise over a row") {
    assertEquals(arrayEval("=LEN(A1:D1)"), Right(row(num(5), num(1), num(1), num(1))))
    assertEquals(
      arrayEval("=UPPER(A1:D1)"),
      Right(row(text("LABEL"), text("X"), text("Y"), text("Z")))
    )
    assertEquals(
      arrayEval("=LEFT(A1:D1,1)"),
      Right(row(text("L"), text("x"), text("y"), text("z")))
    )
    assertEquals(
      arrayEval("=CONCATENATE(A1:D1,\"!\")"),
      Right(row(text("Label!"), text("x!"), text("y!"), text("z!")))
    )
    assertEquals(
      arrayEval("=TEXT(B2:B4,\"0.0\")"),
      Right(column(text("5.0"), text("-2.0"), text("7.0")))
    )
    assertEquals(
      arrayEval("=SEARCH(\"a\",A1:D1)"),
      Right(row(num(2), err(CellError.Value), err(CellError.Value), err(CellError.Value)))
    )
    assertEquals(
      number(arrayEval("=SUMPRODUCT(LEN(IF(B2:B4>0,\"pp\",\"n\")))")),
      BigDecimal(5)
    )
  }

  // ===== Information =====

  test("information functions lift over their value") {
    assertEquals(arrayEval("=ISNUMBER(B2:B4)"), Right(column(bool(true), bool(true), bool(true))))
    assertEquals(
      arrayEval("=ISNUMBER(A1:D1)"),
      Right(row(bool(false), bool(false), bool(false), bool(false)))
    )
    assertEquals(
      arrayEval("=ISBLANK(A1:A3)"),
      Right(column(bool(false), bool(true), bool(true)))
    )
    assertEquals(
      arrayEval("=ISERROR(1/(B2:B4-5))"),
      Right(column(bool(true), bool(false), bool(false)))
    )
    assertEquals(arrayEval("=IFNA(H2:H3,0)"), Right(column(num(1), num(0))))
    assertEquals(arrayEval("=ERROR.TYPE(H2:H3)"), Right(column(err(CellError.NA), num(7))))
    assertEquals(arrayEval("=N(B2:B4)"), Right(column(num(5), num(-2), num(7))))
    assertEquals(number(arrayEval("=SUM(--ISNUMBER(B2:B4))")), BigDecimal(3))
    assertEquals(
      arrayEval("=FILTER(G1:G3,ISNUMBER(SEARCH(\"a\",G1:G3)))"),
      Right(column(text("a"), text("a")))
    )
  }

  test("IFERROR and IFNA lift their value slot only") {
    assertEquals(arrayEval("=IFERROR(5,A1:A3)"), Right(num(5)))
    assertEquals(cellAt("=IFERROR(5,A1:A3)", ref"F2"), Right(num(5)))
    assertEquals(arrayEval("=IFERROR(C2:C4/0,9)"), Right(column(num(9), num(9), num(9))))
  }

  // ===== Dates =====

  test("date functions lift; DATE+1 converts every element to its serial") {
    assertEquals(
      arrayEval("=DATE(2020,B2:B4,1)"),
      Right(column(date(2020, 5, 1), date(2019, 10, 1), date(2020, 7, 1)))
    )
    assertEquals(arrayEval("=YEAR(E2:E4)"), Right(column(num(2020), num(2020), num(2020))))
    assertEquals(arrayEval("=MONTH(E2:E4)"), Right(column(num(1), num(2), num(3))))
    assertEquals(
      number(arrayEval("=SUMPRODUCT(DATE(2026,B2:B4+5,1)-DATE(2026,1,1))")),
      BigDecimal(273 + 59 + 334)
    )
  }

  test("Analysis ToolPak lineage: a multi-cell range reference is #VALUE!, an array lifts") {
    List("=EOMONTH(E2:E4,0)", "=EDATE(E2:E4,1)", "=WORKDAY(E2:E4,1)", "=YEARFRAC(E2:E4,E4)")
      .foreach { f =>
        arrayEval(f) match
          case Left(EvalError.ErrorValue(CellError.Value, _)) => ()
          case other => fail(s"$f: expected #VALUE!, got $other")
        assertEquals(cellAt(f, ref"F3"), Right(err(CellError.Value)), f)
      }
    assertEquals(
      arrayEval("=EOMONTH(+E2:E4,0)"),
      Right(column(date(2020, 1, 31), date(2020, 2, 29), date(2020, 3, 31)))
    )
    assertEquals(arrayEval("=EOMONTH(E2:E2,0)"), Right(LocalDate.of(2020, 1, 31)))
    assertEquals(
      arrayEval("=EDATE(E2,+B2:B3)"),
      Right(column(date(2020, 6, 15), date(2019, 11, 15)))
    )
  }

  // ===== Financial =====

  test("time-value functions lift over an array of periods") {
    val lifted = number(arrayEval("=SUMPRODUCT(PV(0.1,B2:B4+10,-100))"))
    val elementwise = List(15, 8, 17).map(n => number(arrayEval(s"=PV(0.1,$n,-100)"))).sum
    assertEquals(lifted, elementwise)
  }

  // ===== Lookup and criteria (phase B) =====

  test("lookups lift their lookup value") {
    assertEquals(
      arrayEval("=VLOOKUP(B2:B4,B2:D4,2,FALSE)"),
      Right(column(num(-3), num(4), num(1)))
    )
    assertEquals(
      arrayEval("=VLOOKUP(H2:H3+4,B2:D4,2,FALSE)"),
      Right(column(num(-3), err(CellError.NA)))
    )
    assertEquals(arrayEval("=MATCH(B2:B4,B2:B4,0)"), Right(column(num(1), num(2), num(3))))
    // XLOOKUP lifts lookup_value only: a found value never replicates its if_not_found array
    assertEquals(arrayEval("=XLOOKUP(5,B2:B4,C2:C4,D2:D4)"), Right(num(-3)))
  }

  test("criteria functions lift their criteria: the distinct-count idiom") {
    assertEquals(number(arrayEval("=SUMPRODUCT(1/COUNTIF(G1:G5,G1:G5))")), BigDecimal(3))
    assertEquals(number(arrayEval("=SUM(SUMIF(B2:B4,B2:B3,C2:C4))")), BigDecimal(1))
    assertEquals(number(arrayEval("=SUM(COUNTIFS(G1:G5,G1:G2))")), BigDecimal(4))
  }

  test("statistical k-slots lift") {
    assertEquals(arrayEval("=LARGE(B2:D4,SEQUENCE(3))"), Right(column(num(7), num(5), num(4))))
    assertEquals(number(arrayEval("=SUM(LARGE(B2:D4,SEQUENCE(3)))")), BigDecimal(16))
    assertEquals(arrayEval("=RANK(B2:B4,B2:B4)"), Right(column(num(2), num(3), num(1))))
  }

  // ===== Element error table =====

  test("per-element failures: error values carry, type mismatches are #VALUE!") {
    assertEquals(arrayEval("=ABS(A1:A2)"), Right(column(err(CellError.Value), num(0))))
    assertEquals(arrayEval("=ABS(H2:H3)"), Right(column(num(1), err(CellError.NA))))
    assertEquals(
      arrayEval("=MOD(B2:B4,0)"),
      Right(column(err(CellError.Div0), err(CellError.Div0), err(CellError.Div0)))
    )
    arrayEval("=SUMPRODUCT(ABS(A1:A2))") match
      case Left(EvalError.ErrorValue(CellError.Value, _)) => ()
      case other => fail(s"expected #VALUE!, got $other")
  }

  test("a host failure in an element fails the whole call loudly, never as #VALUE!") {
    arrayEval("=VLOOKUP(B2:B4,Missing!A1:B3,2,FALSE)") match
      case Left(EvalError.EvalFailed(msg, _)) => assert(msg.contains("Missing"), msg)
      case other => fail(s"expected a loud failure, got $other")
  }

  test("a text or math domain error is an element's #VALUE!/#NUM!, so IFERROR keeps the rest") {
    // FIND's miss, LEFT's negative length and CEILING's sign rule are Excel error values, not host
    // failures: each element keeps its own result, and IFERROR replaces only the failing ones
    arrayEval("=LEFT(A1:D1,-1)") match
      case Right(ar: ArrayResult) =>
        assert(ar.values.flatten.forall(_ == CellValue.Error(CellError.Value)), ar.toString)
      case other => fail(s"expected an array of #VALUE!, got $other")
    val words = Sheet(SheetName.unsafe("Words"))
      .put(ref"A1", text("banana"))
      .put(ref"A2", text("cherry"))
      .put(ref"A3", text("apple"))
    // FIND("a") over banana/cherry/apple = {2;#VALUE!;1}: 2 + 0 + 1
    assertEquals(
      words.evaluateFormula("=SUMPRODUCT(IFERROR(FIND(\"a\",A1:A3),0))"),
      Right(CellValue.Number(3))
    )
    assertEquals(
      words.evaluateFormula("=SUMPRODUCT(--ISNUMBER(FIND(\"a\",A1:A3)))"),
      Right(CellValue.Number(2))
    )
  }

  test("per-element domain errors: FIND, LEFT and CEILING agree with Excel and LibreOffice") {
    val sheet = List("apple", "Banana", "grape", "kiwi", "pear").zipWithIndex
      .foldLeft(Sheet(SheetName.unsafe("S"))) { case (s, (word, i)) =>
        s.put(ARef.from0(0, i), text(word))
      }
      .put(ref"B1", CellValue.Number(2))
      .put(ref"B2", CellValue.Number(-1))
      .put(ref"B3", CellValue.Number(3))
      .put(ref"B4", CellValue.Number(0))
      .put(ref"B5", CellValue.Number(1))
    List(
      "=SUMPRODUCT(--ISNUMBER(FIND(\"p\",A1:A5)))" -> 3,
      "=SUMPRODUCT(--ISNUMBER(SEARCH(\"p\",A1:A5)))" -> 3,
      "=SUMPRODUCT(IFERROR(FIND(\"p\",A1:A5),0))" -> 7,
      "=SUMPRODUCT(--ISERROR(FIND(\"p\",A1:A5)))" -> 2,
      // LEFT("Banana",-1) is #VALUE!: 2 + 0 + 3 + 0 + 1
      "=SUMPRODUCT(LEN(IFERROR(LEFT(A1:A5,B1:B5),\"\")))" -> 6,
      // Excel 2010+: CEILING(-1.5,1) is -1 and CEILING(-0.5,1) is 0: 2 - 1 + 3 + 0 + 1
      "=SUMPRODUCT(IFERROR(CEILING(B1:B5-0.5,1),0))" -> 5
    ).foreach { (formula, expected) =>
      assertEquals(sheet.evaluateFormula(formula), Right(CellValue.Number(expected)), formula)
    }
  }

  // ===== Single evaluation =====

  private final class CountingRng extends Rng:
    @SuppressWarnings(Array("org.wartremover.warts.Var"))
    var draws: Int = 0
    def nextDouble(): Double =
      draws += 1
      0.5

  test("each lifted slot evaluates once: RAND draws once, not once per element") {
    List("=ROUND(B2:B4*RAND(),0)", "=ABS(ABS(C2:C4*RAND()))", "=ROUND(RAND()+B2:B4,0)").foreach {
      f =>
        val rng = new CountingRng
        val result =
          Evaluator.arrayInstance(rng).eval(parse(f), sheet, Clock.system, Some(workbook), None)
        assert(result.isRight, s"$f: $result")
        assertEquals(rng.draws, 1, f)
    }
  }

  // ===== Contexts =====

  test("lifting inside IF branches and LET bindings") {
    assertEquals(
      arrayEval("=IF(B2:B4>0,ABS(C2:C4),0)"),
      Right(column(num(3), num(0), num(1)))
    )
    assertEquals(cellAt("=LET(x,LEN(A1:D1),SUM(x))", ref"F2"), Right(num(8)))
  }

  test("a defined name bound to a range lifts like the range; a one-cell name is that cell") {
    val named = workbook
      .withDefinedName("vals", "Sheet1!$C$2:$C$4")
      .withDefinedName("dates", "Sheet1!$E$2:$E$4")
      .withDefinedName("one", "Sheet1!$C$3")
      .withDefinedName("gap", "Sheet1!$A$9")
    def arrayNamed(f: String): Either[EvalError, Any] =
      Evaluator.arrayInstance.eval(parse(f), sheet, Clock.system, Some(named), Some(ref"Z1"))
    def cellNamed(f: String, at: ARef) =
      sheet.evaluateFormula(f, Clock.system, Some(named), Some(at))
    assertEquals(arrayNamed("=ABS(vals)"), Right(column(num(3), num(4), num(1))))
    assertEquals(cellNamed("=ABS(vals)", ref"F3"), Right(num(4)))
    assertEquals(cellNamed("=EOMONTH(dates,0)", ref"F3"), Right(err(CellError.Value)))
    assertEquals(arrayNamed("=ABS(one)"), Right(BigDecimal(4)))
    assertEquals(cellNamed("=ABS(one)", ref"F9"), Right(num(4)))
    // a name bound to a blank cell reads as a reference to it: LEN is 0, YEAR 1900 (Excel)
    assertEquals(cellNamed("=LEN(gap)", ref"F9"), cellNamed("=LEN(A9)", ref"F9"))
    assertEquals(cellNamed("=LEN(gap)", ref"F9"), Right(num(0)))
    assertEquals(cellNamed("=YEAR(gap)", ref"F9"), Right(num(1900)))
  }

  // ===== Unchanged behaviour =====

  test("non-lifted functions and scalar calls are unchanged") {
    assertEquals(arrayEval("=ROUND(C2,1)"), Right(BigDecimal(-3)))
    assertEquals(number(arrayEval("=SUMPRODUCT(B2:B4,C2:C4)")), BigDecimal(-15 - 8 + 7))
    assertEquals(arrayEval("=NOT(B2:B4>0)"), Right(column(bool(false), bool(true), bool(false))))
    // a computed array in a plain cell keeps the GH-302 top-left convention
    assertEquals(cellAt("=ABS(C2:C4+D2:D4)", ref"F4"), Right(num(2)))
  }

  // ===== Scalar (Normal-kind) cells: implicit intersection =====

  test("a plain cell intersects a range in a lifted slot with its own row or column") {
    assertEquals(cellAt("=ABS(C2:C4)", ref"F3"), Right(num(4)))
    assertEquals(cellAt("=ABS(C2:C4)", ref"F2"), Right(num(3)))
    assertEquals(cellAt("=LEN(A1:D1)", ref"B9"), Right(num(1)))
    assertEquals(cellAt("=ISNUMBER(B2:B4)", ref"F3"), Right(bool(true)))
    assertEquals(cellAt("=ABS(C2:C4)", ref"F10"), Right(err(CellError.Value)))
    assertEquals(cellAt("=ISERROR(ABS(C2:C4))", ref"F10"), Right(bool(true)))
  }

  test("an intersection without the formula's position is the same loud error as @") {
    Evaluator.eval(parse("=ABS(C2:C4)"), sheet) match
      case Left(EvalError.EvalFailed(msg, _)) => assert(msg.contains("cell position"), msg)
      case other => fail(s"expected a loud failure, got $other")
  }

  test("an uncached precedent read through a range evaluates at its own position") {
    // H2:H4 are plain cells (each intersects C2:C4 in its own row: 3, 4, 1); K9 is a CSE record
    // (array mode, anchor element (0,0): 3). Nothing is recalculated first, so every aggregate
    // and criteria walk must evaluate them on demand exactly as a direct reference does.
    val plain = FormulaKind.Normal()
    val precedents = sheet
      .put(ref"H2", CellValue.Formula("ABS($C$2:$C$4)", None, plain))
      .put(ref"H3", CellValue.Formula("ABS($C$2:$C$4)", None, plain))
      .put(ref"H4", CellValue.Formula("ABS($C$2:$C$4)", None, plain))
      .put(
        ref"K9",
        CellValue.Formula(
          "ABS($C$2:$C$4)",
          None,
          FormulaKind.ArrayFormula(CellRange(ref"K9", ref"K11"))
        )
      )
    def at(formula: String): XLResult[CellValue] = precedents.evaluateFormula(formula)
    assertEquals(at("=H3+0"), Right(num(4)), "the direct-reference path")
    assertEquals(at("=K9+0"), Right(num(3)), "the direct-reference path")
    List(
      "=SUM(H2:H4)" -> num(8),
      "=AVERAGE(H2:H4)" -> dec("2.666666666666666666666666666666667"),
      "=MAX(H2:H4)" -> num(4),
      "=COUNT(H2:H4)" -> num(3),
      "=MEDIAN(H2:H4)" -> num(3),
      "=LARGE(H2:H4,1)" -> num(4),
      "=RANK(4,H2:H4)" -> num(1),
      "=COUNTIF(H2:H4,\">2\")" -> num(2),
      "=SUMIF(H2:H4,\">0\")" -> num(8),
      "=SUMIF(H2:H4,4,C2:C4)" -> num(4),
      "=AVERAGEIF(H2:H4,\">1\")" -> dec("3.5"),
      "=SUMIFS(C2:C4,H2:H4,4)" -> num(4),
      "=COUNTIFS(H2:H4,\">2\")" -> num(2),
      "=MAXIFS(H2:H4,H2:H4,\"<4\")" -> num(3),
      "=AVERAGEIFS(H2:H4,H2:H4,\">1\")" -> dec("3.5"),
      "=SUMPRODUCT(H2:H4)" -> num(8),
      "=SUM(K9)" -> num(3),
      "=SUM(K9:K9)" -> num(3)
    ).foreach { case (formula, expected) =>
      assertEquals(at(formula), Right(expected), formula)
    }
  }

  // ===== Array-kind records =====

  test("an ArrayFormula record evaluates array-aware; its anchor shows element (0,0)") {
    val cse =
      sheet
        .put(
          ref"F2",
          CellValue.Formula(
            "SUM(ABS(C2:C4+D2:D4))",
            None,
            FormulaKind.ArrayFormula(CellRange(ref"F2", ref"F2"))
          )
        )
        .put(
          ref"G6",
          CellValue.Formula(
            "ABS(C2:C4)",
            None,
            FormulaKind.ArrayFormula(CellRange(ref"G6", ref"G8"))
          )
        )
    assertEquals(cse.evaluateCell(ref"F2"), Right(num(7)))
    assertEquals(cse.evaluateCell(ref"G6"), Right(num(3)), "outside the range: not #VALUE!")
    val recalc = Workbook(cse).recalculate()
    assert(recalc.isClean, recalc.errors.toString)
    val out = recalc.workbook.sheets.headOption.getOrElse(fail("no sheet"))
    assertEquals(
      out(ref"F2").value match { case CellValue.Formula(_, c, _) => c; case _ => None },
      Some(num(7))
    )
    assertEquals(
      out(ref"G6").value match { case CellValue.Formula(_, c, _) => c; case _ => None },
      Some(num(3))
    )
  }
