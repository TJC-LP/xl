package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.functions.{EvalContext, FunctionRegistry}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-712: ROW and COLUMN over a reference a function computes. SUMPRODUCT's whole-axis trimming
 * (GH-192) reaches the reference INDEX, a name or another sheet returns, so ROW's array lines up
 * with the operands beside it; the trimming shapes arrays only, never a scalar ROW or COLUMN; and a
 * LET binding read by ROW, COLUMN, ROWS or COLUMNS keeps its reference. The spill and plain-cell
 * cases over literal ranges live in `ArrayLiftingSpec`, the LibreOffice table in
 * `PlainCellIntersectionSpec`.
 */
class RowColumnArraySpec extends FunSuite:
  private def number(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def text(s: String): CellValue = CellValue.Text(s)
  private def column(values: CellValue*): ArrayResult = ArrayResult(values.toVector.map(Vector(_)))
  private val sheet = (1 to 3).foldLeft(Sheet(SheetName.unsafe("Data"))) { (s, i) =>
    s.put(ARef.from0(0, i - 1), number(i))
      .put(ARef.from0(1, i - 1), number(i * 10))
  }

  private def array(formula: String): XLResult[CellValue] =
    sheet.evaluateFormula(formula, Clock.system, Some(Workbook(sheet)), None)

  private def plain(formula: String): XLResult[CellValue] =
    val placed = sheet.put(ref"H9", CellValue.Formula(formula, None))
    placed.evaluateCell(ref"H9", Clock.system, Some(Workbook(placed)))

  private def parse(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(e => fail(s"parse failed: $e"), identity)

  /** `evala`'s raw value: the formula as an array, unspilled. */
  private def evala(formula: String): Either[EvalError, Any] =
    Evaluator.arrayInstance.eval(parse(formula), sheet, Clock.system, Some(Workbook(sheet)), None)

  /** `evala`'s spill: the cells the result occupies when anchored at H1. */
  private def spill(formula: String): XLResult[Vector[CellValue]] =
    sheet.evaluateArrayFormula(formula, ref"H1", Clock.system, Some(Workbook(sheet))).map {
      (spilled, range) => range.cells.map(at => spilled(at).value).toVector
    }

  // ===== SUMPRODUCT bounds computed references (GH-712, GH-192) =====

  test("SUMPRODUCT bounds ROW of a computed whole-column reference consistently") {
    assertEquals(array("=SUMPRODUCT(A:A,ROW(INDEX(A:A,0)))"), Right(number(14)))
  }

  test("SUMPRODUCT bounds COLUMN of a computed whole-row reference consistently") {
    assertEquals(array("=SUMPRODUCT(1:1,COLUMN(INDEX(1:1,0)))"), Right(number(21)))
  }

  test("computed positions preserve leading blanks, empty extents, and the INDEX origin") {
    val offset = Sheet(SheetName.unsafe("Offset"))
      .put(ref"A2", number(2))
      .put(ref"A3", number(3))
    assertEquals(
      offset.evaluateFormula("=SUMPRODUCT(A:A,ROW(INDEX(A:A,0)))"),
      Right(number(13))
    )
    assertEquals(
      offset.evaluateFormula("=SUMPRODUCT(ROW(INDEX(A:A,2)))"),
      Right(number(2))
    )
    assertEquals(
      Sheet(SheetName.unsafe("Empty")).evaluateFormula("=SUMPRODUCT(A:A,ROW(INDEX(A:A,0)))"),
      Right(number(0))
    )
  }

  test("computed positions use SUMPRODUCT's shared cross-sheet extent") {
    val other = Sheet(SheetName.unsafe("Other")).put(ref"A5", number(5))
    val wb = Workbook(sheet, other)
    assertEquals(
      sheet.evaluateFormula("=SUMPRODUCT(A:A,ROW(INDEX(Other!A:A,0)))", Clock.system, Some(wb)),
      Right(number(14))
    )
  }

  test(
    "only ROW and COLUMN take SUMPRODUCT's bounds on their reference; ROWS and COLUMNS keep it"
  ) {
    val flagged = FunctionRegistry.all.filter(_.flags.boundsReferenceAxis).map(_.name).toSet
    assertEquals(flagged, Set("ROW", "COLUMN"))
    val trim = (_: Option[SheetName], range: CellRange) =>
      if range.isFullColumn then CellRange(ref"A5", ref"A7") else range
    assertEquals(TExpr.transformRanges(parse("=ROW(A:A)"), trim), parse("=ROW(A5:A7)"))
    assertEquals(TExpr.transformRanges(parse("=COLUMN(A:A)"), trim), parse("=COLUMN(A5:A7)"))
    assertEquals(TExpr.transformRanges(parse("=ROWS(A:A)"), trim), parse("=ROWS(A:A)"))
    assertEquals(TExpr.transformRanges(parse("=COLUMNS(A:A)"), trim), parse("=COLUMNS(A:A)"))
  }

  // ===== The trimming shapes arrays only: a scalar ROW or COLUMN is the reference's own =====

  /** Data from A5:A7 and C1:E1, so the used range starts at neither row 1 nor column A. */
  private val offsetGrid = List[(ARef, CellValue)](
    ref"A5" -> number(1),
    ref"A6" -> number(2),
    ref"A7" -> number(3),
    ref"C1" -> number(1),
    ref"D1" -> number(2),
    ref"E1" -> number(3)
  ).foldLeft(Sheet(SheetName.unsafe("Grid"))) { case (s, (at, v)) => s.put(at, v) }

  private def gridCell(formula: String, at: ARef): XLResult[CellValue] =
    val placed = offsetGrid.put(at, CellValue.Formula(formula, None))
    placed.evaluateCell(at, Clock.system, Some(Workbook(placed)))

  test("a plain cell's ROW(A:A) and COLUMN(1:1) are 1 wherever the used range starts") {
    assertEquals(gridCell("=ROW(A:A)", ref"H9"), Right(number(1)))
    assertEquals(gridCell("=COLUMN(1:1)", ref"H9"), Right(number(1)))
    assertEquals(gridCell("=ROW(A:A)+0", ref"H9"), Right(number(1)))
    assertEquals(gridCell("=COLUMN(1:1)*1", ref"H9"), Right(number(1)))
  }

  /** Column A filled from row 5 only: the used extent is rows 5..7. */
  private val rowsFrom5 = (5 to 7).foldLeft(Sheet(SheetName.unsafe("Rows"))) { (s, r) =>
    s.put(ARef.from0(0, r - 1), number(r - 4))
  }

  private def rowsCell(sheet: Sheet, formula: String, at: ARef): XLResult[CellValue] =
    val placed = sheet.put(at, CellValue.Formula(formula, None))
    placed.evaluateCell(at, Clock.system, Some(Workbook(placed)))

  test("a precedent's plain ROW(A:A) read inside SUMPRODUCT is 1, not the first used row") {
    // B6 evaluates as its own plain formula while SUMPRODUCT trims A:A to rows 5..7
    val withPrecedent = rowsFrom5.put(ref"B6", CellValue.Formula("=ROW(A:A)", None))
    assertEquals(rowsCell(withPrecedent, "=SUMPRODUCT((A:A>0)*B6)", ref"B7"), Right(number(3)))
    // inside SUMPRODUCT ROW's own array is trimmed to the used rows: 5 + 6 + 7, where Excel
    // sums 1 to 1,048,576
    assertEquals(rowsCell(rowsFrom5, "=SUMPRODUCT(ROW(A:A))", ref"B6"), Right(number(18)))
  }

  test("a scalar ROW or COLUMN ignores SUMPRODUCT's shared bounds") {
    // No formula reaches this today: every evaluator that carries the bounds is in array mode.
    // The invariant is pinned on the context directly so a future scalar sub-evaluation inside
    // SUMPRODUCT cannot pick up the used extent.
    val trimmed = Some((_: CellRange) => CellRange(ref"C5", ref"E7"))
    val scalar = EvalContext(
      offsetGrid,
      Clock.system,
      None,
      [A] => (_: TExpr[A]) => Left(EvalError.EvalFailed("unused", None)),
      _ => Left(EvalError.EvalFailed("unused", None)),
      operands = EvalContext.Operands(
        arrayMode = false,
        referenceArg = None,
        selectsReference = false,
        arrayRangeBounds = trimmed
      )
    )
    def scalarCall(formula: String): Either[EvalError, Any] = parse(formula) match
      case call: TExpr.Call[?] => call.spec.eval(call.args, scalar)
      case other => fail(s"expected a call, got $other")
    assertEquals(scalarCall("=ROW(A:A)"), Right(BigDecimal(1)))
    assertEquals(scalarCall("=COLUMN(1:1)"), Right(BigDecimal(1)))
    // the same bounds in array mode do trim: rows 5..7 and columns C..E
    val arrayMode = scalar.copy(operands = scalar.operands.copy(arrayMode = true))
    def arrayCall(formula: String): Either[EvalError, Any] = parse(formula) match
      case call: TExpr.Call[?] => call.spec.eval(call.args, arrayMode)
      case other => fail(s"expected a call, got $other")
    assertEquals(arrayCall("=ROW(A:A)"), Right(column(number(5), number(6), number(7))))
    assertEquals(
      arrayCall("=COLUMN(1:1)"),
      Right(ArrayResult(Vector(Vector(number(3), number(4), number(5)))))
    )
  }

  // ===== The runtime trimming follows the static path: nested shape readers see every row =====

  test("MAX, MIN and COUNT of ROW(A:A) nested in SUMPRODUCT see the whole column, as in Excel") {
    // (A:A>0) has three TRUEs (rows 5..7); MAX(ROW(A:A)) and COUNT(ROW(A:A)) are 1,048,576 and
    // MIN(ROW(A:A)) is 1 in Excel and LibreOffice, not the used extent's 7, 3 and 5
    val wholeColumn = 3 * 1048576
    def sp(formula: String) = rowsCell(rowsFrom5, formula, ref"B6")
    assertEquals(sp("=SUMPRODUCT((A:A>0)*MAX(ROW(A:A)))"), Right(number(wholeColumn)))
    assertEquals(sp("=SUMPRODUCT((A:A>0)*COUNT(ROW(A:A)))"), Right(number(wholeColumn)))
    assertEquals(sp("=SUMPRODUCT((A:A>0)*MIN(ROW(A:A)))"), Right(number(3)))
    assertEquals(sp("=SUMPRODUCT((A:A>0)*MAX(ROW(INDEX(A:A,0))))"), Right(number(wholeColumn)))
    assertEquals(sp("=SUMPRODUCT(MAX(ROW(A:A)))"), Right(number(1048576)))
    // COLUMN over a whole row: (1:1>0) has three TRUEs (C1:E1), MAX(COLUMN(1:1)) is 16,384
    assertEquals(
      gridCell("=SUMPRODUCT((1:1>0)*MAX(COLUMN(1:1)))", ref"H9"),
      Right(number(3 * 16384))
    )
    // ROW's array as SUMPRODUCT's own element source still lines up: 5 + 6 + 7
    assertEquals(sp("=SUMPRODUCT((A:A>0)*ROW(A:A))"), Right(number(18)))
    assertEquals(sp("=SUMPRODUCT((A:A>0)*ROW(INDEX(A:A,0)))"), Right(number(18)))
  }

  test("a named formula read inside SUMPRODUCT evaluates whole, outside the trimming") {
    def at(formula: String) =
      val placed = rowsFrom5.put(ref"B6", CellValue.Formula(formula, None))
      val named = Workbook(placed)
        .withDefinedName("nmMax", "MAX(ROW(Rows!$A:$A))")
        .withDefinedName("nmCount", "COUNT(ROW(Rows!$A:$A))")
      placed.evaluateCell(ref"B6", Clock.system, Some(named))
    assertEquals(at("=SUMPRODUCT((A:A>0)*nmMax)"), Right(number(3 * 1048576)))
    assertEquals(at("=SUMPRODUCT((A:A>0)*nmCount)"), Right(number(3 * 1048576)))
  }

  // ===== A union is no single reference =====

  test("ROW and COLUMN of a union are #VALUE!, as in Excel") {
    val value = CellValue.Error(CellError.Value)
    assertEquals(plain("=ROW((A1,A3))"), Right(value))
    assertEquals(plain("=COLUMN((A1,B1))"), Right(value))
    assertEquals(plain("=IFERROR(ROW((A1,A3)),\"err\")"), Right(text("err")))
    assertEquals(plain("=ERROR.TYPE(COLUMN((A1,B1)))"), Right(number(3)))
    assertEquals(array("=SUMPRODUCT(ROW((A1,A3)))"), Right(value))
    evala("=ROW((A1:A2,A3))") match
      case Left(EvalError.ErrorValue(CellError.Value, _)) => ()
      case other => fail(s"expected #VALUE!, got $other")
    // an intersection that leaves one area is an ordinary reference
    assertEquals(plain("=ROW(A1:B3 B2:C3)"), Right(number(2)))
  }

  // ===== LET keeps the reference ROW, COLUMN, ROWS and COLUMNS read =====

  test("ROW keeps the reference of an INDEX result bound by LET in array mode") {
    assertEquals(array("=LET(chosen,INDEX(B1:B3,2),ROW(chosen))"), Right(number(2)))
  }

  test("LET keeps references for both coordinates and shape inspection") {
    assertEquals(array("=LET(chosen,INDEX(A1:B3,2,0),COLUMN(chosen))"), Right(number(1)))
    assertEquals(array("=LET(chosen,INDEX(B:B,0),ROWS(chosen))"), Right(number(1048576)))
    assertEquals(array("=LET(chosen,INDEX(1:1,0),COLUMNS(chosen))"), Right(number(16384)))
  }

  test("one LET binding read as a value and as a shape: r*2 + ROWS(r) is {5;7;9}") {
    // Excel: r is A1:A3, r*2 is {2;4;6}, ROWS(r) is 3
    assertEquals(
      evala("=LET(r,OFFSET(A1,0,0,3),r*2+ROWS(r))"),
      Right(column(number(5), number(7), number(9)))
    )
    assertEquals(
      evala("=LET(r,OFFSET(A1,0,0,3),SUM(r)+ROWS(r)+COLUMNS(r))"),
      Right(BigDecimal(10))
    )
    assertEquals(
      evala("=LET(r,INDEX(A1:B3,0,2),r+ROW(r))"),
      Right(column(number(11), number(22), number(33)))
    )
  }

  test("a reference consumer under unary minus, plus, percent or a nested LET keeps the binding") {
    assertEquals(evala("=LET(r,INDEX(A1:A3,2),-ROW(r))"), Right(BigDecimal(-2)))
    assertEquals(
      evala("=LET(r,INDEX(A1:B3,0,1),-ROW(r))"),
      Right(column(number(-1), number(-2), number(-3)))
    )
    assertEquals(evala("=LET(r,INDEX(A1:A3,3),+ROW(r))"), Right(BigDecimal(3)))
    assertEquals(evala("=LET(r,INDEX(A1:A3,3),ROW(r)%)"), Right(BigDecimal("0.03")))
    assertEquals(evala("=LET(r,INDEX(A1:A3,2),n,-COLUMN(r),n)"), Right(BigDecimal(-1)))
    assertEquals(evala("=LET(r,INDEX(A1:A3,2),LET(s,1,ROW(r)+s))"), Right(BigDecimal(3)))
    assertEquals(plain("=LET(r,INDEX(A1:A3,2),-ROW(r))"), Right(number(-2)))
  }

  // ===== Consumers that do not lift over ROW's array =====

  test("ADDRESS, CHOOSE and OFFSET over ROW's array under evala keep its first number") {
    // Excel 365 lifts ADDRESS and CHOOSE over the array ({"$A$1";"$A$2";"$A$3"}, {"a";"b";"c"})
    // and spills #VALUE! for OFFSET's array of references; xl reads the array's top-left (1),
    // the typed collapse of a non-lifted scalar slot, and never fails internally.
    assertEquals(spill("=ADDRESS(ROW(A1:A3),1)"), Right(Vector(text("$A$1"))))
    assertEquals(spill("=CHOOSE(ROW(A1:A3),\"a\",\"b\",\"c\")"), Right(Vector(text("a"))))
    assertEquals(spill("=OFFSET(A1,ROW(A1:A3),0)"), Right(Vector(number(2))))
    assertEquals(spill("=CHOOSE(COLUMN(A1:C1),\"a\",\"b\",\"c\")"), Right(Vector(text("a"))))
  }
