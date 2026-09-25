package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * The references a plain formula cell passes around: IF and CHOOSE return the reference they
 * select, OFFSET a range sized like its base, a dynamic named range the reference it computes, a
 * LET name the reference it was bound to. In a value position a plain cell intersects them, in an
 * aggregate's argument they stay whole, and `@` intersects them exactly as the plain cell does; the
 * evaluation itself runs once (a selector draws RAND once). Array formulas keep array semantics
 * wherever they are evaluated, an iterative cycle included.
 */
class PlainCellReferenceSpec extends FunSuite:

  private def num(n: Double): CellValue = CellValue.Number(BigDecimal(n))

  /** A1:A10 = 1..10, B1:B10 = 2, Other!A1:A10 = 10..100. */
  private val book: Workbook =
    val sheet = (1 to 10).foldLeft(Sheet(SheetName.unsafe("S"))) { (s, i) =>
      s.put(ARef.from0(0, i - 1), num(i.toDouble)).put(ARef.from0(1, i - 1), num(2))
    }
    val other = (1 to 10).foldLeft(Sheet(SheetName.unsafe("Other"))) { (s, i) =>
      s.put(ARef.from0(0, i - 1), num(i * 10.0))
    }
    Workbook(Vector(sheet, other))
      .withDefinedName("dyn", "OFFSET(S!$A$1,0,0,10,1)")
      .withDefinedName("fixed", "S!$A$1:$A$10")

  private def sheetS: Sheet = book.sheets.headOption.getOrElse(fail("no sheet"))

  /** `formula` as a plain cell at `at` on S, evaluated at its position. */
  private def plain(formula: String, at: ARef): XLResult[CellValue] =
    val placed = sheetS.put(at, CellValue.Formula(formula, None))
    placed.evaluateCell(at, Clock.system, Some(book.put(placed)))

  test("OFFSET's height, width and sheet default to its base reference's, as in Excel") {
    assertEquals(plain("=SUM(OFFSET(A1:A10,0,1))", ref"D1"), Right(num(20)))
    assertEquals(plain("=ROWS(OFFSET(A1:A10,0,1))", ref"D1"), Right(num(10)))
    // in a value position the moved range is intersected: B4 from row 4
    assertEquals(plain("=OFFSET(A2:A5,0,1)*10", ref"D4"), Right(num(20)))
    assertEquals(plain("=OFFSET(Other!A1,1,0)", ref"D1"), Right(num(20)))
    assertEquals(plain("=SUM(OFFSET(Other!A1:A10,0,0))", ref"D1"), Right(num(550)))
  }

  test("a dynamic named range is a reference: intersected in a value position, whole in SUM") {
    assertEquals(plain("=dyn*2", ref"D5"), Right(num(10)))
    assertEquals(plain("=dyn*2", ref"D20"), Right(CellValue.Error(CellError.Value)))
    assertEquals(plain("=SUM(dyn)", ref"D20"), Right(num(55)))
    assertEquals(plain("=SUM(IF(TRUE,dyn,0))", ref"D20"), Right(num(55)))
    assertEquals(plain("=fixed*2", ref"D5"), Right(num(10)))
    assertEquals(plain("=SUM(IF(TRUE,fixed,0))", ref"D20"), Right(num(55)))
  }

  test("IF evaluates its condition once and keeps the reference it selects whole") {
    final class CountingRng(value: Double) extends Rng:
      @SuppressWarnings(Array("org.wartremover.warts.Var"))
      var draws: Int = 0
      def nextDouble(): Double =
        draws += 1
        value
    List(0.2 -> num(55), 0.8 -> num(0)).foreach { (draw, expected) =>
      val rng = CountingRng(draw)
      val placed = sheetS.put(ref"D5", CellValue.Formula("SUM(IF(RAND()<0.5,A1:A10,0))", None))
      val result = placed.evaluateCell(ref"D5", Clock.system, rng, Some(book.put(placed)))
      assertEquals(result, Right(expected), s"draw $draw")
      assertEquals(rng.draws, 1, "the condition evaluates once")
    }
  }

  test("an array formula in an iterative cycle evaluates as an array every round") {
    // C2 = {=SUM(A1:A3*1)+0*D2} (a CSE record) and D2 = C2 form a cycle: the array formula sums
    // A1:A3 (6) each round, never the plain cell's row-2 intersection (2)
    val cycle = sheetS
      .put(
        ref"C2",
        CellValue.Formula(
          "SUM(A1:A3*1)+0*D2",
          None,
          FormulaKind.ArrayFormula(CellRange(ref"C2", ref"C2"))
        )
      )
      .put(ref"D2", CellValue.Formula("C2", None))
    val result = book.put(cycle).recalculate(IterativeCalc(20, BigDecimal("0.001")))
    val c2 = result.workbook.sheets.headOption.flatMap(_.cells.get(ref"C2")).map(_.value)
    assertEquals(
      c2.collect { case CellValue.Formula(_, cached, _) => cached },
      Some(Some(num(6)))
    )
  }

  test("evaluateFormula at a cell is that plain cell's evaluateCell") {
    List(
      "=A1:A10*2",
      "=SUM(A1:A10*B1:B10)",
      "=IF(A1:A10>4,\"big\",\"small\")",
      "=SUM(IF(A1:A10>4,A1:A10,0))",
      "=dyn*2",
      "=COUNTIF(A1:A10,\">\"&A1:A10)"
    ).foreach { formula =>
      List(ref"D1", ref"D5", ref"D20").foreach { at =>
        assertEquals(
          sheetS.evaluateFormula(formula, Clock.system, Some(book), Some(at)),
          plain(formula, at),
          s"$formula at ${at.toA1}"
        )
      }
    }
  }

  test("the element law: a plain cell in row k reads element k of the array formula") {
    // for a column reference r and a plain cell in r's row k, plain(f) = arrayEval(f)(k); outside
    // r's rows it is #VALUE!
    List(
      "=A1:A10*3",
      "=ABS(A1:A10-5)",
      "=A1:A10&\"x\"",
      "=IF(A1:A10>4,A1:A10,-1)",
      "=ROUND(A1:A10/3,1)",
      "=COUNTIF(A1:A10,\"<\"&A1:A10)"
    ).foreach { formula =>
      val array = sheetS.evaluateArrayFormula(formula, ref"F1", Clock.system, Some(book)) match
        case Right((spilled, _)) => spilled
        case Left(err) => fail(s"$formula: ${err.message}")
      (1 to 10).foreach { row =>
        val at = ARef.from0(3, row - 1)
        val element = array.cells.get(ARef.from0(5, row - 1)).map(_.value)
        assertEquals(plain(formula, at).toOption, element, s"$formula at row $row")
      }
      assertEquals(plain(formula, ref"D11"), Right(CellValue.Error(CellError.Value)), formula)
    }
  }

  /** The reference-denoting expressions of the laws below, each A1:A10 (or one cell of it). */
  private val references = List(
    "A1:A10",
    "fixed",
    "dyn",
    "OFFSET(A1,0,0,10,1)",
    "INDIRECT(\"A1:A10\")",
    "INDEX(A1:B10,0,1)",
    "IF(TRUE,A1:A10,B1:B10)",
    "CHOOSE(1,A1:A10,B1:B10)",
    "A5"
  )

  test("LET never changes a value: plain(LET(x,e,b)) is plain(b with e for x)") {
    // a binding that denotes a reference binds the reference: intersected where the body reads a
    // value, whole in SUM and ROWS, materialized under SUMPRODUCT
    val bodies = List("SUM(x)", "x*2", "ROWS(x)", "SUMPRODUCT(x*B1:B10)", "AVERAGE(x)", "COUNTA(x)")
    for
      e <- references
      body <- bodies
      at <- List(ref"D5", ref"D20")
    do
      val substituted = plain("=" + body.replace("x", e), at)
      assert(substituted.isRight, s"$body with $e at ${at.toA1}: $substituted")
      assertEquals(plain(s"=LET(x,$e,$body)", at), substituted, s"LET(x,$e,$body) at ${at.toA1}")
  }

  test("a LET name bound to a reference passes it on: SUM(LET(r,e,r)) and a second binding") {
    references.foreach { e =>
      assertEquals(
        plain(s"=SUM(LET(r,$e,r))", ref"D20"),
        Right(num(if e == "A5" then 5 else 55)),
        e
      )
      assertEquals(plain(s"=LET(r,$e,s,r,s*2)", ref"D5"), Right(num(10)), e)
    }
  }

  test("a LET binding that selects a reference evaluates its selector once") {
    final class CountingRng(value: Double) extends Rng:
      @SuppressWarnings(Array("org.wartremover.warts.Var"))
      var draws: Int = 0
      def nextDouble(): Double =
        draws += 1
        value
    val rng = CountingRng(0.2)
    val formula = "LET(x,IF(RAND()<0.5,A1:A10,B1:B10),SUM(x)+ROWS(x)+x)"
    val placed = sheetS.put(ref"D5", CellValue.Formula(formula, None))
    // SUM 55 + ROWS 10 + the row-5 cell 5
    assertEquals(
      placed.evaluateCell(ref"D5", Clock.system, rng, Some(book.put(placed))),
      Right(num(70))
    )
    assertEquals(rng.draws, 1)
  }

  test(
    "@ intersects the reference a function, a name or a LET name returns: plain(@e) is plain(e)"
  ) {
    for
      e <- references.filterNot(_ == "A5")
      at <- List(ref"D5", ref"D20")
    do assertEquals(plain(s"=@$e", at), plain(s"=$e", at), s"@$e at ${at.toA1}")
    assertEquals(plain("=LET(r,dyn,@r)", ref"D5"), Right(num(5)))
    // an array value keeps its top-left
    assertEquals(plain("=@SEQUENCE(3)", ref"D5"), Right(num(1)))
  }

  test("@ in an array formula intersects with the anchor's row too") {
    val array =
      sheetS.evaluateArrayFormula("=@OFFSET(A1,0,0,10,1)", ref"F5", Clock.system, Some(book))
    assertEquals(array.map(_._1.cells.get(ref"F5").map(_.value)), Right(Some(num(5))))
  }

  test("OFFSET's base may be any reference; one that is none is #VALUE!") {
    assertEquals(plain("=OFFSET(dyn,0,1)", ref"D5"), Right(num(2)))
    assertEquals(plain("=SUM(OFFSET(INDEX(A1:B10,0,1),0,1))", ref"D5"), Right(num(20)))
    assertEquals(plain("=SUM(OFFSET(IF(TRUE,A1:A10,B1:B10),0,1))", ref"D5"), Right(num(20)))
    assertEquals(
      plain("=OFFSET(IF(TRUE,5,A1),0,0)", ref"D5"),
      Right(CellValue.Error(CellError.Value))
    )
    assertEquals(
      plain("=OFFSET(INDIRECT(\"nowhere\"),0,0)", ref"D5"),
      Right(CellValue.Error(CellError.Ref))
    )
  }
