package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.eval.FormulaEvaluationError
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import munit.FunSuite

/** GH-741: library callers get the same placement and closure protection as CLI eval. */
class FormulaEvaluationSpec extends FunSuite:
  private val at = ref"D5"
  private val sheet = Sheet("Data").put(at, 99).put(ref"A1", 10)
  private def number(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  private def cycle(result: Either[FormulaEvaluationError, CellValue], path: List[ARef]): Unit =
    result match
      case Left(failure) =>
        assertEquals(failure.cycle, Some(path))
        assert(failure.error.toString.contains("Circular reference detected"), failure.toString)
      case other => fail(s"Expected cycle $path, got $other")

  test("GH-741: placement refuses its old value, including own-sheet qualified reads") {
    List("=D5+1", "=Data!D5+1").foreach { formula =>
      cycle(SheetEvaluator.evaluateAt(Workbook(sheet), sheet, at, formula), List(at, at))
    }
    assertEquals(sheet(at).value, number(99))
  }

  test("GH-741: a static precedent cycle retains the full path") {
    val source = sheet
      .put(ref"B1", CellValue.Formula("C1*2"))
      .put(ref"C1", CellValue.Formula("D5"))
    cycle(
      SheetEvaluator.evaluateAt(Workbook(source), source, at, "B1+1"),
      List(at, ref"B1", ref"C1", at)
    )
  }

  test("GH-741: transitive range cycles include placements outside the used range") {
    val source = Sheet("Data")
      .put(ref"A1", CellValue.Formula("B1+1"))
      .put(ref"B1", CellValue.Formula("SUM(D:D)"))
    cycle(
      SheetEvaluator.evaluateAt(Workbook(source), source, at, "A1"),
      List(at, ref"A1", ref"B1", at)
    )
    val empty = Sheet("Empty")
    cycle(
      SheetEvaluator.evaluateAt(Workbook(empty), empty, ref"XFD1048576", "SUM(XFD:XFD)"),
      List(ref"XFD1048576", ref"XFD1048576")
    )
  }

  List(
    "INDIRECT(\"D5\")+1",
    "IFERROR(INDIRECT(\"D5\"),0)",
    "OFFSET(A1,4,3)",
    "SUM(OFFSET(A1,0,3,9,1))"
  ).foreach { formula =>
    test(s"GH-741: runtime guard survives swallowed failures: $formula") {
      cycle(SheetEvaluator.evaluateAt(Workbook(sheet), sheet, at, formula), List(at))
    }
  }

  test("GH-741: runtime guard also checks precedent evaluations and uncached cross-sheet reads") {
    val source = sheet.put(ref"C1", CellValue.Formula("IFERROR(INDIRECT(\"D5\"),0)"))
    cycle(SheetEvaluator.evaluateAt(Workbook(source), source, at, "C1+1"), List(at))
    val other = Sheet("Other").put(ref"A1", CellValue.Formula("Data!D5*2"))
    cycle(SheetEvaluator.evaluateAt(Workbook(sheet, other), sheet, at, "Other!A1"), List(at))
  }

  List(
    "IFERROR(COUNTIF(col,1),0)",
    "IFERROR(COUNTIFS(col,1),0)",
    "IFERROR(SUMIF(col,\">0\"),0)",
    "IFERROR(AVERAGEIF(col,\">0\"),0)",
    "IFERROR(MATCH(7,col,0),0)",
    "IFERROR(VLOOKUP(7,block,2,FALSE),0)",
    "IFERROR(XLOOKUP(7,col,col),0)",
    "IFERROR(SUMPRODUCT(col),0)",
    "IFERROR(SUM(col),0)",
    "IFERROR(COUNTA(col),0)",
    "IFERROR(COUNTBLANK(col),0)"
  ).foreach { formula =>
    test(s"GH-741: named range reader sees the guard: $formula") {
      val book = Workbook(sheet).copy(metadata =
        Workbook(sheet).metadata.copy(definedNames =
          Vector(
            DefinedName("col", "Data!$D$1:$D$9"),
            DefinedName("block", "Data!$D$1:$E$9")
          )
        )
      )
      cycle(SheetEvaluator.evaluateAt(book, sheet, at, formula), List(at))
    }
  }

  test("GH-741: positionless and positioned calls share override propagation through chains") {
    val source = sheet
      .put(ref"B1", CellValue.Formula("A1*2", Some(number(20))))
      .put(ref"C1", CellValue.Formula("B1+50", Some(number(70))))
      .put(ref"F1", CellValue.Formula("F1")) // unrelated cycle is never visited
    val book = Workbook(source)
    val overrides = Map(ref"A1" -> number(200))
    assertEquals(SheetEvaluator.evaluateAt(book, source, at, "C1", overrides), Right(number(450)))
    assertEquals(
      SheetEvaluator.evaluateWithClosure(book, source, "C1", overrides),
      Right(number(450))
    )
    assertEquals(source(ref"A1").value, number(10))
    assertEquals(source(ref"C1").value, CellValue.Formula("B1+50", Some(number(70))))
  }

  test("GH-741: overrides invalidate transitive readers on other sheets") {
    val other = Sheet("Other")
      .put(ref"A1", CellValue.Formula("Data!A1*2", Some(number(20))))
      .put(ref"A2", CellValue.Formula("A1+1", Some(number(21))))
    val book = Workbook(sheet, other)
    val overrides = Map(ref"A1" -> number(200))
    assertEquals(
      SheetEvaluator.evaluateAt(book, sheet, at, "Other!A2", overrides),
      Right(number(401))
    )
    assertEquals(
      SheetEvaluator.evaluateWithClosure(book, sheet, "Other!A2", overrides),
      Right(number(401))
    )
  }

  test("GH-741: positioned implicit intersection differs from positionless array evaluation") {
    val source = sheet.put(ref"A5", 3).put(ref"B5", 4).put(ref"B1", 2)
    assertEquals(
      SheetEvaluator.evaluateAt(Workbook(source), source, at, "SUM(A1:A5*B1:B5)"),
      Right(number(12))
    )
    assertEquals(
      SheetEvaluator.evaluateWithClosure(Workbook(source), source, "SUM(A1:A5*B1:B5)"),
      Right(number(32))
    )
  }

  test("GH-741: array evaluation reuses precedent spill tracking") {
    val source = sheet.put(
      ref"B1",
      CellValue.Formula("SEQUENCE(3)", None, FormulaKind.ArrayFormula(ref"B1:B3"))
    )
    SheetEvaluator.evaluateArrayWithClosure(Workbook(source), source, "B1#*2", ref"F1") match
      case Right((computed, range)) =>
        assertEquals(range, ref"F1:F3")
        assertEquals(
          List(ref"F1", ref"F2", ref"F3").map(computed(_).value),
          List(number(2), number(4), number(6))
        )
      case Left(error) => fail(error.toString)
  }

  test("GH-741: existing placement content is replaced, even a broken formula") {
    val source = sheet.put(at, CellValue.Formula("NOT_A_FUNCTION(1)"))
    assertEquals(SheetEvaluator.evaluateAt(Workbook(source), source, at, "A1+1"), Right(number(11)))
    assert(
      SheetEvaluator.evaluateAt(Workbook(source), source, at, "1", Map(at -> number(2))).isLeft
    )
  }

  test("GH-741: cycles in the closure are refused even without a position") {
    val source = sheet.put(ref"B1", CellValue.Formula("C1")).put(ref"C1", CellValue.Formula("B1"))
    val result = SheetEvaluator.evaluateWithClosure(Workbook(source), source, "B1")
    assert(result.left.exists(_.cycle.nonEmpty), result.toString)
  }

  test("GH-741: long cycle diagnostics are abbreviated but typed paths are complete") {
    val source = (1 to 20).foldLeft(Sheet("Data")) { (s, row) =>
      s.put(ARef.from0(0, row - 1), CellValue.Formula(if row == 20 then "D5" else s"A${row + 1}"))
    }
    val expected = at :: (1 to 20).map(row => ARef.from0(0, row - 1)).toList ::: List(at)
    val result = SheetEvaluator.evaluateAt(Workbook(source), source, at, "A1")
    cycle(result, expected)
    assert(result.left.exists(_.error.toString.contains("more")), result.toString)
  }

  test("GH-741: parser errors retain their diagnostic and have no cycle") {
    val formula = "=1+"
    val expected = FormulaParser.parse(formula).left.toOption.map(ParseError.toXLError(_, formula))
    val result = SheetEvaluator.evaluateAt(Workbook(sheet), sheet, at, formula)
    assertEquals(result.left.toOption.map(_.error), expected)
    assertEquals(result.left.toOption.flatMap(_.cycle), None)
  }

  test("GH-741: cached cross-sheet cycle gap remains scoped to GH-729") {
    val other = Sheet("Other").put(ref"A1", CellValue.Formula("Data!D5*2", Some(number(14))))
    assertEquals(
      SheetEvaluator.evaluateAt(Workbook(sheet, other), sheet, at, "Other!A1"),
      Right(number(14))
    )
  }

  test("GH-741: low-level currentCell retains its existing position-only behavior") {
    val sheet = Sheet("Data").put(ref"D5", 99)
    assertEquals(
      sheet.evaluateFormula("=D5+1", currentCell = Some(ref"D5")),
      Right(CellValue.Number(BigDecimal(100)))
    )
  }

  test("GH-741: ISBLANK preserves the existing guard on an indirect placement read") {
    cycle(
      SheetEvaluator.evaluateAt(Workbook(sheet), sheet, at, "ISBLANK(INDIRECT(\"D5\"))"),
      List(at)
    )
  }
